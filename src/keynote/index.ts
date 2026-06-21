import { access, mkdir, rm, stat } from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

import type { DeckIR } from "../ir/index.ts";
import { parsePptxToIr, exportIrToPptx } from "../pptx/index.ts";
import { parseNativeKeynoteToIr } from "./native.ts";

export { isKeynoteHtmlExportDir, parseKeynoteHtmlExportToIr, parseKeynoteHtmlToIr } from "./html.ts";
export type { KeynoteHtmlImportOptions, KeynoteHtmlParseOptions } from "./html.ts";
export { detectNativeKeynotePackage, materializeNativeKeynoteAssetFiles, parseNativeKeynoteToIr } from "./native.ts";
export type {
  NativeKeynoteAssetMaterializationRequest,
  NativeKeynoteAssetMaterializationResult,
  NativeIwaCompression,
  NativeIwaFieldSummary,
  NativeIwaStreamMetadata,
  NativeIwaStreamRole,
  NativeKeynoteDetection,
  NativeKeynotePackageFormat
} from "./native.ts";

const execFileAsync = promisify(execFile);

export interface KeynoteImportOptions {
  exportedPptxPath?: string;
  preferNative?: boolean;
  nativeFallback?: boolean;
  workDir?: string;
  automationTimeoutMs?: number;
  allowAutomation?: boolean;
  bridgeExport?: (keynotePath: string, pptxPath: string, options: KeynoteAutomationOptions) => Promise<void>;
}

export interface KeynoteExportOptions {
  intermediatePptxPath?: string;
  automationTimeoutMs?: number;
  allowAutomation?: boolean;
}

export interface KeynoteHtmlExportOptions {
  automationTimeoutMs?: number;
  allowAutomation?: boolean;
}

export type KeynoteMovieExportFormat = "360p" | "540p" | "720p" | "1080p" | "2160p" | "native";
export type KeynoteMovieCodec =
  | "h264"
  | "hevc"
  | "AppleProRes422"
  | "AppleProRes4444"
  | "AppleProRes422LT"
  | "AppleProRes422HQ"
  | "AppleProRes422Proxy";
export type KeynoteMovieFrameRate = "FPS12" | "FPS2398" | "FPS24" | "FPS25" | "FPS2997" | "FPS30" | "FPS50" | "FPS5994" | "FPS60";

export interface KeynoteMovieExportOptions extends KeynoteAutomationOptions {
  format?: KeynoteMovieExportFormat;
  codec?: KeynoteMovieCodec;
  frameRate?: KeynoteMovieFrameRate;
}

export async function parseKeynoteToIr(
  keynotePath: string,
  options: KeynoteImportOptions = {}
): Promise<DeckIR> {
  if (options.preferNative) {
    return parseNativeKeynoteToIr(keynotePath);
  }

  const pptxPath =
    options.exportedPptxPath ??
    path.join(options.workDir ?? path.dirname(keynotePath), `${path.basename(keynotePath, path.extname(keynotePath))}.keynote-bridge.pptx`);

  if (!options.exportedPptxPath) {
    try {
      await (options.bridgeExport ?? exportKeynoteToPptx)(keynotePath, pptxPath, {
        automationTimeoutMs: options.automationTimeoutMs,
        allowAutomation: options.allowAutomation
      });
    } catch (error) {
      if (options.nativeFallback === false) {
        throw error;
      }
      try {
        const deck = await parseNativeKeynoteToIr(keynotePath);
        deck.conversion?.messages.unshift({
          severity: "warning",
          code: "keynote-bridge-fallback-native",
          message: `Keynote PPTX bridge failed, so native package probing was used instead. ${errorMessage(error)}`
        });
        return deck;
      } catch {
        throw error;
      }
    }
  }

  const deck = await parsePptxToIr(pptxPath);
  deck.conversion ??= {
    status: "partial",
    messages: []
  };
  deck.conversion.source = {
    kind: "keynote",
    uri: keynotePath,
    application: "Keynote"
  };
  deck.conversion.uncertainMappings ??= [];
  deck.conversion.uncertainMappings.push({
    code: "keynote-pptx-bridge",
    description: "Imported through Keynote's PPTX export bridge; native Keynote-only animation mappings are best-effort.",
    severity: "warning",
    confidence: 0.72
  });
  deck.conversion.messages.push({
    severity: "info",
    code: "keynote-bridge-import",
    message: "Keynote deck was exported to PPTX locally, then converted to KeyMorph IR."
  });
  return deck;
}

export async function exportIrToKeynote(deck: DeckIR, keynotePath: string, options: KeynoteExportOptions = {}): Promise<void> {
  const pptxPath =
    options.intermediatePptxPath ??
    path.join(path.dirname(keynotePath), `${path.basename(keynotePath, path.extname(keynotePath))}.pptx`);
  await mkdir(path.dirname(pptxPath), { recursive: true });
  await exportIrToPptx(deck, pptxPath);
  await importPptxAndSaveKeynote(pptxPath, keynotePath, {
    automationTimeoutMs: options.automationTimeoutMs,
    allowAutomation: options.allowAutomation
  });
}

export interface KeynoteExportResult {
  pptxPath: string;
  unsupportedCount: number;
}

export async function exportIrToKeynoteBridge(
  deck: DeckIR,
  pptxPath: string,
  exportPptx: (deck: DeckIR, pptxPath: string) => Promise<void>
): Promise<KeynoteExportResult> {
  await exportPptx(deck, pptxPath);
  return {
    pptxPath,
    unsupportedCount:
      (deck.conversion?.unsupportedFeatures?.length ?? 0) + (deck.conversion?.degradedFeatures?.length ?? 0)
  };
}

export interface KeynoteAutomationOptions {
  automationTimeoutMs?: number;
  allowAutomation?: boolean;
}

export async function exportKeynoteToPptx(keynotePath: string, pptxPath: string, options: KeynoteAutomationOptions = {}): Promise<void> {
  await assertReadableFile(keynotePath, "Input Keynote file");
  assertAutomationAllowed(options);
  await mkdir(path.dirname(pptxPath), { recursive: true });
  await runAppleScript(`
set inputFile to POSIX file "${escapeAppleScriptString(path.resolve(keynotePath))}"
set outputFile to POSIX file "${escapeAppleScriptString(path.resolve(pptxPath))}"
tell application "Keynote"
  activate
  set theDocument to open inputFile
  export theDocument to outputFile as Microsoft PowerPoint
  close theDocument saving no
end tell
`, options);
}

export async function exportKeynoteToHtml(keynotePath: string, outputDir: string, options: KeynoteHtmlExportOptions = {}): Promise<void> {
  await assertReadableFile(keynotePath, "Input Keynote file");
  assertAutomationAllowed(options);
  const resolvedOutputDir = path.resolve(outputDir);
  await rm(resolvedOutputDir, { recursive: true, force: true });
  await mkdir(resolvedOutputDir, { recursive: true });
  await runAppleScript(`
set inputFile to POSIX file "${escapeAppleScriptString(path.resolve(keynotePath))}"
set outputFolder to POSIX file "${escapeAppleScriptString(resolvedOutputDir)}"
tell application "Keynote"
  activate
  set theDocument to open inputFile
  export theDocument to outputFolder as HTML
  close theDocument saving no
end tell
`, options);
}

export async function exportKeynoteToMovie(keynotePath: string, moviePath: string, options: KeynoteMovieExportOptions = {}): Promise<void> {
  await assertReadableFile(keynotePath, "Input Keynote file");
  assertAutomationAllowed(options);
  await mkdir(path.dirname(moviePath), { recursive: true });
  const movieFormat = keynoteMovieFormatToken(options.format ?? "2160p");
  const movieCodec = keynoteMovieCodecToken(options.codec ?? "h264");
  const movieFrameRate = keynoteMovieFrameRateToken(options.frameRate ?? "FPS60");
  const automationTimeoutMs = resolveMovieAutomationTimeoutMs(options.automationTimeoutMs);
  const appleEventTimeoutSeconds = Math.max(1, Math.ceil(automationTimeoutMs / 1000));
  try {
    await runAppleScript(`
set inputFile to POSIX file "${escapeAppleScriptString(path.resolve(keynotePath))}"
set outputFile to POSIX file "${escapeAppleScriptString(path.resolve(moviePath))}"
with timeout of ${appleEventTimeoutSeconds} seconds
  tell application "Keynote"
    activate
    set theDocument to open inputFile
    export theDocument to outputFile as QuickTime movie with properties {movie format:${movieFormat}, movie codec:${movieCodec}, movie framerate:${movieFrameRate}}
    close theDocument saving no
  end tell
end timeout
`, { ...options, automationTimeoutMs });
  } catch (error) {
    if (await fileExistsWithContent(moviePath)) return;
    throw error;
  }
}

export async function importPptxAndSaveKeynote(pptxPath: string, keynotePath: string, options: KeynoteAutomationOptions = {}): Promise<void> {
  await assertReadableFile(pptxPath, "Input PPTX file");
  assertAutomationAllowed(options);
  await mkdir(path.dirname(keynotePath), { recursive: true });
  await runAppleScript(`
set inputFile to POSIX file "${escapeAppleScriptString(path.resolve(pptxPath))}"
set outputFile to POSIX file "${escapeAppleScriptString(path.resolve(keynotePath))}"
tell application "Keynote"
  activate
  set theDocument to open inputFile
  save theDocument in outputFile
  close theDocument saving no
end tell
`, options);
}

function assertAutomationAllowed(options: KeynoteAutomationOptions = {}): void {
  if (options.allowAutomation || process.env.KEYMORPH_ALLOW_KEYNOTE_AUTOMATION === "1") {
    return;
  }

  throw new Error(
    "Keynote GUI automation is disabled by default. Pass allowAutomation: true or set KEYMORPH_ALLOW_KEYNOTE_AUTOMATION=1 to run AppleScript."
  );
}

async function runAppleScript(script: string, options: KeynoteAutomationOptions = {}): Promise<void> {
  if (process.platform !== "darwin") {
    throw new Error("Keynote conversion requires macOS with Keynote installed.");
  }
  const timeout = resolveAutomationTimeoutMs(options.automationTimeoutMs);
  try {
    await execFileAsync("osascript", ["-e", script], { maxBuffer: 1024 * 1024 * 8, timeout });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Keynote automation failed. Confirm Keynote is installed and allow automation permission when macOS prompts. ${message}`);
  }
}

function resolveAutomationTimeoutMs(optionTimeoutMs: number | undefined): number | undefined {
  if (optionTimeoutMs !== undefined) {
    return validateAutomationTimeout(optionTimeoutMs, "automationTimeoutMs");
  }

  const envTimeout = process.env.KEYMORPH_KEYNOTE_AUTOMATION_TIMEOUT_MS;
  if (envTimeout === undefined || envTimeout.trim() === "") {
    return 120000;
  }

  return validateAutomationTimeout(Number(envTimeout), "KEYMORPH_KEYNOTE_AUTOMATION_TIMEOUT_MS");
}

function resolveMovieAutomationTimeoutMs(optionTimeoutMs: number | undefined): number {
  if (optionTimeoutMs !== undefined) {
    return validateAutomationTimeout(optionTimeoutMs, "automationTimeoutMs");
  }

  const envTimeout = process.env.KEYMORPH_KEYNOTE_AUTOMATION_TIMEOUT_MS;
  if (envTimeout !== undefined && envTimeout.trim() !== "") {
    return validateAutomationTimeout(Number(envTimeout), "KEYMORPH_KEYNOTE_AUTOMATION_TIMEOUT_MS");
  }

  return 600000;
}

function validateAutomationTimeout(value: number, name: string): number {
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`${name} must be a positive timeout in milliseconds.`);
  }
  return Math.floor(value);
}

async function assertReadableFile(filePath: string, label: string): Promise<void> {
  try {
    await access(filePath);
  } catch {
    throw new Error(`${label} does not exist: ${filePath}`);
  }
}

async function fileExistsWithContent(filePath: string): Promise<boolean> {
  try {
    return (await stat(filePath)).size > 0;
  } catch {
    return false;
  }
}

function escapeAppleScriptString(value: string): string {
  return value.replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

function keynoteMovieFormatToken(format: KeynoteMovieExportFormat): string {
  switch (format) {
    case "360p":
      return "format360p";
    case "540p":
      return "format540p";
    case "720p":
      return "format720p";
    case "1080p":
      return "format1080p";
    case "2160p":
      return "format2160p";
    case "native":
      return "native size";
  }
}

function keynoteMovieCodecToken(codec: KeynoteMovieCodec): string {
  if (codec === "hevc") return "HEVC";
  return codec;
}

function keynoteMovieFrameRateToken(frameRate: KeynoteMovieFrameRate): string {
  return frameRate;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
