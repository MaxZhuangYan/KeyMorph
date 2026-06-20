import { mkdir } from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

import type { DeckIR } from "../ir/index.ts";
import { parsePptxToIr, exportIrToPptx } from "../pptx/index.ts";

const execFileAsync = promisify(execFile);

export interface KeynoteImportOptions {
  exportedPptxPath?: string;
  workDir?: string;
}

export interface KeynoteExportOptions {
  intermediatePptxPath?: string;
}

export async function parseKeynoteToIr(
  keynotePath: string,
  options: KeynoteImportOptions = {}
): Promise<DeckIR> {
  const pptxPath =
    options.exportedPptxPath ??
    path.join(options.workDir ?? path.dirname(keynotePath), `${path.basename(keynotePath, path.extname(keynotePath))}.keynote-bridge.pptx`);

  if (!options.exportedPptxPath) {
    await exportKeynoteToPptx(keynotePath, pptxPath);
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
  await importPptxAndSaveKeynote(pptxPath, keynotePath);
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

export async function exportKeynoteToPptx(keynotePath: string, pptxPath: string): Promise<void> {
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
`);
}

export async function importPptxAndSaveKeynote(pptxPath: string, keynotePath: string): Promise<void> {
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
`);
}

async function runAppleScript(script: string): Promise<void> {
  if (process.platform !== "darwin") {
    throw new Error("Keynote conversion requires macOS with Keynote installed.");
  }
  try {
    await execFileAsync("osascript", ["-e", script], { maxBuffer: 1024 * 1024 * 8 });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Keynote automation failed. Confirm Keynote is installed and allow automation permission when macOS prompts. ${message}`);
  }
}

function escapeAppleScriptString(value: string): string {
  return value.replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}
