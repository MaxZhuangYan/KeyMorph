import { access, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";
import { pathToFileURL } from "node:url";

import type { DeckIR } from "../ir/index.ts";
import { comparePngFiles, type PixelFidelityOptions, type PixelFidelityResult } from "../report/fidelity.ts";
import { createDeckTimeline, renderHtmlDocument, resolveDeckTime, type DeckTimeResolution } from "../runtime/index.ts";

export interface VideoExportOptions {
  fps?: number;
  scale?: number;
  keepFrames?: boolean;
  ffmpegPath?: string;
  framesDir?: string;
}

export interface VideoFrameCaptureOptions extends VideoExportOptions {
  outputDir?: string;
}

export interface VideoFrameCaptureResult {
  plan: VideoExportPlan;
  framesDir: string;
  htmlPath: string;
  frames: CapturedVideoFrame[];
  cleanup?: () => Promise<void>;
}

export interface CapturedVideoFrame extends VideoFramePlan {
  filePath: string;
}

export interface PixelFrameFidelityOptions extends PixelFidelityOptions {
  diffPath?: string;
  includeAA?: boolean;
}

export type PixelFrameFidelityResult = PixelFidelityResult;

export interface VideoFrameFidelityOptions extends PixelFrameFidelityOptions {
  diffDir?: string;
  reportPath?: string;
}

export interface VideoFrameFidelityEntry extends PixelFidelityResult {
  frame: number;
  timeMs: number;
  referencePath: string;
  actualPath: string;
  outputPath: string;
  resolution: DeckTimeResolution;
}

export interface VideoFrameFidelityReport {
  generatedAt: string;
  plan: Omit<VideoExportPlan, "frames">;
  frames: VideoFrameFidelityEntry[];
  summary: {
    frameCount: number;
    totalPixels: number;
    comparedPixels: number;
    mismatchedPixels: number;
    mismatchRatio: number;
    meanPixelFidelityScore: number;
    worstFrame?: VideoFrameFidelityEntry;
  };
  reportPath?: string;
}

export interface VideoDependencyStatus {
  missing: string[];
  available: {
    playwright: boolean;
    browser: boolean;
    ffmpeg: boolean;
    pngFidelity: boolean;
    pixelmatch: boolean;
    pngjs: boolean;
  };
  guidance: string[];
}

export interface VideoExportPlan {
  width: number;
  height: number;
  scale: number;
  fps: number;
  durationMs: number;
  totalFrames: number;
  outputWidth: number;
  outputHeight: number;
  frames: VideoFramePlan[];
}

export interface VideoFramePlan {
  frame: number;
  timeMs: number;
  outputPath: string;
  resolution: DeckTimeResolution;
}

export class VideoExportDependencyError extends Error {
  readonly missing: string[];
  readonly guidance: string[];

  constructor(missing: string[], guidance = videoDependencyGuidance(missing)) {
    super(`Video export requires missing local dependencies: ${missing.join(", ")}. ${guidance.join(" ")}`);
    this.name = "VideoExportDependencyError";
    this.missing = missing;
    this.guidance = guidance;
  }
}

export class PixelFrameFidelityDependencyError extends Error {
  readonly missing: string[];
  readonly guidance: string[];

  constructor(missing: string[], guidance = videoDependencyGuidance(missing)) {
    super(`Pixel frame fidelity comparison requires missing local dependencies: ${missing.join(", ")}. ${guidance.join(" ")}`);
    this.name = "PixelFrameFidelityDependencyError";
    this.missing = missing;
    this.guidance = guidance;
  }
}

export function createVideoExportPlan(deck: DeckIR, options: VideoExportOptions = {}): VideoExportPlan {
  const scale = Math.max(1, options.scale ?? 4);
  const fps = Math.max(1, options.fps ?? 30);
  const timeline = createDeckTimeline(deck);
  const durationMs = timeline.durationMs;
  const totalFrames = Math.max(1, Math.ceil((durationMs / 1000) * fps));
  return {
    width: deck.deck.size.width,
    height: deck.deck.size.height,
    scale,
    fps,
    durationMs,
    totalFrames,
    outputWidth: Math.round(deck.deck.size.width * scale),
    outputHeight: Math.round(deck.deck.size.height * scale),
    frames: createVideoFramePlan(deck, { fps, totalFrames })
  };
}

export function createVideoFramePlan(
  deck: DeckIR,
  options: Pick<VideoExportOptions, "fps"> & { totalFrames?: number; framePathPrefix?: string } = {}
): VideoFramePlan[] {
  const fps = Math.max(1, options.fps ?? 30);
  const durationMs = createDeckTimeline(deck).durationMs;
  const totalFrames = options.totalFrames ?? Math.max(1, Math.ceil((durationMs / 1000) * fps));
  const prefix = options.framePathPrefix ?? "frame";

  return Array.from({ length: totalFrames }, (_, frame) => {
    const timeMs = Math.min(durationMs, (frame / fps) * 1000);
    return {
      frame,
      timeMs,
      outputPath: `${prefix}-${String(frame).padStart(6, "0")}.png`,
      resolution: resolveDeckTime(deck, timeMs)
    };
  });
}

export async function comparePixelFrames(
  referencePngPath: string,
  actualPngPath: string,
  options: PixelFrameFidelityOptions = {}
): Promise<PixelFrameFidelityResult> {
  return comparePngFiles(referencePngPath, actualPngPath, options);
}

export async function captureVideoFrames(
  deck: DeckIR,
  options: VideoFrameCaptureOptions = {}
): Promise<VideoFrameCaptureResult> {
  const { chromium } = await importPlaywright().catch(() => {
    throw new VideoExportDependencyError(["playwright"]);
  });
  const plan = createVideoExportPlan(deck, options);
  const ownsFramesDir = !options.outputDir && !options.framesDir;
  const framesDir = path.resolve(options.outputDir ?? options.framesDir ?? (await mkdtemp(path.join(tmpdir(), "keymorph-frames-"))));
  const htmlPath = path.join(framesDir, "runtime.html");
  let browser: Awaited<ReturnType<typeof chromium.launch>> | undefined;
  let completed = false;

  try {
    await mkdir(framesDir, { recursive: true });
    await writeFile(htmlPath, renderHtmlDocument(deck, { controls: false, stageScale: plan.scale }), "utf8");
    try {
      browser = await chromium.launch({ headless: true });
    } catch {
      throw new VideoExportDependencyError(["playwright chromium browser"]);
    }
    const page = await browser.newPage({
      viewport: { width: plan.outputWidth, height: plan.outputHeight },
      deviceScaleFactor: 1
    });
    await page.goto(pathToFileURL(htmlPath).toString(), { waitUntil: "load" });

    for (const frame of plan.frames) {
      const framePath = path.join(framesDir, frame.outputPath);
      await mkdir(path.dirname(framePath), { recursive: true });
      await page.evaluate(async (time) => {
        window.__keyMorphRuntime?.seekGlobal(time);
        await document.fonts?.ready;
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      }, frame.timeMs);
      await page.screenshot({ path: framePath });
    }

    completed = true;
    return {
      plan,
      framesDir,
      htmlPath,
      frames: plan.frames.map((frame) => ({ ...frame, filePath: path.join(framesDir, frame.outputPath) })),
      cleanup: ownsFramesDir ? () => rm(framesDir, { recursive: true, force: true }) : undefined
    };
  } finally {
    await browser?.close();
    if (!completed && ownsFramesDir) await rm(framesDir, { recursive: true, force: true });
  }
}

export async function createVideoFrameFidelityReport(
  deck: DeckIR,
  referenceFramesDir: string,
  actualFramesDir: string,
  options: VideoFrameFidelityOptions = {}
): Promise<VideoFrameFidelityReport> {
  const plan = createVideoExportPlan(deck, options);
  if (options.diffDir) await mkdir(options.diffDir, { recursive: true });

  const frames: VideoFrameFidelityEntry[] = [];
  for (const frame of plan.frames) {
    const referencePath = path.join(referenceFramesDir, frame.outputPath);
    const actualPath = path.join(actualFramesDir, frame.outputPath);
    const diffPath = options.diffDir ? path.join(options.diffDir, frame.outputPath) : undefined;
    const result = await comparePixelFrames(referencePath, actualPath, { ...options, diffPath });
    frames.push({
      ...result,
      frame: frame.frame,
      timeMs: frame.timeMs,
      referencePath,
      actualPath,
      outputPath: frame.outputPath,
      resolution: frame.resolution
    });
  }

  const report: VideoFrameFidelityReport = {
    generatedAt: new Date().toISOString(),
    plan: summarizePlan(plan),
    frames,
    summary: summarizeFrameFidelity(frames),
    reportPath: options.reportPath
  };
  if (options.reportPath) {
    await mkdir(path.dirname(options.reportPath), { recursive: true });
    await writeFile(options.reportPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");
  }
  return report;
}

export async function exportIrToVideo(
  deck: DeckIR,
  outputPath: string,
  options: VideoExportOptions = {}
): Promise<void> {
  const missing = await missingVideoDependencies(options.ffmpegPath);
  if (missing.length) throw new VideoExportDependencyError(missing);

  const ffmpeg = options.ffmpegPath ?? "ffmpeg";
  const capture = await captureVideoFrames(deck, { ...options, outputDir: options.framesDir });
  try {
    await run(ffmpeg, [
      "-y",
      "-framerate",
      String(capture.plan.fps),
      "-i",
      path.join(capture.framesDir, "frame-%06d.png"),
      "-c:v",
      "libx264",
      "-pix_fmt",
      "yuv420p",
      outputPath
    ]);
  } finally {
    if (!options.keepFrames) await capture.cleanup?.();
  }
}

function summarizePlan(plan: VideoExportPlan): Omit<VideoExportPlan, "frames"> {
  return {
    width: plan.width,
    height: plan.height,
    scale: plan.scale,
    fps: plan.fps,
    durationMs: plan.durationMs,
    totalFrames: plan.totalFrames,
    outputWidth: plan.outputWidth,
    outputHeight: plan.outputHeight
  };
}

function summarizeFrameFidelity(frames: VideoFrameFidelityEntry[]): VideoFrameFidelityReport["summary"] {
  const frameCount = frames.length;
  const totalPixels = frames.reduce((sum, frame) => sum + frame.totalPixels, 0);
  const comparedPixels = frames.reduce((sum, frame) => sum + frame.comparedPixels, 0);
  const mismatchedPixels = frames.reduce((sum, frame) => sum + frame.mismatchedPixels, 0);
  const meanPixelFidelityScore = frameCount
    ? frames.reduce((sum, frame) => sum + frame.pixelFidelityScore, 0) / frameCount
    : 1;
  const worstFrame = frames.reduce<VideoFrameFidelityEntry | undefined>((worst, frame) => {
    if (!worst) return frame;
    if (frame.pixelFidelityScore < worst.pixelFidelityScore) return frame;
    if (frame.pixelFidelityScore === worst.pixelFidelityScore && frame.mismatchRatio > worst.mismatchRatio) return frame;
    return worst;
  }, undefined);

  return {
    frameCount,
    totalPixels,
    comparedPixels,
    mismatchedPixels,
    mismatchRatio: Number((mismatchedPixels / Math.max(1, totalPixels)).toFixed(6)),
    meanPixelFidelityScore: Number(meanPixelFidelityScore.toFixed(4)),
    worstFrame
  };
}

async function missingVideoDependencies(ffmpegPath = "ffmpeg"): Promise<string[]> {
  const missing: string[] = [];
  try {
    await importPlaywright();
  } catch {
    missing.push("playwright");
  }
  if (ffmpegPath.includes(path.sep)) {
    try {
      await access(ffmpegPath);
    } catch {
      missing.push("ffmpeg");
    }
  } else if (!(await commandAvailable(ffmpegPath))) {
    missing.push("ffmpeg");
  }
  return missing;
}

export async function resolveVideoDependencies(ffmpegPath = "ffmpeg"): Promise<string[]> {
  return missingVideoDependencies(ffmpegPath);
}

export async function describeVideoDependencies(ffmpegPath = "ffmpeg"): Promise<VideoDependencyStatus> {
  const [videoMissing, frameCaptureMissing, pixelMissing] = await Promise.all([
    missingVideoDependencies(ffmpegPath),
    missingFrameCaptureDependencies(),
    missingPixelFrameDependencies()
  ]);
  const missing = Array.from(new Set([...videoMissing, ...frameCaptureMissing]));
  return {
    missing,
    available: {
      playwright: !missing.includes("playwright"),
      browser: !missing.includes("playwright") && !missing.includes("playwright chromium browser"),
      ffmpeg: !missing.includes("ffmpeg"),
      pngFidelity: true,
      pixelmatch: !pixelMissing.includes("pixelmatch"),
      pngjs: !pixelMissing.includes("pngjs")
    },
    guidance: videoDependencyGuidance(missing)
  };
}

function videoDependencyGuidance(missing: string[]): string[] {
  const guidance = missing.map((dependency) => {
    if (dependency === "ffmpeg") return "Install ffmpeg on PATH or pass VideoExportOptions.ffmpegPath with an executable ffmpeg path.";
    if (dependency === "playwright") {
      return "Install Playwright in the local project or set KEYMORPH_PLAYWRIGHT_MODULE to a bundled playwright/index.mjs path.";
    }
    if (dependency === "playwright chromium browser") return "Install a Playwright Chromium browser for headless rendering.";
    if (dependency === "pixelmatch") {
      return "Install pixelmatch in the local project or set KEYMORPH_PIXELMATCH_MODULE to a bundled pixelmatch/index.js path.";
    }
    if (dependency === "pngjs") return "Install pngjs in the local project or set KEYMORPH_PNGJS_MODULE to a bundled pngjs/lib/png.js path.";
    return `Install or configure ${dependency}.`;
  });
  return guidance.length ? guidance : ["All optional video dependencies are available."];
}

async function missingPixelFrameDependencies(): Promise<string[]> {
  const missing: string[] = [];
  try {
    await importPixelmatch();
  } catch {
    missing.push("pixelmatch");
  }
  try {
    await importPngjs();
  } catch {
    missing.push("pngjs");
  }
  return missing;
}

async function missingFrameCaptureDependencies(): Promise<string[]> {
  let playwright: Awaited<ReturnType<typeof importPlaywright>>;
  try {
    playwright = await importPlaywright();
  } catch {
    return ["playwright"];
  }

  if (!(await canLaunchChromium(playwright))) return ["playwright chromium browser"];
  return [];
}

async function canLaunchChromium(playwright: Awaited<ReturnType<typeof importPlaywright>>): Promise<boolean> {
  let browser: Awaited<ReturnType<typeof playwright.chromium.launch>> | undefined;
  try {
    browser = await playwright.chromium.launch({ headless: true });
    return true;
  } catch {
    return false;
  } finally {
    await browser?.close().catch(() => undefined);
  }
}

async function importPixelFrameTools(): Promise<{ pixelmatch: Pixelmatch; PNG: PngConstructor }> {
  const [pixelmatch, pngjs] = await Promise.all([importPixelmatch(), importPngjs()]);
  return { pixelmatch, PNG: pngjs.PNG };
}

type Pixelmatch = (
  img1: Uint8Array | Uint8ClampedArray,
  img2: Uint8Array | Uint8ClampedArray,
  output: Uint8Array | Uint8ClampedArray | undefined,
  width: number,
  height: number,
  options?: { threshold?: number; includeAA?: boolean }
) => number;

interface PngImage {
  width: number;
  height: number;
  data: Uint8Array | Uint8ClampedArray;
}

interface PngConstructor {
  new (options: { width: number; height: number }): PngImage;
  sync: {
    read(data: Buffer): PngImage;
    write(image: PngImage): Buffer;
  };
}

async function importPixelmatch(): Promise<Pixelmatch> {
  try {
    const module = await import("pixelmatch");
    return module.default as Pixelmatch;
  } catch {
    for (const candidate of moduleFallbackPaths("KEYMORPH_PIXELMATCH_MODULE", "pixelmatch/index.js")) {
      try {
        await access(candidate);
        const module = await import(pathToFileURL(candidate).toString());
        return module.default as Pixelmatch;
      } catch {
        // Try the next configured runtime path.
      }
    }
    throw new Error("pixelmatch module was not found.");
  }
}

async function importPngjs(): Promise<{ PNG: PngConstructor }> {
  try {
    const module = await import("pngjs");
    return { PNG: (module.PNG ?? module.default?.PNG ?? module.default) as PngConstructor };
  } catch {
    for (const candidate of moduleFallbackPaths("KEYMORPH_PNGJS_MODULE", "pngjs/lib/png.js")) {
      try {
        await access(candidate);
        const module = await import(pathToFileURL(candidate).toString());
        return { PNG: (module.PNG ?? module.default?.PNG ?? module.default) as PngConstructor };
      } catch {
        // Try the next configured runtime path.
      }
    }
    throw new Error("pngjs module was not found.");
  }
}

async function readPng(PNG: PngConstructor, filePath: string): Promise<PngImage> {
  return PNG.sync.read(await readFile(filePath));
}

function normalizePngData(PNG: PngConstructor, image: PngImage, width: number, height: number): Uint8Array | Uint8ClampedArray {
  if (image.width === width && image.height === height) return image.data;
  const normalized = new PNG({ width, height });
  normalized.data.fill(0);
  for (let y = 0; y < image.height; y += 1) {
    const sourceStart = y * image.width * 4;
    const targetStart = y * width * 4;
    normalized.data.set(image.data.subarray(sourceStart, sourceStart + image.width * 4), targetStart);
  }
  return normalized.data;
}

function normalizePixelmatchThreshold(value: number | undefined): number {
  if (value === undefined || !Number.isFinite(value)) return 0.1;
  return Math.max(0, Math.min(1, value));
}

async function importPlaywright(): Promise<typeof import("playwright")> {
  try {
    return await import("playwright");
  } catch {
    for (const candidate of playwrightFallbackPaths()) {
      try {
        await access(candidate);
        return import(pathToFileURL(candidate).toString());
      } catch {
        // Try the next configured runtime path.
      }
    }
    throw new Error("Playwright module was not found.");
  }
}

function playwrightFallbackPaths(): string[] {
  return moduleFallbackPaths("KEYMORPH_PLAYWRIGHT_MODULE", "playwright/index.mjs");
}

function moduleFallbackPaths(envName: string, packageRelativePath: string): string[] {
  return [
    process.env[envName],
    process.env.HOME
      ? path.join(process.env.HOME, ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules", packageRelativePath)
      : undefined
  ].filter((candidate): candidate is string => Boolean(candidate));
}

function commandAvailable(command: string): Promise<boolean> {
  return new Promise((resolve) => {
    const child = spawn(command, ["-version"], { stdio: "ignore" });
    child.once("error", () => resolve(false));
    child.once("exit", (code) => resolve(code === 0));
  });
}

function run(command: string, args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: ["ignore", "ignore", "pipe"] });
    let stderr = "";
    child.stderr.on("data", (chunk: Buffer) => {
      stderr += chunk.toString("utf8");
    });
    child.once("error", reject);
    child.once("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${command} exited with code ${code}: ${stderr}`));
    });
  });
}
