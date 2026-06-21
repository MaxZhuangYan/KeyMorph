import { access, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";
import { pathToFileURL } from "node:url";

import type { DeckIR } from "../ir/index.ts";
import {
  aggregatePixelFidelityResults,
  comparePngFiles,
  type PixelFidelityOptions,
  type PixelFidelityResult
} from "../report/fidelity.ts";
import {
  createDeckTimeline,
  createRuntimeFrameSnapshot,
  renderHtmlDocument,
  resolveDeckTime,
  type DeckTimeResolution,
  type RuntimeFrameSnapshot
} from "../runtime/index.ts";

export interface VideoExportOptions {
  fps?: number;
  scale?: number;
  keepFrames?: boolean;
  ffmpegPath?: string;
  framesDir?: string;
}

export interface VideoFrameCaptureOptions extends VideoExportOptions {
  outputDir?: string;
  captureRuntimeSnapshots?: boolean;
  includeInactiveSlides?: boolean;
  frameSettleAnimationFrames?: number;
  onBeforeFrameCapture?: (frame: VideoFramePlan) => void | Promise<void>;
  onAfterFrameCapture?: (frame: CapturedVideoFrame) => void | Promise<void>;
}

export interface VideoFrameCaptureResult {
  plan: VideoExportPlan;
  framesDir: string;
  htmlPath: string;
  frames: CapturedVideoFrame[];
  cleanup?: () => Promise<void>;
}

export interface VideoFrameExtractionResult {
  plan: VideoExportPlan;
  framesDir: string;
  frames: VideoFramePlan[];
}

export interface CapturedVideoFrame extends VideoFramePlan {
  filePath: string;
  runtimeSnapshot?: RuntimeFrameSnapshot;
  expectedSnapshot: RuntimeFrameSnapshot;
}

export interface VideoFrameDiagnostic extends VideoFramePlan {
  snapshot: RuntimeFrameSnapshot;
}

export interface PixelFrameFidelityOptions extends PixelFidelityOptions {
  diffPath?: string;
  includeAA?: boolean;
}

export type PixelFrameFidelityResult = PixelFidelityResult;

export interface VideoFrameFidelityOptions extends PixelFrameFidelityOptions {
  diffDir?: string;
  reportPath?: string;
  framePassThreshold?: number;
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
    matchedFrames: number;
    mismatchedFrames: number;
    passingFrames: number;
    failingFrames: number;
    mismatchRatio: number;
    meanMismatchRatio: number;
    maxMismatchRatio: number;
    meanPixelFidelityScore: number;
    minPixelFidelityScore: number;
    maxPixelFidelityScore: number;
    frameFidelityPassThreshold: number;
    transitionFrameCount: number;
    contentFrameCount: number;
    worstFrame?: VideoFrameFidelityEntry;
    bestFrame?: VideoFrameFidelityEntry;
    bySlide: VideoFrameFidelitySlideSummary[];
  };
  reportPath?: string;
}

export interface VideoFrameFidelitySlideSummary {
  slideIndex: number;
  slideId: string;
  frameCount: number;
  transitionFrameCount: number;
  totalPixels: number;
  mismatchedPixels: number;
  mismatchRatio: number;
  meanPixelFidelityScore: number;
  worstFrame?: number;
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
  canCaptureFrames: boolean;
  canExportVideo: boolean;
  canComparePng: boolean;
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

export function createVideoFrameDiagnostics(
  deck: DeckIR,
  options: Pick<VideoExportOptions, "fps" | "scale"> & {
    totalFrames?: number;
    framePathPrefix?: string;
    includeInactiveSlides?: boolean;
  } = {}
): VideoFrameDiagnostic[] {
  return createVideoFramePlan(deck, options).map((frame) => ({
    ...frame,
    snapshot: createRuntimeFrameSnapshot(deck, frame.timeMs, {
      effectiveGroupStates: true,
      includeInactiveSlides: options.includeInactiveSlides
    })
  }));
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
  const missing = await missingFrameCaptureDependencies();
  if (missing.length) throw new VideoExportDependencyError(missing);
  const { chromium } = await importPlaywright();
  const plan = createVideoExportPlan(deck, options);
  const ownsFramesDir = !options.outputDir && !options.framesDir;
  const framesDir = path.resolve(options.outputDir ?? options.framesDir ?? (await mkdtemp(path.join(tmpdir(), "keymorph-frames-"))));
  const htmlPath = path.join(framesDir, "runtime.html");
  let browser: Awaited<ReturnType<typeof chromium.launch>> | undefined;
  let completed = false;
  const capturedFrames: CapturedVideoFrame[] = [];

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
      const expectedSnapshot = createRuntimeFrameSnapshot(deck, frame.timeMs, {
        effectiveGroupStates: true,
        includeInactiveSlides: options.includeInactiveSlides
      });
      await mkdir(path.dirname(framePath), { recursive: true });
      await options.onBeforeFrameCapture?.(frame);
      const runtimeSnapshot = await page.evaluate(
        async ({ timeMs, includeInactiveSlides, animationFrames, captureRuntimeSnapshots }) => {
          if (captureRuntimeSnapshots && window.__keyMorphRuntime?.captureFrame) {
            return window.__keyMorphRuntime.captureFrame(timeMs, {
              includeInactiveSlides,
              animationFrames,
              effectiveGroupStates: true
            });
          }
          window.__keyMorphRuntime?.seekGlobal(timeMs);
          await document.fonts?.ready;
          await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
          return captureRuntimeSnapshots ? window.__keyMorphRuntime?.getGlobalSnapshot?.(timeMs, { includeInactiveSlides }) : undefined;
        },
        {
          timeMs: frame.timeMs,
          includeInactiveSlides: options.includeInactiveSlides ?? false,
          animationFrames: options.frameSettleAnimationFrames ?? 2,
          captureRuntimeSnapshots: options.captureRuntimeSnapshots ?? true
        }
      );
      await page.screenshot({ path: framePath });
      const capturedFrame: CapturedVideoFrame = {
        ...frame,
        filePath: framePath,
        runtimeSnapshot: runtimeSnapshot ?? undefined,
        expectedSnapshot
      };
      capturedFrames.push(capturedFrame);
      await options.onAfterFrameCapture?.(capturedFrame);
    }

    completed = true;
    return {
      plan,
      framesDir,
      htmlPath,
      frames: capturedFrames,
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
    summary: summarizeFrameFidelity(frames, normalizeFramePassThreshold(options.framePassThreshold)),
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

export async function extractVideoFramesFromVideo(
  deck: DeckIR,
  inputVideoPath: string,
  options: VideoExportOptions = {}
): Promise<VideoFrameExtractionResult> {
  const ffmpeg = options.ffmpegPath ?? "ffmpeg";
  const missing = await missingFfmpegDependencies(ffmpeg);
  if (missing.length) throw new VideoExportDependencyError(missing);

  const plan = createVideoExportPlan(deck, options);
  const ownsFramesDir = !options.framesDir;
  const framesDir = path.resolve(options.framesDir ?? (await mkdtemp(path.join(tmpdir(), "keymorph-extracted-frames-"))));
  await rm(framesDir, { recursive: true, force: true });
  await mkdir(framesDir, { recursive: true });

  try {
    await run(ffmpeg, [
      "-y",
      "-i",
      inputVideoPath,
      "-vf",
      `fps=${plan.fps},scale=${plan.outputWidth}:${plan.outputHeight}:flags=lanczos`,
      "-frames:v",
      String(plan.totalFrames),
      "-start_number",
      "0",
      path.join(framesDir, "frame-%06d.png")
    ]);
    return { plan, framesDir, frames: plan.frames };
  } catch (error) {
    if (ownsFramesDir) await rm(framesDir, { recursive: true, force: true });
    throw error;
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

function summarizeFrameFidelity(
  frames: VideoFrameFidelityEntry[],
  frameFidelityPassThreshold: number
): VideoFrameFidelityReport["summary"] {
  const aggregate = aggregatePixelFidelityResults(frames, { passThreshold: frameFidelityPassThreshold });
  const frameCount = frames.length;
  const transitionFrameCount = frames.filter((frame) => frame.resolution.inTransition).length;
  const contentFrameCount = frameCount - transitionFrameCount;
  const worstFrame = frames.reduce<VideoFrameFidelityEntry | undefined>((worst, frame) => {
    if (!worst) return frame;
    if (frame.pixelFidelityScore < worst.pixelFidelityScore) return frame;
    if (frame.pixelFidelityScore === worst.pixelFidelityScore && frame.mismatchRatio > worst.mismatchRatio) return frame;
    return worst;
  }, undefined);
  const bestFrame = frames.reduce<VideoFrameFidelityEntry | undefined>((best, frame) => {
    if (!best) return frame;
    if (frame.pixelFidelityScore > best.pixelFidelityScore) return frame;
    if (frame.pixelFidelityScore === best.pixelFidelityScore && frame.mismatchRatio < best.mismatchRatio) return frame;
    return best;
  }, undefined);

  return {
    frameCount,
    totalPixels: aggregate.totalPixels,
    comparedPixels: aggregate.comparedPixels,
    mismatchedPixels: aggregate.mismatchedPixels,
    matchedFrames: aggregate.matchedItems,
    mismatchedFrames: aggregate.mismatchedItems,
    passingFrames: aggregate.passingItems,
    failingFrames: aggregate.failingItems,
    mismatchRatio: aggregate.mismatchRatio,
    meanMismatchRatio: aggregate.meanMismatchRatio,
    maxMismatchRatio: aggregate.maxMismatchRatio,
    meanPixelFidelityScore: aggregate.meanPixelFidelityScore,
    minPixelFidelityScore: aggregate.minPixelFidelityScore,
    maxPixelFidelityScore: aggregate.maxPixelFidelityScore,
    frameFidelityPassThreshold: aggregate.passThreshold,
    transitionFrameCount,
    contentFrameCount,
    worstFrame,
    bestFrame,
    bySlide: summarizeFrameFidelityBySlide(frames)
  };
}

function summarizeFrameFidelityBySlide(frames: VideoFrameFidelityEntry[]): VideoFrameFidelitySlideSummary[] {
  const groups = new Map<string, VideoFrameFidelityEntry[]>();
  for (const frame of frames) {
    const key = `${frame.resolution.slideIndex}:${frame.resolution.slideId}`;
    const list = groups.get(key) ?? [];
    list.push(frame);
    groups.set(key, list);
  }

  return Array.from(groups.values()).map((slideFrames) => {
    const first = slideFrames[0];
    const frameCount = slideFrames.length;
    const totalPixels = slideFrames.reduce((sum, frame) => sum + frame.totalPixels, 0);
    const mismatchedPixels = slideFrames.reduce((sum, frame) => sum + frame.mismatchedPixels, 0);
    const meanPixelFidelityScore = frameCount
      ? slideFrames.reduce((sum, frame) => sum + frame.pixelFidelityScore, 0) / frameCount
      : 1;
    const worst = slideFrames.reduce<VideoFrameFidelityEntry | undefined>((candidate, frame) => {
      if (!candidate) return frame;
      if (frame.pixelFidelityScore < candidate.pixelFidelityScore) return frame;
      if (frame.pixelFidelityScore === candidate.pixelFidelityScore && frame.mismatchRatio > candidate.mismatchRatio) return frame;
      return candidate;
    }, undefined);

    return {
      slideIndex: first?.resolution.slideIndex ?? 0,
      slideId: first?.resolution.slideId ?? "",
      frameCount,
      transitionFrameCount: slideFrames.filter((frame) => frame.resolution.inTransition).length,
      totalPixels,
      mismatchedPixels,
      mismatchRatio: Number((mismatchedPixels / Math.max(1, totalPixels)).toFixed(6)),
      meanPixelFidelityScore: Number(meanPixelFidelityScore.toFixed(4)),
      worstFrame: worst?.frame
    };
  });
}

async function missingVideoDependencies(ffmpegPath = "ffmpeg"): Promise<string[]> {
  const missing: string[] = [];
  try {
    await importPlaywright();
  } catch {
    missing.push("playwright");
  }
  missing.push(...(await missingFfmpegDependencies(ffmpegPath)));
  return missing;
}

async function missingFfmpegDependencies(ffmpegPath = "ffmpeg"): Promise<string[]> {
  if (ffmpegPath.includes(path.sep)) {
    try {
      await access(ffmpegPath);
    } catch {
      return ["ffmpeg"];
    }
  } else if (!(await commandAvailable(ffmpegPath))) {
    return ["ffmpeg"];
  }
  return [];
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
  const available = {
    playwright: !missing.includes("playwright"),
    browser: !missing.includes("playwright") && !missing.includes("playwright chromium browser"),
    ffmpeg: !missing.includes("ffmpeg"),
    pngFidelity: true,
    pixelmatch: !pixelMissing.includes("pixelmatch"),
    pngjs: !pixelMissing.includes("pngjs")
  };
  return {
    missing,
    available,
    canCaptureFrames: available.playwright && available.browser,
    canExportVideo: available.playwright && available.browser && available.ffmpeg,
    canComparePng: available.pngFidelity,
    guidance: videoDependencyGuidance(missing)
  };
}

function normalizeFramePassThreshold(value: number | undefined): number {
  if (value === undefined || !Number.isFinite(value)) return 0.995;
  return Math.max(0, Math.min(1, value));
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
