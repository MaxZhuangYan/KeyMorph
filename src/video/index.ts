import { access, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";
import { pathToFileURL } from "node:url";

import type { DeckIR } from "../ir/index.ts";
import { createDeckTimeline, renderHtmlDocument, resolveDeckTime, type DeckTimeResolution } from "../runtime/index.ts";

export interface VideoExportOptions {
  fps?: number;
  scale?: number;
  keepFrames?: boolean;
  ffmpegPath?: string;
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

  constructor(missing: string[]) {
    super(`Video export requires missing local dependencies: ${missing.join(", ")}.`);
    this.name = "VideoExportDependencyError";
    this.missing = missing;
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

export async function exportIrToVideo(
  deck: DeckIR,
  outputPath: string,
  options: VideoExportOptions = {}
): Promise<void> {
  const missing = await missingVideoDependencies(options.ffmpegPath);
  if (missing.length) throw new VideoExportDependencyError(missing);

  const { chromium } = await importPlaywright().catch(() => {
    throw new VideoExportDependencyError(["playwright"]);
  });
  const ffmpeg = options.ffmpegPath ?? "ffmpeg";
  const plan = createVideoExportPlan(deck, options);
  const workspace = await mkdtemp(path.join(tmpdir(), "keymorph-video-"));
  const htmlPath = path.join(workspace, "runtime.html");

  let browser: Awaited<ReturnType<typeof chromium.launch>> | undefined;
  try {
    await writeFile(htmlPath, renderHtmlDocument(deck, { controls: false, stageScale: plan.scale }), "utf8");
    try {
      browser = await chromium.launch({ headless: true });
    } catch (error) {
      throw new VideoExportDependencyError(["playwright chromium browser"]);
    }
    const page = await browser.newPage({
      viewport: { width: plan.outputWidth, height: plan.outputHeight },
      deviceScaleFactor: 1
    });
    await page.goto(pathToFileURL(htmlPath).toString(), { waitUntil: "load" });

    for (const frame of plan.frames) {
      await page.evaluate((time) => window.__keyMorphRuntime?.seekGlobal(time), frame.timeMs);
      await page.screenshot({ path: path.join(workspace, frame.outputPath) });
    }

    await run(ffmpeg, [
      "-y",
      "-framerate",
      String(plan.fps),
      "-i",
      path.join(workspace, "frame-%06d.png"),
      "-c:v",
      "libx264",
      "-pix_fmt",
      "yuv420p",
      outputPath
    ]);
  } finally {
    await browser?.close();
    if (!options.keepFrames) await rm(workspace, { recursive: true, force: true });
  }
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
  return [
    process.env.KEYMORPH_PLAYWRIGHT_MODULE,
    process.env.HOME
      ? path.join(process.env.HOME, ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs")
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
