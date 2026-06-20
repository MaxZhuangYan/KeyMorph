import { access, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";
import { pathToFileURL } from "node:url";

import type { DeckIR } from "../ir/index.ts";
import { renderHtmlDocument } from "../runtime/index.ts";

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
  const durationMs = deck.deck.slides.reduce(
    (total, slide) => total + Math.max(1, slide.timeline?.durationMs ?? 2500),
    0
  );
  return {
    width: deck.deck.size.width,
    height: deck.deck.size.height,
    scale,
    fps,
    durationMs,
    totalFrames: Math.max(1, Math.ceil((durationMs / 1000) * fps)),
    outputWidth: Math.round(deck.deck.size.width * scale),
    outputHeight: Math.round(deck.deck.size.height * scale)
  };
}

export async function exportIrToVideo(
  deck: DeckIR,
  outputPath: string,
  options: VideoExportOptions = {}
): Promise<void> {
  const missing = await missingVideoDependencies(options.ffmpegPath);
  if (missing.length) throw new VideoExportDependencyError(missing);

  const { chromium } = await import("playwright").catch((error: unknown) => {
    throw new VideoExportDependencyError(["playwright"]);
  });
  const ffmpeg = options.ffmpegPath ?? "ffmpeg";
  const plan = createVideoExportPlan(deck, options);
  const workspace = await mkdtemp(path.join(tmpdir(), "keymorph-video-"));
  const htmlPath = path.join(workspace, "runtime.html");

  let browser: Awaited<ReturnType<typeof chromium.launch>> | undefined;
  try {
    await writeFile(htmlPath, renderHtmlDocument(deck, { controls: false }), "utf8");
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({
      viewport: { width: plan.outputWidth, height: plan.outputHeight },
      deviceScaleFactor: 1
    });
    await page.goto(pathToFileURL(htmlPath).toString(), { waitUntil: "load" });

    for (let frame = 0; frame < plan.totalFrames; frame += 1) {
      const timeMs = (frame / plan.fps) * 1000;
      await page.evaluate((time) => window.__keyMorphRuntime?.seek(time), timeMs);
      await page.screenshot({ path: path.join(workspace, `frame-${String(frame).padStart(6, "0")}.png`) });
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
    await import("playwright");
  } catch {
    missing.push("playwright");
  }
  if (ffmpegPath.includes(path.sep)) {
    try {
      await access(ffmpegPath);
    } catch {
      missing.push("ffmpeg");
    }
  }
  return missing;
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
