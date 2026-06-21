import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import { decodePng, encodePng, type RgbaImage } from "../../src/report/fidelity.ts";
import {
  captureVideoFrames,
  comparePixelFrames,
  createSegmentPlan,
  createVideoFrameDiagnostics,
  createVideoFrameFidelityReport,
  createVideoExportPlan,
  createVideoFramePlan,
  describeVideoDependencies,
  resolveVideoDependencies,
  splitVideoIntoSegments,
  VideoExportDependencyError
} from "../../src/video/index.ts";

describe("video export planning", () => {
  test("creates a deterministic frame plan from the deck timeline", () => {
    const plan = createVideoExportPlan(createDemoDeck(), { fps: 10, scale: 2 });

    assert.equal(plan.width, 1280);
    assert.equal(plan.height, 720);
    assert.equal(plan.outputWidth, 2560);
    assert.equal(plan.outputHeight, 1440);
    assert.equal(plan.durationMs, 5300);
    assert.equal(plan.totalFrames, 53);
    assert.equal(plan.frames.length, 53);
    assert.deepEqual(
      plan.frames.slice(0, 3).map((frame) => [frame.frame, frame.timeMs, frame.outputPath]),
      [
        [0, 0, "frame-000000.png"],
        [1, 100, "frame-000001.png"],
        [2, 200, "frame-000002.png"]
      ]
    );
  });

  test("marks frames that occur during slide transitions", () => {
    const frames = createVideoFramePlan(createDemoDeck(), { fps: 10 });
    const transitionFrame = frames.find((frame) => frame.timeMs === 3000);

    assert.equal(transitionFrame?.resolution.slideIndex, 1);
    assert.equal(transitionFrame?.resolution.previousSlideIndex, 0);
    assert.equal(transitionFrame?.resolution.inTransition, true);
    assert.equal(Number(transitionFrame?.resolution.transitionProgress.toFixed(3)), 0.444);
  });

  test("creates deterministic runtime diagnostics for planned frames", () => {
    const diagnostics = createVideoFrameDiagnostics(createDemoDeck(), { fps: 10, totalFrames: 36, includeInactiveSlides: true });
    const transitionFrame = diagnostics.find((frame) => frame.resolution.inTransition);

    assert.equal(diagnostics.length, 36);
    assert.equal(diagnostics[0]?.snapshot.globalTimeMs, 0);
    assert.equal(diagnostics[0]?.snapshot.slides.some((slide) => slide.phase === "inactive"), true);
    assert.equal(transitionFrame?.snapshot.resolution.inTransition, true);
    assert.equal(transitionFrame?.snapshot.transition?.currentSlideIndex, 1);
    assert.equal(transitionFrame?.snapshot.slides.some((slide) => slide.phase === "previous"), true);
  });

  test("creates slide video segments from content and timeline event boundaries", () => {
    const segments = createSegmentPlan(createDemoDeck());

    assert.deepEqual(
      segments.map((segment) => [segment.id, segment.slideIndex, segment.clickIndex, segment.startMs, segment.endMs, segment.outputName]),
      [
        ["slide-1-segment-1", 0, 0, 0, 350, "slide-1-segment-1.mp4"],
        ["slide-1-segment-2", 0, 1, 350, 1050, "slide-1-segment-2.mp4"],
        ["slide-1-segment-3", 0, 2, 1050, 1600, "slide-1-segment-3.mp4"],
        ["slide-1-segment-4", 0, 3, 1600, 2500, "slide-1-segment-4.mp4"],
        ["slide-1-segment-5", 0, 4, 2500, 2600, "slide-1-segment-5.mp4"],
        ["slide-2-segment-1", 1, 0, 3500, 3700, "slide-2-segment-1.mp4"],
        ["slide-2-segment-2", 1, 1, 3700, 4320, "slide-2-segment-2.mp4"],
        ["slide-2-segment-3", 1, 2, 4320, 5300, "slide-2-segment-3.mp4"]
      ]
    );
    assert.deepEqual(
      segments.map((segment) => segment.durationMs),
      [350, 700, 550, 900, 100, 200, 620, 980]
    );
  });

  test("creates dry-run commands for ffmpeg and avconvert segment splitting", async () => {
    const segments = createSegmentPlan(createDemoDeck()).slice(0, 1);
    const ffmpeg = await splitVideoIntoSegments("/tmp/keynote movie.m4v", segments, "/tmp/segments", {
      dryRun: true,
      ffmpegPath: "/opt/bin/ffmpeg"
    });
    const avconvert = await splitVideoIntoSegments("/tmp/keynote movie.m4v", segments, "/tmp/segments", {
      dryRun: true,
      avconvertPath: "/usr/bin/avconvert"
    });

    assert.deepEqual(ffmpeg.commands[0]?.args, [
      "-ss",
      "0",
      "-i",
      "/tmp/keynote movie.m4v",
      "-t",
      "0.35",
      "-c:v",
      "libx264",
      "-pix_fmt",
      "yuv420p",
      "-c:a",
      "aac",
      "/tmp/segments/slide-1-segment-1.mp4"
    ]);
    assert.equal(ffmpeg.commands[0]?.tool, "ffmpeg");
    assert.equal(ffmpeg.commands[0]?.command, "/opt/bin/ffmpeg");
    assert.deepEqual(avconvert.commands[0]?.args, [
      "--source",
      "/tmp/keynote movie.m4v",
      "--output",
      "/tmp/segments/slide-1-segment-1.mp4",
      "--replace",
      "--preset",
      "PresetHighestQuality",
      "--start",
      "0",
      "--duration",
      "0.35"
    ]);
    assert.equal(avconvert.commands[0]?.tool, "avconvert");
    assert.equal(avconvert.commands[0]?.command, "/usr/bin/avconvert");
  });

  test("exposes clear dependency error details", () => {
    const error = new VideoExportDependencyError(["playwright", "ffmpeg"]);

    assert.equal(error.name, "VideoExportDependencyError");
    assert.deepEqual(error.missing, ["playwright", "ffmpeg"]);
    assert.match(error.message, /playwright, ffmpeg/);
    assert.match(error.message, /KEYMORPH_PLAYWRIGHT_MODULE/);
    assert.match(error.message, /ffmpegPath/);
    assert.ok(error.guidance.length >= 2);
  });

  test("reports ffmpeg separately from Playwright availability", async () => {
    const missing = await resolveVideoDependencies("keymorph-missing-ffmpeg");

    assert.deepEqual(missing, ["ffmpeg"]);
  });

  test("describes optional video and pixel-frame dependencies with guidance", async () => {
    const status = await describeVideoDependencies("keymorph-missing-ffmpeg");

    assert.equal(status.available.ffmpeg, false);
    assert.equal(status.available.pngFidelity, true);
    assert.equal(status.canExportVideo, false);
    assert.equal(status.canComparePng, true);
    assert.equal(status.missing.includes("ffmpeg"), true);
    assert.ok(status.guidance.some((item) => item.includes("ffmpegPath")));
  });

  test("compares rendered frame PNGs and writes an optional diff without external pixel tools", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-video-pixel-"));
    const referencePath = path.join(dir, "reference.png");
    const actualPath = path.join(dir, "actual.png");
    const diffPath = path.join(dir, "diff.png");
    await writeFile(referencePath, RED_RED_PNG);
    await writeFile(actualPath, RED_BLUE_PNG);

    const result = await comparePixelFrames(referencePath, actualPath, { threshold: 0, diffPath });

    assert.equal(result.dimensionsMatch, true);
    assert.equal(result.totalPixels, 2);
    assert.equal(result.mismatchedPixels, 1);
    assert.equal(result.mismatchRatio, 0.5);
    assert.equal(result.diffPath, diffPath);
    assert.deepEqual(Array.from(decodePng(await readFile(diffPath)).data.slice(4, 8)), [255, 40, 40, 255]);
  });

  test("creates a frame-level fidelity report with optional per-frame diffs", async () => {
    const deck = createDemoDeck();
    const framePlan = createVideoExportPlan(deck, { fps: 1 }).frames;
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-video-report-"));
    const referenceDir = path.join(dir, "reference");
    const actualDir = path.join(dir, "actual");
    const diffDir = path.join(dir, "diff");
    const reportPath = path.join(dir, "frame-fidelity.json");
    await mkdir(referenceDir, { recursive: true });
    await mkdir(actualDir, { recursive: true });

    for (const frame of framePlan) {
      await writeFile(path.join(referenceDir, frame.outputPath), encodePng(image(1, 1, [255, 255, 255, 255])));
      await writeFile(
        path.join(actualDir, frame.outputPath),
        encodePng(frame.frame === 1 ? image(1, 1, [0, 0, 0, 255]) : image(1, 1, [255, 255, 255, 255]))
      );
    }

    const report = await createVideoFrameFidelityReport(deck, referenceDir, actualDir, {
      fps: 1,
      diffDir,
      reportPath
    });
    const persisted = JSON.parse(await readFile(reportPath, "utf8"));

    assert.equal(report.frames.length, framePlan.length);
    assert.equal(report.summary.frameCount, framePlan.length);
    assert.equal(report.summary.mismatchedPixels, 1);
    assert.equal(report.summary.mismatchedFrames, 1);
    assert.equal(report.summary.matchedFrames, framePlan.length - 1);
    assert.equal(report.summary.failingFrames, 1);
    assert.equal(report.summary.passingFrames, framePlan.length - 1);
    assert.equal(report.summary.mismatchRatio, Number((1 / framePlan.length).toFixed(6)));
    assert.equal(report.summary.meanMismatchRatio, Number((1 / framePlan.length).toFixed(6)));
    assert.equal(report.summary.maxMismatchRatio, 1);
    assert.equal(report.summary.minPixelFidelityScore, report.summary.worstFrame?.pixelFidelityScore);
    assert.equal(report.summary.maxPixelFidelityScore, 1);
    assert.equal(report.summary.worstFrame?.frame, 1);
    assert.equal(report.summary.bestFrame?.frame, 0);
    assert.equal(report.summary.bySlide.length, 2);
    assert.equal(report.summary.bySlide[0]?.mismatchedPixels, 1);
    assert.equal(persisted.summary.mismatchedPixels, 1);
    assert.equal(persisted.summary.bySlide[0].worstFrame, 1);
    assert.equal(decodePng(await readFile(path.join(diffDir, framePlan[1]?.outputPath ?? ""))).width, 1);
  });

  test("capture reports missing browser dependencies gracefully", async (context) => {
    const status = await describeVideoDependencies("keymorph-missing-ffmpeg");
    if (status.available.playwright && status.available.browser) {
      context.skip("Playwright browser is available; missing-browser behavior is not active in this environment.");
      return;
    }

    await assert.rejects(() => captureVideoFrames(createDemoDeck(), { fps: 1, scale: 1 }), (error) => {
      assert.ok(error instanceof VideoExportDependencyError);
      assert.ok(error.missing.includes("playwright") || error.missing.includes("playwright chromium browser"));
      return true;
    });
  });
});

function image(width: number, height: number, data: number[]): RgbaImage {
  return { width, height, data: new Uint8Array(data) };
}

const RED_RED_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAD0lEQVR4AWP8z8DwnwEIAA0FAgA+gZVNAAAAAElFTkSuQmCC",
  "base64"
);
const RED_BLUE_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4AWP4z8Dwn4Hh/38AD/kD/dtv74kAAAAASUVORK5CYII=",
  "base64"
);
