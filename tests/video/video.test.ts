import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import {
  comparePixelFrames,
  createVideoExportPlan,
  createVideoFramePlan,
  describeVideoDependencies,
  resolveVideoDependencies,
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
    assert.equal(status.missing.includes("ffmpeg"), true);
    assert.ok(status.guidance.some((item) => item.includes("ffmpegPath")));
  });

  test("compares rendered frame PNGs with bundled pixel tools", async (context) => {
    const status = await describeVideoDependencies("keymorph-missing-ffmpeg");
    if (!status.available.pixelmatch || !status.available.pngjs) {
      context.skip("pixelmatch/pngjs are not available in local or bundled runtime paths.");
      return;
    }

    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-video-pixel-"));
    const referencePath = path.join(dir, "reference.png");
    const actualPath = path.join(dir, "actual.png");
    const diffPath = path.join(dir, "diff.png");
    await writeFile(referencePath, RED_RED_PNG);
    await writeFile(actualPath, RED_BLUE_PNG);

    const result = await comparePixelFrames(referencePath, actualPath, { threshold: 0, diffPath });

    assert.equal(result.dimensionsMatch, true);
    assert.equal(result.comparedPixels, 2);
    assert.equal(result.mismatchedPixels, 1);
    assert.equal(result.mismatchRatio, 0.5);
    assert.equal(result.diffPath, diffPath);
  });
});

const RED_RED_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAD0lEQVR4AWP8z8DwnwEIAA0FAgA+gZVNAAAAAElFTkSuQmCC",
  "base64"
);
const RED_BLUE_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4AWP4z8Dwn4Hh/38AD/kD/dtv74kAAAAASUVORK5CYII=",
  "base64"
);
