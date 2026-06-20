import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import {
  createVideoExportPlan,
  createVideoFramePlan,
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
  });

  test("reports ffmpeg separately from Playwright availability", async () => {
    const missing = await resolveVideoDependencies("keymorph-missing-ffmpeg");

    assert.deepEqual(missing, ["ffmpeg"]);
  });
});
