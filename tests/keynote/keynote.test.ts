import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { exportKeynoteToPptx } from "../../src/keynote/index.ts";

describe("Keynote bridge", () => {
  test("fails with a clear local automation error for missing input or unavailable Keynote", async () => {
    await assert.rejects(
      () => exportKeynoteToPptx("/tmp/missing.key", "/tmp/missing.pptx"),
      /Keynote automation failed|Keynote conversion requires macOS/
    );
  });
});
