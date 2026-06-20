import { mkdtemp, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import { exportIrToPptx } from "../../src/pptx/index.ts";

describe("PPTX export", () => {
  test("writes a PPTX file for the demo deck", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "deck.pptx");

    await exportIrToPptx(createDemoDeck(), out);

    assert.ok((await stat(out)).size > 0);
  });
});
