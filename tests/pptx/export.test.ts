import { mkdtemp, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import { exportIrToPptx, parsePptxToIr } from "../../src/pptx/index.ts";

describe("PPTX export", () => {
  test("writes a PPTX file for the demo deck", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "deck.pptx");

    await exportIrToPptx(createDemoDeck(), out);

    assert.ok((await stat(out)).size > 0);
  });

  test("parses static slides from a KeyMorph-exported PPTX", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "deck.pptx");

    await exportIrToPptx(createDemoDeck(), out);
    const parsed = await parsePptxToIr(out);

    assert.equal(parsed.deck.slides.length, 2);
    assert.equal(parsed.conversion?.status, "partial");
    assert.ok(parsed.deck.slides[0].objects.some((object) => object.type === "text"));
    assert.match(JSON.stringify(parsed), /KeyMorph/);
  });
});
