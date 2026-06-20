import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import { keyframeEventToCssFrames, renderHtmlDocument, renderSlideMarkup } from "../../src/runtime/index.ts";

describe("HTML runtime rendering", () => {
  test("renders a self-contained runtime document", () => {
    const html = renderHtmlDocument(createDemoDeck());

    assert.match(html, /window\.__KEYMORPH_DECK__/);
    assert.match(html, /KeyMorph/);
    assert.match(html, /__keyMorphRuntime/);
  });

  test("renders slide objects as positioned markup", () => {
    const deck = createDemoDeck();
    const markup = renderSlideMarkup(deck.deck.slides[0]);

    assert.match(markup, /data-object-id="title-slide-1"/);
    assert.match(markup, /Any slide deck becomes/);
  });

  test("converts keyframe event tracks to CSS frames", () => {
    const event = createDemoDeck().deck.slides[0].timeline?.events[0];

    assert.equal(event?.kind, "keyframes");
    if (event?.kind !== "keyframes") throw new Error("Expected keyframes event.");

    const frames = keyframeEventToCssFrames(event);

    assert.deepEqual({ offset: frames[0]?.offset, opacity: frames[0]?.opacity }, { offset: 0, opacity: "0" });
    assert.deepEqual({ offset: frames[1]?.offset, opacity: frames[1]?.opacity }, { offset: 1, opacity: "1" });
  });
});
