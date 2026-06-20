import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import {
  createDeckTimeline,
  keyframeEventToCssFrames,
  renderHtmlDocument,
  renderSlideMarkup,
  resolveDeckTime
} from "../../src/runtime/index.ts";

describe("HTML runtime rendering", () => {
  test("renders a self-contained runtime document", () => {
    const html = renderHtmlDocument(createDemoDeck());

    assert.match(html, /window\.__KEYMORPH_DECK__/);
    assert.match(html, /KeyMorph/);
    assert.match(html, /__keyMorphRuntime/);
    assert.match(html, /seekGlobal/);
    assert.match(html, /nextSlide/);
    assert.match(html, /activeTimelineTransition/);
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
    assert.match(String(frames[0]?.transform), /scale\(0\.98,0\.98\)/);
  });

  test("builds a deck-level timeline that includes incoming slide transitions", () => {
    const deck = createDemoDeck();
    const timeline = createDeckTimeline(deck);

    assert.equal(timeline.durationMs, 5300);
    assert.deepEqual(timeline.slides.map((slide) => slide.transitionDurationMs), [0, 900]);
    assert.deepEqual(
      timeline.slides.map((slide) => [slide.startMs, slide.contentStartMs, slide.endMs]),
      [
        [0, 0, 2600],
        [2600, 3500, 5300]
      ]
    );
  });

  test("resolves global scrub time across slides and transitions", () => {
    const deck = createDemoDeck();

    assert.deepEqual(
      pickResolution(resolveDeckTime(deck, 2500)),
      { slideIndex: 0, slideTimeMs: 2500, inTransition: false, transitionProgress: 1 }
    );
    assert.deepEqual(
      pickResolution(resolveDeckTime(deck, 3050)),
      { slideIndex: 1, slideTimeMs: 0, inTransition: true, transitionProgress: 0.5 }
    );
    assert.deepEqual(
      pickResolution(resolveDeckTime(deck, 3600)),
      { slideIndex: 1, slideTimeMs: 100, inTransition: false, transitionProgress: 1 }
    );
  });
});

function pickResolution(resolution: ReturnType<typeof resolveDeckTime>) {
  return {
    slideIndex: resolution.slideIndex,
    slideTimeMs: resolution.slideTimeMs,
    inTransition: resolution.inTransition,
    transitionProgress: resolution.transitionProgress
  };
}
