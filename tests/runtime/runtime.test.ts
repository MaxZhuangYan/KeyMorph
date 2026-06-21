import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import {
  createDeckTimeline,
  createRuntimeFrameSnapshot,
  createSlideTimingPlan,
  easeProgressValue,
  keyframeEventToCssFrames,
  renderHtmlDocument,
  renderSlideMarkup,
  resolveMorphTransitionObjectStates,
  resolveTimelineEventDiagnostics,
  resolveSlideObjectStates,
  resolveDeckTime
} from "../../src/runtime/index.ts";
import type { DeckIR } from "../../src/ir/index.ts";

describe("HTML runtime rendering", () => {
  test("renders a self-contained runtime document", () => {
    const html = renderHtmlDocument(createDemoDeck());

    assert.match(html, /window\.__KEYMORPH_DECK__/);
    assert.match(html, /KeyMorph/);
    assert.match(html, /__keyMorphRuntime/);
    assert.match(html, /seekGlobal/);
    assert.match(html, /nextSlide/);
    assert.match(html, /activeTimelineTransition/);
    assert.match(html, /getCurrentFrameState/);
    assert.match(html, /getGlobalSnapshot/);
    assert.match(html, /getTimelineEventDiagnostics/);
    assert.match(html, /captureFrame/);
    assert.match(html, /seekFrame/);
    assert.match(html, /createTimingPlan/);
    assert.match(html, /setPlaybackRate/);
  });

  test("renders slide objects as positioned markup", () => {
    const deck = createDemoDeck();
    const markup = renderSlideMarkup(deck.deck.slides[0]);

    assert.match(markup, /data-object-id="title-slide-1"/);
    assert.match(markup, /Any slide deck becomes/);
  });

  test("renders native character-build text as grapheme spans", () => {
    const deck = createCharacterBuildDeck();
    const markup = renderSlideMarkup(deck.deck.slides[0]!, deck);

    assert.match(markup, /class="km-text-char"/);
    assert.equal((markup.match(/data-char-index=/g) ?? []).length, 4);
    assert.match(markup, />逐</);
    assert.match(renderHtmlDocument(deck), /applyCharacterTextState/);
  });

  test("does not split non-dissolve character metadata into grapheme spans", () => {
    const deck = createCharacterBuildDeck("blur-in");
    const markup = renderSlideMarkup(deck.deck.slides[0]!, deck);

    assert.doesNotMatch(markup, /class="km-text-char"/);
  });

  test("renders shape text with text layout class and styles", () => {
    const deck = createShapeTextDeck();
    const markup = renderSlideMarkup(deck.deck.slides[0]!, deck);

    assert.match(markup, /class="km-object km-shape km-shape-text"/);
    assert.match(markup, /font-size:22px/);
    assert.match(markup, /line-height:1\.35/);
    assert.match(markup, /color:#123456/);
    assert.match(markup, /white-space:pre-wrap/);
    assert.match(markup, /First\nSecond/);
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

  test("converts rotation and uniform scale keyframes to CSS transforms", () => {
    const frames = keyframeEventToCssFrames({
      id: "spin",
      kind: "keyframes",
      targetId: "box",
      durationMs: 1000,
      tracks: [
        {
          property: "scale",
          keyframes: [
            { offset: 0, value: 1 },
            { offset: 1, value: 1.5 }
          ]
        },
        {
          property: "rotation",
          keyframes: [
            { offset: 0, value: 0 },
            { offset: 1, value: 45 }
          ]
        }
      ]
    });

    assert.match(String(frames[1]?.transform), /scale\(1\.5,1\.5\)/);
    assert.match(String(frames[1]?.transform), /rotate\(45deg\)/);
  });

  test("converts blur keyframes to CSS filters", () => {
    const frames = keyframeEventToCssFrames({
      id: "blur",
      kind: "keyframes",
      targetId: "box",
      durationMs: 1000,
      tracks: [
        {
          property: "filter.blurPx",
          interpolation: "number",
          keyframes: [
            { offset: 0, value: 16 },
            { offset: 1, value: 0 }
          ]
        }
      ]
    });

    assert.equal(frames[0]?.filter, "blur(16px)");
    assert.equal(frames[1]?.filter, "none");
  });

  test("evaluates property and keyframe events for rotation and scale aliases", () => {
    const deck = createRuntimeEvalDeck();
    const states = resolveSlideObjectStates(deck, deck.deck.slides[0], 500);
    const transform = states.get("box")?.transform;

    assert.equal(transform?.rotateDeg, 45);
    assert.equal(transform?.scaleX, 1.25);
    assert.equal(transform?.scaleY, 1.25);
  });

  test("evaluates blur filter aliases in runtime state", () => {
    const deck = createFilterDeck();
    const state = resolveSlideObjectStates(deck, deck.deck.slides[0], 500).get("box");

    assert.equal(state?.filter?.blurPx, 8);
  });

  test("interpolates structured bounds, fills, strokes, crops, and alpha colors", () => {
    const deck = createInterpolationDeck();
    const state = resolveSlideObjectStates(deck, deck.deck.slides[0], 500).get("box");

    assert.deepEqual(state?.bounds, { x: 50, y: 25, width: 150, height: 75 });
    assert.deepEqual(state?.crop, { x: 0.1, y: 0.2, width: 0.65, height: 0.75, unit: "ratio" });
    assert.deepEqual(state?.style?.fill, { type: "solid", color: "#808000" });
    assert.deepEqual(state?.style?.stroke, { color: "#808080", width: 6 });
  });

  test("supports cubic-bezier, steps, and named easing during evaluation", () => {
    assert.equal(easeProgressValue({ type: "steps", count: 4, position: "end" }, 0.49), 0.25);
    assert.equal(easeProgressValue({ type: "steps", count: 4, position: "start" }, 0.01), 0.25);
    assert.ok(easeProgressValue("backOut", 0.5) > 0.5);
    assert.ok(easeProgressValue({ type: "cubicBezier", x1: 0.42, y1: 0, x2: 1, y2: 1 }, 0.5) < 0.5);
  });

  test("resolves effective child state through animated group opacity, bounds, and transform", () => {
    const deck = createGroupDeck();
    const raw = resolveSlideObjectStates(deck, deck.deck.slides[0], 500);
    const effective = resolveSlideObjectStates(deck, deck.deck.slides[0], 500, { effectiveGroupStates: true });

    assert.equal(raw.get("child")?.bounds?.x, 5);
    assert.equal(effective.get("child")?.bounds?.x, 105);
    assert.equal(effective.get("child")?.bounds?.y, 55);
    assert.equal(effective.get("child")?.opacity, 0.25);
    assert.equal(effective.get("child")?.visible, true);
    assert.equal(effective.get("child")?.transform?.translateX, 10);
    assert.equal(effective.get("child")?.transform?.scaleX, 2);
    assert.equal(effective.get("child")?.filter?.blurPx, 6);
  });

  test("resolves timing dependency graph by graph nodes and reports cycles", () => {
    const deck = createDependencyDeck();
    const slide = deck.deck.slides[0];
    const plan = createSlideTimingPlan(slide);

    assert.equal(plan.starts.intro, 100);
    assert.equal(plan.starts.follow, 450);
    assert.equal(plan.starts.before, 350);
    assert.deepEqual(plan.warnings, []);

    const cyclicPlan = createSlideTimingPlan({
      ...slide,
      timeline: {
        events: slide.timeline?.events ?? [],
        dependencyGraph: {
          nodes: [],
          edges: [
            { from: "intro", to: "follow", relation: "after" },
            { from: "follow", to: "intro", relation: "after" }
          ]
        }
      }
    });

    assert.equal(cyclicPlan.warnings[0]?.code, "cycle");
    assert.ok(cyclicPlan.order.includes("intro"));
    assert.ok(cyclicPlan.order.includes("follow"));
  });

  test("resolves morph transition states with matching and interpolated compatible properties", () => {
    const deck = createMorphTransitionDeck();
    const previousSlide = deck.deck.slides[0];
    const currentSlide = deck.deck.slides[1];
    const resolved = resolveMorphTransitionObjectStates(deck, previousSlide, currentSlide, currentSlide.transition, 0.5);
    const state = resolved.states.get("shape-b");

    assert.deepEqual(resolved.pairs, [{ fromObjectId: "shape-a", toObjectId: "shape-b", morphKey: "shared" }]);
    assert.deepEqual(state?.bounds, { x: 60, y: 70, width: 150, height: 100 });
    assert.equal(state?.transform?.rotateDeg, 45);
    assert.ok(Math.abs((state?.opacity ?? 0) - 0.6) < 0.000001);
    assert.deepEqual(state?.style?.fill, { type: "solid", color: "#800080" });
  });

  test("does not pair incompatible object types by geometry fallback", () => {
    const deck = createGeometryFallbackDeck();
    const previousSlide = deck.deck.slides[0];
    const currentSlide = deck.deck.slides[1];
    const resolved = resolveMorphTransitionObjectStates(deck, previousSlide, currentSlide, currentSlide.transition, 0.5);

    assert.deepEqual(resolved.pairs, []);
    assert.equal(resolved.states.size, 0);
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

  test("reports timeline event diagnostics with active, filled, and future phases", () => {
    const deck = createRuntimeEvalDeck();
    const diagnostics = resolveTimelineEventDiagnostics(deck.deck.slides[0], 500);
    const rotate = diagnostics.find((event) => event.eventId === "rotate");
    const scale = diagnostics.find((event) => event.eventId === "scale");

    assert.equal(rotate?.phase, "active");
    assert.equal(rotate?.active, true);
    assert.equal(rotate?.applied, true);
    assert.equal(rotate?.targetId, "box");
    assert.equal(rotate?.property, "transform.rotation");
    assert.equal(rotate?.rawProgress, 0.5);
    assert.equal(rotate?.appliedProgress, 0.5);
    assert.equal(scale?.kind, "keyframes");
    assert.equal(scale?.durationMs, 1000);
  });

  test("creates a global runtime snapshot for transition and object state debugging", () => {
    const deck = createMorphTransitionDeck();
    const snapshot = createRuntimeFrameSnapshot(deck, 1500, { includeInactiveSlides: true });

    assert.equal(snapshot.deckId, "morph-transition");
    assert.equal(snapshot.resolution.inTransition, true);
    assert.equal(snapshot.resolution.slideIndex, 1);
    assert.equal(snapshot.transition?.previousSlideId, "a");
    assert.equal(snapshot.transition?.currentSlideId, "b");
    assert.deepEqual(snapshot.transition?.pairs, [{ fromObjectId: "shape-a", toObjectId: "shape-b", morphKey: "shared" }]);
    assert.equal(snapshot.slides.length, 2);
    assert.equal(snapshot.slides[0]?.phase, "previous");
    assert.equal(snapshot.slides[1]?.phase, "current");
    assert.equal(snapshot.slides[0]?.timeMs, 1000);
    assert.equal(snapshot.objects["shape-a"]?.bounds?.x, 10);
    assert.equal(snapshot.objects["shape-b"]?.bounds?.x, 110);
    assert.equal(snapshot.transition?.states["shape-b"]?.bounds?.x, 60);
  });
});

function createRuntimeEvalDeck(): DeckIR {
  return {
    irVersion: "keymorph.ir.v1",
    deck: {
      id: "runtime-eval",
      size: { width: 100, height: 100, unit: "px" },
      slides: [
        {
          id: "slide",
          objects: [
            {
              id: "box",
              type: "shape",
              shape: "rect",
              bounds: { x: 0, y: 0, width: 20, height: 20 },
              transform: { scaleX: 1, scaleY: 1, rotateDeg: 0 }
            }
          ],
          timeline: {
            durationMs: 1000,
            events: [
              {
                id: "rotate",
                kind: "property",
                targetId: "box",
                property: "transform.rotation",
                from: 0,
                to: 90,
                durationMs: 1000,
                fill: "both"
              },
              {
                id: "scale",
                kind: "keyframes",
                targetId: "box",
                durationMs: 1000,
                fill: "both",
                tracks: [
                  {
                    property: "scale",
                    keyframes: [
                      { offset: 0, value: 1 },
                      { offset: 1, value: 1.5 }
                    ]
                  }
                ]
              }
            ]
          }
        }
      ]
    }
  };
}

function createFilterDeck(): DeckIR {
  return {
    irVersion: "keymorph.ir.v1",
    deck: {
      id: "filter-eval",
      size: { width: 100, height: 100, unit: "px" },
      slides: [
        {
          id: "slide",
          objects: [
            {
              id: "box",
              type: "shape",
              shape: "rect",
              bounds: { x: 0, y: 0, width: 20, height: 20 }
            }
          ],
          timeline: {
            durationMs: 1000,
            events: [
              {
                id: "blur",
                kind: "keyframes",
                targetId: "box",
                durationMs: 1000,
                fill: "both",
                tracks: [
                  {
                    property: "custom.keynote.filters.gaussianBlur",
                    interpolation: "number",
                    keyframes: [
                      { offset: 0, value: 16 },
                      { offset: 1, value: 0 }
                    ]
                  }
                ]
              }
            ]
          }
        }
      ]
    }
  };
}

function createInterpolationDeck(): DeckIR {
  return {
    irVersion: "keymorph.ir.v1",
    deck: {
      id: "interpolation",
      size: { width: 200, height: 200, unit: "px" },
      slides: [
        {
          id: "slide",
          objects: [
            {
              id: "box",
              type: "image",
              source: { dataUri: "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==" },
              bounds: { x: 0, y: 0, width: 100, height: 50 },
              crop: { x: 0, y: 0, width: 1, height: 1, unit: "ratio" },
              style: {
                fill: { type: "solid", color: "#ff0000" },
                stroke: { color: "#000000", width: 2 }
              }
            }
          ],
          timeline: {
            durationMs: 1000,
            events: [
              {
                id: "bounds",
                kind: "property",
                targetId: "box",
                property: "bounds",
                to: { x: 100, y: 50, width: 200, height: 100 },
                durationMs: 1000,
                fill: "both"
              },
              {
                id: "fill",
                kind: "property",
                targetId: "box",
                property: "style.fill",
                to: { type: "solid", color: "#00ff00" },
                durationMs: 1000,
                fill: "both"
              },
              {
                id: "stroke",
                kind: "property",
                targetId: "box",
                property: "style.stroke",
                to: { color: "#ffffff", width: 10 },
                durationMs: 1000,
                fill: "both"
              },
              {
                id: "crop",
                kind: "property",
                targetId: "box",
                property: "crop",
                to: { x: 0.2, y: 0.4, width: 0.3, height: 0.5, unit: "ratio" },
                durationMs: 1000,
                fill: "both"
              }
            ]
          }
        }
      ]
    }
  };
}

function createGroupDeck(): DeckIR {
  return {
    irVersion: "keymorph.ir.v1",
    deck: {
      id: "groups",
      size: { width: 200, height: 200, unit: "px" },
      slides: [
        {
          id: "slide",
          objects: [
            {
              id: "group",
              type: "group",
              bounds: { x: 20, y: 20, width: 120, height: 120 },
              opacity: 0.5,
              transform: { translateX: 0, scaleX: 1, scaleY: 1 },
              filter: { blurPx: 6 },
              children: [
                {
                  id: "child",
                  type: "shape",
                  shape: "rect",
                  bounds: { x: 5, y: 5, width: 20, height: 20 },
                  opacity: 0.5,
                  transform: { translateX: 0, scaleX: 1, scaleY: 1 }
                }
              ]
            }
          ],
          timeline: {
            durationMs: 1000,
            events: [
              {
                id: "move-group",
                kind: "property",
                targetId: "group",
                property: "bounds",
                from: { x: 20, y: 20, width: 120, height: 120 },
                to: { x: 100, y: 50, width: 120, height: 120 },
                durationMs: 500,
                fill: "forwards"
              },
              {
                id: "transform-group",
                kind: "property",
                targetId: "group",
                property: "transform",
                from: { translateX: 0, scaleX: 1, scaleY: 1 },
                to: { translateX: 10, scaleX: 2, scaleY: 2 },
                durationMs: 500,
                fill: "forwards"
              }
            ]
          }
        }
      ]
    }
  };
}

function createDependencyDeck(): DeckIR {
  return {
    irVersion: "keymorph.ir.v1",
    deck: {
      id: "dependency",
      size: { width: 100, height: 100, unit: "px" },
      slides: [
        {
          id: "slide",
          objects: [
            {
              id: "box",
              type: "shape",
              shape: "rect",
              bounds: { x: 0, y: 0, width: 10, height: 10 }
            }
          ],
          timeline: {
            durationMs: 1000,
            events: [
              { id: "intro", kind: "visibility", targetId: "box", visible: true, start: { type: "absolute", atMs: 100 }, durationMs: 200 },
              { id: "follow", kind: "visibility", targetId: "box", visible: false, durationMs: 100 },
              { id: "before", kind: "visibility", targetId: "box", visible: true, durationMs: 50 }
            ],
            dependencyGraph: {
              nodes: [
                { id: "n-intro", eventId: "intro" },
                { id: "n-follow", eventId: "follow" },
                { id: "n-before", eventId: "before" }
              ],
              edges: [
                { from: "n-intro", to: "n-follow", relation: "after", offsetMs: 150 },
                { from: "n-follow", to: "n-before", relation: "before", offsetMs: 50 }
              ]
            }
          }
        }
      ]
    }
  };
}

function createMorphTransitionDeck(): DeckIR {
  return {
    irVersion: "keymorph.ir.v1",
    deck: {
      id: "morph-transition",
      size: { width: 300, height: 200, unit: "px" },
      slides: [
        {
          id: "a",
          objects: [
            {
              id: "shape-a",
              type: "shape",
              shape: "rect",
              morphKey: "shared",
              bounds: { x: 10, y: 20, width: 100, height: 80 },
              opacity: 0.2,
              transform: { rotateDeg: 0 },
              style: { fill: { type: "solid", color: "#ff0000" } }
            }
          ],
          timeline: { durationMs: 1000, events: [] }
        },
        {
          id: "b",
          objects: [
            {
              id: "shape-b",
              type: "shape",
              shape: "rect",
              morphKey: "shared",
              bounds: { x: 110, y: 120, width: 200, height: 120 },
              opacity: 1,
              transform: { rotateDeg: 90 },
              style: { fill: { type: "solid", color: "#0000ff" } }
            }
          ],
          transition: {
            type: "morph",
            durationMs: 1000,
            morph: {
              strategy: "morph",
              matching: { matchBy: ["morphKey"] },
              properties: ["bounds", "transform", "opacity", "fill"]
            }
          },
          timeline: { durationMs: 1000, events: [] }
        }
      ]
    }
  };
}

function createGeometryFallbackDeck(): DeckIR {
  return {
    irVersion: "keymorph.ir.v1",
    deck: {
      id: "geometry-fallback",
      size: { width: 300, height: 200, unit: "px" },
      slides: [
        {
          id: "a",
          objects: [
            {
              id: "title",
              type: "text",
              bounds: { x: 10, y: 20, width: 100, height: 80 },
              text: { plainText: "Title", runs: [{ text: "Title", style: { fontSize: 24 } }] }
            }
          ],
          timeline: { durationMs: 1000, events: [] }
        },
        {
          id: "b",
          objects: [
            {
              id: "placeholder",
              type: "placeholder",
              placeholderType: "custom",
              bounds: { x: 10, y: 20, width: 100, height: 80 }
            }
          ],
          transition: {
            type: "morph",
            durationMs: 1000,
            morph: {
              strategy: "morph",
              matching: { matchBy: ["geometry"], fallback: "geometry", tolerance: 0.1 },
              properties: ["bounds", "opacity"]
            }
          },
          timeline: { durationMs: 1000, events: [] }
        }
      ]
    }
  };
}

function createCharacterBuildDeck(fallback = "dissolve-in"): DeckIR {
  return {
    irVersion: "keymorph.ir.v1",
    deck: {
      id: "character-build",
      size: { width: 400, height: 200, unit: "px" },
      slides: [
        {
          id: "slide",
          objects: [
            {
              id: "title",
              type: "text",
              bounds: { x: 0, y: 0, width: 300, height: 80 },
              text: {
                plainText: "逐字溶解",
                runs: [{ text: "逐字溶解", style: { fontSize: 32 } }]
              }
            }
          ],
          timeline: {
            durationMs: 1000,
            events: [
              {
                id: "character-dissolve",
                kind: "keyframes",
                targetId: "title",
                durationMs: 1000,
                fill: "both",
                tracks: [
                  {
                    property: "opacity",
                    interpolation: "number",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 1 }
                    ]
                  }
                ],
                metadata: {
                  nativeBuildGranularity: "character",
                  nativeBuildFallback: fallback,
                  nativeBuildDirection: "In"
                }
              }
            ]
          }
        }
      ]
    }
  };
}

function createShapeTextDeck(): DeckIR {
  return {
    irVersion: "keymorph.ir.v1",
    deck: {
      id: "shape-text",
      size: { width: 400, height: 200, unit: "px" },
      slides: [
        {
          id: "slide",
          objects: [
            {
              id: "shape-title",
              type: "shape",
              shape: "rect",
              bounds: { x: 10, y: 20, width: 260, height: 90 },
              text: {
                plainText: "First\nSecond",
                runs: [
                  {
                    text: "First\nSecond",
                    style: { fontSize: 22, lineHeight: 1.35, color: "#123456" }
                  }
                ]
              }
            }
          ],
          timeline: { durationMs: 1000, events: [] }
        }
      ]
    }
  };
}

function pickResolution(resolution: ReturnType<typeof resolveDeckTime>) {
  return {
    slideIndex: resolution.slideIndex,
    slideTimeMs: resolution.slideTimeMs,
    inTransition: resolution.inTransition,
    transitionProgress: resolution.transitionProgress
  };
}
