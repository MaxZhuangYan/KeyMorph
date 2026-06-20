import { mkdtemp, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import { exportIrToPptx, parsePptxToIr } from "../../src/pptx/index.ts";
import { IR_VERSION, type DeckIR } from "../../src/ir/index.ts";

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

  test("round-trips simple keyframe and visibility timing through PPTX XML", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-pptx-test-"));
    const out = path.join(dir, "animated.pptx");

    await exportIrToPptx(createAnimatedDeck(), out);
    const parsed = await parsePptxToIr(out);
    const events = parsed.deck.slides[0].timeline?.events ?? [];

    assert.equal(parsed.conversion?.statistics?.animationCount, 5);
    assert.equal(parsed.conversion?.statistics?.unsupportedFeatureCount, 0);
    assert.ok(events.some((event) => event.kind === "visibility" && event.visible === false && event.start?.type === "absolute" && event.start.atMs === 1100));

    const fade = events.find((event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "opacity"));
    assert.equal(fade?.kind, "keyframes");
    if (fade?.kind !== "keyframes") throw new Error("Expected imported fade keyframes.");
    assert.equal(fade.durationMs, 700);
    assert.deepEqual(fade.tracks.find((track) => track.property === "opacity")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 1, value: 1 }
    ]);

    const motion = events.find(
      (event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "transform.translateX")
    );
    assert.equal(motion?.kind, "keyframes");
    if (motion?.kind !== "keyframes") throw new Error("Expected imported motion keyframes.");
    assert.equal(motion.durationMs, 600);
    assert.deepEqual(motion.tracks.find((track) => track.property === "transform.translateX")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 1, value: 120 }
    ]);
    assert.deepEqual(motion.tracks.find((track) => track.property === "transform.translateY")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 1, value: 45 }
    ]);

    const scale = events.find(
      (event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "transform.scaleX")
    );
    assert.equal(scale?.kind, "keyframes");
    if (scale?.kind !== "keyframes") throw new Error("Expected imported scale keyframes.");
    assert.deepEqual(scale.tracks.find((track) => track.property === "transform.scaleX")?.keyframes, [
      { offset: 0, value: 1 },
      { offset: 1, value: 1.25 }
    ]);
    assert.deepEqual(scale.tracks.find((track) => track.property === "transform.scaleY")?.keyframes, [
      { offset: 0, value: 1 },
      { offset: 1, value: 0.75 }
    ]);

    const rotation = events.find(
      (event) => event.kind === "keyframes" && event.tracks.some((track) => track.property === "transform.rotateDeg")
    );
    assert.equal(rotation?.kind, "keyframes");
    if (rotation?.kind !== "keyframes") throw new Error("Expected imported rotation keyframes.");
    assert.deepEqual(rotation.tracks.find((track) => track.property === "transform.rotateDeg")?.keyframes, [
      { offset: 0, value: 0 },
      { offset: 1, value: 45 }
    ]);
  });
});

function createAnimatedDeck(): DeckIR {
  return {
    irVersion: IR_VERSION,
    metadata: { title: "Animated PPTX fixture", sourceApplication: "KeyMorph tests" },
    deck: {
      id: "animated-fixture",
      title: "Animated PPTX fixture",
      size: { width: 1280, height: 720, unit: "px" },
      slides: [
        {
          id: "slide-1",
          index: 0,
          name: "Animations",
          background: { type: "solid", color: "#ffffff" },
          objects: [
            {
              id: "headline",
              type: "text",
              name: "Headline",
              bounds: { x: 96, y: 96, width: 500, height: 96 },
              opacity: 1,
              text: {
                plainText: "Animated headline",
                runs: [{ text: "Animated headline", style: { fontFamily: "Arial", fontSize: 36, color: "#111827" } }]
              }
            },
            {
              id: "badge",
              type: "shape",
              name: "Badge",
              shape: "roundRect",
              bounds: { x: 120, y: 260, width: 180, height: 72 },
              opacity: 1,
              style: {
                fill: { type: "solid", color: "#0f766e" },
                stroke: { color: "#0f766e", width: 0 }
              }
            }
          ],
          timeline: {
            durationMs: 1800,
            events: [
              {
                id: "headline-fade",
                kind: "keyframes",
                targetId: "headline",
                start: { type: "absolute", atMs: 250 },
                durationMs: 700,
                fill: "both",
                tracks: [
                  {
                    property: "opacity",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 1 }
                    ]
                  }
                ]
              },
              {
                id: "headline-hide",
                kind: "visibility",
                targetId: "headline",
                start: { type: "absolute", atMs: 1100 },
                durationMs: 0,
                fill: "both",
                visible: false
              },
              {
                id: "badge-slide",
                kind: "keyframes",
                targetId: "badge",
                start: { type: "absolute", atMs: 500 },
                durationMs: 600,
                fill: "both",
                tracks: [
                  {
                    property: "transform.translateX",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 120 }
                    ]
                  },
                  {
                    property: "transform.translateY",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 45 }
                    ]
                  }
                ]
              },
              {
                id: "badge-scale",
                kind: "keyframes",
                targetId: "badge",
                start: { type: "absolute", atMs: 950 },
                durationMs: 400,
                fill: "both",
                tracks: [
                  {
                    property: "transform.scaleX",
                    keyframes: [
                      { offset: 0, value: 1 },
                      { offset: 1, value: 1.25 }
                    ]
                  },
                  {
                    property: "transform.scaleY",
                    keyframes: [
                      { offset: 0, value: 1 },
                      { offset: 1, value: 0.75 }
                    ]
                  }
                ]
              },
              {
                id: "badge-rotate",
                kind: "keyframes",
                targetId: "badge",
                start: { type: "absolute", atMs: 1250 },
                durationMs: 300,
                fill: "both",
                tracks: [
                  {
                    property: "transform.rotateDeg",
                    keyframes: [
                      { offset: 0, value: 0 },
                      { offset: 1, value: 45 }
                    ]
                  }
                ]
              }
            ],
            dependencyGraph: { edges: [] }
          }
        }
      ]
    },
    conversion: { status: "success", messages: [] }
  };
}
