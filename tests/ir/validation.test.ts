import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { IR_VERSION, type KeyMorphIR, validateIR } from "../../src/ir/index.ts";

const validIR: KeyMorphIR = {
  irVersion: IR_VERSION,
  deck: {
    id: "deck-1",
    size: {
      width: 1920,
      height: 1080,
      unit: "px",
    },
    assets: [
      {
        id: "asset-logo",
        kind: "image",
        uri: "https://example.test/logo.png",
        mimeType: "image/png",
      },
    ],
    slides: [
      {
        id: "slide-1",
        objects: [
          {
            id: "title",
            type: "text",
            bounds: {
              x: 100,
              y: 120,
              width: 800,
              height: 120,
            },
            text: {
              plainText: "KeyMorph",
            },
          },
          {
            id: "logo-group",
            type: "group",
            children: [
              {
                id: "logo",
                type: "image",
                morphKey: "brand-logo",
                source: {
                  assetId: "asset-logo",
                },
              },
            ],
          },
        ],
        states: [
          {
            id: "title-exit",
            targetId: "title",
            properties: {
              opacity: 0,
              transform: {
                translateY: -40,
              },
            },
          },
        ],
        transition: {
          type: "magicMove",
          durationMs: 650,
          morph: {
            strategy: "magicMove",
            properties: ["bounds", "transform", "opacity"],
          },
        },
        timeline: {
          durationMs: 1200,
          defaultEasing: "easeInOut",
          events: [
            {
              id: "fade-title",
              kind: "keyframes",
              targetId: "title",
              durationMs: 400,
              tracks: [
                {
                  property: "opacity",
                  keyframes: [
                    {
                      offset: 0,
                      value: 1,
                    },
                    {
                      offset: 1,
                      value: 0,
                    },
                  ],
                },
              ],
            },
            {
              id: "apply-title-state",
              kind: "setState",
              targetId: "title",
              stateId: "title-exit",
              start: {
                type: "after",
                eventId: "fade-title",
              },
            },
            {
              id: "move-logo",
              kind: "morph",
              strategy: "magicMove",
              durationMs: 600,
              from: {
                objectId: "logo",
              },
              to: {
                snapshot: {
                  bounds: {
                    x: 1500,
                    y: 80,
                    width: 180,
                    height: 180,
                  },
                },
              },
              pairs: [
                {
                  fromObjectId: "logo",
                  toObjectId: "logo",
                  morphKey: "brand-logo",
                },
              ],
              dependencies: [
                {
                  eventId: "fade-title",
                  relation: "with",
                },
              ],
            },
          ],
          dependencyGraph: {
            nodes: [
              {
                id: "fade-title",
                kind: "event",
              },
              {
                id: "move-logo",
                kind: "event",
              },
            ],
            edges: [
              {
                from: "fade-title",
                to: "move-logo",
                relation: "with",
              },
            ],
          },
        },
      },
    ],
  },
  conversion: {
    status: "partial",
    messages: [
      {
        severity: "warning",
        code: "UNSUPPORTED_EFFECT",
        message: "A source shadow effect was approximated.",
        path: "$.deck.slides[0].objects[0]",
      },
    ],
  },
};

describe("validateIR", () => {
  test("accepts a representative IR v1 document", () => {
    const result = validateIR(validIR);

    assert.equal(result.valid, true);
    assert.equal(result.errors.length, 0);
    assert.equal(result.value, validIR);
  });

  test("returns structured errors instead of throwing", () => {
    const result = validateIR({
      irVersion: "wrong",
      deck: {
        id: "deck-1",
        size: {
          width: 1920,
          height: "1080",
          unit: "px",
        },
        slides: [
          {
            id: "slide-1",
            objects: [
              {
                id: "title",
                type: "text",
              },
            ],
            timeline: {
              events: [
                {
                  id: "bad-event",
                  kind: "keyframes",
                  targetId: "missing-object",
                  tracks: [
                    {
                      property: "opacity",
                      keyframes: [
                        {
                          offset: 1,
                          value: 0,
                        },
                        {
                          offset: 0,
                          value: 1,
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    });

    const codes = result.errors.map((error) => error.code);

    assert.equal(result.valid, false);
    assert.equal(result.value, undefined);
    assert.ok(codes.includes("INVALID_IR_VERSION"));
    assert.ok(codes.includes("NUMBER_REQUIRED"));
    assert.ok(codes.includes("TEXT_REQUIRED"));
    assert.ok(codes.includes("UNKNOWN_TARGET_ID"));
    assert.ok(codes.includes("KEYFRAME_OFFSETS_UNSORTED"));
    assert.equal(result.errors.every((error) => error.path.startsWith("$.") || error.path === "$"), true);
  });

  test("detects timing dependency cycles", () => {
    const cyclicIR: KeyMorphIR = {
      ...validIR,
      deck: {
        ...validIR.deck,
        slides: [
          {
            ...validIR.deck.slides[0],
            timeline: {
              events: [
                {
                  id: "a",
                  kind: "visibility",
                  targetId: "title",
                  visible: true,
                  dependencies: [
                    {
                      eventId: "b",
                    },
                  ],
                },
                {
                  id: "b",
                  kind: "visibility",
                  targetId: "title",
                  visible: false,
                  dependencies: [
                    {
                      eventId: "a",
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    };

    const result = validateIR(cyclicIR);

    assert.equal(result.valid, false);
    assert.equal(result.errors.some((error) => error.code === "TIMING_DEPENDENCY_CYCLE"), true);
  });
});
