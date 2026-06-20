import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { scoreConversion } from "../../src/report/index.ts";

describe("conversion report scoring", () => {
  test("scores unsupported, degraded, and uncertain mappings", () => {
    const score = scoreConversion({
      status: "partial",
      messages: [],
      unsupportedFeatures: [
        {
          code: "anim",
          severity: "warning",
          area: "animation",
          description: "Unsupported animation",
          fallback: "morph"
        }
      ],
      degradedFeatures: [
        {
          code: "timing",
          severity: "warning",
          area: "animation",
          description: "Timing degraded",
          fallback: "split slides"
        }
      ],
      uncertainMappings: [
        {
          code: "shape-map",
          severity: "warning",
          description: "Shape mapping is approximate",
          confidence: 0.7
        }
      ]
    });

    assert.equal(score.animationLostCount, 1);
    assert.equal(score.degradedAnimationCount, 1);
    assert.equal(score.uncertainMappingCount, 1);
    assert.ok(score.fidelityScore < 1);
  });
});
