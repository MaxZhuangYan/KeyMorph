import { describe, test } from "node:test";
import assert from "node:assert/strict";

import { createLossReport, scoreConversion } from "../../src/report/index.ts";

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
    assert.equal(score.riskLevel, "medium");
    assert.ok(score.penaltyBreakdown.length >= 3);
    assert.ok(score.fidelityScore < 1);
  });

  test("creates a structured loss report", () => {
    const loss = createLossReport({
      status: "partial",
      messages: [],
      unsupportedFeatures: [
        {
          code: "pptx-anim",
          severity: "warning",
          area: "animation",
          description: "Unsupported timeline node",
          fallback: "morph"
        },
        {
          code: "slide-transition",
          severity: "warning",
          area: "transition",
          description: "Unsupported transition",
          fallback: "video"
        }
      ],
      degradedFeatures: [
        {
          code: "fade-easing",
          severity: "warning",
          area: "animation",
          description: "Approximated easing",
          fallback: "linear"
        }
      ]
    });

    assert.equal(loss.animationLostCount, 1);
    assert.equal(loss.unsupportedTransitions.length, 1);
    assert.equal(loss.unsupportedFeatures.length, 2);
    assert.equal(loss.degradedFeatures.length, 1);
    assert.match(loss.degradedAnimations[0], /fade-easing/);
  });
});
