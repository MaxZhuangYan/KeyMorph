import type { ConversionReport } from "../ir/index.ts";

export interface ConversionScore {
  animationLostCount: number;
  degradedAnimationCount: number;
  uncertainMappingCount: number;
  fidelityScore: number;
  recommendedFixes: string[];
  report: ConversionReport;
}

export interface ConversionLossReport extends ConversionScore {
  degradedAnimations: string[];
  unsupportedAnimations: string[];
  unsupportedTransitions: string[];
  generatedAt: string;
}

export function scoreConversion(report: ConversionReport): ConversionScore {
  const unsupported = report.unsupportedFeatures ?? [];
  const degraded = report.degradedFeatures ?? [];
  const uncertain = report.uncertainMappings ?? [];
  const animationLostCount = unsupported.filter((item) => item.area === "animation").length;
  const degradedAnimationCount = degraded.filter((item) => item.area === "animation").length;
  const uncertainMappingCount = uncertain.length;

  const penalty =
    animationLostCount * 0.12 + degradedAnimationCount * 0.07 + uncertainMappingCount * 0.03;
  const fidelityScore = Math.max(0, Math.min(1, Number((1 - penalty).toFixed(3))));

  const recommendedFixes = new Set<string>();
  if (animationLostCount > 0) recommendedFixes.add("Use morph replacement for unsupported animations.");
  if (degradedAnimationCount > 0) recommendedFixes.add("Split dense timelines into additional slides.");
  if (unsupported.some((item) => item.area === "transition")) {
    recommendedFixes.add("Use video fallback for transitions that cannot be expressed in PPTX.");
  }
  if (recommendedFixes.size === 0) recommendedFixes.add("No manual fixes recommended for this report.");

  return {
    animationLostCount,
    degradedAnimationCount,
    uncertainMappingCount,
    fidelityScore,
    recommendedFixes: Array.from(recommendedFixes),
    report
  };
}

export function createLossReport(report: ConversionReport): ConversionLossReport {
  const score = scoreConversion(report);
  const unsupported = report.unsupportedFeatures ?? [];
  const degraded = report.degradedFeatures ?? [];

  return {
    ...score,
    degradedAnimations: degraded
      .filter((item) => item.area === "animation")
      .map((item) => `${item.code}: ${item.description}`),
    unsupportedAnimations: unsupported
      .filter((item) => item.area === "animation")
      .map((item) => `${item.code}: ${item.description}`),
    unsupportedTransitions: unsupported
      .filter((item) => item.area === "transition")
      .map((item) => `${item.code}: ${item.description}`),
    generatedAt: new Date().toISOString()
  };
}
