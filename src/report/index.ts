import type { ConversionReport } from "../ir/index.ts";
export * from "./fidelity.ts";

export interface ConversionScore {
  animationLostCount: number;
  degradedAnimationCount: number;
  uncertainMappingCount: number;
  fidelityScore: number;
  riskLevel: "low" | "medium" | "high";
  penaltyBreakdown: ConversionPenalty[];
  recommendedFixes: string[];
  report: ConversionReport;
}

export interface ConversionLossReport extends ConversionScore {
  unsupportedFeatures: ConversionFeatureSummary[];
  degradedFeatures: ConversionFeatureSummary[];
  uncertainMappings: ConversionFeatureSummary[];
  degradedAnimations: string[];
  unsupportedAnimations: string[];
  unsupportedTransitions: string[];
  generatedAt: string;
}

export interface ConversionPenalty {
  code: string;
  area: string;
  severity: string;
  penalty: number;
  reason: string;
}

export interface ConversionFeatureSummary {
  code: string;
  area: string;
  severity: string;
  description: string;
  fallback?: string;
  confidence?: number;
}

export function scoreConversion(report: ConversionReport): ConversionScore {
  const unsupported = report.unsupportedFeatures ?? [];
  const degraded = report.degradedFeatures ?? [];
  const uncertain = report.uncertainMappings ?? [];
  const animationLostCount = unsupported.filter((item) => item.area === "animation").length;
  const degradedAnimationCount = degraded.filter((item) => item.area === "animation").length;
  const uncertainMappingCount = uncertain.length;

  const penaltyBreakdown: ConversionPenalty[] = [
    ...unsupported.map((item) => ({
      code: item.code,
      area: item.area ?? "unknown",
      severity: item.severity,
      penalty: featurePenalty(item.area ?? "unknown", item.severity, "unsupported"),
      reason: item.description
    })),
    ...degraded.map((item) => ({
      code: item.code,
      area: item.area ?? "unknown",
      severity: item.severity,
      penalty: featurePenalty(item.area ?? "unknown", item.severity, "degraded"),
      reason: item.description
    })),
    ...uncertain.map((item) => ({
      code: item.code,
      area: "mapping",
      severity: item.severity,
      penalty: uncertaintyPenalty(item.confidence),
      reason: item.description
    }))
  ];

  const penalty = penaltyBreakdown.reduce((total, item) => total + item.penalty, 0);
  const fidelityScore = Math.max(0, Math.min(1, Number((1 - penalty).toFixed(3))));
  const riskLevel = fidelityScore >= 0.9 ? "low" : fidelityScore >= 0.72 ? "medium" : "high";

  const recommendedFixes = new Set<string>();
  if (animationLostCount > 0) recommendedFixes.add("Use morph replacement for unsupported animations.");
  if (degradedAnimationCount > 0) recommendedFixes.add("Split dense timelines into additional slides.");
  if (unsupported.some((item) => item.area === "media")) {
    recommendedFixes.add("Use HTML runtime or video fallback for media-heavy slides.");
  }
  if (uncertainMappingCount > 0) {
    recommendedFixes.add("Review uncertain object mappings before final export.");
  }
  if (unsupported.some((item) => item.area === "transition")) {
    recommendedFixes.add("Use video fallback for transitions that cannot be expressed in PPTX.");
  }
  if (recommendedFixes.size === 0) recommendedFixes.add("No manual fixes recommended for this report.");

  return {
    animationLostCount,
    degradedAnimationCount,
    uncertainMappingCount,
    fidelityScore,
    riskLevel,
    penaltyBreakdown,
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
    unsupportedFeatures: unsupported.map((item) => ({
      code: item.code,
      area: item.area ?? "unknown",
      severity: item.severity,
      description: item.description,
      fallback: item.fallback
    })),
    degradedFeatures: degraded.map((item) => ({
      code: item.code,
      area: item.area ?? "unknown",
      severity: item.severity,
      description: item.description,
      fallback: item.fallback
    })),
    uncertainMappings: (report.uncertainMappings ?? []).map((item) => ({
      code: item.code,
      area: "mapping",
      severity: item.severity,
      description: item.description,
      confidence: item.confidence
    })),
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

function featurePenalty(area: string, severity: string, kind: "unsupported" | "degraded"): number {
  const areaWeight: Record<string, number> = {
    animation: 0.12,
    transition: 0.1,
    media: 0.09,
    layout: 0.08,
    text: 0.05,
    image: 0.05,
    shape: 0.04,
    asset: 0.04,
    unknown: 0.04
  };
  const severityMultiplier: Record<string, number> = {
    error: 1.3,
    warning: 1,
    info: 0.45
  };
  const kindMultiplier = kind === "unsupported" ? 1 : 0.58;
  return Number(((areaWeight[area] ?? areaWeight.unknown) * (severityMultiplier[severity] ?? 1) * kindMultiplier).toFixed(4));
}

function uncertaintyPenalty(confidence: number | undefined): number {
  const normalized = typeof confidence === "number" ? Math.max(0, Math.min(1, confidence)) : 0.65;
  return Number(((1 - normalized) * 0.08).toFixed(4));
}
