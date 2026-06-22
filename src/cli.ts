import { copyFile, cp, mkdir, readFile, rename, rm, stat, writeFile } from "node:fs/promises";
import { spawn } from "node:child_process";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { createDemoDeck } from "./demo/createDemoDeck.ts";
import { IR_VERSION, validateIR, type DeckIR, type Slide } from "./ir/index.ts";
import { renderHtmlDocument } from "./runtime/index.ts";
import { createLossReport, scoreConversion, type ConversionLossReport } from "./report/index.ts";
import { comparePngFiles } from "./report/fidelity.ts";
import { exportIrToPptx, parsePptxToIr } from "./pptx/index.ts";
import {
  exportIrToKeynote,
  exportKeynoteToHtml,
  exportKeynoteToMovie,
  materializeNativeKeynoteAssetFiles,
  parseKeynoteHtmlExportToIr,
  parseKeynoteToIr,
  type KeynoteAutomationOptions
} from "./keynote/index.ts";
import {
  captureVideoFrames,
  createVideoExportPlan,
  createSegmentPlan,
  createVideoFrameFidelityReport,
  describeVideoDependencies,
  exportIrToVideo,
  extractVideoFramesFromVideo,
  readMovieDurationMs,
  scaleSegmentPlanToDuration,
  splitVideoIntoSegments,
  type VideoDependencyStatus,
  type VideoExportOptions,
  type VideoFrameFidelityReport,
  type VideoExportPlan,
  type SplitVideoSegmentsOptions,
  type SplitVideoSegmentsResult,
  type VideoSegmentPlanEntry,
  VideoExportDependencyError
} from "./video/index.ts";

export type ProductInputKind = "pptx" | "keynote" | "ir";

export interface ProductBundleOptions {
  sourceName?: string;
  jobId?: string;
  allowKeynoteAutomation?: boolean;
  keynoteAutomationTimeoutMs?: number;
  keynoteBridgeExport?: (keynotePath: string, pptxPath: string, options: KeynoteAutomationOptions) => Promise<void>;
  keynoteHtmlExport?: (keynotePath: string, outputDir: string, options: KeynoteAutomationOptions) => Promise<void>;
  keynoteMovieExport?: (keynotePath: string, moviePath: string, options: KeynoteAutomationOptions) => Promise<void>;
  segmentVideoSplit?: (
    inputMoviePath: string,
    segments: VideoSegmentPlanEntry[],
    outputDir: string,
    options: SplitVideoSegmentsOptions
  ) => Promise<SplitVideoSegmentsResult>;
  segmentPosterCreate?: (videoPath: string, durationMs: number, outputDir: string, id: string) => Promise<string | string[] | undefined>;
  video?: Pick<VideoExportOptions, "fps" | "scale" | "ffmpegPath">;
}

export interface ProductKeyBenchmarkOptions extends ProductBundleOptions {
  outputDir?: string;
  runBaseline?: boolean;
  baseline?: ProductBaselineExportOptions;
}

export interface ProductKeyBenchmarkResult {
  inputPath: string;
  copiedSourcePath: string;
  bundleDir: string;
  summaryPath: string;
  bundle: ProductBundleResult;
  baseline?: ProductBaselineExportResult;
  summary: ProductKeyBenchmarkSummary;
}

export interface ProductKeyBenchmarkSummary {
  generatedAt: string;
  inputPath: string;
  copiedSourcePath: string;
  bundleDir: string;
  sourceName: string;
  slideCount: number;
  objectCount: number;
  runtime: ProductRuntimeSummary;
  conversion: ProductBundleManifest["report"];
  videoPlan: Omit<VideoExportPlan, "frames">;
  videoDependencies: VideoDependencyStatus;
  baseline: {
    status: ProductBaselineExportResult["status"] | "not-run";
    message: string;
    reportPath: string | null;
    diffPath: string | null;
    referenceFramesPath: string | null;
    actualFramesPath: string | null;
    summary?: ProductFrameFidelitySummary;
    worstFrames: ProductBenchmarkFrameSummary[];
    worstSlides: ProductBenchmarkSlideSummary[];
  };
  nextRecommendedActions: string[];
}

export interface ProductBenchmarkFrameSummary {
  frame: number;
  slideIndex: number;
  slideId: string;
  timeMs: number;
  pixelFidelityScore: number;
  mismatchRatio: number;
  referencePath: string;
  actualPath: string;
  diffPath?: string;
}

export interface ProductBenchmarkSlideSummary {
  slideIndex: number;
  slideId: string;
  frameCount: number;
  meanPixelFidelityScore: number;
  mismatchRatio: number;
  worstFrame?: number;
}

export interface ProductFrameFidelityInsights {
  worstFrames: ProductBenchmarkFrameSummary[];
  worstSlides: ProductBenchmarkSlideSummary[];
}

export interface ProductBaselineExportOptions {
  allowKeynoteAutomation?: boolean;
  keynoteAutomationTimeoutMs?: number;
  keynoteMovieExport?: (keynotePath: string, moviePath: string, options: KeynoteAutomationOptions) => Promise<void>;
  keynoteFrameExport?: (
    keynotePath: string,
    framesDir: string,
    plan: VideoExportPlan,
    options: KeynoteAutomationOptions & VideoExportOptions
  ) => Promise<void>;
  runtimeFrameExport?: (deck: DeckIR, framesDir: string, plan: VideoExportPlan, options: VideoExportOptions) => Promise<void>;
  video?: Pick<VideoExportOptions, "fps" | "scale" | "ffmpegPath">;
}

export interface ProductBundlePaths {
  jobDir: string;
  source: string;
  deckIr: string;
  runtimeHtml: string;
  irRuntimeHtml: string;
  keynoteHtmlDir: string;
  keynoteHtmlIndex: string;
  keynoteMovie: string;
  segmentsDir: string;
  segmentPlan: string;
  segmentPostersDir: string;
  nativeAssetsDir: string;
  rebuiltPptx: string;
  segmentedPptx: string;
  staticStepsPptx: string;
  hybridPptx: string;
  lossReport: string;
  manifest: string;
  videoPlan: string;
  videoStatus: string;
  rebuiltKeynote: string;
  renderVideo: string;
  framesLatest: string;
  frameFidelity: string;
  frameDiffs: string;
  baselineStatus: string;
  baselineMovie: string;
  baselineFrames: string;
  baselineActualFrames: string;
  baselineFidelity: string;
  baselineDiffs: string;
}

export interface ProductBundleManifest {
  jobId: string;
  sourceName: string;
  sourceKind: ProductInputKind;
  createdAt: string;
  slideCount: number;
  objectCount: number;
  artifacts: {
    source: string;
    deckIr: string;
    runtimeHtml: string;
    irRuntimeHtml: string | null;
    keynoteHtml: string | null;
    keynoteMovie: string | null;
    segmentPlan: string | null;
    segmentedPptx: string | null;
    staticStepsPptx: string | null;
    hybridPptx: string | null;
    nativeAssets: string | null;
    rebuiltPptx: string;
    lossReport: string;
    videoPlan: string;
    videoStatus: string;
    rebuiltKeynote: string | null;
    renderVideo: string | null;
    frameFidelity: string | null;
    baselineStatus: string;
    keynoteBaselineMovie: string | null;
    baselineFrames: string | null;
    baselineActualFrames: string | null;
    baselineFrameFidelity: string | null;
    baselineFrameDiffs: string | null;
  };
  report: {
    fidelityScore: number;
    riskLevel: ConversionLossReport["riskLevel"];
    animationLostCount: number;
    degradedAnimationCount: number;
    uncertainMappingCount: number;
    recommendedFixes: string[];
  };
  runtime: ProductRuntimeSummary;
  video: ProductVideoSummary;
  baseline: ProductBaselineSummary;
  keynote: ProductDeferredExportSummary;
}

export interface ProductRuntimeSummary {
  mode: "keymorph-ir" | "keynote-html" | "keynote-movie";
  fidelity: "ir-reconstructed" | "keynote-native" | "keynote-rendered-video";
  message: string;
  htmlPath: string;
  irHtmlPath: string | null;
  moviePath: string | null;
  keynoteHtmlPath: string | null;
}

export interface ProductVideoSummary {
  plan: Omit<VideoExportPlan, "frames">;
  dependencies: VideoDependencyStatus;
  endpoint: string | null;
  statusPath: string;
}

export interface ProductBaselineSummary {
  available: boolean | null;
  endpoint: string | null;
  message: string;
  statusPath: string;
  referenceMoviePath: string | null;
  referenceFramesPath: string | null;
  actualFramesPath: string | null;
  fidelityReportPath: string | null;
  diffPath: string | null;
  summary?: ProductFrameFidelitySummary;
}

export interface ProductDeferredExportSummary {
  available: boolean | null;
  endpoint: string | null;
  message: string;
  outputPath: string;
}

export interface ProductBundleResult {
  jobId: string;
  sourceName: string;
  sourceKind: ProductInputKind;
  deck: DeckIR;
  lossReport: ConversionLossReport;
  videoPlan: VideoExportPlan;
  videoDependencies: VideoDependencyStatus;
  manifest: ProductBundleManifest;
  paths: ProductBundlePaths;
}

export interface ProductVideoExportResult {
  status: "ready";
  outputPath: string;
  plan: Omit<VideoExportPlan, "frames">;
  framesDir: string;
  frameFidelity: {
    reportPath: string;
    diffDir: string;
    summary: ProductFrameFidelitySummary;
  } | null;
}

export type ProductFrameFidelitySummary = VideoFrameFidelityReport["summary"];

export interface ProductBaselineExportReady {
  status: "ready";
  source: "keynote-movie" | "keynote-reference-frames";
  referenceMoviePath: string | null;
  referenceFramesDir: string;
  actualFramesDir: string;
  reportPath: string;
  diffDir: string;
  plan: Omit<VideoExportPlan, "frames">;
  summary: ProductFrameFidelitySummary;
}

export interface ProductBaselineExportUnavailable {
  status: "unsupported" | "dependency-missing";
  message: string;
  missing?: string[];
  guidance?: string[];
}

export type ProductBaselineExportResult = ProductBaselineExportReady | ProductBaselineExportUnavailable;

export interface ProductDeferredExportUnavailable {
  status: "dependency-missing";
  message: string;
}

export interface ProductKeynoteExportReady {
  status: "ready";
  outputPath: string;
  message: string;
}

type ProductKeynoteExportResult = ProductKeynoteExportReady | ProductDeferredExportUnavailable;

export async function createProductBundle(inputPath: string, outputDir: string, options: ProductBundleOptions = {}): Promise<ProductBundleResult> {
  const resolvedInput = path.resolve(inputPath);
  const sourceName = sanitizeFileName(options.sourceName ?? path.basename(resolvedInput));
  const sourceKind = detectInputKind(sourceName);
  const jobId = options.jobId ?? createJobId(sourceName);
  const jobDir = path.resolve(outputDir);
  const paths = createBundlePaths(jobDir, sourceName);
  const createdAt = new Date().toISOString();

  await mkdir(jobDir, { recursive: true });
  await copySource(resolvedInput, paths.source);

  let deck = await inputToIr(paths.source, sourceName, sourceKind, {
    allowKeynoteAutomation: options.allowKeynoteAutomation,
    keynoteAutomationTimeoutMs: options.keynoteAutomationTimeoutMs,
    keynoteBridgeExport: options.keynoteBridgeExport
  });
  let validation = validateIR(deck);
  if (!validation.valid && sourceKind === "keynote" && options.allowKeynoteAutomation) {
    deck = await fallbackKeynoteBridgeIr(paths.source, validation);
    validation = validateIR(deck);
  }
  if (!validation.valid) {
    throw new Error(`Converted IR is invalid: ${validation.errors.map((error) => `${error.path} ${error.message}`).join("; ")}`);
  }

  const materializedNativeAssetCount = sourceKind === "keynote" ? await materializeBundleNativeAssets(paths.source, deck, paths) : 0;
  const lossReport = createLossReport(deck.conversion ?? { status: "success", messages: [] });
  const videoPlan = createVideoExportPlan(deck, options.video);
  const videoDependencies = await describeVideoDependencies(options.video?.ffmpegPath);
  const manifest = createBundleManifest({
    jobId,
    sourceName,
    sourceKind,
    createdAt,
    deck,
    lossReport,
    videoPlan,
    videoDependencies,
    paths
  });
  manifest.artifacts.nativeAssets = materializedNativeAssetCount > 0 ? "assets/native" : null;
  if (materializedNativeAssetCount > 0) {
    manifest.report.recommendedFixes = Array.from(
      new Set([
        ...manifest.report.recommendedFixes,
        "Use materialized native Keynote image assets for HTML preview instead of low-resolution text fallback."
      ])
    );
  }

  await writeJson(paths.deckIr, deck);
  await writeFile(paths.irRuntimeHtml, renderHtmlDocument(deck), "utf8");
  const runtime = await prepareBundleRuntime({
    sourcePath: paths.source,
    sourceKind,
    deck,
    paths,
    allowKeynoteAutomation: options.allowKeynoteAutomation,
    keynoteAutomationTimeoutMs: options.keynoteAutomationTimeoutMs,
    keynoteHtmlExport: options.keynoteHtmlExport,
    keynoteMovieExport: options.keynoteMovieExport
  });
  manifest.runtime = runtime;
  manifest.artifacts.runtimeHtml = "runtime.html";
  manifest.artifacts.irRuntimeHtml = runtime.irHtmlPath ? "runtime-ir.html" : null;
  manifest.artifacts.keynoteHtml = runtime.keynoteHtmlPath;
  manifest.artifacts.keynoteMovie = runtime.moviePath;
  manifest.artifacts.renderVideo = runtime.mode === "keynote-movie" ? runtime.moviePath : manifest.artifacts.renderVideo;
  await exportIrToPptx(deck, paths.rebuiltPptx);
  const segmented = await createSegmentedMoviePptx({
    deck,
    paths,
    createdAt,
    runtime,
    splitVideo: options.segmentVideoSplit,
    createPoster: options.segmentPosterCreate,
    splitOptions: { ffmpegPath: options.video?.ffmpegPath }
  });
  manifest.artifacts.segmentPlan = segmented.segmentPlanPath;
  manifest.artifacts.segmentedPptx = segmented.pptxPath;
  manifest.artifacts.staticStepsPptx = segmented.staticPptxPath;
  manifest.artifacts.hybridPptx = segmented.hybridPptxPath;
  if (segmented.pptxPath) {
    manifest.report.recommendedFixes = Array.from(
      new Set([
        ...manifest.report.recommendedFixes,
        "Use segmented.pptx for the smoothest high-fidelity animated playback; use static-steps.pptx only when presenter-controlled still states are preferred."
      ])
    );
  } else if (segmented.message) {
    deck.conversion?.messages.push({
      severity: "warning",
      code: "segmented-pptx-unavailable",
      message: segmented.message
    });
    await writeJson(paths.deckIr, deck);
  }
  await writeJson(paths.lossReport, lossReport);
  await writeJson(paths.videoPlan, {
    generatedAt: createdAt,
    plan: summarizeVideoPlan(videoPlan),
    frames: videoPlan.frames
  });
  await writeJson(paths.videoStatus, {
    generatedAt: createdAt,
    status: videoDependencies.missing.length ? "dependency-missing" : "ready",
    dependencies: videoDependencies,
    plan: summarizeVideoPlan(videoPlan),
    outputPath: paths.renderVideo
  });
  await writeJson(paths.baselineStatus, {
    generatedAt: createdAt,
    status: sourceKind === "keynote" ? "not-run" : "unsupported",
    message: manifest.baseline.message,
    plan: summarizeVideoPlan(videoPlan)
  });
  await writeJson(paths.manifest, manifest);

  return {
    jobId,
    sourceName,
    sourceKind,
    deck,
    lossReport,
    videoPlan,
    videoDependencies,
    manifest,
    paths
  };
}

export async function inspectProductInput(inputPath: string): Promise<{
  sourceName: string;
  sourceKind: ProductInputKind;
  slideCount: number;
  objectCount: number;
  validation: ReturnType<typeof validateIR>;
  report: ConversionLossReport;
  videoPlan: Omit<VideoExportPlan, "frames">;
  videoDependencies: VideoDependencyStatus;
}> {
  const resolvedInput = path.resolve(inputPath);
  const sourceName = sanitizeFileName(path.basename(resolvedInput));
  const sourceKind = detectInputKind(sourceName);
  const deck = await inputToIr(resolvedInput, sourceName, sourceKind, { allowKeynoteAutomation: false });
  const validation = validateIR(deck);
  const report = createLossReport(deck.conversion ?? { status: "success", messages: [] });
  const videoPlan = createVideoExportPlan(deck);
  const videoDependencies = await describeVideoDependencies();

  return {
    sourceName,
    sourceKind,
    slideCount: deck.deck.slides.length,
    objectCount: countObjects(deck),
    validation,
    report,
    videoPlan: summarizeVideoPlan(videoPlan),
    videoDependencies
  };
}

export async function benchmarkKeynoteDeck(inputPath: string, options: ProductKeyBenchmarkOptions = {}): Promise<ProductKeyBenchmarkResult> {
  const resolvedInput = path.resolve(inputPath);
  const sourceName = sanitizeFileName(options.sourceName ?? path.basename(resolvedInput));
  if (detectInputKind(sourceName) !== "keynote") {
    throw new Error("benchmark-key requires an original .key source deck.");
  }

  const outputDir = path.resolve(options.outputDir ?? path.join("demo", "out", "benchmarks", createJobId(sourceName)));
  const sourceCopyDir = path.join(outputDir, "source-copy");
  const copiedSourcePath = path.join(sourceCopyDir, sourceName);
  const bundleDir = path.join(outputDir, "bundle");
  const summaryPath = path.join(outputDir, "benchmark-summary.json");

  await rm(sourceCopyDir, { recursive: true, force: true });
  await mkdir(sourceCopyDir, { recursive: true });
  await copySource(resolvedInput, copiedSourcePath);

  const bundle = await createProductBundle(copiedSourcePath, bundleDir, {
    ...options,
    sourceName,
    jobId: options.jobId ?? createJobId(sourceName)
  });
  const runBaseline = options.runBaseline ?? Boolean(options.allowKeynoteAutomation);
  const baseline = runBaseline
    ? await exportProductBundleBaseline(bundleDir, {
        ...options.baseline,
        allowKeynoteAutomation: options.baseline?.allowKeynoteAutomation ?? options.allowKeynoteAutomation,
        keynoteAutomationTimeoutMs: options.baseline?.keynoteAutomationTimeoutMs ?? options.keynoteAutomationTimeoutMs,
        video: {
          ...options.video,
          ...options.baseline?.video
        }
      })
    : undefined;

  const summary = await createKeyBenchmarkSummary({
    generatedAt: new Date().toISOString(),
    inputPath: resolvedInput,
    copiedSourcePath,
    bundle,
    bundleDir,
    baseline
  });
  await writeJson(summaryPath, summary);
  return { inputPath: resolvedInput, copiedSourcePath, bundleDir, summaryPath, bundle, baseline, summary };
}

export function createProductFrameFidelityInsights(
  report: VideoFrameFidelityReport,
  options: { frameLimit?: number; slideLimit?: number } = {}
): ProductFrameFidelityInsights {
  return {
    worstFrames: summarizeWorstBenchmarkFrames(report, options.frameLimit ?? 12),
    worstSlides: summarizeWorstBenchmarkSlides(report, options.slideLimit ?? 8)
  };
}

export async function exportProductBundleKeynote(
  jobDir: string,
  options: Pick<ProductBundleOptions, "allowKeynoteAutomation" | "keynoteAutomationTimeoutMs"> = {}
): Promise<ProductKeynoteExportResult> {
  const paths = createExistingBundlePaths(jobDir);
  const deck = await readJson<DeckIR>(paths.deckIr);
  try {
    await exportIrToKeynote(deck, paths.rebuiltKeynote, {
      intermediatePptxPath: path.join(paths.jobDir, "rebuilt.key-bridge.pptx"),
      allowAutomation: options.allowKeynoteAutomation,
      automationTimeoutMs: options.keynoteAutomationTimeoutMs
    });
    await patchBundleManifest(paths.jobDir, (manifest) => ({
      ...manifest,
      artifacts: { ...manifest.artifacts, rebuiltKeynote: "rebuilt.key" },
      keynote: {
        ...manifest.keynote,
        available: true,
        message: "Keynote file generated."
      }
    }));
    return { status: "ready", outputPath: paths.rebuiltKeynote, message: "Keynote file generated." };
  } catch (error) {
    const message = `Keynote export unavailable: ${errorMessage(error)}`;
    await patchBundleManifest(paths.jobDir, (manifest) => ({
      ...manifest,
      keynote: {
        ...manifest.keynote,
        available: false,
        message
      }
    }));
    return { status: "dependency-missing", message };
  }
}

export async function exportProductBundleVideo(
  jobDir: string,
  options: VideoExportOptions = {}
): Promise<ProductVideoExportResult> {
  const paths = createExistingBundlePaths(jobDir);
  const deck = await readJson<DeckIR>(paths.deckIr);
  const renderOptions: VideoExportOptions = {
    fps: options.fps ?? 30,
    scale: options.scale ?? 4,
    ffmpegPath: options.ffmpegPath,
    keepFrames: true,
    framesDir: path.join(paths.jobDir, "frames", "rendering")
  };
  const plan = createVideoExportPlan(deck, renderOptions);
  const previousFramesDir = path.join(paths.jobDir, "frames", "previous");
  const renderingFramesDir = renderOptions.framesDir ?? path.join(paths.jobDir, "frames", "rendering");

  await rm(renderingFramesDir, { recursive: true, force: true });
  await rm(paths.frameDiffs, { recursive: true, force: true });
  await rm(previousFramesDir, { recursive: true, force: true });
  await renameIfExists(paths.framesLatest, previousFramesDir);

  try {
    await exportIrToVideo(deck, paths.renderVideo, renderOptions);
    await rename(renderingFramesDir, paths.framesLatest);
    const frameFidelity = (await directoryExists(previousFramesDir))
      ? await createVideoFrameFidelityReport(deck, previousFramesDir, paths.framesLatest, {
          fps: renderOptions.fps,
          scale: renderOptions.scale,
          diffDir: paths.frameDiffs,
          reportPath: paths.frameFidelity
        })
      : null;
    const result: ProductVideoExportResult = {
      status: "ready",
      outputPath: paths.renderVideo,
      plan: summarizeVideoPlan(plan),
      framesDir: paths.framesLatest,
      frameFidelity: frameFidelity
        ? {
            reportPath: paths.frameFidelity,
            diffDir: paths.frameDiffs,
            summary: frameFidelity.summary
          }
        : null
    };
    await writeJson(paths.videoStatus, {
      generatedAt: new Date().toISOString(),
      status: "ready",
      outputPath: paths.renderVideo,
      framesDir: paths.framesLatest,
      frameFidelity: result.frameFidelity,
      plan: result.plan,
      dependencies: await describeVideoDependencies(renderOptions.ffmpegPath)
    });
    await patchBundleManifest(paths.jobDir, (manifest) => ({
      ...manifest,
      artifacts: {
        ...manifest.artifacts,
        renderVideo: "render.mp4",
        frameFidelity: result.frameFidelity ? "frame-fidelity.json" : manifest.artifacts.frameFidelity
      }
    }));
    return result;
  } catch (error) {
    if (error instanceof VideoExportDependencyError) {
      await writeJson(paths.videoStatus, {
        generatedAt: new Date().toISOString(),
        status: "dependency-missing",
        missing: error.missing,
        message: error.message,
        guidance: error.guidance,
        plan: summarizeVideoPlan(plan)
      });
    }
    throw error;
  }
}

export async function exportProductBundleBaseline(
  jobDir: string,
  options: ProductBaselineExportOptions = {}
): Promise<ProductBaselineExportResult> {
  const paths = createExistingBundlePaths(jobDir);
  const manifest = await readJson<ProductBundleManifest>(paths.manifest);
  const sourcePath = path.join(paths.jobDir, manifest.artifacts.source);
  const renderOptions: VideoExportOptions = {
    fps: options.video?.fps ?? 30,
    scale: options.video?.scale ?? 4,
    ffmpegPath: options.video?.ffmpegPath
  };
  const deck = await readJson<DeckIR>(paths.deckIr);
  const plan = createVideoExportPlan(deck, renderOptions);

  if (manifest.sourceKind !== "keynote") {
    const message = "Keynote golden baseline is only available for original .key source decks.";
    await writeBaselineUnavailable(paths, "unsupported", message, summarizeVideoPlan(plan));
    return { status: "unsupported", message };
  }

  await rm(paths.baselineMovie, { force: true });
  await rm(paths.baselineFrames, { recursive: true, force: true });
  await rm(paths.baselineActualFrames, { recursive: true, force: true });
  await rm(paths.baselineDiffs, { recursive: true, force: true });

  try {
    const source: ProductBaselineExportReady["source"] = options.keynoteFrameExport ? "keynote-reference-frames" : "keynote-movie";
    if (options.keynoteFrameExport) {
      await options.keynoteFrameExport(sourcePath, paths.baselineFrames, plan, {
        allowAutomation: options.allowKeynoteAutomation,
        automationTimeoutMs: options.keynoteAutomationTimeoutMs,
        ...renderOptions
      });
    } else {
      await (options.keynoteMovieExport ?? exportKeynoteToMovie)(sourcePath, paths.baselineMovie, {
        allowAutomation: options.allowKeynoteAutomation,
        automationTimeoutMs: options.keynoteAutomationTimeoutMs
      });
      await extractVideoFramesFromVideo(deck, paths.baselineMovie, {
        ...renderOptions,
        framesDir: paths.baselineFrames
      });
    }

    if (options.runtimeFrameExport) {
      await options.runtimeFrameExport(deck, paths.baselineActualFrames, plan, renderOptions);
    } else {
      await captureVideoFrames(deck, {
        ...renderOptions,
        outputDir: paths.baselineActualFrames
      });
    }

    const report = await createVideoFrameFidelityReport(deck, paths.baselineFrames, paths.baselineActualFrames, {
      ...renderOptions,
      diffDir: paths.baselineDiffs,
      reportPath: paths.baselineFidelity
    });
    const referenceMoviePath = source === "keynote-movie" ? paths.baselineMovie : null;
    const result: ProductBaselineExportReady = {
      status: "ready",
      source,
      referenceMoviePath,
      referenceFramesDir: paths.baselineFrames,
      actualFramesDir: paths.baselineActualFrames,
      reportPath: paths.baselineFidelity,
      diffDir: paths.baselineDiffs,
      plan: summarizeVideoPlan(plan),
      summary: report.summary
    };

    await writeJson(paths.baselineStatus, {
      generatedAt: new Date().toISOString(),
      status: "ready",
      source,
      referenceMoviePath,
      referenceFramesDir: paths.baselineFrames,
      actualFramesDir: paths.baselineActualFrames,
      reportPath: paths.baselineFidelity,
      diffDir: paths.baselineDiffs,
      plan: result.plan,
      summary: report.summary
    });
    await patchBundleManifest(paths.jobDir, (current) => ({
      ...current,
      artifacts: {
        ...current.artifacts,
        keynoteBaselineMovie: referenceMoviePath ? "baseline/keynote-reference.m4v" : current.artifacts.keynoteBaselineMovie,
        baselineFrames: "frames/baseline",
        baselineActualFrames: "frames/keymorph-baseline",
        baselineFrameFidelity: "baseline-fidelity.json",
        baselineFrameDiffs: "baseline-diffs"
      },
      baseline: {
        ...current.baseline,
        available: true,
        message: "Keynote golden baseline pixel fidelity report is ready.",
        referenceMoviePath: referenceMoviePath ? "baseline/keynote-reference.m4v" : null,
        referenceFramesPath: "frames/baseline",
        actualFramesPath: "frames/keymorph-baseline",
        fidelityReportPath: "baseline-fidelity.json",
        diffPath: "baseline-diffs",
        summary: report.summary
      }
    }));
    return result;
  } catch (error) {
    const dependencyError = error instanceof VideoExportDependencyError ? error : null;
    const message = `Keynote golden baseline unavailable: ${errorMessage(error)}`;
    await writeBaselineUnavailable(paths, "dependency-missing", message, summarizeVideoPlan(plan), {
      missing: dependencyError?.missing,
      guidance: dependencyError?.guidance
    });
    return {
      status: "dependency-missing",
      message,
      missing: dependencyError?.missing,
      guidance: dependencyError?.guidance
    };
  }
}

export function createProductApiResponse(bundle: ProductBundleResult, basePath: string): Record<string, unknown> {
  const base = stripTrailingSlash(basePath);
  return {
    jobId: bundle.jobId,
    sourceName: bundle.sourceName,
    sourceKind: bundle.sourceKind,
    slideCount: bundle.manifest.slideCount,
    objectCount: bundle.manifest.objectCount,
    fidelityScore: bundle.lossReport.fidelityScore,
    animationLostCount: bundle.lossReport.animationLostCount,
    degradedAnimationCount: bundle.lossReport.degradedAnimationCount,
    uncertainMappingCount: bundle.lossReport.uncertainMappingCount,
    recommendedFixes: bundle.lossReport.recommendedFixes,
    previewUrl: `${base}/runtime.html`,
    manifestUrl: `${base}/manifest.json`,
    runtime: bundle.manifest.runtime,
    runtimeMode: bundle.manifest.runtime.mode,
    runtimeFidelity: bundle.manifest.runtime.fidelity,
    runtimeMessage: bundle.manifest.runtime.message,
    videoPlan: summarizeVideoPlan(bundle.videoPlan),
    videoDependencies: bundle.videoDependencies,
    videoEndpoint: `/api/jobs/${bundle.jobId}/video`,
    baselineEndpoint: `/api/jobs/${bundle.jobId}/baseline`,
    keynoteEndpoint: `/api/jobs/${bundle.jobId}/keynote`,
    downloads: {
      source: `${base}/${encodeURIComponent(bundle.sourceName)}`,
      html: `${base}/runtime.html`,
      ir: `${base}/deck.ir.json`,
      pptx: `${base}/rebuilt.pptx`,
      segmentedPptx: bundle.manifest.artifacts.segmentedPptx ? `${base}/${bundle.manifest.artifacts.segmentedPptx}` : null,
      staticStepsPptx: bundle.manifest.artifacts.staticStepsPptx ? `${base}/${bundle.manifest.artifacts.staticStepsPptx}` : null,
      hybridPptx: bundle.manifest.artifacts.hybridPptx ? `${base}/${bundle.manifest.artifacts.hybridPptx}` : null,
      key: null,
      report: `${base}/loss-report.json`,
      manifest: `${base}/manifest.json`,
      videoPlan: `${base}/video-plan.json`,
      videoStatus: `${base}/video-status.json`,
      video: bundle.manifest.artifacts.renderVideo ? `${base}/${bundle.manifest.artifacts.renderVideo}` : null,
      baselineStatus: `${base}/baseline-status.json`,
      baselineFidelity: bundle.manifest.artifacts.baselineFrameFidelity ? `${base}/${bundle.manifest.artifacts.baselineFrameFidelity}` : null,
      baselineDiffs: bundle.manifest.artifacts.baselineFrameDiffs ? `${base}/${bundle.manifest.artifacts.baselineFrameDiffs}/` : null
    },
    baselineCanRun: bundle.sourceKind === "keynote",
    baselineAvailable: bundle.manifest.baseline.available,
    baselineMessage: bundle.manifest.baseline.message,
    baselineSummary: bundle.manifest.baseline.summary ?? null,
    keynoteAvailable: null,
    keynoteMessage: bundle.manifest.keynote.message
  };
}

async function createKeyBenchmarkSummary(input: {
  generatedAt: string;
  inputPath: string;
  copiedSourcePath: string;
  bundle: ProductBundleResult;
  bundleDir: string;
  baseline?: ProductBaselineExportResult;
}): Promise<ProductKeyBenchmarkSummary> {
  const fidelityReport = await readOptionalFrameFidelityReport(path.join(input.bundleDir, "baseline-fidelity.json"));
  const baselineStatus = input.baseline?.status ?? "not-run";
  const baselineMessage =
    input.baseline?.message ??
    (fidelityReport ? "Keynote golden baseline pixel fidelity report is ready." : "Baseline was not run. Pass --allow-keynote to run it.");
  const insights = fidelityReport ? createProductFrameFidelityInsights(fidelityReport, { frameLimit: 12, slideLimit: 8 }) : null;
  const worstFrames = insights?.worstFrames ?? [];
  const worstSlides = insights?.worstSlides ?? [];
  const nextRecommendedActions = createBenchmarkRecommendations(input.bundle, baselineStatus, fidelityReport);

  return {
    generatedAt: input.generatedAt,
    inputPath: input.inputPath,
    copiedSourcePath: input.copiedSourcePath,
    bundleDir: input.bundleDir,
    sourceName: input.bundle.sourceName,
    slideCount: input.bundle.manifest.slideCount,
    objectCount: input.bundle.manifest.objectCount,
    runtime: input.bundle.manifest.runtime,
    conversion: input.bundle.manifest.report,
    videoPlan: summarizeVideoPlan(input.bundle.videoPlan),
    videoDependencies: input.bundle.videoDependencies,
    baseline: {
      status: baselineStatus,
      message: baselineMessage,
      reportPath: fidelityReport?.reportPath ?? null,
      diffPath: fidelityReport ? path.join(input.bundleDir, "baseline-diffs") : null,
      referenceFramesPath: fidelityReport ? path.join(input.bundleDir, "frames", "baseline") : null,
      actualFramesPath: fidelityReport ? path.join(input.bundleDir, "frames", "keymorph-baseline") : null,
      summary: fidelityReport?.summary,
      worstFrames,
      worstSlides
    },
    nextRecommendedActions
  };
}

async function readOptionalFrameFidelityReport(reportPath: string): Promise<VideoFrameFidelityReport | null> {
  try {
    return await readJson<VideoFrameFidelityReport>(reportPath);
  } catch {
    return null;
  }
}

function summarizeWorstBenchmarkFrames(report: VideoFrameFidelityReport, limit: number): ProductBenchmarkFrameSummary[] {
  return [...report.frames]
    .sort((left, right) => left.pixelFidelityScore - right.pixelFidelityScore || right.mismatchRatio - left.mismatchRatio)
    .slice(0, limit)
    .map((frame) => ({
      frame: frame.frame,
      slideIndex: frame.resolution.slideIndex,
      slideId: frame.resolution.slideId,
      timeMs: frame.timeMs,
      pixelFidelityScore: frame.pixelFidelityScore,
      mismatchRatio: frame.mismatchRatio,
      referencePath: frame.referencePath,
      actualPath: frame.actualPath,
      diffPath: frame.diffPath
    }));
}

function summarizeWorstBenchmarkSlides(report: VideoFrameFidelityReport, limit: number): ProductBenchmarkSlideSummary[] {
  return [...report.summary.bySlide]
    .sort((left, right) => left.meanPixelFidelityScore - right.meanPixelFidelityScore || right.mismatchRatio - left.mismatchRatio)
    .slice(0, limit)
    .map((slide) => ({
      slideIndex: slide.slideIndex,
      slideId: slide.slideId,
      frameCount: slide.frameCount,
      meanPixelFidelityScore: slide.meanPixelFidelityScore,
      mismatchRatio: slide.mismatchRatio,
      worstFrame: slide.worstFrame
    }));
}

function createBenchmarkRecommendations(
  bundle: ProductBundleResult,
  baselineStatus: ProductKeyBenchmarkSummary["baseline"]["status"],
  fidelityReport: VideoFrameFidelityReport | null
): string[] {
  const recommendations = new Set<string>(bundle.lossReport.recommendedFixes);
  if (baselineStatus === "not-run") {
    recommendations.add("Run benchmark-key with --allow-keynote to generate Keynote reference frames and pixel diffs.");
  }
  if (baselineStatus === "dependency-missing") {
    recommendations.add("Install or authorize the missing local dependencies reported in baseline-status.json, then rerun benchmark-key.");
  }
  if (fidelityReport?.summary.worstFrame !== undefined) {
    recommendations.add("Start fixes from baseline.worstFrames[0]; it identifies the lowest-scoring frame, slide, and diff PNG.");
  }
  if ((fidelityReport?.summary.meanPixelFidelityScore ?? 1) < 0.92) {
    recommendations.add("Prioritize HTML asset resolution, crop handling, text metrics, and native animation mappings before PPTX export.");
  }
  return Array.from(recommendations);
}

async function runDemo(): Promise<void> {
  const outDir = path.resolve("demo/out");
  await mkdir(outDir, { recursive: true });

  const sourceDeck = createDemoDeck();
  const sourceIrPath = path.join(outDir, "source.ir.json");
  const originalPptxPath = path.join(outDir, "original.pptx");
  const importedIrPath = path.join(outDir, "imported.ir.json");
  const htmlPath = path.join(outDir, "runtime.html");
  const pptxPath = path.join(outDir, "rebuilt.pptx");
  const reportPath = path.join(outDir, "conversion-report.json");

  await writeJson(sourceIrPath, sourceDeck);
  await exportIrToPptx(sourceDeck, originalPptxPath);

  const importedDeck = await parsePptxToIr(originalPptxPath);
  await writeJson(importedIrPath, importedDeck);
  await writeFile(htmlPath, renderHtmlDocument(importedDeck), "utf8");
  await exportIrToPptx(importedDeck, pptxPath);
  await writeJson(reportPath, scoreConversion(importedDeck.conversion ?? { status: "success", messages: [] }));

  console.log("Demo generated:");
  console.log(`  source IR: ${sourceIrPath}`);
  console.log(`  original PPTX: ${originalPptxPath}`);
  console.log(`  imported IR: ${importedIrPath}`);
  console.log(`  HTML runtime: ${pathToFileURL(htmlPath).toString()}`);
  console.log(`  rebuilt PPTX: ${pptxPath}`);
  console.log(`  report: ${reportPath}`);
}

async function main(): Promise<void> {
  const [command, input, output, ...args] = process.argv.slice(2);

  switch (command) {
    case "demo":
      await runDemo();
      return;
    case "convert": {
      if (!input || !output) throw new Error("Usage: convert <input.pptx|input.key|input.ir.json> <output-dir> [--allow-keynote] [--fps N] [--scale N] [--ffmpeg PATH]");
      const flags = parseFlags(args);
      const bundle = await createProductBundle(input, output, {
        allowKeynoteAutomation: Boolean(flags["allow-keynote"]),
        video: {
          fps: parsePositiveNumber(flags.fps),
          scale: parsePositiveNumber(flags.scale),
          ffmpegPath: typeof flags.ffmpeg === "string" ? flags.ffmpeg : undefined
        }
      });
      printBundleSummary(bundle);
      return;
    }
    case "inspect": {
      if (!input) throw new Error("Usage: inspect <input.pptx|input.key|input.ir.json>");
      console.log(JSON.stringify(await inspectProductInput(input), null, 2));
      return;
    }
    case "benchmark-key": {
      if (!input) {
        throw new Error("Usage: benchmark-key <input.key> [output-dir] [--allow-keynote] [--fps N] [--scale N] [--ffmpeg PATH]");
      }
      const flags = parseFlags(args);
      const result = await benchmarkKeynoteDeck(input, {
        outputDir: output,
        allowKeynoteAutomation: flags["allow-keynote"] === true,
        runBaseline: flags["allow-keynote"] === true,
        video: {
          fps: parsePositiveNumber(flags.fps),
          scale: parsePositiveNumber(flags.scale),
          ffmpegPath: typeof flags.ffmpeg === "string" ? flags.ffmpeg : undefined
        }
      });
      printBenchmarkSummary(result);
      return;
    }
    case "bundle-keynote": {
      if (!input) throw new Error("Usage: bundle-keynote <job-dir> [--allow-keynote]");
      console.log(JSON.stringify(await exportProductBundleKeynote(input, { allowKeynoteAutomation: parseFlags(args)["allow-keynote"] === true }), null, 2));
      return;
    }
    case "bundle-video": {
      if (!input) throw new Error("Usage: bundle-video <job-dir> [--fps N] [--scale N] [--ffmpeg PATH]");
      const flags = parseFlags(args);
      console.log(
        JSON.stringify(
          await exportProductBundleVideo(input, {
            fps: parsePositiveNumber(flags.fps),
            scale: parsePositiveNumber(flags.scale),
            ffmpegPath: typeof flags.ffmpeg === "string" ? flags.ffmpeg : undefined
          }),
          null,
          2
        )
      );
      return;
    }
    case "bundle-baseline": {
      if (!input) {
        throw new Error("Usage: bundle-baseline <job-dir> [--allow-keynote] [--fps N] [--scale N] [--ffmpeg PATH]");
      }
      const flags = parseFlags(args);
      console.log(
        JSON.stringify(
          await exportProductBundleBaseline(input, {
            allowKeynoteAutomation: flags["allow-keynote"] === true,
            video: {
              fps: parsePositiveNumber(flags.fps),
              scale: parsePositiveNumber(flags.scale),
              ffmpegPath: typeof flags.ffmpeg === "string" ? flags.ffmpeg : undefined
            }
          }),
          null,
          2
        )
      );
      return;
    }
    case "pptx-to-ir": {
      if (!input || !output) throw new Error("Usage: pptx-to-ir <input.pptx> <output.ir.json>");
      await writeJson(output, await parsePptxToIr(input));
      return;
    }
    case "key-to-ir": {
      if (!input || !output) throw new Error("Usage: key-to-ir <input.key> <output.ir.json>");
      await writeJson(output, await parseKeynoteToIr(input, { workDir: path.dirname(path.resolve(output)) }));
      return;
    }
    case "keyhtml-to-ir": {
      if (!input || !output) throw new Error("Usage: keyhtml-to-ir <keynote-html-export-dir> <output.ir.json>");
      await writeJson(output, await parseKeynoteHtmlExportToIr(input));
      return;
    }
    case "ir-to-html": {
      if (!input || !output) throw new Error("Usage: ir-to-html <input.ir.json> <output.html>");
      const deck = await readJson<DeckIR>(input);
      await mkdir(path.dirname(path.resolve(output)), { recursive: true });
      await writeFile(output, renderHtmlDocument(deck), "utf8");
      return;
    }
    case "ir-to-pptx": {
      if (!input || !output) throw new Error("Usage: ir-to-pptx <input.ir.json> <output.pptx>");
      await exportIrToPptx(await readJson<DeckIR>(input), output);
      return;
    }
    case "ir-to-key": {
      if (!input || !output) throw new Error("Usage: ir-to-key <input.ir.json> <output.key>");
      await exportIrToKeynote(await readJson<DeckIR>(input), output);
      return;
    }
    case "ir-to-video": {
      if (!input || !output) throw new Error("Usage: ir-to-video <input.ir.json> <output.mp4>");
      await exportIrToVideo(await readJson<DeckIR>(input), output);
      return;
    }
    case "ir-report": {
      if (!input || !output) throw new Error("Usage: ir-report <input.ir.json> <output.report.json>");
      const deck = await readJson<DeckIR>(input);
      await writeJson(output, createLossReport(deck.conversion ?? { status: "success", messages: [] }));
      return;
    }
    case "png-fidelity": {
      if (!input || !output) throw new Error("Usage: png-fidelity <reference.png> <actual.png> [output.report.json]");
      const report = await comparePngFiles(input, output);
      const reportPath = process.argv[5];
      if (reportPath) {
        await writeJson(reportPath, report);
      } else {
        console.log(JSON.stringify(report, null, 2));
      }
      return;
    }
    default:
      throw new Error(
        "Usage: keymorph <demo|convert|inspect|benchmark-key|bundle-keynote|bundle-video|bundle-baseline|pptx-to-ir|key-to-ir|keyhtml-to-ir|ir-to-html|ir-to-pptx|ir-to-key|ir-to-video|ir-report|png-fidelity> [input] [output]"
      );
  }
}

async function inputToIr(
  sourcePath: string,
  fileName: string,
  sourceKind = detectInputKind(fileName),
  options: Pick<ProductBundleOptions, "allowKeynoteAutomation" | "keynoteAutomationTimeoutMs" | "keynoteBridgeExport"> = {}
): Promise<DeckIR> {
  if (sourceKind === "pptx") return parsePptxToIr(sourcePath);
  if (sourceKind === "ir") return readJson<DeckIR>(sourcePath);
  if (sourceKind === "keynote") {
    return parseKeynoteToIr(sourcePath, {
      workDir: path.dirname(sourcePath),
      allowAutomation: options.allowKeynoteAutomation,
      automationTimeoutMs: options.keynoteAutomationTimeoutMs,
      bridgeExport: options.keynoteBridgeExport
    });
  }
  throw new Error("Unsupported file type. Use .pptx, .key, or .ir.json.");
}

async function fallbackKeynoteBridgeIr(sourcePath: string, validation: ReturnType<typeof validateIR>): Promise<DeckIR> {
  const deck = await parseKeynoteToIr(sourcePath, { preferNative: true });
  deck.conversion ??= { status: "partial", messages: [] };
  deck.conversion.status = "partial";
  deck.conversion.messages.unshift({
    severity: "warning",
    code: "keynote-bridge-ir-validation-fallback",
    message:
      "Keynote PPTX bridge produced unsupported timing dependency semantics, so native IR probing was used for analysis while the KeyMorph IR runtime remains the default preview."
  });
  deck.conversion.degradedFeatures ??= [];
  deck.conversion.degradedFeatures.push({
    code: "keynote-bridge-timing-cycle",
    severity: "warning",
    area: "animation",
    description: `Keynote-exported PPTX timing graph did not validate: ${validation.errors.map((error) => error.message).join(" ")}`,
    fallback: "Use KeyMorph IR runtime for preview and keep native IR probing for analysis artifacts."
  });
  return deck;
}

async function prepareBundleRuntime(input: {
  sourcePath: string;
  sourceKind: ProductInputKind;
  deck: DeckIR;
  paths: ProductBundlePaths;
  allowKeynoteAutomation?: boolean;
  keynoteAutomationTimeoutMs?: number;
  keynoteHtmlExport?: (keynotePath: string, outputDir: string, options: KeynoteAutomationOptions) => Promise<void>;
  keynoteMovieExport?: (keynotePath: string, moviePath: string, options: KeynoteAutomationOptions) => Promise<void>;
}): Promise<ProductRuntimeSummary> {
  if (input.sourceKind === "keynote" && input.allowKeynoteAutomation) {
    try {
      await (input.keynoteMovieExport ?? exportKeynoteToMovie)(input.sourcePath, input.paths.keynoteMovie, {
        allowAutomation: input.allowKeynoteAutomation,
        automationTimeoutMs: input.keynoteAutomationTimeoutMs
      });
      return writeMovieRuntime(input.paths);
    } catch (error) {
      if (await fileExistsWithContent(input.paths.keynoteMovie)) {
        return writeMovieRuntime(input.paths);
      }
      input.deck.conversion?.messages.push({
        severity: "warning",
        code: "keynote-movie-export-unavailable",
        message: `Keynote movie export failed, so Keynote HTML export was tried next. ${errorMessage(error)}`
      });
    }

    try {
      await (input.keynoteHtmlExport ?? exportKeynoteToHtml)(input.sourcePath, input.paths.keynoteHtmlDir, {
        allowAutomation: input.allowKeynoteAutomation,
        automationTimeoutMs: input.keynoteAutomationTimeoutMs
      });
      await copyFile(input.paths.irRuntimeHtml, input.paths.runtimeHtml);
      return {
        mode: "keymorph-ir",
        fidelity: "ir-reconstructed",
        message:
          "Runtime preview uses the KeyMorph IR. Keynote's native HTML export was preserved separately for comparison because it can expose Keynote template/check pages.",
        htmlPath: "runtime.html",
        irHtmlPath: null,
        moviePath: null,
        keynoteHtmlPath: "keynote-html/index.html"
      };
    } catch (error) {
      input.deck.conversion?.messages.push({
        severity: "warning",
        code: "keynote-html-export-unavailable",
        message: `Keynote HTML export failed, so KeyMorph IR runtime was used instead. ${errorMessage(error)}`
      });
    }
  }

  await copyFile(input.paths.irRuntimeHtml, input.paths.runtimeHtml);
  return {
    mode: "keymorph-ir",
    fidelity: "ir-reconstructed",
    message: "Runtime preview was rendered from the KeyMorph IR.",
    htmlPath: "runtime.html",
    irHtmlPath: null,
    moviePath: null,
    keynoteHtmlPath: null
  };
}

async function materializeBundleNativeAssets(sourcePath: string, deck: DeckIR, paths: ProductBundlePaths): Promise<number> {
  const requests = (deck.deck.assets ?? [])
    .map((asset) => {
      const sourcePath = typeof asset.metadata?.nativeSourcePath === "string" ? asset.metadata.nativeSourcePath : undefined;
      if (!sourcePath || asset.uri || asset.dataUri) {
        return undefined;
      }
      if (!nativeAssetIsUsedBySlide(deck, asset.id)) {
        return undefined;
      }
      return { assetId: asset.id, sourcePath };
    })
    .filter((request): request is { assetId: string; sourcePath: string } => Boolean(request));

  const materialized = await materializeNativeKeynoteAssetFiles(sourcePath, requests, paths.nativeAssetsDir, "assets/native");
  if (materialized.length === 0) {
    return 0;
  }

  const byAssetId = new Map(materialized.map((asset) => [asset.assetId, asset]));
  for (const asset of deck.deck.assets ?? []) {
    const match = byAssetId.get(asset.id);
    if (!match) {
      continue;
    }
    asset.uri = match.uri;
    asset.metadata = {
      ...(asset.metadata ?? {}),
      nativeMaterializedAsset: true,
      nativeMaterializedUri: match.uri,
      nativeMaterializedPath: path.relative(paths.jobDir, match.outputPath).split(path.sep).join(path.posix.sep),
      nativeMaterializedByteLength: match.byteLength
    };
  }

  deck.conversion ??= { status: "partial", messages: [] };
  deck.conversion.status = deck.conversion.status === "success" ? "partial" : deck.conversion.status;
  deck.conversion.messages.push({
    severity: "info",
    code: "keynote-native-assets-materialized",
    message: `Materialized ${materialized.length} native Keynote package asset(s) into bundle-local files for HTML runtime fidelity.`
  });
  deck.conversion.metadata = {
    ...(deck.conversion.metadata ?? {}),
    materializedNativeAssetCount: materialized.length
  };
  return materialized.length;
}

function nativeAssetIsUsedBySlide(deck: DeckIR, assetId: string): boolean {
  return deck.deck.slides.some((slide) => slideUsesAsset(slide, assetId));
}

function slideUsesAsset(slide: DeckIR["deck"]["slides"][number], assetId: string): boolean {
  if (slide.background?.type === "image" && slide.background.source.assetId === assetId) {
    return true;
  }
  return slide.objects.some((object) => objectUsesAsset(object, assetId));
}

function objectUsesAsset(object: DeckIR["deck"]["slides"][number]["objects"][number], assetId: string): boolean {
  if ((object.type === "image" || object.type === "media") && object.source.assetId === assetId) {
    return true;
  }
  if (object.type === "group") {
    return object.children.some((child) => objectUsesAsset(child, assetId));
  }
  return false;
}

async function createSegmentedMoviePptx(input: {
  deck: DeckIR;
  paths: ProductBundlePaths;
  createdAt: string;
  runtime: ProductRuntimeSummary;
  splitVideo?: ProductBundleOptions["segmentVideoSplit"];
  createPoster?: ProductBundleOptions["segmentPosterCreate"];
  splitOptions: SplitVideoSegmentsOptions;
}): Promise<{ segmentPlanPath: string | null; pptxPath: string | null; staticPptxPath: string | null; hybridPptxPath: string | null; message?: string }> {
  if (input.runtime.mode !== "keynote-movie" || !input.runtime.moviePath) {
    return { segmentPlanPath: null, pptxPath: null, staticPptxPath: null, hybridPptxPath: null };
  }

  const rawSegments = createSegmentPlan(input.deck);
  const movieDurationMs = await readOptionalMovieDurationMs(input.paths.keynoteMovie);
  const segments = movieDurationMs ? scaleSegmentPlanToDuration(rawSegments, movieDurationMs) : rawSegments;
  await writeJson(input.paths.segmentPlan, {
    generatedAt: input.createdAt,
    sourceMovie: input.runtime.moviePath,
    sourceTimelineDurationMs: rawSegments.length ? Math.max(...rawSegments.map((segment) => segment.endMs)) : 0,
    movieDurationMs: movieDurationMs ?? null,
    scaledToMovieDuration: Boolean(movieDurationMs),
    mode: "movie-segment-pptx",
    advanceMode: "click-after-playback",
    message:
      "Each entry maps one Keynote-rendered movie interval to one PPTX slide with an embedded full-slide video. Slides play automatically when entered, hold on the segment poster/final frame, and wait for the presenter click before advancing.",
    segments
  });

  if (segments.length === 0) {
    return { segmentPlanPath: "segment-plan.json", pptxPath: null, staticPptxPath: null, hybridPptxPath: null, message: "No playable timeline segments were found for segmented PPTX export." };
  }

  try {
    const split = await (input.splitVideo ?? splitVideoIntoSegments)(
      input.paths.keynoteMovie,
      segments,
      input.paths.segmentsDir,
      input.splitOptions
    );
    const posters = new Map<string, string>();
    let staticFrames: StaticStepFrame[];
    if (input.createPoster) {
      await mkdir(input.paths.segmentPostersDir, { recursive: true });
      const posterFrames: StaticStepFrame[] = [];
      for (const command of split.commands) {
        const posterPaths = normalizePosterPaths(
          await input.createPoster(command.outputPath, command.segment.durationMs, input.paths.segmentPostersDir, command.segment.id)
        );
        if (posterPaths[0]) posters.set(command.segment.id, posterPaths[0]);
        for (const [posterIndex, posterPath] of posterPaths.entries()) {
          posterFrames.push({
            id: `${command.segment.id}-poster-${posterIndex + 1}`,
            imagePath: posterPath,
            segment: command.segment,
            sampleMs: posterSampleMs(command.segment.durationMs, posterIndex, posterPaths.length)
          });
        }
      }
      staticFrames = posterFrames.length ? posterFrames : createStaticFramesFromPosters(split.commands, posters);
    } else {
      for (const [id, posterPath] of await createSegmentPosters(split.commands, input.paths.segmentPostersDir)) {
        posters.set(id, posterPath);
      }
      staticFrames = await createSegmentStaticFrames(split.commands, input.paths.segmentPostersDir);
    }
    const segmentedDeck = createSegmentedMovieDeck(input.deck, split.commands, posters);
    await exportIrToPptx(segmentedDeck, input.paths.segmentedPptx);
    const staticStepsDeck = createStaticStepDeck(input.deck, staticFrames);
    const staticPptxPath = staticStepsDeck ? "static-steps.pptx" : null;
    if (staticStepsDeck) await exportIrToPptx(staticStepsDeck, input.paths.staticStepsPptx);
    const hybridFrames = await selectMeaningfulHybridHoldFrames(staticFrames);
    const hybridDeck = createHybridStepDeck(input.deck, split.commands, posters, hybridFrames);
    const hybridPptxPath = hybridDeck ? "hybrid.pptx" : null;
    if (hybridDeck) await exportIrToPptx(hybridDeck, input.paths.hybridPptx);
    return { segmentPlanPath: "segment-plan.json", pptxPath: "segmented.pptx", staticPptxPath, hybridPptxPath };
  } catch (error) {
    return {
      segmentPlanPath: "segment-plan.json",
      pptxPath: null,
      staticPptxPath: null,
      hybridPptxPath: null,
      message: `Segmented high-fidelity PPTX export failed. Install ffmpeg or use macOS avconvert, then rerun conversion. ${errorMessage(error)}`
    };
  }
}

async function readOptionalMovieDurationMs(moviePath: string): Promise<number | undefined> {
  try {
    return await readMovieDurationMs(moviePath);
  } catch {
    return undefined;
  }
}

async function createSegmentPosters(commands: SplitVideoSegmentsResult["commands"], outputDir: string): Promise<Map<string, string>> {
  const posters = new Map<string, string>();
  await mkdir(outputDir, { recursive: true });
  for (const command of commands) {
    const posterPath = await createSegmentPoster(command.outputPath, command.segment.durationMs, outputDir, command.segment.id);
    if (posterPath) posters.set(command.segment.id, posterPath);
  }
  return posters;
}

async function createSegmentPoster(videoPath: string, durationMs: number, outputDir: string, id: string): Promise<string | undefined> {
  const sampleMs = Math.max(500, Math.min(durationMs * 0.9, durationMs - 250));
  return createSegmentPosterAt(videoPath, sampleMs, outputDir, `${id}-poster`);
}

function posterSampleMs(durationMs: number, index: number, count: number): number {
  const duration = Math.max(1, Math.round(durationMs));
  if (count <= 1) return Math.max(0, Math.min(duration - 1, Math.round(duration * 0.9)));
  if (index >= count - 1) return Math.max(0, duration - 250);
  const ratio = (index + 1) / count;
  return Math.max(0, Math.min(duration - 1, Math.round(duration * ratio)));
}

interface StaticStepFrame {
  id: string;
  imagePath: string;
  segment: VideoSegmentPlanEntry;
  sampleMs: number;
}

async function createSegmentStaticFrames(commands: SplitVideoSegmentsResult["commands"], outputDir: string): Promise<StaticStepFrame[]> {
  const candidates: StaticStepFrame[] = [];
  await mkdir(outputDir, { recursive: true });
  for (const command of commands) {
    for (const sampleMs of segmentStaticSampleTimes(command.segment.durationMs)) {
      const imagePath = await createSegmentPosterAt(command.outputPath, sampleMs, outputDir, `${command.segment.id}-static-${sampleMs}`);
      if (imagePath) {
        candidates.push({
          id: `${command.segment.id}-static-${sampleMs}`,
          imagePath,
          segment: command.segment,
          sampleMs
        });
      }
    }
  }
  return dedupeStaticFrames(candidates);
}

function segmentStaticSampleTimes(durationMs: number): number[] {
  const duration = Math.max(1, Math.round(durationMs));
  const times = new Set<number>([Math.max(0, Math.min(duration - 250, Math.round(duration * 0.9)))]);
  for (let timeMs = 1000; timeMs < duration - 250; timeMs += 2000) {
    times.add(timeMs);
  }
  if (duration > 1500) times.add(Math.max(500, duration - 250));
  return Array.from(times)
    .filter((timeMs) => timeMs >= 0 && timeMs < duration)
    .sort((a, b) => a - b);
}

async function createSegmentPosterAt(videoPath: string, sampleMs: number, outputDir: string, id: string): Promise<string | undefined> {
  const sampleVideo = path.join(outputDir, `${id}-source.mp4`);
  try {
    await runCommand("/usr/bin/avconvert", [
      "--source",
      videoPath,
      "--output",
      sampleVideo,
      "--replace",
      "--preset",
      "PresetHighestQuality",
      "--start",
      formatSeconds(sampleMs / 1000),
      "--duration",
      "0.5"
    ]);
    await runCommand("/usr/bin/qlmanage", ["-t", "-s", "5120", "-o", outputDir, sampleVideo]);
    const generated = `${sampleVideo}.png`;
    if (await fileExistsWithContent(generated)) return generated;
  } catch {
    // Static step extraction is best-effort; keep the rest of the bundle usable.
  }
  return undefined;
}

async function dedupeStaticFrames(frames: StaticStepFrame[]): Promise<StaticStepFrame[]> {
  const kept: StaticStepFrame[] = [];
  for (const frame of frames) {
    const previous = kept[kept.length - 1];
    if (previous) {
      const comparison = await comparePngFiles(previous.imagePath, frame.imagePath, { threshold: 0.02 });
      if (comparison.pixelFidelityScore >= 0.995 || comparison.mismatchRatio <= 0.001) continue;
    }
    kept.push(frame);
  }
  const last = frames[frames.length - 1];
  if (last && kept[kept.length - 1]?.imagePath !== last.imagePath) kept.push(last);
  return kept;
}

async function selectMeaningfulHybridHoldFrames(frames: StaticStepFrame[]): Promise<StaticStepFrame[]> {
  const framesBySegment = groupStaticFramesBySegment(frames);
  const selected: StaticStepFrame[] = [];
  for (const segmentFrames of framesBySegment.values()) {
    selected.push(...(await selectMeaningfulSegmentHoldFrames(segmentFrames)));
  }
  return selected;
}

async function selectMeaningfulSegmentHoldFrames(frames: StaticStepFrame[]): Promise<StaticStepFrame[]> {
  const sorted = [...frames].sort((a, b) => a.sampleMs - b.sampleMs);
  const finalFrame = sorted[sorted.length - 1];
  if (!finalFrame) return [];
  const selected: StaticStepFrame[] = [];
  for (const frame of sorted.slice(0, -1).reverse()) {
    if (frame.sampleMs < finalFrame.sampleMs - 500 && (await frameLooksMeaningfullyDifferent(frame, finalFrame))) {
      selected.push(frame);
      break;
    }
  }
  selected.push(finalFrame);
  return selected;
}

async function frameLooksMeaningfullyDifferent(a: StaticStepFrame, b: StaticStepFrame): Promise<boolean> {
  if (a.imagePath === b.imagePath) return false;
  try {
    const comparison = await comparePngFiles(a.imagePath, b.imagePath, { threshold: 0.02 });
    return comparison.mismatchRatio >= 0.035 || comparison.pixelFidelityScore <= 0.965;
  } catch {
    return true;
  }
}

function groupStaticFramesBySegment(frames: StaticStepFrame[]): Map<string, StaticStepFrame[]> {
  const framesBySegment = new Map<string, StaticStepFrame[]>();
  for (const frame of frames) {
    const list = framesBySegment.get(frame.segment.id) ?? [];
    list.push(frame);
    framesBySegment.set(frame.segment.id, list);
  }
  return framesBySegment;
}

function normalizePosterPaths(value: string | string[] | undefined): string[] {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function createStaticFramesFromPosters(
  commands: SplitVideoSegmentsResult["commands"],
  posters: Map<string, string>
): StaticStepFrame[] {
  return commands.flatMap((command) => {
    const imagePath = posters.get(command.segment.id);
    return imagePath
      ? [
          {
            id: `${command.segment.id}-poster`,
            imagePath,
            segment: command.segment,
            sampleMs: Math.max(0, Math.round(command.segment.durationMs * 0.9))
          }
        ]
      : [];
  });
}

function createSegmentedMovieDeck(sourceDeck: DeckIR, commands: SplitVideoSegmentsResult["commands"], posters: Map<string, string> = new Map()): DeckIR {
  const slides: Slide[] = commands.map((command, index) => {
    const sourceSlide = sourceDeck.deck.slides[command.segment.slideIndex];
    const objectId = `segment-video-${index + 1}`;
    const posterPath = posters.get(command.segment.id);
    return {
      id: `segment-slide-${index + 1}`,
      index,
      name: `${sourceSlide?.name ?? `Slide ${command.segment.slideIndex + 1}`} · segment ${command.segment.clickIndex + 1}`,
      background: { type: "solid", color: "#000000" },
      objects: [
        {
          id: objectId,
          type: "media",
          mediaType: "video",
          name: command.segment.outputName,
          bounds: { x: 0, y: 0, width: sourceDeck.deck.size.width, height: sourceDeck.deck.size.height },
          opacity: 1,
          source: { uri: command.outputPath },
          posterSource: posterPath ? { uri: posterPath } : undefined,
          playback: { autoplay: true, startMs: 0, endMs: command.segment.durationMs },
          metadata: {
            sourceSlideId: sourceSlide?.id ?? null,
            sourceSlideIndex: command.segment.slideIndex,
            sourceClickIndex: command.segment.clickIndex,
            sourceMovieStartMs: command.segment.startMs,
            sourceMovieEndMs: command.segment.endMs,
            sourceMovieDurationMs: command.segment.durationMs,
            keymorphSegmentedMovie: true
          }
        }
      ],
      timeline: {
        durationMs: command.segment.durationMs,
        events: [
          {
            id: `segment-play-${index + 1}`,
            kind: "media",
            targetId: objectId,
            action: "play",
            seekMs: 0,
            start: { type: "absolute", atMs: 0 },
            durationMs: 1,
            fill: "forwards"
          }
        ]
      },
      transition: {
        type: "cut",
        trigger: "click",
        durationMs: 0,
        metadata: {
          keymorphSegmentedMovie: true,
          advanceMode: "click-after-playback",
          holdFinalFrame: true,
          playbackDurationMs: command.segment.durationMs
        }
      },
      metadata: {
        sourceSlideId: sourceSlide?.id ?? null,
        sourceSlideIndex: command.segment.slideIndex,
        sourceClickIndex: command.segment.clickIndex,
        keymorphSegmentedMovie: true
      }
    };
  });

  return {
    irVersion: IR_VERSION,
    metadata: {
      ...(sourceDeck.metadata ?? {}),
      title: `${sourceDeck.metadata?.title ?? sourceDeck.deck.title ?? "Keynote"} segmented playback`,
      sourceApplication: "KeyMorph segmented movie PPTX"
    },
    deck: {
      id: `${sourceDeck.deck.id}-segmented-movie`,
      title: `${sourceDeck.deck.title ?? "Keynote"} segmented playback`,
      size: sourceDeck.deck.size,
      slides
    },
    conversion: {
      status: "partial",
      messages: [
        {
          severity: "info",
          code: "segmented-movie-pptx",
          message:
            "Generated from Keynote's rendered movie, split at KeyMorph timeline boundaries, and embedded as one click-advanced full-slide video step per PPTX slide."
        }
      ],
      degradedFeatures: [
        {
          code: "segmented-movie-pptx-not-editable",
          severity: "info",
          area: "media",
          description: "Segmented PPTX preserves visual animation playback but replaces editable slide objects with full-slide videos.",
          fallback: "Edit the source Keynote deck and rerun conversion to refresh the segmented playback draft."
        }
      ]
    }
  };
}

function createHybridStepDeck(
  sourceDeck: DeckIR,
  commands: SplitVideoSegmentsResult["commands"],
  posters: Map<string, string>,
  frames: StaticStepFrame[]
): DeckIR | undefined {
  if (commands.length === 0) return undefined;
  const slides: Slide[] = [];
  const framesBySegment = groupStaticFramesBySegment(frames);

  for (const command of commands) {
    const segmentFrames = (framesBySegment.get(command.segment.id) ?? []).sort((a, b) => a.sampleMs - b.sampleMs);
    const videoSlideIndex = slides.length;
    slides.push(createHybridVideoSlide(sourceDeck, command, posters.get(command.segment.id), videoSlideIndex));
    for (const frame of segmentFrames) {
      const holdSlideIndex = slides.length;
      slides.push(createHybridHoldSlide(sourceDeck, frame, holdSlideIndex));
    }
  }

  if (slides.length === 0) return undefined;
  return {
    irVersion: IR_VERSION,
    metadata: {
      ...(sourceDeck.metadata ?? {}),
      title: `${sourceDeck.metadata?.title ?? sourceDeck.deck.title ?? "Keynote"} hybrid playback`,
      sourceApplication: "KeyMorph hybrid video/static PPTX"
    },
    deck: {
      id: `${sourceDeck.deck.id}-hybrid`,
      title: `${sourceDeck.deck.title ?? "Keynote"} hybrid playback`,
      size: sourceDeck.deck.size,
      slides
    },
    conversion: {
      status: "partial",
      messages: [
        {
          severity: "info",
          code: "hybrid-video-static-pptx",
          message:
            "Generated from Keynote-rendered movie segments, preserving each smooth video interval and adding click-held visual states only when the rendered movie shows a meaningful completed state."
        }
      ],
      degradedFeatures: [
        {
          code: "hybrid-video-static-pptx-not-editable",
          severity: "info",
          area: "media",
          description: "Hybrid PPTX preserves rendered animation intervals and completed visual states, but replaces editable slide objects with videos and images.",
          fallback: "Edit the source Keynote deck and rerun conversion to refresh the hybrid draft."
        }
      ]
    }
  };
}

function createHybridVideoSlide(
  sourceDeck: DeckIR,
  command: SplitVideoSegmentsResult["commands"][number],
  posterPath: string | undefined,
  index: number
): Slide {
  const sourceSlide = sourceDeck.deck.slides[command.segment.slideIndex];
  const objectId = `hybrid-video-${index + 1}`;
  const durationMs = command.segment.durationMs;
  return {
    id: `hybrid-video-slide-${index + 1}`,
    index,
    name: `${sourceSlide?.name ?? `Slide ${command.segment.slideIndex + 1}`} · smooth segment ${command.segment.clickIndex + 1}`,
    background: { type: "solid", color: "#000000" },
    objects: [
      {
        id: objectId,
        type: "media",
        mediaType: "video",
        name: command.segment.outputName,
        bounds: { x: 0, y: 0, width: sourceDeck.deck.size.width, height: sourceDeck.deck.size.height },
        opacity: 1,
        source: { uri: command.outputPath },
        posterSource: posterPath ? { uri: posterPath } : undefined,
        playback: { autoplay: true, startMs: 0, endMs: durationMs },
        metadata: {
          sourceSlideId: sourceSlide?.id ?? null,
          sourceSlideIndex: command.segment.slideIndex,
          sourceClickIndex: command.segment.clickIndex,
          sourceSegmentStartMs: command.segment.startMs,
          sourceSegmentSeekMs: 0,
          sourceSegmentDurationMs: durationMs,
          keymorphHybridVideo: true
        }
      }
    ],
    timeline: {
      durationMs,
      events: [
        {
          id: `hybrid-play-${index + 1}`,
          kind: "media",
          targetId: objectId,
          action: "play",
          seekMs: 0,
          start: { type: "absolute", atMs: 0 },
          durationMs: 1,
          fill: "forwards"
        }
      ]
    },
    transition: {
      type: "cut",
      trigger: "auto",
      durationMs: 0,
      metadata: {
        keymorphSegmentedMovie: true,
        autoAdvanceAfterMs: durationMs,
        disableClickAdvanceUntilMs: durationMs,
        keymorphHybridVideo: true
      }
    },
    metadata: {
      sourceSlideId: sourceSlide?.id ?? null,
      sourceSlideIndex: command.segment.slideIndex,
      sourceClickIndex: command.segment.clickIndex,
      keymorphHybridVideo: true
    }
  };
}

function createHybridHoldSlide(sourceDeck: DeckIR, frame: StaticStepFrame, index: number): Slide {
  const sourceSlide = sourceDeck.deck.slides[frame.segment.slideIndex];
  return {
    id: `hybrid-hold-slide-${index + 1}`,
    index,
    name: `${sourceSlide?.name ?? `Slide ${frame.segment.slideIndex + 1}`} · hold ${Math.round(frame.sampleMs)}ms`,
    background: { type: "solid", color: "#000000" },
    objects: [
      {
        id: `hybrid-hold-image-${index + 1}`,
        type: "image",
        name: `${frame.id} hold`,
        bounds: { x: 0, y: 0, width: sourceDeck.deck.size.width, height: sourceDeck.deck.size.height },
        opacity: 1,
        source: { uri: frame.imagePath },
        metadata: {
          sourceSlideId: sourceSlide?.id ?? null,
          sourceSlideIndex: frame.segment.slideIndex,
          sourceClickIndex: frame.segment.clickIndex,
          sourceMovieStartMs: frame.segment.startMs,
          sourceMovieEndMs: frame.segment.endMs,
          sourceSegmentSampleMs: frame.sampleMs,
          sourceMovieSampleMs: frame.segment.startMs + frame.sampleMs,
          keymorphHybridHold: true
        }
      }
    ],
    timeline: { durationMs: 1, events: [], dependencyGraph: { edges: [] } },
    transition: { type: "cut", trigger: "click", durationMs: 0 },
    metadata: {
      sourceSlideId: sourceSlide?.id ?? null,
      sourceSlideIndex: frame.segment.slideIndex,
      sourceClickIndex: frame.segment.clickIndex,
      sourceSegmentSampleMs: frame.sampleMs,
      sourceMovieSampleMs: frame.segment.startMs + frame.sampleMs,
      keymorphHybridHold: true
    }
  };
}

function createStaticStepDeck(
  sourceDeck: DeckIR,
  frames: StaticStepFrame[]
): DeckIR | undefined {
  const slides: Slide[] = [];
  for (const [index, frame] of frames.entries()) {
    const sourceSlide = sourceDeck.deck.slides[frame.segment.slideIndex];
    slides.push({
      id: `static-step-slide-${index + 1}`,
      index: slides.length,
      name: `${sourceSlide?.name ?? `Slide ${frame.segment.slideIndex + 1}`} · static step ${index + 1}`,
      background: { type: "solid", color: "#000000" },
      objects: [
        {
          id: `static-step-image-${index + 1}`,
          type: "image",
          name: `${frame.id} frame`,
          bounds: { x: 0, y: 0, width: sourceDeck.deck.size.width, height: sourceDeck.deck.size.height },
          opacity: 1,
          source: { uri: frame.imagePath },
          metadata: {
            sourceSlideId: sourceSlide?.id ?? null,
            sourceSlideIndex: frame.segment.slideIndex,
            sourceClickIndex: frame.segment.clickIndex,
            sourceMovieStartMs: frame.segment.startMs,
            sourceMovieEndMs: frame.segment.endMs,
            sourceSegmentSampleMs: frame.sampleMs,
            sourceMovieSampleMs: frame.segment.startMs + frame.sampleMs,
            keymorphStaticStep: true
          }
        }
      ],
      timeline: { durationMs: 1, events: [], dependencyGraph: { edges: [] } },
      transition: { type: "cut", trigger: "click", durationMs: 0 },
      metadata: {
        sourceSlideId: sourceSlide?.id ?? null,
        sourceSlideIndex: frame.segment.slideIndex,
        sourceClickIndex: frame.segment.clickIndex,
        sourceSegmentSampleMs: frame.sampleMs,
        sourceMovieSampleMs: frame.segment.startMs + frame.sampleMs,
        keymorphStaticStep: true
      }
    });
  }
  if (slides.length === 0) return undefined;

  return {
    irVersion: IR_VERSION,
    metadata: {
      ...(sourceDeck.metadata ?? {}),
      title: `${sourceDeck.metadata?.title ?? sourceDeck.deck.title ?? "Keynote"} static build steps`,
      sourceApplication: "KeyMorph static step PPTX"
    },
    deck: {
      id: `${sourceDeck.deck.id}-static-steps`,
      title: `${sourceDeck.deck.title ?? "Keynote"} static build steps`,
      size: sourceDeck.deck.size,
      slides
    },
    conversion: {
      status: "partial",
      messages: [
        {
          severity: "info",
          code: "static-step-pptx",
          message:
            "Generated from deduplicated Keynote-rendered movie samples, with one static full-slide image per discovered visual step."
        }
      ],
      degradedFeatures: [
        {
          code: "static-step-pptx-not-editable",
          severity: "info",
          area: "media",
          description: "Static step PPTX preserves each build's completed visual state but replaces editable slide objects and animation playback with full-slide images.",
          fallback: "Edit the source Keynote deck and rerun conversion to refresh the static step draft."
        }
      ]
    }
  };
}

async function writeMovieRuntime(paths: ProductBundlePaths): Promise<ProductRuntimeSummary> {
  await writeFile(paths.runtimeHtml, renderMovieRuntimeHtml("keynote-movie.m4v"), "utf8");
  return {
    mode: "keynote-movie",
    fidelity: "keynote-rendered-video",
    message: "Runtime preview uses Keynote's rendered QuickTime movie export for high-fidelity animated playback.",
    htmlPath: "runtime.html",
    irHtmlPath: "runtime-ir.html",
    moviePath: "keynote-movie.m4v",
    keynoteHtmlPath: null
  };
}

function renderMovieRuntimeHtml(moviePath: string): string {
  const escapedPath = escapeHtmlAttr(moviePath);
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>KeyMorph Keynote Movie Runtime</title>
  <style>
    html, body { width: 100%; height: 100%; margin: 0; background: #000; overflow: hidden; }
    video { width: 100%; height: 100%; object-fit: contain; background: #000; display: block; }
  </style>
</head>
<body>
  <video src="${escapedPath}" controls autoplay playsinline preload="metadata"></video>
</body>
</html>
`;
}

function escapeHtmlAttr(value: string): string {
  return value.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char] ?? char);
}

function detectInputKind(fileName: string): ProductInputKind {
  const lowerName = fileName.toLowerCase();
  if (lowerName.endsWith(".pptx")) return "pptx";
  if (lowerName.endsWith(".key")) return "keynote";
  if (lowerName.endsWith(".ir.json") || lowerName.endsWith(".json")) return "ir";
  throw new Error("Unsupported file type. Use .pptx, .key, or .ir.json.");
}

function createBundlePaths(jobDir: string, sourceName: string): ProductBundlePaths {
  return {
    jobDir,
    source: path.join(jobDir, sourceName),
    deckIr: path.join(jobDir, "deck.ir.json"),
    runtimeHtml: path.join(jobDir, "runtime.html"),
    irRuntimeHtml: path.join(jobDir, "runtime-ir.html"),
    keynoteHtmlDir: path.join(jobDir, "keynote-html"),
    keynoteHtmlIndex: path.join(jobDir, "keynote-html", "index.html"),
    keynoteMovie: path.join(jobDir, "keynote-movie.m4v"),
    segmentsDir: path.join(jobDir, "segments"),
    segmentPlan: path.join(jobDir, "segment-plan.json"),
    segmentPostersDir: path.join(jobDir, "posters"),
    nativeAssetsDir: path.join(jobDir, "assets", "native"),
    rebuiltPptx: path.join(jobDir, "rebuilt.pptx"),
    segmentedPptx: path.join(jobDir, "segmented.pptx"),
    staticStepsPptx: path.join(jobDir, "static-steps.pptx"),
    hybridPptx: path.join(jobDir, "hybrid.pptx"),
    lossReport: path.join(jobDir, "loss-report.json"),
    manifest: path.join(jobDir, "manifest.json"),
    videoPlan: path.join(jobDir, "video-plan.json"),
    videoStatus: path.join(jobDir, "video-status.json"),
    rebuiltKeynote: path.join(jobDir, "rebuilt.key"),
    renderVideo: path.join(jobDir, "render.mp4"),
    framesLatest: path.join(jobDir, "frames", "latest"),
    frameFidelity: path.join(jobDir, "frame-fidelity.json"),
    frameDiffs: path.join(jobDir, "frame-diffs"),
    baselineStatus: path.join(jobDir, "baseline-status.json"),
    baselineMovie: path.join(jobDir, "baseline", "keynote-reference.m4v"),
    baselineFrames: path.join(jobDir, "frames", "baseline"),
    baselineActualFrames: path.join(jobDir, "frames", "keymorph-baseline"),
    baselineFidelity: path.join(jobDir, "baseline-fidelity.json"),
    baselineDiffs: path.join(jobDir, "baseline-diffs")
  };
}

function createExistingBundlePaths(jobDir: string): ProductBundlePaths {
  return createBundlePaths(path.resolve(jobDir), "source");
}

function createBundleManifest(input: {
  jobId: string;
  sourceName: string;
  sourceKind: ProductInputKind;
  createdAt: string;
  deck: DeckIR;
  lossReport: ConversionLossReport;
  videoPlan: VideoExportPlan;
  videoDependencies: VideoDependencyStatus;
  paths: ProductBundlePaths;
}): ProductBundleManifest {
  return {
    jobId: input.jobId,
    sourceName: input.sourceName,
    sourceKind: input.sourceKind,
    createdAt: input.createdAt,
    slideCount: input.deck.deck.slides.length,
    objectCount: countObjects(input.deck),
    artifacts: {
      source: input.sourceName,
      deckIr: "deck.ir.json",
      runtimeHtml: "runtime.html",
      irRuntimeHtml: null,
      keynoteHtml: null,
      keynoteMovie: null,
      segmentPlan: null,
      segmentedPptx: null,
      staticStepsPptx: null,
      hybridPptx: null,
      nativeAssets: null,
      rebuiltPptx: "rebuilt.pptx",
      lossReport: "loss-report.json",
      videoPlan: "video-plan.json",
      videoStatus: "video-status.json",
      rebuiltKeynote: null,
      renderVideo: null,
      frameFidelity: null,
      baselineStatus: "baseline-status.json",
      keynoteBaselineMovie: null,
      baselineFrames: null,
      baselineActualFrames: null,
      baselineFrameFidelity: null,
      baselineFrameDiffs: null
    },
    report: {
      fidelityScore: input.lossReport.fidelityScore,
      riskLevel: input.lossReport.riskLevel,
      animationLostCount: input.lossReport.animationLostCount,
      degradedAnimationCount: input.lossReport.degradedAnimationCount,
      uncertainMappingCount: input.lossReport.uncertainMappingCount,
      recommendedFixes: input.lossReport.recommendedFixes
    },
    runtime: {
      mode: "keymorph-ir",
      fidelity: "ir-reconstructed",
      message: "HTML runtime was rendered from the KeyMorph IR.",
      htmlPath: "runtime.html",
      irHtmlPath: null,
      moviePath: null,
      keynoteHtmlPath: null
    },
    video: {
      plan: summarizeVideoPlan(input.videoPlan),
      dependencies: input.videoDependencies,
      endpoint: null,
      statusPath: "video-status.json"
    },
    baseline: {
      available: null,
      endpoint: null,
      message:
        input.sourceKind === "keynote"
          ? "Keynote golden baseline is available on demand and may ask macOS for Keynote automation permission."
          : "Keynote golden baseline requires an original .key source deck.",
      statusPath: "baseline-status.json",
      referenceMoviePath: null,
      referenceFramesPath: null,
      actualFramesPath: null,
      fidelityReportPath: null,
      diffPath: null
    },
    keynote: {
      available: null,
      endpoint: null,
      message: "Keynote export is available on demand and may ask macOS for Keynote automation permission.",
      outputPath: input.paths.rebuiltKeynote
    }
  };
}

function summarizeVideoPlan(plan: VideoExportPlan): Omit<VideoExportPlan, "frames"> {
  return {
    width: plan.width,
    height: plan.height,
    scale: plan.scale,
    fps: plan.fps,
    durationMs: plan.durationMs,
    totalFrames: plan.totalFrames,
    outputWidth: plan.outputWidth,
    outputHeight: plan.outputHeight
  };
}

function countObjects(deck: DeckIR): number {
  return deck.deck.slides.reduce((sum, slide) => sum + slide.objects.reduce((slideSum, object) => slideSum + countObjectTree(object), 0), 0);
}

function countObjectTree(object: DeckIR["deck"]["slides"][number]["objects"][number]): number {
  if (object.type !== "group") return 1;
  return 1 + object.children.reduce((sum, child) => sum + countObjectTree(child), 0);
}

async function copySource(from: string, to: string): Promise<void> {
  if (path.resolve(from) === path.resolve(to)) return;
  await mkdir(path.dirname(to), { recursive: true });
  const stats = await stat(from);
  if (stats.isDirectory()) {
    await rm(to, { recursive: true, force: true });
    await cp(from, to, { recursive: true });
    return;
  }
  await copyFile(from, to);
}

async function patchBundleManifest(jobDir: string, patch: (manifest: ProductBundleManifest) => ProductBundleManifest): Promise<void> {
  const manifestPath = path.join(jobDir, "manifest.json");
  const manifest = await readJson<ProductBundleManifest>(manifestPath);
  await writeJson(manifestPath, patch(manifest));
}

async function writeBaselineUnavailable(
  paths: ProductBundlePaths,
  status: ProductBaselineExportUnavailable["status"],
  message: string,
  plan: Omit<VideoExportPlan, "frames">,
  details: Pick<ProductBaselineExportUnavailable, "missing" | "guidance"> = {}
): Promise<void> {
  await writeJson(paths.baselineStatus, {
    generatedAt: new Date().toISOString(),
    status,
    message,
    missing: details.missing,
    guidance: details.guidance,
    plan
  });
  await patchBundleManifest(paths.jobDir, (manifest) => ({
    ...manifest,
    baseline: {
      ...manifest.baseline,
      available: false,
      message
    }
  }));
}

async function readJson<T>(filePath: string): Promise<T> {
  const source = await readFile(filePath, "utf8");
  return JSON.parse(source) as T;
}

async function writeJson(filePath: string, value: unknown): Promise<void> {
  await mkdir(path.dirname(filePath), { recursive: true });
  await writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

async function directoryExists(dirPath: string): Promise<boolean> {
  try {
    return (await stat(dirPath)).isDirectory();
  } catch {
    return false;
  }
}

async function fileExistsWithContent(filePath: string): Promise<boolean> {
  try {
    return (await stat(filePath)).size > 0;
  } catch {
    return false;
  }
}

async function renameIfExists(from: string, to: string): Promise<boolean> {
  if (!(await directoryExists(from))) return false;
  await mkdir(path.dirname(to), { recursive: true });
  await rename(from, to);
  return true;
}

function formatSeconds(value: number): string {
  if (!Number.isFinite(value)) return "0";
  const rounded = Math.round(Math.max(0, value) * 1000) / 1000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

function runCommand(command: string, args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: ["ignore", "ignore", "pipe"] });
    let stderr = "";
    child.stderr.on("data", (chunk: Buffer) => {
      stderr += chunk.toString("utf8");
    });
    child.once("error", reject);
    child.once("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${command} exited with code ${code}: ${stderr}`));
    });
  });
}

function sanitizeFileName(fileName: string): string {
  return path.basename(fileName).replace(/[^a-zA-Z0-9._-]/g, "_") || "deck";
}

function createJobId(sourceName: string): string {
  const base = sourceName.replace(/\.[^.]+$/, "").replace(/[^a-zA-Z0-9_-]/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
  return `${base || "deck"}-${new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 14)}`;
}

function parsePositiveNumber(value: unknown): number | undefined {
  if (value === undefined || value === true) return undefined;
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return undefined;
  return number;
}

function parseFlags(args: string[]): Record<string, string | true> {
  const flags: Record<string, string | true> = {};
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (!arg?.startsWith("--")) continue;
    const equalsIndex = arg.indexOf("=");
    if (equalsIndex !== -1) {
      flags[arg.slice(2, equalsIndex)] = arg.slice(equalsIndex + 1);
      continue;
    }
    const key = arg.slice(2);
    const next = args[index + 1];
    if (next && !next.startsWith("--")) {
      flags[key] = next;
      index += 1;
    } else {
      flags[key] = true;
    }
  }
  return flags;
}

function stripTrailingSlash(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

function printBundleSummary(bundle: ProductBundleResult): void {
  console.log("Bundle generated:");
  console.log(`  job: ${bundle.jobId}`);
  console.log(`  source: ${bundle.paths.source}`);
  console.log(`  IR: ${bundle.paths.deckIr}`);
  console.log(`  HTML runtime: ${pathToFileURL(bundle.paths.runtimeHtml).toString()}`);
  console.log(`  rebuilt PPTX: ${bundle.paths.rebuiltPptx}`);
  if (bundle.manifest.artifacts.segmentedPptx) {
    console.log(`  segmented PPTX: ${bundle.paths.segmentedPptx}`);
  }
  if (bundle.manifest.artifacts.hybridPptx) {
    console.log(`  hybrid PPTX: ${bundle.paths.hybridPptx}`);
  }
  if (bundle.manifest.artifacts.staticStepsPptx) {
    console.log(`  static steps PPTX: ${bundle.paths.staticStepsPptx}`);
  }
  if (bundle.manifest.artifacts.segmentPlan) {
    console.log(`  segment plan: ${bundle.paths.segmentPlan}`);
  }
  console.log(`  loss report: ${bundle.paths.lossReport}`);
  console.log(`  manifest: ${bundle.paths.manifest}`);
  console.log(`  video plan: ${bundle.paths.videoPlan}`);
  console.log(`  video deps: ${bundle.videoDependencies.missing.length ? `missing ${bundle.videoDependencies.missing.join(", ")}` : "ready"}`);
}

function printBenchmarkSummary(result: ProductKeyBenchmarkResult): void {
  console.log("Keynote benchmark generated:");
  console.log(`  copied source: ${result.copiedSourcePath}`);
  console.log(`  bundle: ${result.bundleDir}`);
  console.log(`  HTML runtime: ${pathToFileURL(result.bundle.paths.runtimeHtml).toString()}`);
  console.log(`  summary: ${result.summaryPath}`);
  console.log(`  baseline: ${result.summary.baseline.status}`);
  if (result.summary.baseline.reportPath) {
    console.log(`  baseline report: ${result.summary.baseline.reportPath}`);
  }
  const worstFrame = result.summary.baseline.worstFrames[0];
  if (worstFrame) {
    console.log(
      `  worst frame: #${worstFrame.frame} slide ${worstFrame.slideIndex + 1} score ${worstFrame.pixelFidelityScore} diff ${worstFrame.diffPath ?? "n/a"}`
    );
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(new URL(import.meta.url).pathname)) {
  main().catch((error: unknown) => {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  });
}
