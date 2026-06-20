import { copyFile, cp, mkdir, readFile, rename, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { createDemoDeck } from "./demo/createDemoDeck.ts";
import { validateIR, type DeckIR } from "./ir/index.ts";
import { renderHtmlDocument } from "./runtime/index.ts";
import { createLossReport, scoreConversion, type ConversionLossReport } from "./report/index.ts";
import { comparePngFiles } from "./report/fidelity.ts";
import { exportIrToPptx, parsePptxToIr } from "./pptx/index.ts";
import { exportIrToKeynote, parseKeynoteToIr } from "./keynote/index.ts";
import {
  createVideoExportPlan,
  createVideoFrameFidelityReport,
  describeVideoDependencies,
  exportIrToVideo,
  type VideoDependencyStatus,
  type VideoExportOptions,
  type VideoFrameFidelityReport,
  type VideoExportPlan,
  VideoExportDependencyError
} from "./video/index.ts";

export type ProductInputKind = "pptx" | "keynote" | "ir";

export interface ProductBundleOptions {
  sourceName?: string;
  jobId?: string;
  allowKeynoteAutomation?: boolean;
  keynoteAutomationTimeoutMs?: number;
  video?: Pick<VideoExportOptions, "fps" | "scale" | "ffmpegPath">;
}

export interface ProductBundlePaths {
  jobDir: string;
  source: string;
  deckIr: string;
  runtimeHtml: string;
  rebuiltPptx: string;
  lossReport: string;
  manifest: string;
  videoPlan: string;
  videoStatus: string;
  rebuiltKeynote: string;
  renderVideo: string;
  framesLatest: string;
  frameFidelity: string;
  frameDiffs: string;
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
    rebuiltPptx: string;
    lossReport: string;
    videoPlan: string;
    videoStatus: string;
    rebuiltKeynote: string | null;
    renderVideo: string | null;
    frameFidelity: string | null;
  };
  report: {
    fidelityScore: number;
    riskLevel: ConversionLossReport["riskLevel"];
    animationLostCount: number;
    degradedAnimationCount: number;
    uncertainMappingCount: number;
    recommendedFixes: string[];
  };
  video: ProductVideoSummary;
  keynote: ProductDeferredExportSummary;
}

export interface ProductVideoSummary {
  plan: Omit<VideoExportPlan, "frames">;
  dependencies: VideoDependencyStatus;
  endpoint: string | null;
  statusPath: string;
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

  const deck = await inputToIr(paths.source, sourceName, sourceKind, {
    allowKeynoteAutomation: options.allowKeynoteAutomation,
    keynoteAutomationTimeoutMs: options.keynoteAutomationTimeoutMs
  });
  const validation = validateIR(deck);
  if (!validation.valid) {
    throw new Error(`Converted IR is invalid: ${validation.errors.map((error) => `${error.path} ${error.message}`).join("; ")}`);
  }

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

  await writeJson(paths.deckIr, deck);
  await writeFile(paths.runtimeHtml, renderHtmlDocument(deck), "utf8");
  await exportIrToPptx(deck, paths.rebuiltPptx);
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
    videoPlan: summarizeVideoPlan(bundle.videoPlan),
    videoDependencies: bundle.videoDependencies,
    videoEndpoint: `/api/jobs/${bundle.jobId}/video`,
    keynoteEndpoint: `/api/jobs/${bundle.jobId}/keynote`,
    downloads: {
      source: `${base}/${encodeURIComponent(bundle.sourceName)}`,
      html: `${base}/runtime.html`,
      ir: `${base}/deck.ir.json`,
      pptx: `${base}/rebuilt.pptx`,
      key: null,
      report: `${base}/loss-report.json`,
      manifest: `${base}/manifest.json`,
      videoPlan: `${base}/video-plan.json`,
      videoStatus: `${base}/video-status.json`,
      video: null
    },
    keynoteAvailable: null,
    keynoteMessage: bundle.manifest.keynote.message
  };
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
        "Usage: keymorph <demo|convert|inspect|bundle-keynote|bundle-video|pptx-to-ir|key-to-ir|ir-to-html|ir-to-pptx|ir-to-key|ir-to-video|ir-report|png-fidelity> [input] [output]"
      );
  }
}

async function inputToIr(
  sourcePath: string,
  fileName: string,
  sourceKind = detectInputKind(fileName),
  options: Pick<ProductBundleOptions, "allowKeynoteAutomation" | "keynoteAutomationTimeoutMs"> = {}
): Promise<DeckIR> {
  if (sourceKind === "pptx") return parsePptxToIr(sourcePath);
  if (sourceKind === "ir") return readJson<DeckIR>(sourcePath);
  if (sourceKind === "keynote") {
    return parseKeynoteToIr(sourcePath, {
      workDir: path.dirname(sourcePath),
      allowAutomation: options.allowKeynoteAutomation,
      automationTimeoutMs: options.keynoteAutomationTimeoutMs
    });
  }
  throw new Error("Unsupported file type. Use .pptx, .key, or .ir.json.");
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
    rebuiltPptx: path.join(jobDir, "rebuilt.pptx"),
    lossReport: path.join(jobDir, "loss-report.json"),
    manifest: path.join(jobDir, "manifest.json"),
    videoPlan: path.join(jobDir, "video-plan.json"),
    videoStatus: path.join(jobDir, "video-status.json"),
    rebuiltKeynote: path.join(jobDir, "rebuilt.key"),
    renderVideo: path.join(jobDir, "render.mp4"),
    framesLatest: path.join(jobDir, "frames", "latest"),
    frameFidelity: path.join(jobDir, "frame-fidelity.json"),
    frameDiffs: path.join(jobDir, "frame-diffs")
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
      rebuiltPptx: "rebuilt.pptx",
      lossReport: "loss-report.json",
      videoPlan: "video-plan.json",
      videoStatus: "video-status.json",
      rebuiltKeynote: null,
      renderVideo: null,
      frameFidelity: null
    },
    report: {
      fidelityScore: input.lossReport.fidelityScore,
      riskLevel: input.lossReport.riskLevel,
      animationLostCount: input.lossReport.animationLostCount,
      degradedAnimationCount: input.lossReport.degradedAnimationCount,
      uncertainMappingCount: input.lossReport.uncertainMappingCount,
      recommendedFixes: input.lossReport.recommendedFixes
    },
    video: {
      plan: summarizeVideoPlan(input.videoPlan),
      dependencies: input.videoDependencies,
      endpoint: null,
      statusPath: "video-status.json"
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

async function renameIfExists(from: string, to: string): Promise<boolean> {
  if (!(await directoryExists(from))) return false;
  await mkdir(path.dirname(to), { recursive: true });
  await rename(from, to);
  return true;
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
  console.log(`  loss report: ${bundle.paths.lossReport}`);
  console.log(`  manifest: ${bundle.paths.manifest}`);
  console.log(`  video plan: ${bundle.paths.videoPlan}`);
  console.log(`  video deps: ${bundle.videoDependencies.missing.length ? `missing ${bundle.videoDependencies.missing.join(", ")}` : "ready"}`);
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
