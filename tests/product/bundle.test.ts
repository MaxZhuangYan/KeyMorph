import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdir, mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { createDemoDeck } from "../../src/demo/createDemoDeck.ts";
import { validateIR } from "../../src/ir/index.ts";
import { exportIrToPptx } from "../../src/pptx/index.ts";
import { encodePng, type RgbaImage } from "../../src/report/fidelity.ts";
import {
  createProductApiResponse,
  createProductBundle,
  exportProductBundleBaseline,
  exportProductBundleKeynote,
  inspectProductInput,
  type ProductBundleManifest
} from "../../src/cli.ts";

const execFileAsync = promisify(execFile);

describe("product bundle workflow", () => {
  test("creates the concrete local bundle artifacts from IR input", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-bundle-"));
    const inputPath = path.join(dir, "source.ir.json");
    const bundleDir = path.join(dir, "bundle");
    await writeFile(inputPath, `${JSON.stringify(createDemoDeck(), null, 2)}\n`, "utf8");

    const bundle = await createProductBundle(inputPath, bundleDir, { jobId: "fixture-job" });
    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;
    const deck = JSON.parse(await readFile(path.join(bundleDir, "deck.ir.json"), "utf8"));
    const report = JSON.parse(await readFile(path.join(bundleDir, "loss-report.json"), "utf8"));
    const videoPlan = JSON.parse(await readFile(path.join(bundleDir, "video-plan.json"), "utf8"));
    const videoStatus = JSON.parse(await readFile(path.join(bundleDir, "video-status.json"), "utf8"));

    assert.equal(bundle.jobId, "fixture-job");
    assert.equal(bundle.sourceKind, "ir");
    assert.equal(validateIR(deck).valid, true);
    assert.equal(manifest.artifacts.source, "source.ir.json");
    assert.equal(manifest.artifacts.deckIr, "deck.ir.json");
    assert.equal(manifest.artifacts.runtimeHtml, "runtime.html");
    assert.equal(manifest.artifacts.rebuiltPptx, "rebuilt.pptx");
    assert.equal(manifest.artifacts.lossReport, "loss-report.json");
    assert.equal(manifest.artifacts.videoPlan, "video-plan.json");
    assert.equal(manifest.artifacts.videoStatus, "video-status.json");
    assert.equal(manifest.artifacts.baselineStatus, "baseline-status.json");
    assert.equal(manifest.artifacts.baselineFrameFidelity, null);
    assert.equal(manifest.slideCount, createDemoDeck().deck.slides.length);
    assert.equal(report.generatedAt.length > 0, true);
    assert.equal(videoPlan.frames.length, bundle.videoPlan.totalFrames);
    assert.deepEqual(videoStatus.dependencies.available, bundle.videoDependencies.available);
    assert.match(await readFile(path.join(bundleDir, "runtime.html"), "utf8"), /window\.__KEYMORPH_DECK__/);
    assert.ok((await stat(path.join(bundleDir, "rebuilt.pptx"))).size > 0);
  });

  test("builds the local API response from the same bundle manifest", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-api-"));
    const inputPath = path.join(dir, "source.ir.json");
    const bundleDir = path.join(dir, "bundle");
    await writeFile(inputPath, `${JSON.stringify(createDemoDeck(), null, 2)}\n`, "utf8");

    const bundle = await createProductBundle(inputPath, bundleDir, { jobId: "api-job" });
    const response = createProductApiResponse(bundle, "/demo/out/jobs/api-job") as {
      previewUrl: string;
      manifestUrl: string;
      downloads: Record<string, string | null>;
      videoEndpoint: string;
      baselineEndpoint: string;
      keynoteEndpoint: string;
    };

    assert.equal(response.previewUrl, "/demo/out/jobs/api-job/runtime.html");
    assert.equal(response.manifestUrl, "/demo/out/jobs/api-job/manifest.json");
    assert.equal(response.downloads.source, "/demo/out/jobs/api-job/source.ir.json");
    assert.equal(response.downloads.videoPlan, "/demo/out/jobs/api-job/video-plan.json");
    assert.equal(response.downloads.videoStatus, "/demo/out/jobs/api-job/video-status.json");
    assert.equal(response.downloads.baselineStatus, "/demo/out/jobs/api-job/baseline-status.json");
    assert.equal(response.downloads.baselineFidelity, null);
    assert.equal(response.videoEndpoint, "/api/jobs/api-job/video");
    assert.equal(response.baselineEndpoint, "/api/jobs/api-job/baseline");
    assert.equal(response.keynoteEndpoint, "/api/jobs/api-job/keynote");
  });

  test("inspects input without launching Keynote automation", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-inspect-"));
    const inputPath = path.join(dir, "source.ir.json");
    await writeFile(inputPath, `${JSON.stringify(createDemoDeck(), null, 2)}\n`, "utf8");

    const result = await inspectProductInput(inputPath);

    assert.equal(result.sourceKind, "ir");
    assert.equal(result.validation.valid, true);
    assert.equal(result.slideCount, 2);
    assert.equal(result.videoPlan.totalFrames, 159);
  });

  test("reports on-demand Keynote export as unavailable without launching Keynote by default", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-keynote-"));
    const inputPath = path.join(dir, "source.ir.json");
    const bundleDir = path.join(dir, "bundle");
    await writeFile(inputPath, `${JSON.stringify(createDemoDeck(), null, 2)}\n`, "utf8");
    await createProductBundle(inputPath, bundleDir, { jobId: "keynote-job" });

    const result = await exportProductBundleKeynote(bundleDir);
    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;

    assert.equal(result.status, "dependency-missing");
    assert.match(result.message, /Keynote GUI automation is disabled by default|Keynote conversion requires macOS/);
    assert.equal(manifest.keynote.available, false);
  });

  test("uses Keynote movie export as the main runtime when automation is allowed", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-keynote-movie-"));
    const inputPath = path.join(dir, "source.key");
    const bundleDir = path.join(dir, "bundle");
    await writeFile(inputPath, "stub keynote package", "utf8");

    const bundle = await createProductBundle(inputPath, bundleDir, {
      jobId: "keynote-movie-job",
      allowKeynoteAutomation: true,
      keynoteBridgeExport: async (_keynotePath, pptxPath) => {
        await exportIrToPptx(createDemoDeck(), pptxPath);
      },
      keynoteMovieExport: async (_keynotePath, moviePath) => {
        await writeFile(moviePath, "stub movie", "utf8");
      },
      keynoteHtmlExport: async (_keynotePath, outputDir) => {
        await mkdir(outputDir, { recursive: true });
        await writeFile(path.join(outputDir, "index.html"), "<!doctype html><title>Keynote Native Runtime</title>", "utf8");
      }
    });

    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;

    assert.equal(bundle.sourceKind, "keynote");
    assert.equal(manifest.runtime.mode, "keynote-movie");
    assert.equal(manifest.runtime.fidelity, "keynote-rendered-video");
    assert.equal(manifest.artifacts.runtimeHtml, "runtime.html");
    assert.equal(manifest.artifacts.irRuntimeHtml, "runtime-ir.html");
    assert.equal(manifest.artifacts.keynoteHtml, null);
    assert.equal(manifest.artifacts.keynoteMovie, "keynote-movie.m4v");
    assert.equal(manifest.artifacts.renderVideo, "keynote-movie.m4v");
    const runtimeHtml = await readFile(path.join(bundleDir, "runtime.html"), "utf8");
    assert.match(runtimeHtml, /<video src="keynote-movie\.m4v"/);
    assert.match(await readFile(path.join(bundleDir, "runtime-ir.html"), "utf8"), /window\.__KEYMORPH_DECK__/);
    assert.match(await readFile(path.join(bundleDir, "keynote-movie.m4v"), "utf8"), /stub movie/);
  });

  test("materializes used native Keynote package images into bundle-local runtime assets", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-native-assets-"));
    const keyPath = path.join(dir, "native.key");
    const bundleDir = path.join(dir, "bundle");
    await mkdir(path.join(keyPath, "Data"), { recursive: true });
    await mkdir(path.join(keyPath, "Index"), { recursive: true });
    const largeImage = concat([pngBytes(1300, 1300), new Uint8Array(2 * 1024 * 1024 + 1)]);
    await writeFile(path.join(keyPath, "Data", "hero large.png"), largeImage);
    await writeFile(path.join(keyPath, "Index", "Slide-1.iwa"), concat([protoString("Data/hero large.png")]));

    await createProductBundle(keyPath, bundleDir, { jobId: "native-assets-job" });

    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;
    const deck = JSON.parse(await readFile(path.join(bundleDir, "deck.ir.json"), "utf8"));
    const runtimeHtml = await readFile(path.join(bundleDir, "runtime.html"), "utf8");
    const materializedAsset = deck.deck.assets.find((asset) => asset.name === "hero large.png");

    assert.equal(manifest.artifacts.nativeAssets, "assets/native");
    assert.match(materializedAsset.uri, /^assets\/native\/hero-large-[a-f0-9]{12}\.png$/);
    assert.equal(materializedAsset.metadata.nativeMaterializedAsset, true);
    assert.match(runtimeHtml, /assets\/native\/hero-large-[a-f0-9]{12}\.png/);
    assert.equal((await stat(path.join(bundleDir, materializedAsset.uri))).size, largeImage.byteLength);
    assert.equal(deck.conversion.messages.some((message) => message.code === "keynote-native-assets-materialized"), true);
  });

  test("falls back to Keynote HTML runtime when movie export is unavailable", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-keynote-html-"));
    const inputPath = path.join(dir, "source.key");
    const bundleDir = path.join(dir, "bundle");
    await writeFile(inputPath, "stub keynote package", "utf8");

    await createProductBundle(inputPath, bundleDir, {
      jobId: "keynote-html-job",
      allowKeynoteAutomation: true,
      keynoteBridgeExport: async (_keynotePath, pptxPath) => {
        await exportIrToPptx(createDemoDeck(), pptxPath);
      },
      keynoteMovieExport: async () => {
        throw new Error("Synthetic movie export failure");
      },
      keynoteHtmlExport: async (_keynotePath, outputDir) => {
        await mkdir(outputDir, { recursive: true });
        await writeFile(path.join(outputDir, "index.html"), "<!doctype html><title>Keynote Native Runtime</title>", "utf8");
      }
    });

    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;
    const runtimeHtml = await readFile(path.join(bundleDir, "runtime.html"), "utf8");

    assert.equal(manifest.runtime.mode, "keynote-html");
    assert.equal(manifest.artifacts.keynoteMovie, null);
    assert.equal(manifest.artifacts.keynoteHtml, "keynote-html/index.html");
    assert.match(runtimeHtml, /location\.replace\("keynote-html\/index\.html"\)/);
    assert.match(await readFile(path.join(bundleDir, "keynote-html", "index.html"), "utf8"), /Keynote Native Runtime/);
  });

  test("keeps movie runtime when exporter reports an error after writing a movie", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-keynote-movie-late-error-"));
    const inputPath = path.join(dir, "source.key");
    const bundleDir = path.join(dir, "bundle");
    await writeFile(inputPath, "stub keynote package", "utf8");

    await createProductBundle(inputPath, bundleDir, {
      jobId: "keynote-movie-late-error-job",
      allowKeynoteAutomation: true,
      keynoteBridgeExport: async (_keynotePath, pptxPath) => {
        await exportIrToPptx(createDemoDeck(), pptxPath);
      },
      keynoteMovieExport: async (_keynotePath, moviePath) => {
        await writeFile(moviePath, "movie completed before timeout", "utf8");
        throw new Error("Synthetic late AppleEvent timeout");
      },
      keynoteHtmlExport: async () => {
        throw new Error("HTML fallback should not run");
      }
    });

    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;
    assert.equal(manifest.runtime.mode, "keynote-movie");
    assert.equal(manifest.artifacts.keynoteMovie, "keynote-movie.m4v");
    assert.match(await readFile(path.join(bundleDir, "keynote-movie.m4v"), "utf8"), /movie completed/);
  });

  test("exports a Keynote golden baseline from the bundle copy and compares runtime frames", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-baseline-"));
    const inputPath = path.join(dir, "original-source.key");
    const bundleDir = path.join(dir, "bundle");
    await writeFile(inputPath, "stub keynote package", "utf8");

    await createProductBundle(inputPath, bundleDir, {
      jobId: "baseline-job",
      allowKeynoteAutomation: true,
      keynoteBridgeExport: async (_keynotePath, pptxPath) => {
        await exportIrToPptx(createDemoDeck(), pptxPath);
      },
      keynoteMovieExport: async (_keynotePath, moviePath) => {
        await writeFile(moviePath, "stub movie", "utf8");
      }
    });

    let keynoteFrameSource = "";
    const result = await exportProductBundleBaseline(bundleDir, {
      allowKeynoteAutomation: true,
      video: { fps: 1, scale: 1 },
      keynoteFrameExport: async (keynotePath, framesDir, plan) => {
        keynoteFrameSource = keynotePath;
        await mkdir(framesDir, { recursive: true });
        for (const frame of plan.frames) {
          await writeFile(path.join(framesDir, frame.outputPath), encodePng(image(1, 1, [255, 255, 255, 255])));
        }
      },
      runtimeFrameExport: async (_deck, framesDir, plan) => {
        await mkdir(framesDir, { recursive: true });
        for (const frame of plan.frames) {
          await writeFile(
            path.join(framesDir, frame.outputPath),
            encodePng(frame.frame === 1 ? image(1, 1, [0, 0, 0, 255]) : image(1, 1, [255, 255, 255, 255]))
          );
        }
      }
    });

    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;
    const status = JSON.parse(await readFile(path.join(bundleDir, "baseline-status.json"), "utf8"));
    const report = JSON.parse(await readFile(path.join(bundleDir, "baseline-fidelity.json"), "utf8"));

    assert.equal(result.status, "ready");
    assert.equal(keynoteFrameSource, path.join(bundleDir, "original-source.key"));
    assert.notEqual(keynoteFrameSource, inputPath);
    assert.equal(manifest.baseline.available, true);
    assert.equal(manifest.artifacts.baselineFrames, "frames/baseline");
    assert.equal(manifest.artifacts.baselineActualFrames, "frames/keymorph-baseline");
    assert.equal(manifest.artifacts.baselineFrameFidelity, "baseline-fidelity.json");
    assert.equal(manifest.artifacts.baselineFrameDiffs, "baseline-diffs");
    assert.equal(status.status, "ready");
    assert.equal(report.summary.mismatchedFrames, 1);
  });

  test("reports Keynote baseline as unsupported for non-Keynote bundles", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-baseline-unsupported-"));
    const inputPath = path.join(dir, "source.ir.json");
    const bundleDir = path.join(dir, "bundle");
    await writeFile(inputPath, `${JSON.stringify(createDemoDeck(), null, 2)}\n`, "utf8");
    await createProductBundle(inputPath, bundleDir, { jobId: "baseline-unsupported-job" });

    const result = await exportProductBundleBaseline(bundleDir, { video: { fps: 1, scale: 1 } });
    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;

    assert.equal(result.status, "unsupported");
    assert.match(result.message, /original \.key source/);
    assert.equal(manifest.baseline.available, false);
  });

  test("CLI convert writes the same bundle shape", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-cli-"));
    const inputPath = path.join(dir, "source.ir.json");
    const bundleDir = path.join(dir, "cli-bundle");
    await writeFile(inputPath, `${JSON.stringify(createDemoDeck(), null, 2)}\n`, "utf8");

    const result = await execFileAsync(process.execPath, ["--experimental-transform-types", "src/cli.ts", "convert", inputPath, bundleDir], {
      cwd: path.resolve("."),
      timeout: 30000
    });
    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;

    assert.match(result.stdout, /Bundle generated:/);
    assert.equal(manifest.sourceKind, "ir");
    assert.equal(manifest.artifacts.runtimeHtml, "runtime.html");
  });

  test("dev UI includes Simplified and Traditional Chinese language options", async () => {
    const source = await readFile(path.resolve("scripts/dev.mjs"), "utf8");

    assert.match(source, /<option value="zh-Hans">简体中文<\/option>/);
    assert.match(source, /<option value="zh-Hant">繁體中文<\/option>/);
    assert.match(source, /把演示文稿拖到这里/);
    assert.match(source, /把簡報拖到這裡/);
    assert.match(source, /运行 Keynote 基准对比/);
    assert.match(source, /執行 Keynote 基準對比/);
  });
});

function image(width: number, height: number, data: number[]): RgbaImage {
  return { width, height, data: new Uint8Array(data) };
}

function protoString(value: string, fieldNumber = 1): Uint8Array {
  const encoded = new TextEncoder().encode(value);
  return concat([varint((fieldNumber << 3) | 2), varint(encoded.length), encoded]);
}

function varint(value: number): Uint8Array {
  const bytes: number[] = [];
  let remaining = value;
  while (remaining > 0x7f) {
    bytes.push((remaining & 0x7f) | 0x80);
    remaining = Math.floor(remaining / 128);
  }
  bytes.push(remaining);
  return new Uint8Array(bytes);
}

function concat(chunks: Uint8Array[]): Uint8Array {
  const length = chunks.reduce((total, chunk) => total + chunk.byteLength, 0);
  const out = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return out;
}

function pngBytes(width = 1, height = 1): Uint8Array {
  const bytes = new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82
  ]);
  new DataView(bytes.buffer).setUint32(16, width, false);
  new DataView(bytes.buffer).setUint32(20, height, false);
  return bytes;
}
