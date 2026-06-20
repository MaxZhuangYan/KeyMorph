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
import {
  createProductApiResponse,
  createProductBundle,
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
      keynoteEndpoint: string;
    };

    assert.equal(response.previewUrl, "/demo/out/jobs/api-job/runtime.html");
    assert.equal(response.manifestUrl, "/demo/out/jobs/api-job/manifest.json");
    assert.equal(response.downloads.source, "/demo/out/jobs/api-job/source.ir.json");
    assert.equal(response.downloads.videoPlan, "/demo/out/jobs/api-job/video-plan.json");
    assert.equal(response.downloads.videoStatus, "/demo/out/jobs/api-job/video-status.json");
    assert.equal(response.videoEndpoint, "/api/jobs/api-job/video");
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

  test("uses Keynote native HTML export as the main runtime when automation is allowed", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-product-keynote-html-"));
    const inputPath = path.join(dir, "source.key");
    const bundleDir = path.join(dir, "bundle");
    await writeFile(inputPath, "stub keynote package", "utf8");

    const bundle = await createProductBundle(inputPath, bundleDir, {
      jobId: "keynote-html-job",
      allowKeynoteAutomation: true,
      keynoteBridgeExport: async (_keynotePath, pptxPath) => {
        await exportIrToPptx(createDemoDeck(), pptxPath);
      },
      keynoteHtmlExport: async (_keynotePath, outputDir) => {
        await mkdir(outputDir, { recursive: true });
        await writeFile(path.join(outputDir, "index.html"), "<!doctype html><title>Keynote Native Runtime</title>", "utf8");
      }
    });

    const manifest = JSON.parse(await readFile(path.join(bundleDir, "manifest.json"), "utf8")) as ProductBundleManifest;

    assert.equal(bundle.sourceKind, "keynote");
    assert.equal(manifest.runtime.mode, "keynote-html");
    assert.equal(manifest.runtime.fidelity, "keynote-native");
    assert.equal(manifest.artifacts.runtimeHtml, "runtime.html");
    assert.equal(manifest.artifacts.irRuntimeHtml, "runtime-ir.html");
    assert.equal(manifest.artifacts.keynoteHtml, "keynote-html/index.html");
    const runtimeHtml = await readFile(path.join(bundleDir, "runtime.html"), "utf8");
    assert.match(runtimeHtml, /location\.replace\("keynote-html\/index\.html"\)/);
    assert.doesNotMatch(runtimeHtml, /<iframe/);
    assert.match(await readFile(path.join(bundleDir, "runtime-ir.html"), "utf8"), /window\.__KEYMORPH_DECK__/);
    assert.match(await readFile(path.join(bundleDir, "keynote-html", "index.html"), "utf8"), /Keynote Native Runtime/);
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
  });
});
