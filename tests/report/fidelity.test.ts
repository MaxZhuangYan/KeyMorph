import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

import {
  aggregatePixelFidelityResults,
  comparePngFiles,
  compareRgbaImages,
  decodePng,
  encodePng,
  type RgbaImage
} from "../../src/report/fidelity.ts";

const execFileAsync = promisify(execFile);

describe("pixel fidelity scoring", () => {
  test("scores identical images as perfect fidelity", () => {
    const reference = image(2, 2, [255, 255, 255, 255, 0, 0, 0, 255, 10, 20, 30, 255, 90, 80, 70, 255]);
    const result = compareRgbaImages(reference, reference);

    assert.equal(result.pixelFidelityScore, 1);
    assert.equal(result.mismatchedPixels, 0);
    assert.equal(result.dimensionsMatch, true);
  });

  test("reports mismatch ratio and error for changed pixels", () => {
    const reference = image(2, 1, [255, 255, 255, 255, 0, 0, 0, 255]);
    const actual = image(2, 1, [255, 255, 255, 255, 255, 0, 0, 255]);
    const result = compareRgbaImages(reference, actual, { threshold: 0.01 });

    assert.equal(result.totalPixels, 2);
    assert.equal(result.mismatchedPixels, 1);
    assert.equal(result.mismatchRatio, 0.5);
    assert.ok(result.pixelFidelityScore < 1);
    assert.ok(result.meanAbsoluteError > 0);
  });

  test("encodes and decodes RGBA PNG images for file-based comparisons", () => {
    const source = image(2, 1, [12, 34, 56, 255, 200, 180, 160, 128]);
    const decoded = decodePng(encodePng(source));

    assert.equal(decoded.width, source.width);
    assert.equal(decoded.height, source.height);
    assert.deepEqual(Array.from(decoded.data), Array.from(source.data));
    assert.equal(compareRgbaImages(source, decoded).pixelFidelityScore, 1);
  });

  test("writes a PNG fidelity report through the CLI", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-fidelity-cli-"));
    const referencePath = path.join(dir, "reference.png");
    const actualPath = path.join(dir, "actual.png");
    const reportPath = path.join(dir, "fidelity.json");

    await writeFile(referencePath, encodePng(image(1, 1, [255, 255, 255, 255])));
    await writeFile(actualPath, encodePng(image(1, 1, [0, 0, 0, 255])));
    await execFileAsync(process.execPath, ["--experimental-transform-types", "src/cli.ts", "png-fidelity", referencePath, actualPath, reportPath], {
      cwd: path.resolve(path.dirname(new URL(import.meta.url).pathname), "../..")
    });

    const report = JSON.parse(await readFile(reportPath, "utf8"));
    assert.equal(report.totalPixels, 1);
    assert.equal(report.mismatchedPixels, 1);
    assert.ok(report.pixelFidelityScore < 1);
  });

  test("writes optional PNG diff output without external pixel dependencies", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "keymorph-fidelity-diff-"));
    const referencePath = path.join(dir, "reference.png");
    const actualPath = path.join(dir, "actual.png");
    const diffPath = path.join(dir, "diff.png");

    await writeFile(referencePath, encodePng(image(2, 1, [255, 255, 255, 255, 0, 0, 0, 255])));
    await writeFile(actualPath, encodePng(image(2, 1, [255, 255, 255, 255, 255, 0, 0, 255])));

    const result = await comparePngFiles(referencePath, actualPath, { threshold: 0.01, diffPath });
    const diff = decodePng(await readFile(diffPath));

    assert.equal(result.diffPath, diffPath);
    assert.equal(result.mismatchedPixels, 1);
    assert.equal(diff.width, 2);
    assert.equal(diff.height, 1);
    assert.deepEqual(Array.from(diff.data.slice(4, 8)), [255, 40, 40, 255]);
  });

  test("aggregates frame fidelity results with pass and worst-item metrics", () => {
    const perfect = compareRgbaImages(image(1, 1, [255, 255, 255, 255]), image(1, 1, [255, 255, 255, 255]));
    const changed = compareRgbaImages(image(1, 1, [255, 255, 255, 255]), image(1, 1, [0, 0, 0, 255]), {
      threshold: 0
    });

    const aggregate = aggregatePixelFidelityResults([perfect, changed], { passThreshold: 0.9 });

    assert.equal(aggregate.count, 2);
    assert.equal(aggregate.totalPixels, 2);
    assert.equal(aggregate.mismatchedPixels, 1);
    assert.equal(aggregate.matchedItems, 1);
    assert.equal(aggregate.mismatchedItems, 1);
    assert.equal(aggregate.passingItems, 1);
    assert.equal(aggregate.failingItems, 1);
    assert.equal(aggregate.mismatchRatio, 0.5);
    assert.equal(aggregate.meanMismatchRatio, 0.5);
    assert.equal(aggregate.maxMismatchRatio, 1);
    assert.equal(aggregate.bestIndex, 0);
    assert.equal(aggregate.worstIndex, 1);
  });
});

function image(width: number, height: number, data: number[]): RgbaImage {
  return { width, height, data: new Uint8Array(data) };
}
