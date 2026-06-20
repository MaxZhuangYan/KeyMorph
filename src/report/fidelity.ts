import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { deflateSync, inflateSync } from "node:zlib";

export interface RgbaImage {
  width: number;
  height: number;
  data: Uint8Array | Uint8ClampedArray;
}

export interface PixelFidelityOptions {
  threshold?: number;
  includeAlpha?: boolean;
  diffPath?: string;
}

export interface PixelFidelityResult {
  width: number;
  height: number;
  totalPixels: number;
  comparedPixels: number;
  mismatchedPixels: number;
  missingPixels: number;
  mismatchRatio: number;
  meanAbsoluteError: number;
  rootMeanSquareError: number;
  maxDelta: number;
  threshold: number;
  dimensionsMatch: boolean;
  pixelFidelityScore: number;
  diffPath?: string;
}

export interface RgbaDiffOptions extends PixelFidelityOptions {
  matchOpacity?: number;
}

export function compareRgbaImages(
  reference: RgbaImage,
  actual: RgbaImage,
  options: PixelFidelityOptions = {}
): PixelFidelityResult {
  assertImage(reference, "reference");
  assertImage(actual, "actual");

  const threshold = normalizeThreshold(options.threshold);
  const includeAlpha = options.includeAlpha ?? true;
  const width = Math.max(reference.width, actual.width);
  const height = Math.max(reference.height, actual.height);
  const overlapWidth = Math.min(reference.width, actual.width);
  const overlapHeight = Math.min(reference.height, actual.height);
  const totalPixels = Math.max(1, width * height);
  const comparedPixels = overlapWidth * overlapHeight;
  const missingPixels = totalPixels - comparedPixels;
  const channels = includeAlpha ? 4 : 3;

  let mismatchedPixels = missingPixels;
  let absoluteError = missingPixels * channels;
  let squaredError = missingPixels * channels;
  let maxDelta = missingPixels > 0 ? 1 : 0;

  for (let y = 0; y < overlapHeight; y += 1) {
    for (let x = 0; x < overlapWidth; x += 1) {
      const refOffset = (y * reference.width + x) * 4;
      const actualOffset = (y * actual.width + x) * 4;
      let channelSquared = 0;
      let channelAbsolute = 0;

      for (let channel = 0; channel < channels; channel += 1) {
        const delta = Math.abs(Number(reference.data[refOffset + channel]) - Number(actual.data[actualOffset + channel])) / 255;
        channelAbsolute += delta;
        channelSquared += delta * delta;
      }

      const normalizedDistance = Math.sqrt(channelSquared / channels);
      if (normalizedDistance > threshold) {
        mismatchedPixels += 1;
      }
      absoluteError += channelAbsolute;
      squaredError += channelSquared;
      maxDelta = Math.max(maxDelta, normalizedDistance);
    }
  }

  const denominator = totalPixels * channels;
  const mismatchRatio = mismatchedPixels / totalPixels;
  const meanAbsoluteError = absoluteError / denominator;
  const rootMeanSquareError = Math.sqrt(squaredError / denominator);
  const pixelFidelityScore = Number(
    Math.max(0, Math.min(1, 1 - mismatchRatio * 0.85 - meanAbsoluteError * 0.15)).toFixed(4)
  );

  return {
    width,
    height,
    totalPixels,
    comparedPixels,
    mismatchedPixels,
    missingPixels,
    mismatchRatio: Number(mismatchRatio.toFixed(6)),
    meanAbsoluteError: Number(meanAbsoluteError.toFixed(6)),
    rootMeanSquareError: Number(rootMeanSquareError.toFixed(6)),
    maxDelta: Number(maxDelta.toFixed(6)),
    threshold,
    dimensionsMatch: reference.width === actual.width && reference.height === actual.height,
    pixelFidelityScore
  };
}

export async function comparePngFiles(
  referencePath: string,
  actualPath: string,
  options: PixelFidelityOptions = {}
): Promise<PixelFidelityResult> {
  const [reference, actual] = await Promise.all([readPngFile(referencePath), readPngFile(actualPath)]);
  const result = compareRgbaImages(reference, actual, options);
  if (options.diffPath) {
    await writePngFile(options.diffPath, createRgbaDiffImage(reference, actual, options));
    return { ...result, diffPath: options.diffPath };
  }
  return result;
}

export async function readPngFile(filePath: string): Promise<RgbaImage> {
  return decodePng(await readFile(filePath));
}

export async function writePngFile(filePath: string, image: RgbaImage): Promise<void> {
  await mkdir(path.dirname(filePath), { recursive: true });
  await writeFile(filePath, encodePng(image));
}

export function createRgbaDiffImage(
  reference: RgbaImage,
  actual: RgbaImage,
  options: RgbaDiffOptions = {}
): RgbaImage {
  assertImage(reference, "reference");
  assertImage(actual, "actual");

  const threshold = normalizeThreshold(options.threshold);
  const includeAlpha = options.includeAlpha ?? true;
  const width = Math.max(reference.width, actual.width);
  const height = Math.max(reference.height, actual.height);
  const channels = includeAlpha ? 4 : 3;
  const matchOpacity = Math.max(0, Math.min(1, options.matchOpacity ?? 0.24));
  const data = new Uint8Array(width * height * 4);

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const outputOffset = (y * width + x) * 4;
      const refOffset = pixelOffset(reference, x, y);
      const actualOffset = pixelOffset(actual, x, y);

      if (refOffset === undefined || actualOffset === undefined) {
        data[outputOffset] = 255;
        data[outputOffset + 1] = 0;
        data[outputOffset + 2] = 255;
        data[outputOffset + 3] = 255;
        continue;
      }

      const distance = normalizedPixelDistance(reference.data, actual.data, refOffset, actualOffset, channels);
      if (distance > threshold) {
        data[outputOffset] = 255;
        data[outputOffset + 1] = 40;
        data[outputOffset + 2] = 40;
        data[outputOffset + 3] = 255;
      } else {
        data[outputOffset] = Math.round(Number(actual.data[actualOffset]) * matchOpacity + 255 * (1 - matchOpacity));
        data[outputOffset + 1] = Math.round(Number(actual.data[actualOffset + 1]) * matchOpacity + 255 * (1 - matchOpacity));
        data[outputOffset + 2] = Math.round(Number(actual.data[actualOffset + 2]) * matchOpacity + 255 * (1 - matchOpacity));
        data[outputOffset + 3] = 255;
      }
    }
  }

  return { width, height, data };
}

export function decodePng(input: Uint8Array): RgbaImage {
  const data = input instanceof Buffer ? input : Buffer.from(input);
  assertPngSignature(data);

  let offset = 8;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let interlace = 0;
  const idatChunks: Buffer[] = [];

  while (offset + 12 <= data.length) {
    const length = data.readUInt32BE(offset);
    const type = data.subarray(offset + 4, offset + 8).toString("ascii");
    const chunk = data.subarray(offset + 8, offset + 8 + length);
    offset += 12 + length;

    if (type === "IHDR") {
      width = chunk.readUInt32BE(0);
      height = chunk.readUInt32BE(4);
      bitDepth = chunk[8] ?? 0;
      colorType = chunk[9] ?? 0;
      interlace = chunk[12] ?? 0;
    } else if (type === "IDAT") {
      idatChunks.push(Buffer.from(chunk));
    } else if (type === "IEND") {
      break;
    }
  }

  if (width <= 0 || height <= 0) throw new Error("PNG is missing a valid IHDR chunk.");
  if (bitDepth !== 8) throw new Error(`Only 8-bit PNG files are supported, got bit depth ${bitDepth}.`);
  if (interlace !== 0) throw new Error("Interlaced PNG files are not supported.");

  const channels = channelsForPngColorType(colorType);
  const rowBytes = width * channels;
  const inflated = inflateSync(Buffer.concat(idatChunks));
  const expectedLength = (rowBytes + 1) * height;
  if (inflated.length < expectedLength) {
    throw new Error("PNG IDAT payload is shorter than expected.");
  }

  const rows = unfilterPngRows(inflated, width, height, channels);
  return {
    width,
    height,
    data: pngRowsToRgba(rows, width, height, colorType)
  };
}

export function encodePng(image: RgbaImage): Uint8Array {
  assertImage(image, "image");
  const scanlineLength = image.width * 4 + 1;
  const raw = Buffer.alloc(scanlineLength * image.height);
  for (let y = 0; y < image.height; y += 1) {
    const rowStart = y * scanlineLength;
    raw[rowStart] = 0;
    const sourceStart = y * image.width * 4;
    raw.set(image.data.subarray(sourceStart, sourceStart + image.width * 4), rowStart + 1);
  }

  const chunks = [
    pngChunk("IHDR", concatBuffers([u32be(image.width), u32be(image.height), Buffer.from([8, 6, 0, 0, 0])])),
    pngChunk("IDAT", deflateSync(raw)),
    pngChunk("IEND", Buffer.alloc(0))
  ];

  return concatBuffers([PNG_SIGNATURE, ...chunks]);
}

function unfilterPngRows(data: Buffer, width: number, height: number, channels: number): Buffer {
  const rowBytes = width * channels;
  const output = Buffer.alloc(rowBytes * height);
  let inputOffset = 0;

  for (let y = 0; y < height; y += 1) {
    const filter = data[inputOffset];
    inputOffset += 1;
    const rowStart = y * rowBytes;
    const previousRowStart = rowStart - rowBytes;

    for (let x = 0; x < rowBytes; x += 1) {
      const raw = data[inputOffset + x] ?? 0;
      const left = x >= channels ? output[rowStart + x - channels] ?? 0 : 0;
      const up = y > 0 ? output[previousRowStart + x] ?? 0 : 0;
      const upLeft = y > 0 && x >= channels ? output[previousRowStart + x - channels] ?? 0 : 0;

      output[rowStart + x] = (raw + filterValue(filter ?? 0, left, up, upLeft)) & 0xff;
    }
    inputOffset += rowBytes;
  }

  return output;
}

function pngRowsToRgba(rows: Buffer, width: number, height: number, colorType: number): Uint8Array {
  const rgba = new Uint8Array(width * height * 4);
  const channels = channelsForPngColorType(colorType);

  for (let pixel = 0; pixel < width * height; pixel += 1) {
    const source = pixel * channels;
    const target = pixel * 4;
    if (colorType === 0) {
      rgba[target] = rows[source] ?? 0;
      rgba[target + 1] = rows[source] ?? 0;
      rgba[target + 2] = rows[source] ?? 0;
      rgba[target + 3] = 255;
    } else if (colorType === 2) {
      rgba[target] = rows[source] ?? 0;
      rgba[target + 1] = rows[source + 1] ?? 0;
      rgba[target + 2] = rows[source + 2] ?? 0;
      rgba[target + 3] = 255;
    } else if (colorType === 4) {
      rgba[target] = rows[source] ?? 0;
      rgba[target + 1] = rows[source] ?? 0;
      rgba[target + 2] = rows[source] ?? 0;
      rgba[target + 3] = rows[source + 1] ?? 255;
    } else {
      rgba[target] = rows[source] ?? 0;
      rgba[target + 1] = rows[source + 1] ?? 0;
      rgba[target + 2] = rows[source + 2] ?? 0;
      rgba[target + 3] = rows[source + 3] ?? 255;
    }
  }

  return rgba;
}

function filterValue(filter: number, left: number, up: number, upLeft: number): number {
  if (filter === 0) return 0;
  if (filter === 1) return left;
  if (filter === 2) return up;
  if (filter === 3) return Math.floor((left + up) / 2);
  if (filter === 4) return paeth(left, up, upLeft);
  throw new Error(`Unsupported PNG filter type ${filter}.`);
}

function paeth(left: number, up: number, upLeft: number): number {
  const estimate = left + up - upLeft;
  const leftDistance = Math.abs(estimate - left);
  const upDistance = Math.abs(estimate - up);
  const upLeftDistance = Math.abs(estimate - upLeft);
  if (leftDistance <= upDistance && leftDistance <= upLeftDistance) return left;
  if (upDistance <= upLeftDistance) return up;
  return upLeft;
}

function channelsForPngColorType(colorType: number): number {
  if (colorType === 0) return 1;
  if (colorType === 2) return 3;
  if (colorType === 4) return 2;
  if (colorType === 6) return 4;
  throw new Error(`Unsupported PNG color type ${colorType}.`);
}

function assertImage(image: RgbaImage, name: string): void {
  if (!Number.isInteger(image.width) || image.width <= 0 || !Number.isInteger(image.height) || image.height <= 0) {
    throw new Error(`${name} image must have a positive integer width and height.`);
  }
  if (image.data.length !== image.width * image.height * 4) {
    throw new Error(`${name} image data must be RGBA with width * height * 4 bytes.`);
  }
}

function pixelOffset(image: RgbaImage, x: number, y: number): number | undefined {
  if (x < 0 || y < 0 || x >= image.width || y >= image.height) return undefined;
  return (y * image.width + x) * 4;
}

function normalizedPixelDistance(
  referenceData: Uint8Array | Uint8ClampedArray,
  actualData: Uint8Array | Uint8ClampedArray,
  referenceOffset: number,
  actualOffset: number,
  channels: number
): number {
  let channelSquared = 0;
  for (let channel = 0; channel < channels; channel += 1) {
    const delta = Math.abs(Number(referenceData[referenceOffset + channel]) - Number(actualData[actualOffset + channel])) / 255;
    channelSquared += delta * delta;
  }
  return Math.sqrt(channelSquared / channels);
}

function assertPngSignature(data: Buffer): void {
  if (data.length < PNG_SIGNATURE.length || !data.subarray(0, PNG_SIGNATURE.length).equals(PNG_SIGNATURE)) {
    throw new Error("Input is not a PNG file.");
  }
}

function normalizeThreshold(value: number | undefined): number {
  if (value === undefined) return 0.04;
  if (!Number.isFinite(value)) return 0.04;
  return Math.max(0, Math.min(1, value));
}

function pngChunk(type: string, data: Buffer): Buffer {
  const typeBuffer = Buffer.from(type, "ascii");
  return concatBuffers([u32be(data.length), typeBuffer, data, u32be(crc32(concatBuffers([typeBuffer, data])))]);
}

function u32be(value: number): Buffer {
  const buffer = Buffer.alloc(4);
  buffer.writeUInt32BE(value >>> 0);
  return buffer;
}

function concatBuffers(chunks: Array<Uint8Array | Buffer>): Buffer {
  return Buffer.concat(chunks.map((chunk) => (chunk instanceof Buffer ? chunk : Buffer.from(chunk))));
}

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const CRC_TABLE = new Uint32Array(256).map((_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit += 1) {
    value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  }
  return value >>> 0;
});

function crc32(data: Buffer): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
