import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";
import { inflateRawSync } from "node:zlib";

import { IR_VERSION, type DeckIR, type IRObject, type JSONRecord, type Slide } from "../ir/index.ts";

const DEFAULT_DECK_SIZE = { width: 1280, height: 720 };
const MAX_TEXT_CANDIDATES_PER_STREAM = 80;
const MAX_TEXT_OBJECTS_PER_SLIDE = 24;
const MAX_SNAPPY_OUTPUT_BYTES = 64 * 1024 * 1024;

export interface NativeKeynoteDetection {
  isNative: boolean;
  container: "directory" | "zip";
  hasIndexZip: boolean;
  iwaPaths: string[];
  metadataPaths: string[];
  assetPaths: string[];
}

interface KeynotePackage {
  container: "directory" | "zip";
  entries: Map<string, Uint8Array>;
}

interface NativeKeynoteParts {
  hasIndexZip: boolean;
  iwaEntries: Array<{ path: string; data: Uint8Array }>;
  metadataPaths: string[];
  assetPaths: string[];
  indexZipEntryCount: number;
}

interface NativeMetadata {
  title?: string;
  author?: string;
  createdAt?: string;
  updatedAt?: string;
  size?: { width: number; height: number };
  values: JSONRecord;
}

export async function detectNativeKeynotePackage(keynotePath: string): Promise<NativeKeynoteDetection> {
  const keynotePackage = await readKeynotePackage(keynotePath);
  const parts = readNativeKeynoteParts(keynotePackage.entries);
  return {
    isNative: parts.hasIndexZip || parts.iwaEntries.length > 0,
    container: keynotePackage.container,
    hasIndexZip: parts.hasIndexZip,
    iwaPaths: parts.iwaEntries.map((entry) => entry.path),
    metadataPaths: parts.metadataPaths,
    assetPaths: parts.assetPaths
  };
}

export async function parseNativeKeynoteToIr(keynotePath: string): Promise<DeckIR> {
  const keynotePackage = await readKeynotePackage(keynotePath);
  const parts = readNativeKeynoteParts(keynotePackage.entries);
  if (!parts.hasIndexZip && parts.iwaEntries.length === 0) {
    throw new Error("Native Keynote package was not detected. Expected Index.zip or Index/*.iwa streams.");
  }

  const metadata = readNativeMetadata(keynotePackage.entries);
  const deckSize = metadata.size ?? DEFAULT_DECK_SIZE;
  const slideEntries = chooseSlideIwaEntries(parts.iwaEntries);
  const slides = buildNativeSlides(slideEntries, parts.iwaEntries, deckSize);
  const objectCount = slides.reduce((total, slide) => total + slide.objects.length, 0);
  const deckTitle = metadata.title ?? (path.basename(keynotePath, path.extname(keynotePath)) || "Imported Keynote");

  return {
    irVersion: IR_VERSION,
    metadata: {
      title: deckTitle,
      author: metadata.author,
      createdAt: metadata.createdAt,
      updatedAt: metadata.updatedAt,
      sourceApplication: "Keynote",
      custom: {
        ...metadata.values,
        nativeContainer: keynotePackage.container,
        nativeIndexZip: parts.hasIndexZip,
        nativeIwaStreamCount: parts.iwaEntries.length
      }
    },
    deck: {
      id: "keynote-native-import",
      title: deckTitle,
      size: { width: deckSize.width, height: deckSize.height, unit: "px" },
      slides
    },
    conversion: {
      source: { kind: "keynote", uri: keynotePath, application: "Keynote" },
      status: "partial",
      generatedAt: new Date().toISOString(),
      tool: "keymorph-keynote-native-probe",
      messages: [
        {
          severity: "warning",
          code: "keynote-native-static-probe",
          message:
            "Parsed a native Keynote package directly by inspecting Index.zip/Index/*.iwa streams; layout and animation fidelity are not available without Keynote."
        }
      ],
      unsupportedFeatures: [
        {
          code: "keynote-native-layout-schema",
          severity: "warning",
          area: "layout",
          description:
            "Modern Keynote stores slide layout in private IWA protobuf schemas. The native parser only infers slide streams and text-like payloads.",
          fallback: "Create approximate text objects or placeholders and preserve source paths in metadata."
        },
        {
          code: "keynote-native-animation-schema",
          severity: "warning",
          area: "animation",
          description: "Keynote builds, transitions, Magic Move, and timing data are not mapped by the native fallback.",
          fallback: "Use the Keynote PPTX bridge for the best available animation downgrade path."
        }
      ],
      degradedFeatures: [
        {
          code: "keynote-native-text-positioning",
          severity: "warning",
          area: "text",
          description: "Text recovered from IWA payloads is placed in approximate reading-order boxes.",
          fallback: "Use Keynote's PPTX export bridge when visual placement is required."
        }
      ],
      uncertainMappings: [
        {
          code: "keynote-iwa-protobuf-string-scan",
          severity: "warning",
          description:
            "Text objects are inferred from protobuf length-delimited UTF-8 strings and raw string scans, not from public Keynote object schemas.",
          confidence: 0.38
        }
      ],
      statistics: {
        slideCount: slides.length,
        objectCount,
        animationCount: 0,
        assetCount: parts.assetPaths.length,
        unsupportedFeatureCount: 2,
        degradedFeatureCount: 1,
        uncertainMappingCount: 1
      },
      metadata: {
        container: keynotePackage.container,
        hasIndexZip: parts.hasIndexZip,
        indexZipEntryCount: parts.indexZipEntryCount,
        iwaPathCount: parts.iwaEntries.length,
        metadataPathCount: parts.metadataPaths.length
      }
    }
  };
}

async function readKeynotePackage(keynotePath: string): Promise<KeynotePackage> {
  const fileStat = await stat(keynotePath);
  if (fileStat.isDirectory()) {
    return {
      container: "directory",
      entries: await readDirectoryEntries(keynotePath)
    };
  }

  const data = await readFile(keynotePath);
  return {
    container: "zip",
    entries: readZipEntries(data)
  };
}

async function readDirectoryEntries(root: string): Promise<Map<string, Uint8Array>> {
  const entries = new Map<string, Uint8Array>();

  async function walk(current: string): Promise<void> {
    const dirents = await readdir(current, { withFileTypes: true });
    await Promise.all(
      dirents.map(async (dirent) => {
        const absolutePath = path.join(current, dirent.name);
        if (dirent.isDirectory()) {
          await walk(absolutePath);
          return;
        }
        if (!dirent.isFile()) {
          return;
        }
        const relativePath = normalizePartPath(path.relative(root, absolutePath));
        entries.set(relativePath, await readFile(absolutePath));
      })
    );
  }

  await walk(root);
  return entries;
}

function readNativeKeynoteParts(entries: Map<string, Uint8Array>): NativeKeynoteParts {
  const combinedEntries = new Map<string, Uint8Array>();
  let hasIndexZip = false;
  let indexZipEntryCount = 0;

  for (const [entryPath, data] of entries) {
    const normalized = normalizePartPath(entryPath);
    combinedEntries.set(normalized, data);
    if (normalized.toLowerCase() !== "index.zip") {
      continue;
    }

    hasIndexZip = true;
    const indexEntries = readZipEntries(data);
    indexZipEntryCount += indexEntries.size;
    for (const [innerPath, innerData] of indexEntries) {
      const normalizedInner = normalizePartPath(innerPath);
      combinedEntries.set(normalizedInner.startsWith("Index/") ? normalizedInner : `Index/${normalizedInner}`, innerData);
    }
  }

  const allPaths = Array.from(combinedEntries.keys()).sort(comparePartPaths);
  return {
    hasIndexZip,
    iwaEntries: allPaths
      .filter((entryPath) => entryPath.toLowerCase().endsWith(".iwa"))
      .map((entryPath) => ({ path: entryPath, data: combinedEntries.get(entryPath)! })),
    metadataPaths: allPaths.filter((entryPath) => isMetadataPath(entryPath)),
    assetPaths: allPaths.filter((entryPath) => entryPath.startsWith("Data/")),
    indexZipEntryCount
  };
}

function readNativeMetadata(entries: Map<string, Uint8Array>): NativeMetadata {
  const values: JSONRecord = {};
  for (const [entryPath, data] of entries) {
    if (!isMetadataPath(entryPath)) {
      continue;
    }
    const text = decodeUtf8(data);
    if (!text || !text.includes("<plist")) {
      continue;
    }
    const plistValues = parseXmlPlistValues(text);
    for (const [key, value] of plistValues) {
      values[key] = value;
    }
  }

  return {
    title: stringValue(values, ["title", "Title", "documentTitle", "DocumentTitle", "name", "Name"]),
    author: stringValue(values, ["author", "Author", "creator", "Creator"]),
    createdAt: stringValue(values, ["createdAt", "creationDate", "CreationDate", "dateCreated"]),
    updatedAt: stringValue(values, ["updatedAt", "modificationDate", "ModificationDate", "lastModifiedDate"]),
    size: readDeckSize(values),
    values
  };
}

function buildNativeSlides(
  slideEntries: Array<{ path: string; data: Uint8Array }>,
  allIwaEntries: Array<{ path: string; data: Uint8Array }>,
  deckSize: { width: number; height: number }
): Slide[] {
  const entries = slideEntries.length > 0 ? slideEntries : chooseFallbackIwaEntries(allIwaEntries);
  if (entries.length === 0) {
    return [createPlaceholderSlide(0, "Native Keynote placeholder", undefined, deckSize)];
  }

  const slides = entries.map((entry, index) => {
    const slideId = `slide-${index + 1}`;
    const textCandidates = extractIwaTextCandidates(entry.data).slice(0, MAX_TEXT_OBJECTS_PER_SLIDE);
    const objects =
      textCandidates.length > 0
        ? textCandidates.map((text, objectIndex) => createTextObject(slideId, text, objectIndex, entry.path, deckSize))
        : [createPlaceholderObject(slideId, entry.path, deckSize)];

    return {
      id: slideId,
      index,
      name: readSlideName(entry.path, index),
      background: { type: "solid" as const, color: "#ffffff" },
      objects,
      timeline: { durationMs: 2500, events: [], dependencyGraph: { edges: [] } },
      metadata: {
        nativeSourcePath: entry.path,
        nativeTextCandidateCount: textCandidates.length,
        nativeParser: "iwa-protobuf-string-scan"
      }
    };
  });

  return slides.length > 0 ? slides : [createPlaceholderSlide(0, "Native Keynote placeholder", undefined, deckSize)];
}

function chooseSlideIwaEntries(iwaEntries: Array<{ path: string; data: Uint8Array }>): Array<{ path: string; data: Uint8Array }> {
  return iwaEntries.filter((entry) => isSlideIwaPath(entry.path)).sort((left, right) => comparePartPaths(left.path, right.path));
}

function chooseFallbackIwaEntries(iwaEntries: Array<{ path: string; data: Uint8Array }>): Array<{ path: string; data: Uint8Array }> {
  const documentEntry = iwaEntries.find((entry) => /(^|\/)document\.iwa$/i.test(entry.path));
  if (documentEntry) {
    return [documentEntry];
  }
  return iwaEntries.slice(0, 1);
}

function createPlaceholderSlide(
  index: number,
  name: string,
  sourcePath: string | undefined,
  deckSize: { width: number; height: number }
): Slide {
  const slideId = `slide-${index + 1}`;
  return {
    id: slideId,
    index,
    name,
    background: { type: "solid", color: "#ffffff" },
    objects: [createPlaceholderObject(slideId, sourcePath, deckSize)],
    timeline: { durationMs: 2500, events: [], dependencyGraph: { edges: [] } },
    metadata: {
      nativeSourcePath: sourcePath,
      nativeParser: "iwa-protobuf-string-scan"
    }
  };
}

function createTextObject(
  slideId: string,
  text: string,
  objectIndex: number,
  sourcePath: string,
  deckSize: { width: number; height: number }
): IRObject {
  const marginX = Math.round(deckSize.width * 0.075);
  const top = Math.round(deckSize.height * 0.12);
  const lineHeight = objectIndex === 0 ? 76 : 54;
  const y = top + objectIndex * (lineHeight + 18);

  return {
    id: `${slideId}-text-${objectIndex + 1}`,
    type: "text",
    name: objectIndex === 0 ? "Native text" : `Native text ${objectIndex + 1}`,
    bounds: {
      x: marginX,
      y,
      width: deckSize.width - marginX * 2,
      height: Math.max(44, lineHeight)
    },
    opacity: 1,
    text: {
      plainText: text,
      runs: [
        {
          text,
          style: {
            fontFamily: "Helvetica Neue",
            fontSize: objectIndex === 0 ? 36 : 24,
            color: "#111827"
          }
        }
      ]
    },
    metadata: {
      nativeSourcePath: sourcePath,
      nativeExtraction: "protobuf-length-delimited-or-raw-string"
    }
  };
}

function createPlaceholderObject(
  slideId: string,
  sourcePath: string | undefined,
  deckSize: { width: number; height: number }
): IRObject {
  const marginX = Math.round(deckSize.width * 0.075);
  return {
    id: `${slideId}-native-placeholder`,
    type: "placeholder",
    name: "Native Keynote slide placeholder",
    placeholderType: "custom",
    bounds: {
      x: marginX,
      y: Math.round(deckSize.height * 0.12),
      width: deckSize.width - marginX * 2,
      height: 96
    },
    opacity: 1,
    metadata: {
      nativeSourcePath: sourcePath,
      reason: "No decodable slide text was found in this IWA stream."
    }
  };
}

function extractIwaTextCandidates(data: Uint8Array): string[] {
  const candidates = new Set<string>();
  const payloads = expandIwaPayloads(data);

  for (const payload of payloads) {
    for (const text of extractProtobufStrings(payload)) {
      addTextCandidate(candidates, text);
      if (candidates.size >= MAX_TEXT_CANDIDATES_PER_STREAM) {
        return Array.from(candidates);
      }
    }
    for (const text of extractRawStrings(payload)) {
      addTextCandidate(candidates, text);
      if (candidates.size >= MAX_TEXT_CANDIDATES_PER_STREAM) {
        return Array.from(candidates);
      }
    }
  }

  return Array.from(candidates);
}

function addTextCandidate(candidates: Set<string>, text: string): void {
  const cleaned = cleanTextCandidate(text);
  if (cleaned) {
    candidates.add(cleaned);
  }
}

function expandIwaPayloads(data: Uint8Array): Uint8Array[] {
  const payloads: Uint8Array[] = [data];
  const framed = decodeSnappyFramedStream(data);
  if (framed) {
    payloads.push(framed);
  }

  const rawSnappy = decodeSnappyBlock(data);
  if (rawSnappy) {
    payloads.push(rawSnappy);
  }

  for (const decoded of decodeLengthPrefixedSnappyBlocks(data)) {
    payloads.push(decoded);
  }

  return dedupeByteArrays(payloads);
}

function decodeSnappyFramedStream(data: Uint8Array): Uint8Array | undefined {
  if (data.length < 10 || data[0] !== 0xff || readUint24(data, 1) !== 6 || decodeAscii(data.subarray(4, 10)) !== "sNaPpY") {
    return undefined;
  }

  const chunks: Uint8Array[] = [];
  let offset = 10;
  while (offset + 4 <= data.length) {
    const chunkType = data[offset];
    const chunkLength = readUint24(data, offset + 1);
    const chunkStart = offset + 4;
    const chunkEnd = chunkStart + chunkLength;
    if (chunkEnd > data.length) {
      return undefined;
    }

    const chunk = data.subarray(chunkStart, chunkEnd);
    if (chunkType === 0x00) {
      if (chunk.length < 4) return undefined;
      const decoded = decodeSnappyBlock(chunk.subarray(4));
      if (!decoded) return undefined;
      chunks.push(decoded);
    } else if (chunkType === 0x01) {
      if (chunk.length < 4) return undefined;
      chunks.push(chunk.subarray(4));
    } else if (chunkType === 0xff || chunkType >= 0x80) {
      // Stream identifiers and skippable chunks carry no protobuf payload.
    } else {
      return undefined;
    }

    offset = chunkEnd;
  }

  return chunks.length > 0 ? concat(chunks) : undefined;
}

function decodeLengthPrefixedSnappyBlocks(data: Uint8Array): Uint8Array[] {
  const outputs: Uint8Array[] = [];
  for (const littleEndian of [true, false]) {
    const chunks: Uint8Array[] = [];
    let offset = 0;
    while (offset + 4 <= data.length) {
      const length = new DataView(data.buffer, data.byteOffset + offset, 4).getUint32(0, littleEndian);
      offset += 4;
      if (length === 0 || length > data.length - offset) {
        chunks.length = 0;
        break;
      }
      const decoded = decodeSnappyBlock(data.subarray(offset, offset + length));
      if (!decoded) {
        chunks.length = 0;
        break;
      }
      chunks.push(decoded);
      offset += length;
    }
    if (chunks.length > 0 && offset === data.length) {
      outputs.push(concat(chunks));
    }
  }
  return outputs;
}

function decodeSnappyBlock(data: Uint8Array): Uint8Array | undefined {
  try {
    const lengthInfo = readVarint(data, 0);
    if (!lengthInfo || lengthInfo.value <= 0 || lengthInfo.value > MAX_SNAPPY_OUTPUT_BYTES) {
      return undefined;
    }

    const output = new Uint8Array(lengthInfo.value);
    let inputOffset = lengthInfo.nextOffset;
    let outputOffset = 0;

    while (inputOffset < data.length && outputOffset < output.length) {
      const tag = data[inputOffset];
      if (tag === undefined) return undefined;
      inputOffset += 1;
      const tagType = tag & 0x03;

      if (tagType === 0) {
        let literalLength = tag >>> 2;
        if (literalLength < 60) {
          literalLength += 1;
        } else {
          const lengthByteCount = literalLength - 59;
          if (inputOffset + lengthByteCount > data.length) return undefined;
          literalLength = 1;
          for (let index = 0; index < lengthByteCount; index += 1) {
            literalLength += (data[inputOffset + index] ?? 0) << (8 * index);
          }
          inputOffset += lengthByteCount;
        }
        if (inputOffset + literalLength > data.length || outputOffset + literalLength > output.length) return undefined;
        output.set(data.subarray(inputOffset, inputOffset + literalLength), outputOffset);
        inputOffset += literalLength;
        outputOffset += literalLength;
        continue;
      }

      let copyLength: number;
      let copyOffset: number;
      if (tagType === 1) {
        if (inputOffset >= data.length) return undefined;
        copyLength = ((tag >>> 2) & 0x07) + 4;
        copyOffset = ((tag & 0xe0) << 3) | (data[inputOffset] ?? 0);
        inputOffset += 1;
      } else if (tagType === 2) {
        if (inputOffset + 2 > data.length) return undefined;
        copyLength = (tag >>> 2) + 1;
        copyOffset = (data[inputOffset] ?? 0) | ((data[inputOffset + 1] ?? 0) << 8);
        inputOffset += 2;
      } else {
        if (inputOffset + 4 > data.length) return undefined;
        copyLength = (tag >>> 2) + 1;
        copyOffset =
          (data[inputOffset] ?? 0) |
          ((data[inputOffset + 1] ?? 0) << 8) |
          ((data[inputOffset + 2] ?? 0) << 16) |
          ((data[inputOffset + 3] ?? 0) << 24);
        inputOffset += 4;
      }

      if (copyOffset <= 0 || copyOffset > outputOffset || outputOffset + copyLength > output.length) {
        return undefined;
      }
      for (let index = 0; index < copyLength; index += 1) {
        output[outputOffset] = output[outputOffset - copyOffset];
        outputOffset += 1;
      }
    }

    return outputOffset === output.length ? output : undefined;
  } catch {
    return undefined;
  }
}

function extractProtobufStrings(data: Uint8Array): string[] {
  const strings = new Set<string>();
  for (let offset = 0; offset < data.length && strings.size < MAX_TEXT_CANDIDATES_PER_STREAM; offset += 1) {
    const key = readVarint(data, offset);
    if (!key) {
      continue;
    }
    const wireType = key.value & 0x07;
    const fieldNumber = key.value >>> 3;
    if (wireType !== 2 || fieldNumber <= 0 || fieldNumber > 8191) {
      continue;
    }
    const lengthInfo = readVarint(data, key.nextOffset);
    if (!lengthInfo || lengthInfo.value <= 0 || lengthInfo.value > 16_384) {
      continue;
    }
    const valueStart = lengthInfo.nextOffset;
    const valueEnd = valueStart + lengthInfo.value;
    if (valueEnd > data.length) {
      continue;
    }

    const text = decodeTextCandidate(data.subarray(valueStart, valueEnd), 2);
    if (text) {
      strings.add(text);
    }
  }

  return Array.from(strings);
}

function extractRawStrings(data: Uint8Array): string[] {
  const strings = new Set<string>();
  let start = -1;

  for (let index = 0; index <= data.length; index += 1) {
    const byte = data[index];
    const printable = byte !== undefined && (byte === 0x09 || byte === 0x0a || byte === 0x0d || byte >= 0x20);
    if (printable) {
      if (start < 0) start = index;
      continue;
    }
    if (start >= 0 && index - start >= 4) {
      const text = decodeTextCandidate(data.subarray(start, index), 4);
      if (text) {
        strings.add(text);
        if (strings.size >= MAX_TEXT_CANDIDATES_PER_STREAM) {
          return Array.from(strings);
        }
      }
    }
    start = -1;
  }

  return Array.from(strings);
}

function decodeTextCandidate(bytes: Uint8Array, minLength: number): string | undefined {
  const decoded = decodeUtf8(bytes);
  if (!decoded || decoded.includes("\ufffd")) {
    return undefined;
  }

  const text = cleanTextCandidate(decoded);
  if (!text) {
    return undefined;
  }
  if (text.length < minLength || text.length > 500 || !looksLikePresentationText(text)) {
    return undefined;
  }
  return text;
}

function cleanTextCandidate(value: string): string | undefined {
  let text = value.replace(/\u0000/g, "").replace(/\s+/g, " ").trim();
  for (const marker of CONTAINER_MARKERS) {
    text = text.replaceAll(marker, " ");
  }
  text = text.replace(/^[|:;,\-_.\s]+|[|:;,\-_.\s]+$/g, "").replace(/\s+/g, " ").trim();
  if (!text || isContainerMarker(text)) {
    return undefined;
  }
  return text;
}

function looksLikePresentationText(text: string): boolean {
  const printable = Array.from(text).filter((char) => {
    const code = char.charCodeAt(0);
    return code >= 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;
  }).length;
  if (printable / Math.max(1, text.length) < 0.92) {
    return false;
  }
  if (/^(TSP|TSK|TSD|TSWP|KN|KPF|SFA|SFU|NS)\./.test(text) || /^com\.apple\./.test(text)) {
    return false;
  }
  if (isContainerMarker(text)) {
    return false;
  }
  if (/\.(iwa|plist|jpg|jpeg|png|gif|tiff|mov|m4v|mp4|pdf|key)$/i.test(text)) {
    return false;
  }
  if (/^[0-9a-f]{8}-[0-9a-f-]{27,}$/i.test(text)) {
    return false;
  }
  if (text.length > 80 && !/\s/.test(text) && /[./:_-]/.test(text)) {
    return false;
  }
  return /[\p{L}\p{N}]/u.test(text);
}

function isContainerMarker(text: string): boolean {
  return CONTAINER_MARKERS.has(text);
}

const CONTAINER_MARKERS = new Set(["sNaPpY", "IWA", "Index.zip"]);

function parseXmlPlistValues(source: string): Map<string, string | number | boolean> {
  const values = new Map<string, string | number | boolean>();
  const pairPattern =
    /<key>([\s\S]*?)<\/key>\s*(?:<string>([\s\S]*?)<\/string>|<date>([\s\S]*?)<\/date>|<integer>([\s\S]*?)<\/integer>|<real>([\s\S]*?)<\/real>|<(true|false)\s*\/>)/g;

  for (const match of source.matchAll(pairPattern)) {
    const rawKey = match[1];
    if (!rawKey) continue;
    const key = unescapeXml(rawKey).trim();
    if (!key) continue;

    if (match[2] !== undefined) {
      values.set(key, unescapeXml(match[2]).trim());
    } else if (match[3] !== undefined) {
      values.set(key, unescapeXml(match[3]).trim());
    } else if (match[4] !== undefined) {
      values.set(key, Number(match[4]));
    } else if (match[5] !== undefined) {
      values.set(key, Number(match[5]));
    } else if (match[6] !== undefined) {
      values.set(key, match[6] === "true");
    }
  }

  return values;
}

function readDeckSize(values: JSONRecord): { width: number; height: number } | undefined {
  const width = numberValue(values, ["slideWidth", "SlideWidth", "documentWidth", "DocumentWidth", "width", "Width"]);
  const height = numberValue(values, ["slideHeight", "SlideHeight", "documentHeight", "DocumentHeight", "height", "Height"]);
  if (width === undefined || height === undefined || width < 100 || height < 100 || width > 10000 || height > 10000) {
    return undefined;
  }
  return { width, height };
}

function stringValue(values: JSONRecord, keys: string[]): string | undefined {
  for (const key of keys) {
    const value = values[key];
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }
  return undefined;
}

function numberValue(values: JSONRecord, keys: string[]): number | undefined {
  for (const key of keys) {
    const value = values[key];
    if (typeof value === "number" && Number.isFinite(value)) {
      return value;
    }
    if (typeof value === "string" && value.trim() && Number.isFinite(Number(value))) {
      return Number(value);
    }
  }
  return undefined;
}

function isSlideIwaPath(entryPath: string): boolean {
  const baseName = path.posix.basename(entryPath).toLowerCase();
  if (!baseName.endsWith(".iwa") || /master|layout|style|theme/.test(baseName)) {
    return false;
  }
  return /^slide(?:[-_.\d]|$)/.test(baseName) || /^slide[0-9]+\.iwa$/.test(baseName);
}

function readSlideName(entryPath: string, index: number): string {
  const baseName = path.posix.basename(entryPath, path.extname(entryPath));
  return baseName && baseName.toLowerCase() !== "slide" ? baseName : `Slide ${index + 1}`;
}

function isMetadataPath(entryPath: string): boolean {
  return entryPath.startsWith("Metadata/") || /^metadata\//i.test(entryPath);
}

function readZipEntries(data: Uint8Array): Map<string, Uint8Array> {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const entries = new Map<string, Uint8Array>();
  let offset = 0;

  while (offset + 30 <= data.length && view.getUint32(offset, true) === 0x04034b50) {
    const flags = view.getUint16(offset + 6, true);
    const compression = view.getUint16(offset + 8, true);
    const compressedSize = view.getUint32(offset + 18, true);
    const uncompressedSize = view.getUint32(offset + 22, true);
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const nameStart = offset + 30;
    const dataStart = nameStart + nameLength + extraLength;
    const dataEnd = dataStart + compressedSize;
    const name = decodeUtf8(data.subarray(nameStart, nameStart + nameLength)) ?? "";

    if ((flags & 0x08) !== 0) {
      throw new Error("Keynote ZIP entries using data descriptors are not supported by the native reader.");
    }
    if (dataEnd > data.length) {
      throw new Error(`Invalid Keynote ZIP entry size for ${name}.`);
    }

    const compressed = data.subarray(dataStart, dataEnd);
    let content: Uint8Array;
    if (compression === 0) {
      content = compressed;
    } else if (compression === 8) {
      content = inflateRawSync(compressed);
    } else {
      throw new Error(`Unsupported ZIP compression method ${compression} for ${name}.`);
    }

    if (uncompressedSize !== 0 && content.length !== uncompressedSize) {
      throw new Error(`Invalid uncompressed size for ${name}.`);
    }

    entries.set(normalizePartPath(name), content);
    offset = dataEnd;
  }

  if (entries.size === 0) {
    throw new Error("Keynote file is not a readable ZIP package.");
  }

  return entries;
}

function comparePartPaths(left: string, right: string): number {
  return left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
}

function normalizePartPath(partPath: string): string {
  const parts: string[] = [];
  for (const part of partPath.replaceAll("\\", "/").replace(/^\/+/, "").split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }
  return parts.join("/");
}

function readUint24(data: Uint8Array, offset: number): number {
  return (data[offset] ?? 0) | ((data[offset + 1] ?? 0) << 8) | ((data[offset + 2] ?? 0) << 16);
}

function readVarint(data: Uint8Array, offset: number): { value: number; nextOffset: number } | undefined {
  let value = 0;
  let shift = 0;
  let cursor = offset;
  while (cursor < data.length && shift <= 49) {
    const byte = data[cursor];
    if (byte === undefined) return undefined;
    value += (byte & 0x7f) * 2 ** shift;
    cursor += 1;
    if ((byte & 0x80) === 0) {
      return { value, nextOffset: cursor };
    }
    shift += 7;
  }
  return undefined;
}

function decodeUtf8(data: Uint8Array): string | undefined {
  try {
    return new TextDecoder().decode(data);
  } catch {
    return undefined;
  }
}

function decodeAscii(data: Uint8Array): string {
  return Array.from(data, (byte) => String.fromCharCode(byte)).join("");
}

function unescapeXml(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function dedupeByteArrays(arrays: Uint8Array[]): Uint8Array[] {
  const seen = new Set<string>();
  const deduped: Uint8Array[] = [];
  for (const array of arrays) {
    const key = `${array.byteLength}:${array[0] ?? 0}:${array[1] ?? 0}:${array[array.byteLength - 1] ?? 0}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    deduped.push(array);
  }
  return deduped;
}

function concat(chunks: Uint8Array[]): Uint8Array {
  const length = chunks.reduce((total, chunk) => total + chunk.length, 0);
  const out = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}
