import { createHash } from "node:crypto";
import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";
import { inflateRawSync } from "node:zlib";

import { IR_VERSION, type Asset, type DeckIR, type IRObject, type JSONRecord, type Slide } from "../ir/index.ts";

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
  iwaStreams?: NativeIwaStreamMetadata[];
}

interface KeynotePackage {
  container: "directory" | "zip";
  entries: Map<string, Uint8Array>;
}

interface NativeKeynoteParts {
  hasIndexZip: boolean;
  entries: Map<string, Uint8Array>;
  iwaEntries: Array<{ path: string; data: Uint8Array }>;
  metadataPaths: string[];
  assetPaths: string[];
  indexZipEntryCount: number;
}

export interface NativeIwaStreamMetadata {
  path: string;
  role: NativeIwaStreamRole;
  compression: NativeIwaCompression[];
  byteLength: number;
  expandedByteLength: number;
  textCandidateCount: number;
  assetReferenceCount: number;
}

export type NativeIwaStreamRole = "slide" | "document" | "theme" | "master" | "layout" | "style" | "unknown";
export type NativeIwaCompression = "none" | "snappy-framed" | "snappy-block" | "length-prefixed-snappy";

interface NativeAssetReference {
  assetId: string;
  path: string;
  name: string;
  kind: Asset["kind"];
  mimeType?: string;
}

interface NativeAssetCatalog {
  assets: Asset[];
  byPath: Map<string, NativeAssetReference>;
  byName: Map<string, NativeAssetReference[]>;
  byStem: Map<string, NativeAssetReference[]>;
}

interface NativeAssetMatch {
  asset: NativeAssetReference;
  evidence: string;
  confidence: number;
}

interface ExpandedIwaPayload {
  data: Uint8Array;
  compression: NativeIwaCompression;
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
  const assets = createNativeAssetCatalog(parts.assetPaths, parts.entries);
  return {
    isNative: parts.hasIndexZip || parts.iwaEntries.length > 0,
    container: keynotePackage.container,
    hasIndexZip: parts.hasIndexZip,
    iwaPaths: parts.iwaEntries.map((entry) => entry.path),
    metadataPaths: parts.metadataPaths,
    assetPaths: parts.assetPaths,
    iwaStreams: classifyIwaStreams(parts.iwaEntries, assets)
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
  const assets = createNativeAssetCatalog(parts.assetPaths, parts.entries);
  const iwaStreams = classifyIwaStreams(parts.iwaEntries, assets);
  const slideEntries = chooseSlideIwaEntries(parts.iwaEntries);
  const slides = buildNativeSlides(slideEntries, parts.iwaEntries, deckSize, assets);
  const objectCount = slides.reduce((total, slide) => total + slide.objects.length, 0);
  const recoveredTextObjectCount = slides.reduce(
    (total, slide) => total + slide.objects.filter((object) => object.type === "text").length,
    0
  );
  const recoveredAssetObjectCount = slides.reduce(
    (total, slide) => total + slide.objects.filter((object) => object.type === "image" || object.type === "media").length,
    0
  );
  const unrecoveredAssetCount = Math.max(0, assets.assets.length - countUniqueReferencedAssets(slides));
  const hasDetectedAssets = assets.assets.length > 0;
  const unsupportedFeatures = [
    {
      code: "keynote-native-layout-schema",
      severity: "warning" as const,
      area: "layout" as const,
      description:
        "Modern Keynote stores slide layout in private IWA protobuf schemas. The native parser only classifies streams and creates approximate text and asset objects from visible payload hints.",
      fallback: "Create approximate text or asset placeholders and preserve source paths in metadata."
    },
    {
      code: "keynote-native-animation-schema",
      severity: "warning" as const,
      area: "animation" as const,
      description: "Keynote builds, transitions, Magic Move, and timing data are not mapped by the native fallback.",
      fallback: "Use the Keynote PPTX bridge for the best available animation downgrade path."
    },
    ...(hasDetectedAssets
      ? [
          {
            code: "keynote-native-asset-layout-loss",
            severity: "warning" as const,
            area: "layout" as const,
            description:
              "Image and media files were detected in the Keynote package, but their exact slide geometry, crop, stacking, and styling are not recoverable from the private IWA layout schema.",
            fallback: "Create approximate asset objects only when slide streams contain matching asset references; preserve remaining assets in deck.assets."
          }
        ]
      : [])
  ];
  const degradedFeatures = [
    {
      code: "keynote-native-text-recovery",
      severity: "warning" as const,
      area: "text" as const,
      description:
        recoveredTextObjectCount > 0
          ? "Text-like strings were recovered from IWA payloads and placed in approximate reading-order boxes."
          : "No slide text was recovered from the visible IWA payloads.",
      fallback: "Use Keynote's PPTX export bridge when visual text placement is required."
    },
    ...(recoveredAssetObjectCount > 0
      ? [
          {
            code: "keynote-native-asset-reference-recovery",
            severity: "warning" as const,
            area: "asset" as const,
            description:
              "Some package asset references were matched from slide IWA string payloads and inserted as approximate image/media objects.",
            fallback: "Inspect the source Keynote deck or bridge through PPTX for exact placement."
          }
        ]
      : [])
  ];
  const uncertainMappings = [
    {
      code: "keynote-iwa-protobuf-string-scan",
      severity: "warning" as const,
      description:
        "Text and asset references are inferred from protobuf length-delimited UTF-8 strings and raw string scans, not from public Keynote object schemas.",
      confidence: 0.38
    },
    ...(recoveredAssetObjectCount > 0
      ? [
          {
            code: "keynote-native-asset-reference-scan",
            severity: "warning" as const,
            description:
              "Asset-to-slide associations are based on matching Data/ asset filenames found in IWA payload strings; this does not prove exact placement or usage.",
            confidence: 0.46
          }
        ]
      : [])
  ];
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
        nativeIwaStreamCount: parts.iwaEntries.length,
        nativeAssetCount: assets.assets.length
      }
    },
    deck: {
      id: "keynote-native-import",
      title: deckTitle,
      size: { width: deckSize.width, height: deckSize.height, unit: "px" },
      ...(assets.assets.length > 0 ? { assets: assets.assets } : {}),
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
        },
        {
          severity: recoveredTextObjectCount > 0 ? "info" : "warning",
          code: recoveredTextObjectCount > 0 ? "keynote-native-text-recovered" : "keynote-native-text-not-recovered",
          message:
            recoveredTextObjectCount > 0
              ? `Recovered ${recoveredTextObjectCount} text object(s) from visible IWA string payloads.`
              : "No native text objects were recovered from visible IWA string payloads."
        },
        {
          severity: hasDetectedAssets ? (recoveredAssetObjectCount > 0 ? "info" : "warning") : "info",
          code: hasDetectedAssets
            ? recoveredAssetObjectCount > 0
              ? "keynote-native-assets-referenced"
              : "keynote-native-assets-unplaced"
            : "keynote-native-no-assets",
          message: hasDetectedAssets
            ? recoveredAssetObjectCount > 0
              ? `Detected ${assets.assets.length} package asset(s) and matched ${recoveredAssetObjectCount} approximate slide object(s).`
              : `Detected ${assets.assets.length} package asset(s), but no slide stream references were matched for placement.`
            : "No package image or media assets were detected under Data/."
        }
      ],
      unsupportedFeatures,
      degradedFeatures,
      uncertainMappings,
      statistics: {
        slideCount: slides.length,
        objectCount,
        animationCount: 0,
        assetCount: assets.assets.length,
        unsupportedFeatureCount: unsupportedFeatures.length,
        degradedFeatureCount: degradedFeatures.length,
        uncertainMappingCount: uncertainMappings.length
      },
      metadata: {
        container: keynotePackage.container,
        hasIndexZip: parts.hasIndexZip,
        indexZipEntryCount: parts.indexZipEntryCount,
        iwaPathCount: parts.iwaEntries.length,
        metadataPathCount: parts.metadataPaths.length,
        assetPathCount: parts.assetPaths.length,
        recoveredTextObjectCount,
        recoveredAssetObjectCount,
        unrecoveredAssetCount,
        iwaStreams: iwaStreams.map((stream) => ({
          path: stream.path,
          role: stream.role,
          compression: stream.compression,
          byteLength: stream.byteLength,
          expandedByteLength: stream.expandedByteLength,
          textCandidateCount: stream.textCandidateCount,
          assetReferenceCount: stream.assetReferenceCount
        }))
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
    entries: combinedEntries,
    iwaEntries: allPaths
      .filter((entryPath) => entryPath.toLowerCase().endsWith(".iwa"))
      .map((entryPath) => ({ path: entryPath, data: combinedEntries.get(entryPath)! })),
    metadataPaths: allPaths.filter((entryPath) => isMetadataPath(entryPath)),
    assetPaths: allPaths.filter((entryPath) => entryPath.startsWith("Data/")),
    indexZipEntryCount
  };
}

function createNativeAssetCatalog(assetPaths: string[], entries: Map<string, Uint8Array>): NativeAssetCatalog {
  const assets: Asset[] = [];
  const byPath = new Map<string, NativeAssetReference>();
  const byName = new Map<string, NativeAssetReference[]>();
  const byStem = new Map<string, NativeAssetReference[]>();

  for (const assetPath of assetPaths) {
    const data = entries.get(assetPath);
    const name = path.posix.basename(assetPath);
    if (!name || !data) {
      continue;
    }

    const classification = classifyAssetPath(assetPath);
    const assetId = stableAssetId(assetPath);
    const asset: Asset = {
      id: assetId,
      kind: classification.kind,
      name,
      mimeType: classification.mimeType,
      checksum: `sha256:${sha256Hex(data)}`,
      metadata: {
        nativeSourcePath: assetPath,
        byteLength: data.byteLength
      }
    };
    const reference: NativeAssetReference = {
      assetId,
      path: assetPath,
      name,
      kind: asset.kind,
      mimeType: asset.mimeType
    };

    assets.push(asset);
    byPath.set(assetPath.toLowerCase(), reference);
    addMultiMapValue(byName, name.toLowerCase(), reference);
    addMultiMapValue(byStem, path.posix.basename(name, path.posix.extname(name)).toLowerCase(), reference);
  }

  assets.sort((left, right) => comparePartPaths(String(left.metadata?.nativeSourcePath ?? left.name ?? left.id), String(right.metadata?.nativeSourcePath ?? right.name ?? right.id)));
  return { assets, byPath, byName, byStem };
}

function classifyIwaStreams(
  iwaEntries: Array<{ path: string; data: Uint8Array }>,
  assets: NativeAssetCatalog
): NativeIwaStreamMetadata[] {
  return iwaEntries.map((entry) => {
    const payloads = expandIwaPayloads(entry.data);
    const compression = Array.from(new Set(payloads.map((payload) => payload.compression)));
    return {
      path: entry.path,
      role: classifyIwaRole(entry.path),
      compression,
      byteLength: entry.data.byteLength,
      expandedByteLength: payloads.reduce((max, payload) => Math.max(max, payload.data.byteLength), entry.data.byteLength),
      textCandidateCount: extractIwaTextCandidates(entry.data).length,
      assetReferenceCount: findIwaAssetMatches(entry.data, assets).length
    };
  });
}

function classifyIwaRole(entryPath: string): NativeIwaStreamRole {
  const baseName = path.posix.basename(entryPath).toLowerCase();
  if (isSlideIwaPath(entryPath)) return "slide";
  if (/(^|[-_.])document[-_.]?/.test(baseName) || baseName === "document.iwa") return "document";
  if (/master/.test(baseName)) return "master";
  if (/layout/.test(baseName)) return "layout";
  if (/theme/.test(baseName)) return "theme";
  if (/style/.test(baseName)) return "style";
  return "unknown";
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
  deckSize: { width: number; height: number },
  assets: NativeAssetCatalog
): Slide[] {
  const entries = slideEntries.length > 0 ? slideEntries : chooseFallbackIwaEntries(allIwaEntries);
  if (entries.length === 0) {
    return [createPlaceholderSlide(0, "Native Keynote placeholder", undefined, deckSize)];
  }

  const slides = entries.map((entry, index) => {
    const slideId = `slide-${index + 1}`;
    const textCandidates = extractIwaTextCandidates(entry.data).slice(0, MAX_TEXT_OBJECTS_PER_SLIDE);
    const assetMatches = findIwaAssetMatches(entry.data, assets);
    const textObjects = textCandidates.map((text, objectIndex) => createTextObject(slideId, text, objectIndex, entry.path, deckSize));
    const assetObjects = assetMatches.map((match, assetIndex) =>
      createAssetObject(slideId, match, textObjects.length + assetIndex, entry.path, deckSize)
    );
    const objects =
      textObjects.length > 0 || assetObjects.length > 0
        ? [...textObjects, ...assetObjects]
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
        nativeAssetReferenceCount: assetMatches.length,
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

function createAssetObject(
  slideId: string,
  match: NativeAssetMatch,
  objectIndex: number,
  sourcePath: string,
  deckSize: { width: number; height: number }
): IRObject {
  const asset = match.asset;
  const marginX = Math.round(deckSize.width * 0.075);
  const width = Math.round(deckSize.width * 0.38);
  const height = Math.round(deckSize.height * 0.38);
  const column = objectIndex % 2;
  const row = Math.floor(objectIndex / 2);
  const bounds = {
    x: marginX + column * Math.round(deckSize.width * 0.42),
    y: Math.round(deckSize.height * 0.48) + row * Math.round(deckSize.height * 0.12),
    width,
    height
  };
  const metadata = {
    nativeSourcePath: sourcePath,
    nativeAssetPath: asset.path,
    nativeAssetEvidence: match.evidence,
    nativeAssetMatchConfidence: match.confidence,
    nativeExtraction: "asset-filename-string-scan"
  };

  if (asset.kind === "image") {
    return {
      id: `${slideId}-asset-${objectIndex + 1}`,
      type: "image",
      name: asset.name,
      bounds,
      opacity: 1,
      source: {
        assetId: asset.assetId,
        metadata: {
          nativeAssetPath: asset.path,
          mimeType: asset.mimeType
        }
      },
      metadata
    };
  }

  if (asset.kind === "video" || asset.kind === "audio") {
    return {
      id: `${slideId}-asset-${objectIndex + 1}`,
      type: "media",
      name: asset.name,
      bounds,
      opacity: 1,
      mediaType: asset.kind,
      source: {
        assetId: asset.assetId,
        metadata: {
          nativeAssetPath: asset.path,
          mimeType: asset.mimeType
        }
      },
      metadata
    };
  }

  return {
    id: `${slideId}-asset-${objectIndex + 1}`,
    type: "placeholder",
    name: asset.name,
    placeholderType: "custom",
    bounds,
    opacity: 1,
    metadata: {
      ...metadata,
      reason: `Native package asset kind "${asset.kind}" cannot be represented as a direct IR image/media object.`
    }
  };
}

function extractIwaTextCandidates(data: Uint8Array): string[] {
  const candidates = new Set<string>();
  const payloads = expandIwaPayloads(data);

  for (const payload of payloads) {
    for (const text of extractProtobufStrings(payload.data)) {
      addTextCandidate(candidates, text);
      if (candidates.size >= MAX_TEXT_CANDIDATES_PER_STREAM) {
        return Array.from(candidates);
      }
    }
    for (const text of extractRawStrings(payload.data)) {
      addTextCandidate(candidates, text);
      if (candidates.size >= MAX_TEXT_CANDIDATES_PER_STREAM) {
        return Array.from(candidates);
      }
    }
  }

  return Array.from(candidates);
}

function findIwaAssetMatches(data: Uint8Array, assets: NativeAssetCatalog): NativeAssetMatch[] {
  if (assets.assets.length === 0) {
    return [];
  }

  const strings = new Set<string>();
  for (const payload of expandIwaPayloads(data)) {
    for (const value of extractReferenceStrings(payload.data)) strings.add(value);
  }

  const matches = new Map<string, NativeAssetMatch>();
  for (const value of strings) {
    for (const match of matchAssetReference(value, assets)) {
      const existing = matches.get(match.asset.assetId);
      if (!existing || match.confidence > existing.confidence) {
        matches.set(match.asset.assetId, match);
      }
    }
  }

  return Array.from(matches.values()).sort((left, right) => comparePartPaths(left.asset.path, right.asset.path));
}

function extractReferenceStrings(data: Uint8Array): string[] {
  const strings = new Set<string>();
  for (let offset = 0; offset < data.length; offset += 1) {
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
    if (!lengthInfo || lengthInfo.value <= 0 || lengthInfo.value > 4096) {
      continue;
    }
    const valueStart = lengthInfo.nextOffset;
    const valueEnd = valueStart + lengthInfo.value;
    if (valueEnd > data.length) {
      continue;
    }
    addReferenceString(strings, data.subarray(valueStart, valueEnd));
  }

  let start = -1;
  for (let index = 0; index <= data.length; index += 1) {
    const byte = data[index];
    const printable = byte !== undefined && byte >= 0x20 && byte <= 0x7e;
    if (printable) {
      if (start < 0) start = index;
      continue;
    }
    if (start >= 0 && index - start >= 4) {
      addReferenceString(strings, data.subarray(start, index));
    }
    start = -1;
  }

  return Array.from(strings);
}

function addReferenceString(strings: Set<string>, bytes: Uint8Array): void {
  const decoded = decodeUtf8(bytes);
  if (!decoded || decoded.includes("\ufffd")) {
    return;
  }
  const text = decoded.replace(/\u0000/g, "").trim();
  if (text.length >= 4 && text.length <= 1024 && /[\w.-]+\.[a-z0-9]{2,5}/i.test(text)) {
    strings.add(text);
  }
}

function matchAssetReference(value: string, assets: NativeAssetCatalog): NativeAssetMatch[] {
  const normalizedValue = normalizePartPath(value).toLowerCase();
  const exact = assets.byPath.get(normalizedValue);
  if (exact) {
    return [{ asset: exact, evidence: value, confidence: 0.82 }];
  }

  const baseName = path.posix.basename(normalizedValue);
  const nameMatches = baseName ? assets.byName.get(baseName) ?? [] : [];
  if (nameMatches.length === 1) {
    return [{ asset: nameMatches[0]!, evidence: value, confidence: 0.64 }];
  }
  if (nameMatches.length > 1) {
    return nameMatches.map((asset) => ({ asset, evidence: value, confidence: 0.5 }));
  }

  const stem = path.posix.basename(baseName, path.posix.extname(baseName));
  const stemMatches = stem ? assets.byStem.get(stem) ?? [] : [];
  if (stemMatches.length === 1 && stem.length >= 5) {
    return [{ asset: stemMatches[0]!, evidence: value, confidence: 0.42 }];
  }

  return [];
}

function countUniqueReferencedAssets(slides: Slide[]): number {
  const assetIds = new Set<string>();
  for (const slide of slides) {
    for (const object of slide.objects) {
      if ((object.type === "image" || object.type === "media") && object.source.assetId) {
        assetIds.add(object.source.assetId);
      }
    }
  }
  return assetIds.size;
}

function addTextCandidate(candidates: Set<string>, text: string): void {
  const cleaned = cleanTextCandidate(text);
  if (cleaned) {
    candidates.add(cleaned);
  }
}

function expandIwaPayloads(data: Uint8Array): ExpandedIwaPayload[] {
  const payloads: ExpandedIwaPayload[] = [{ data, compression: "none" }];
  const framed = decodeSnappyFramedStream(data);
  if (framed) {
    payloads.push({ data: framed, compression: "snappy-framed" });
  }

  const rawSnappy = decodeSnappyBlock(data);
  if (rawSnappy) {
    payloads.push({ data: rawSnappy, compression: "snappy-block" });
  }

  for (const decoded of decodeLengthPrefixedSnappyBlocks(data)) {
    payloads.push({ data: decoded, compression: "length-prefixed-snappy" });
  }

  return dedupeExpandedPayloads(payloads);
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

function classifyAssetPath(assetPath: string): { kind: Asset["kind"]; mimeType?: string } {
  const extension = path.posix.extname(assetPath).toLowerCase();
  switch (extension) {
    case ".png":
      return { kind: "image", mimeType: "image/png" };
    case ".jpg":
    case ".jpeg":
      return { kind: "image", mimeType: "image/jpeg" };
    case ".gif":
      return { kind: "image", mimeType: "image/gif" };
    case ".tif":
    case ".tiff":
      return { kind: "image", mimeType: "image/tiff" };
    case ".heic":
      return { kind: "image", mimeType: "image/heic" };
    case ".webp":
      return { kind: "image", mimeType: "image/webp" };
    case ".mp4":
      return { kind: "video", mimeType: "video/mp4" };
    case ".m4v":
      return { kind: "video", mimeType: "video/x-m4v" };
    case ".mov":
      return { kind: "video", mimeType: "video/quicktime" };
    case ".mp3":
      return { kind: "audio", mimeType: "audio/mpeg" };
    case ".m4a":
      return { kind: "audio", mimeType: "audio/mp4" };
    case ".wav":
      return { kind: "audio", mimeType: "audio/wav" };
    case ".aac":
      return { kind: "audio", mimeType: "audio/aac" };
    case ".pdf":
      return { kind: "data", mimeType: "application/pdf" };
    default:
      return { kind: "other" };
  }
}

function stableAssetId(assetPath: string): string {
  return `asset-${sha256Hex(new TextEncoder().encode(assetPath)).slice(0, 16)}`;
}

function sha256Hex(data: Uint8Array): string {
  return createHash("sha256").update(data).digest("hex");
}

function addMultiMapValue<T>(map: Map<string, T[]>, key: string, value: T): void {
  const values = map.get(key);
  if (values) {
    values.push(value);
  } else {
    map.set(key, [value]);
  }
}

function dedupeExpandedPayloads(payloads: ExpandedIwaPayload[]): ExpandedIwaPayload[] {
  const seen = new Set<string>();
  const deduped: ExpandedIwaPayload[] = [];
  for (const payload of payloads) {
    const array = payload.data;
    const key = `${array.byteLength}:${array[0] ?? 0}:${array[1] ?? 0}:${array[array.byteLength - 1] ?? 0}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    deduped.push(payload);
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
