import { createHash } from "node:crypto";
import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";
import { inflateRawSync } from "node:zlib";

import { IR_VERSION, type Asset, type DeckIR, type IRObject, type JSONRecord, type Slide } from "../ir/index.ts";

const DEFAULT_DECK_SIZE = { width: 1280, height: 720 };
const MAX_TEXT_CANDIDATES_PER_STREAM = 80;
const MAX_TEXT_OBJECTS_PER_SLIDE = 24;
const MAX_REFERENCE_CANDIDATES_PER_STREAM = 160;
const MAX_PROTOBUF_FIELDS_PER_PAYLOAD = 4000;
const MAX_NESTED_PROTOBUF_DEPTH = 3;
const MAX_NESTED_PROTOBUF_BYTES = 256 * 1024;
const MAX_FIELD_SUMMARIES_PER_STREAM = 48;
const MAX_SNAPPY_OUTPUT_BYTES = 64 * 1024 * 1024;

export interface NativeKeynoteDetection {
  isNative: boolean;
  container: "directory" | "zip";
  packageFormat: NativeKeynotePackageFormat;
  entryCount: number;
  hasIndexZip: boolean;
  hasLooseIndexDirectory: boolean;
  hasDataDirectory: boolean;
  hasQuickLookPreview: boolean;
  indexZipEntryCount: number;
  iwaPaths: string[];
  metadataPaths: string[];
  assetPaths: string[];
  quickLookPaths: string[];
  iwaStreams?: NativeIwaStreamMetadata[];
}

export type NativeKeynotePackageFormat =
  | "directory-index-zip"
  | "directory-loose-index"
  | "directory-mixed"
  | "zip-index-zip"
  | "zip-loose-index"
  | "zip-mixed"
  | "unknown";

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
  quickLookPaths: string[];
  indexZipEntryCount: number;
  hasLooseIndexDirectory: boolean;
  hasDataDirectory: boolean;
  hasQuickLookPreview: boolean;
}

export interface NativeIwaStreamMetadata {
  path: string;
  role: NativeIwaStreamRole;
  compression: NativeIwaCompression[];
  byteLength: number;
  expandedByteLength: number;
  textCandidateCount: number;
  referenceCandidateCount: number;
  assetReferenceCount: number;
  protobufFieldCount: number;
  protobufFieldPathCount: number;
  nestedMessageCount: number;
  rawStringCount: number;
  fieldSummaries: NativeIwaFieldSummary[];
}

export type NativeIwaStreamRole = "slide" | "document" | "theme" | "master" | "layout" | "style" | "unknown";
export type NativeIwaCompression = "none" | "snappy-framed" | "snappy-block" | "length-prefixed-snappy";

export interface NativeIwaFieldSummary {
  fieldPath: string;
  fieldNumber: number;
  wireType: number;
  occurrences: number;
  textCandidateCount: number;
  referenceCandidateCount: number;
  nestedMessageCount: number;
  minLength?: number;
  maxLength?: number;
  sampleText?: string;
  sampleReference?: string;
}

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
  fieldPath?: string;
  source: "protobuf" | "raw";
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
  documentIdentifier?: string;
  paths: string[];
  values: JSONRecord;
}

interface IwaTextCandidate {
  text: string;
  source: "protobuf" | "raw";
  confidence: number;
  order: number;
  fieldPath?: string;
}

interface IwaReferenceCandidate {
  value: string;
  source: "protobuf" | "raw";
  confidence: number;
  order: number;
  fieldPath?: string;
}

interface IwaScanResult {
  textCandidates: IwaTextCandidate[];
  referenceCandidates: IwaReferenceCandidate[];
  fieldSummaries: NativeIwaFieldSummary[];
  protobufFieldCount: number;
  nestedMessageCount: number;
  rawStringCount: number;
  expandedByteLength: number;
}

export async function detectNativeKeynotePackage(keynotePath: string): Promise<NativeKeynoteDetection> {
  const keynotePackage = await readKeynotePackage(keynotePath);
  const parts = readNativeKeynoteParts(keynotePackage.entries);
  const assets = createNativeAssetCatalog(parts.assetPaths, parts.entries);
  return {
    isNative: parts.hasIndexZip || parts.iwaEntries.length > 0,
    container: keynotePackage.container,
    packageFormat: classifyPackageFormat(keynotePackage.container, parts),
    entryCount: parts.entries.size,
    hasIndexZip: parts.hasIndexZip,
    hasLooseIndexDirectory: parts.hasLooseIndexDirectory,
    hasDataDirectory: parts.hasDataDirectory,
    hasQuickLookPreview: parts.hasQuickLookPreview,
    indexZipEntryCount: parts.indexZipEntryCount,
    iwaPaths: parts.iwaEntries.map((entry) => entry.path),
    metadataPaths: parts.metadataPaths,
    assetPaths: parts.assetPaths,
    quickLookPaths: parts.quickLookPaths,
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
  const packageFormat = classifyPackageFormat(keynotePackage.container, parts);
  const totalProtobufFieldCount = iwaStreams.reduce((total, stream) => total + stream.protobufFieldCount, 0);
  const totalNestedMessageCount = iwaStreams.reduce((total, stream) => total + stream.nestedMessageCount, 0);
  const totalReferenceCandidateCount = iwaStreams.reduce((total, stream) => total + stream.referenceCandidateCount, 0);
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
    {
      code: "keynote-native-protobuf-schema-private",
      severity: "warning" as const,
      area: "unknown" as const,
      description:
        totalProtobufFieldCount > 0
          ? `Scanned ${totalProtobufFieldCount} protobuf-like field(s) across ${parts.iwaEntries.length} IWA stream(s), including ${totalNestedMessageCount} nested length-delimited message candidate(s), but Keynote's private field meanings are not public.`
          : "No protobuf-like fields were confidently decoded from the IWA streams.",
      fallback: "Preserve field-path summaries and recover only visible strings and package references."
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
        totalProtobufFieldCount > 0
          ? "Text and asset references are inferred from protobuf field-path scans, nested length-delimited message candidates, and raw string scans, not from public Keynote object schemas."
          : "Text and asset references are inferred from raw string scans because no protobuf-like field structure was confidently decoded.",
      confidence: totalProtobufFieldCount > 0 ? 0.42 : 0.28
    },
    {
      code: "keynote-native-container-detection",
      severity: packageFormat === "unknown" ? ("warning" as const) : ("info" as const),
      description:
        packageFormat === "unknown"
          ? "The file could be read, but it does not match common native Keynote package container patterns."
          : `Detected native Keynote container pattern "${packageFormat}" with ${parts.entries.size} package entr${parts.entries.size === 1 ? "y" : "ies"}.`,
      confidence: packageFormat === "unknown" ? 0.25 : 0.76
    },
    ...(recoveredAssetObjectCount > 0
      ? [
          {
            code: "keynote-native-asset-reference-scan",
            severity: "warning" as const,
            description:
              `Asset-to-slide associations are based on ${totalReferenceCandidateCount} reference-like string candidate(s) found in IWA payloads; this does not prove exact placement or usage.`,
            confidence: 0.5
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
        nativePackageFormat: packageFormat,
        nativeIndexZip: parts.hasIndexZip,
        nativeIwaStreamCount: parts.iwaEntries.length,
        nativeAssetCount: assets.assets.length,
        nativeMetadataPaths: metadata.paths,
        ...(metadata.documentIdentifier ? { nativeDocumentIdentifier: metadata.documentIdentifier } : {})
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
            `Parsed a native Keynote package directly by inspecting ${packageFormat} IWA streams; layout and animation fidelity are not available without Keynote.`
        },
        {
          severity: totalProtobufFieldCount > 0 ? "info" : "warning",
          code: totalProtobufFieldCount > 0 ? "keynote-native-iwa-fields-scanned" : "keynote-native-iwa-fields-not-decoded",
          message:
            totalProtobufFieldCount > 0
              ? `Scanned ${totalProtobufFieldCount} protobuf-like IWA field(s), including ${totalNestedMessageCount} nested message candidate(s).`
              : "No protobuf-like IWA fields were confidently decoded; recovery used raw string scanning only."
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
        packageFormat,
        entryCount: parts.entries.size,
        hasIndexZip: parts.hasIndexZip,
        hasLooseIndexDirectory: parts.hasLooseIndexDirectory,
        hasDataDirectory: parts.hasDataDirectory,
        hasQuickLookPreview: parts.hasQuickLookPreview,
        indexZipEntryCount: parts.indexZipEntryCount,
        quickLookPathCount: parts.quickLookPaths.length,
        iwaPathCount: parts.iwaEntries.length,
        metadataPathCount: parts.metadataPaths.length,
        assetPathCount: parts.assetPaths.length,
        totalProtobufFieldCount,
        totalNestedMessageCount,
        totalReferenceCandidateCount,
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
          referenceCandidateCount: stream.referenceCandidateCount,
          assetReferenceCount: stream.assetReferenceCount,
          protobufFieldCount: stream.protobufFieldCount,
          protobufFieldPathCount: stream.protobufFieldPathCount,
          nestedMessageCount: stream.nestedMessageCount,
          rawStringCount: stream.rawStringCount,
          fieldSummaries: stream.fieldSummaries
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
  let hasLooseIndexDirectory = false;

  for (const [entryPath, data] of entries) {
    const normalized = normalizePartPath(entryPath);
    combinedEntries.set(normalized, data);
    if (/^Index\/.+\.iwa$/i.test(normalized)) {
      hasLooseIndexDirectory = true;
    }
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
  const quickLookPaths = allPaths.filter((entryPath) => /^QuickLook\//i.test(entryPath));
  return {
    hasIndexZip,
    entries: combinedEntries,
    iwaEntries: allPaths
      .filter((entryPath) => entryPath.toLowerCase().endsWith(".iwa"))
      .map((entryPath) => ({ path: entryPath, data: combinedEntries.get(entryPath)! })),
    metadataPaths: allPaths.filter((entryPath) => isMetadataPath(entryPath)),
    assetPaths: allPaths.filter((entryPath) => entryPath.startsWith("Data/")),
    quickLookPaths,
    indexZipEntryCount,
    hasLooseIndexDirectory,
    hasDataDirectory: allPaths.some((entryPath) => entryPath.startsWith("Data/")),
    hasQuickLookPreview: quickLookPaths.length > 0
  };
}

function classifyPackageFormat(container: KeynotePackage["container"], parts: NativeKeynoteParts): NativeKeynotePackageFormat {
  const prefix = container === "directory" ? "directory" : "zip";
  if (parts.hasIndexZip && parts.hasLooseIndexDirectory) {
    return `${prefix}-mixed`;
  }
  if (parts.hasIndexZip) {
    return `${prefix}-index-zip`;
  }
  if (parts.hasLooseIndexDirectory || parts.iwaEntries.length > 0) {
    return `${prefix}-loose-index`;
  }
  return "unknown";
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
    addMultiMapValue(byStem, normalizeAssetStem(path.posix.basename(name, path.posix.extname(name))), reference);
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
    const scan = scanIwaPayloads(payloads);
    const compression = Array.from(new Set(payloads.map((payload) => payload.compression)));
    return {
      path: entry.path,
      role: classifyIwaRole(entry.path),
      compression,
      byteLength: entry.data.byteLength,
      expandedByteLength: scan.expandedByteLength,
      textCandidateCount: scan.textCandidates.length,
      referenceCandidateCount: scan.referenceCandidates.length,
      assetReferenceCount: findIwaAssetMatchesFromScan(scan, assets).length,
      protobufFieldCount: scan.protobufFieldCount,
      protobufFieldPathCount: scan.fieldSummaries.length,
      nestedMessageCount: scan.nestedMessageCount,
      rawStringCount: scan.rawStringCount,
      fieldSummaries: scan.fieldSummaries
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
  const paths: string[] = [];
  for (const [entryPath, data] of entries) {
    if (!isMetadataPath(entryPath)) {
      continue;
    }
    const text = decodeUtf8(data);
    if (!text || !text.includes("<plist")) {
      continue;
    }
    paths.push(entryPath);
    const plistValues = parseXmlPlistValues(text);
    for (const [key, value] of plistValues) {
      values[key] = value;
    }
  }

  return {
    title: stringValue(values, ["title", "Title", "documentTitle", "DocumentTitle", "name", "Name", "kMDItemTitle"]),
    author: stringValue(values, ["author", "Author", "creator", "Creator", "kMDItemAuthors", "NSHumanReadableCopyright"]),
    createdAt: stringValue(values, ["createdAt", "creationDate", "CreationDate", "dateCreated", "kMDItemFSCreationDate"]),
    updatedAt: stringValue(values, [
      "updatedAt",
      "modificationDate",
      "ModificationDate",
      "lastModifiedDate",
      "kMDItemFSContentChangeDate",
      "kMDItemLastUsedDate"
    ]),
    size: readDeckSize(values),
    documentIdentifier: stringValue(values, [
      "documentIdentifier",
      "DocumentIdentifier",
      "documentUUID",
      "DocumentUUID",
      "NSDocumentIdentifier"
    ]),
    paths: paths.sort(comparePartPaths),
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
    const scan = scanIwaEntry(entry.data);
    const textCandidates = scan.textCandidates.slice(0, MAX_TEXT_OBJECTS_PER_SLIDE);
    const assetMatches = findIwaAssetMatchesFromScan(scan, assets);
    const textObjects = textCandidates.map((candidate, objectIndex) =>
      createTextObject(slideId, candidate, objectIndex, entry.path, deckSize)
    );
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
        nativeReferenceCandidateCount: scan.referenceCandidates.length,
        nativeAssetReferenceCount: assetMatches.length,
        nativeProtobufFieldCount: scan.protobufFieldCount,
        nativeNestedMessageCount: scan.nestedMessageCount,
        nativeFieldSummaries: scan.fieldSummaries.slice(0, 12),
        nativeParser: "iwa-protobuf-field-scan"
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
      nativeParser: "iwa-protobuf-field-scan"
    }
  };
}

function createTextObject(
  slideId: string,
  candidate: IwaTextCandidate,
  objectIndex: number,
  sourcePath: string,
  deckSize: { width: number; height: number }
): IRObject {
  const text = candidate.text;
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
      nativeExtraction: candidate.source === "protobuf" ? "protobuf-field-string" : "raw-string",
      nativeFieldPath: candidate.fieldPath,
      nativeTextConfidence: candidate.confidence
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
    nativeAssetFieldPath: match.fieldPath,
    nativeExtraction: match.source === "protobuf" ? "asset-protobuf-field-string-scan" : "asset-raw-string-scan"
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

function scanIwaEntry(data: Uint8Array): IwaScanResult {
  return scanIwaPayloads(expandIwaPayloads(data));
}

function scanIwaPayloads(payloads: ExpandedIwaPayload[]): IwaScanResult {
  const textCandidates = new Map<string, IwaTextCandidate>();
  const referenceCandidates = new Map<string, IwaReferenceCandidate>();
  const fieldSummaries = new Map<string, NativeIwaFieldSummary>();
  let protobufFieldCount = 0;
  let nestedMessageCount = 0;
  let rawStringCount = 0;
  let orderBase = 0;

  for (const payload of payloads) {
    const protobufScan = scanProtobufPayload(payload.data, orderBase);
    protobufFieldCount += protobufScan.protobufFieldCount;
    nestedMessageCount += protobufScan.nestedMessageCount;
    for (const summary of protobufScan.fieldSummaries) {
      mergeFieldSummary(fieldSummaries, summary);
    }
    for (const candidate of protobufScan.textCandidates) {
      addBestTextCandidate(textCandidates, candidate);
    }
    for (const candidate of protobufScan.referenceCandidates) {
      addBestReferenceCandidate(referenceCandidates, candidate);
    }

    const rawScan = scanRawStrings(payload.data, orderBase + payload.data.byteLength + 1);
    rawStringCount += rawScan.rawStringCount;
    for (const candidate of rawScan.textCandidates) {
      addBestTextCandidate(textCandidates, candidate);
    }
    for (const candidate of rawScan.referenceCandidates) {
      addBestReferenceCandidate(referenceCandidates, candidate);
    }
    orderBase += payload.data.byteLength * 2 + 2;
  }

  const fieldSummaryList = Array.from(fieldSummaries.values()).sort(compareFieldSummaries).slice(0, MAX_FIELD_SUMMARIES_PER_STREAM);
  return {
    textCandidates: Array.from(textCandidates.values()).sort(compareTextCandidates).slice(0, MAX_TEXT_CANDIDATES_PER_STREAM),
    referenceCandidates: Array.from(referenceCandidates.values())
      .sort(compareReferenceCandidates)
      .slice(0, MAX_REFERENCE_CANDIDATES_PER_STREAM),
    fieldSummaries: fieldSummaryList,
    protobufFieldCount,
    nestedMessageCount,
    rawStringCount,
    expandedByteLength: payloads.reduce((max, payload) => Math.max(max, payload.data.byteLength), 0)
  };
}

function findIwaAssetMatchesFromScan(scan: IwaScanResult, assets: NativeAssetCatalog): NativeAssetMatch[] {
  if (assets.assets.length === 0) {
    return [];
  }

  const matches = new Map<string, NativeAssetMatch>();
  for (const candidate of scan.referenceCandidates) {
    for (const match of matchAssetReference(candidate, assets)) {
      const existing = matches.get(match.asset.assetId);
      if (!existing || match.confidence > existing.confidence) {
        matches.set(match.asset.assetId, match);
      }
    }
  }

  return Array.from(matches.values()).sort((left, right) => comparePartPaths(left.asset.path, right.asset.path));
}

function scanProtobufPayload(
  data: Uint8Array,
  orderBase: number
): Pick<
  IwaScanResult,
  "textCandidates" | "referenceCandidates" | "fieldSummaries" | "protobufFieldCount" | "nestedMessageCount"
> {
  const state = {
    textCandidates: new Map<string, IwaTextCandidate>(),
    referenceCandidates: new Map<string, IwaReferenceCandidate>(),
    fieldSummaries: new Map<string, NativeIwaFieldSummary>(),
    protobufFieldCount: 0,
    nestedMessageCount: 0,
    orderBase
  };
  scanProtobufFields(data, "", 0, state);
  return {
    textCandidates: Array.from(state.textCandidates.values()),
    referenceCandidates: Array.from(state.referenceCandidates.values()),
    fieldSummaries: Array.from(state.fieldSummaries.values()),
    protobufFieldCount: state.protobufFieldCount,
    nestedMessageCount: state.nestedMessageCount
  };
}

function scanProtobufFields(
  data: Uint8Array,
  prefix: string,
  depth: number,
  state: {
    textCandidates: Map<string, IwaTextCandidate>;
    referenceCandidates: Map<string, IwaReferenceCandidate>;
    fieldSummaries: Map<string, NativeIwaFieldSummary>;
    protobufFieldCount: number;
    nestedMessageCount: number;
    orderBase: number;
  }
): void {
  for (
    let offset = 0;
    offset < data.length && state.protobufFieldCount < MAX_PROTOBUF_FIELDS_PER_PAYLOAD;
    offset += 1
  ) {
    const field = readProtobufFieldAt(data, offset);
    if (!field) {
      continue;
    }

    state.protobufFieldCount += 1;
    const fieldPath = prefix ? `${prefix}.${field.fieldNumber}` : String(field.fieldNumber);
    const summary = getFieldSummary(state.fieldSummaries, fieldPath, field.fieldNumber, field.wireType);
    summary.occurrences += 1;

    if (field.wireType !== 2 || !field.value) {
      continue;
    }

    summary.minLength = summary.minLength === undefined ? field.value.byteLength : Math.min(summary.minLength, field.value.byteLength);
    summary.maxLength = summary.maxLength === undefined ? field.value.byteLength : Math.max(summary.maxLength, field.value.byteLength);

    const text = decodeTextCandidate(field.value, 2);
    if (text) {
      summary.textCandidateCount += 1;
      summary.sampleText ??= text;
      addBestTextCandidate(state.textCandidates, {
        text,
        source: "protobuf",
        confidence: depth > 0 ? 0.68 : 0.62,
        order: state.orderBase + offset,
        fieldPath
      });
    }

    const reference = decodeReferenceCandidate(field.value);
    if (reference) {
      summary.referenceCandidateCount += 1;
      summary.sampleReference ??= reference;
      addBestReferenceCandidate(state.referenceCandidates, {
        value: reference,
        source: "protobuf",
        confidence: depth > 0 ? 0.76 : 0.7,
        order: state.orderBase + offset,
        fieldPath
      });
    }

    if (depth >= MAX_NESTED_PROTOBUF_DEPTH || field.value.byteLength > MAX_NESTED_PROTOBUF_BYTES) {
      continue;
    }
    if (!looksLikeNestedProtobufMessage(field.value)) {
      continue;
    }

    summary.nestedMessageCount += 1;
    state.nestedMessageCount += 1;
    scanProtobufFields(field.value, fieldPath, depth + 1, state);
  }
}

function readProtobufFieldAt(
  data: Uint8Array,
  offset: number
): { fieldNumber: number; wireType: number; nextOffset: number; value?: Uint8Array } | undefined {
  const key = readVarint(data, offset);
  if (!key || key.nextOffset <= offset) {
    return undefined;
  }
  const wireType = key.value & 0x07;
  const fieldNumber = key.value >>> 3;
  if (fieldNumber <= 0 || fieldNumber > 8191 || ![0, 1, 2, 5].includes(wireType)) {
    return undefined;
  }

  if (wireType === 0) {
    const value = readVarint(data, key.nextOffset);
    if (!value) return undefined;
    return { fieldNumber, wireType, nextOffset: value.nextOffset };
  }

  if (wireType === 1) {
    const nextOffset = key.nextOffset + 8;
    return nextOffset <= data.length ? { fieldNumber, wireType, nextOffset } : undefined;
  }

  if (wireType === 5) {
    const nextOffset = key.nextOffset + 4;
    return nextOffset <= data.length ? { fieldNumber, wireType, nextOffset } : undefined;
  }

  const lengthInfo = readVarint(data, key.nextOffset);
  if (!lengthInfo || lengthInfo.value <= 0 || lengthInfo.value > MAX_NESTED_PROTOBUF_BYTES) {
    return undefined;
  }
  const valueStart = lengthInfo.nextOffset;
  const valueEnd = valueStart + lengthInfo.value;
  if (valueEnd > data.length) {
    return undefined;
  }
  return { fieldNumber, wireType, nextOffset: valueEnd, value: data.subarray(valueStart, valueEnd) };
}

function looksLikeNestedProtobufMessage(data: Uint8Array): boolean {
  if (data.byteLength < 4) {
    return false;
  }
  let offset = 0;
  let fieldCount = 0;
  let lengthDelimitedCount = 0;

  while (offset < data.length && fieldCount < 80) {
    const field = readProtobufFieldAt(data, offset);
    if (!field || field.nextOffset <= offset) {
      return false;
    }
    fieldCount += 1;
    if (field.wireType === 2) {
      lengthDelimitedCount += 1;
    }
    offset = field.nextOffset;
  }

  return offset === data.length && fieldCount >= 2 && lengthDelimitedCount > 0;
}

function scanRawStrings(data: Uint8Array, orderBase: number): {
  textCandidates: IwaTextCandidate[];
  referenceCandidates: IwaReferenceCandidate[];
  rawStringCount: number;
} {
  const textCandidates = new Map<string, IwaTextCandidate>();
  const referenceCandidates = new Map<string, IwaReferenceCandidate>();
  let rawStringCount = 0;
  let start = -1;

  for (let index = 0; index <= data.length; index += 1) {
    const byte = data[index];
    const printable = byte !== undefined && (byte === 0x09 || byte === 0x0a || byte === 0x0d || byte >= 0x20);
    if (printable) {
      if (start < 0) start = index;
      continue;
    }

    if (start >= 0 && index - start >= 4) {
      rawStringCount += 1;
      const bytes = data.subarray(start, index);
      const text = decodeTextCandidate(bytes, 4);
      if (text) {
        addBestTextCandidate(textCandidates, { text, source: "raw", confidence: 0.36, order: orderBase + start });
      }

      const reference = decodeReferenceCandidate(bytes);
      if (reference) {
        addBestReferenceCandidate(referenceCandidates, { value: reference, source: "raw", confidence: 0.48, order: orderBase + start });
      }
    }
    start = -1;
  }

  return {
    textCandidates: Array.from(textCandidates.values()),
    referenceCandidates: Array.from(referenceCandidates.values()),
    rawStringCount
  };
}

function decodeReferenceCandidate(bytes: Uint8Array): string | undefined {
  const decoded = decodeUtf8(bytes);
  if (!decoded || decoded.includes("\ufffd")) {
    return undefined;
  }

  const text = decoded.replace(/\u0000/g, "").trim();
  if (text.length < 4 || text.length > 1024) {
    return undefined;
  }
  if (!/[\w .@()%-]+\.[a-z0-9]{2,5}/i.test(text)) {
    return undefined;
  }
  return text;
}

function matchAssetReference(candidate: IwaReferenceCandidate, assets: NativeAssetCatalog): NativeAssetMatch[] {
  for (const key of assetPathKeysFromReference(candidate.value)) {
    const exact = assets.byPath.get(key);
    if (exact) {
      return [
        {
          asset: exact,
          evidence: candidate.value,
          confidence: Math.min(0.95, candidate.confidence + (key.startsWith("data/") ? 0.14 : 0.08)),
          fieldPath: candidate.fieldPath,
          source: candidate.source
        }
      ];
    }
  }

  const normalizedValue = normalizeReferenceText(candidate.value).toLowerCase();
  const baseName = path.posix.basename(normalizedValue);
  const nameMatches = baseName ? assets.byName.get(baseName) ?? [] : [];
  if (nameMatches.length === 1) {
    return [
      {
        asset: nameMatches[0]!,
        evidence: candidate.value,
        confidence: Math.min(0.82, candidate.confidence + 0.08),
        fieldPath: candidate.fieldPath,
        source: candidate.source
      }
    ];
  }
  if (nameMatches.length > 1) {
    return nameMatches.map((asset) => ({
      asset,
      evidence: candidate.value,
      confidence: Math.min(0.68, candidate.confidence),
      fieldPath: candidate.fieldPath,
      source: candidate.source
    }));
  }

  const stem = normalizeAssetStem(path.posix.basename(baseName, path.posix.extname(baseName)));
  const stemMatches = stem ? assets.byStem.get(stem) ?? [] : [];
  if (stemMatches.length === 1 && stem.length >= 5) {
    return [
      {
        asset: stemMatches[0]!,
        evidence: candidate.value,
        confidence: Math.min(0.64, candidate.confidence - 0.04),
        fieldPath: candidate.fieldPath,
        source: candidate.source
      }
    ];
  }

  return [];
}

function assetPathKeysFromReference(value: string): string[] {
  const keys = new Set<string>();
  const normalized = normalizeReferenceText(value);
  for (const candidate of [normalized, safeDecodeUriComponent(normalized)]) {
    const withoutQuery = candidate.split(/[?#]/, 1)[0] ?? candidate;
    const partPath = normalizePartPath(withoutQuery).toLowerCase();
    if (partPath) {
      keys.add(partPath);
    }
    const dataIndex = partPath.toLowerCase().lastIndexOf("/data/");
    if (dataIndex >= 0) {
      keys.add(partPath.slice(dataIndex + 1));
    }
    if (partPath.toLowerCase().startsWith("data/")) {
      keys.add(partPath);
    }
  }
  return Array.from(keys);
}

function normalizeReferenceText(value: string): string {
  return value.replace(/\u0000/g, "").trim().replace(/^['"]+|['"]+$/g, "").replaceAll("\\", "/");
}

function safeDecodeUriComponent(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
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

function getFieldSummary(
  summaries: Map<string, NativeIwaFieldSummary>,
  fieldPath: string,
  fieldNumber: number,
  wireType: number
): NativeIwaFieldSummary {
  const existing = summaries.get(fieldPath);
  if (existing) {
    return existing;
  }
  const summary: NativeIwaFieldSummary = {
    fieldPath,
    fieldNumber,
    wireType,
    occurrences: 0,
    textCandidateCount: 0,
    referenceCandidateCount: 0,
    nestedMessageCount: 0
  };
  summaries.set(fieldPath, summary);
  return summary;
}

function mergeFieldSummary(summaries: Map<string, NativeIwaFieldSummary>, incoming: NativeIwaFieldSummary): void {
  const summary = getFieldSummary(summaries, incoming.fieldPath, incoming.fieldNumber, incoming.wireType);
  summary.occurrences += incoming.occurrences;
  summary.textCandidateCount += incoming.textCandidateCount;
  summary.referenceCandidateCount += incoming.referenceCandidateCount;
  summary.nestedMessageCount += incoming.nestedMessageCount;
  summary.minLength =
    summary.minLength === undefined
      ? incoming.minLength
      : incoming.minLength === undefined
        ? summary.minLength
        : Math.min(summary.minLength, incoming.minLength);
  summary.maxLength =
    summary.maxLength === undefined
      ? incoming.maxLength
      : incoming.maxLength === undefined
        ? summary.maxLength
        : Math.max(summary.maxLength, incoming.maxLength);
  summary.sampleText ??= incoming.sampleText;
  summary.sampleReference ??= incoming.sampleReference;
}

function addBestTextCandidate(candidates: Map<string, IwaTextCandidate>, candidate: IwaTextCandidate): void {
  const cleaned = cleanTextCandidate(candidate.text);
  if (!cleaned) {
    return;
  }
  const key = cleaned.toLowerCase();
  const normalized = { ...candidate, text: cleaned };
  const existing = candidates.get(key);
  if (!existing || compareTextCandidates(normalized, existing) < 0) {
    candidates.set(key, normalized);
  }
}

function addBestReferenceCandidate(candidates: Map<string, IwaReferenceCandidate>, candidate: IwaReferenceCandidate): void {
  const reference = normalizeReferenceText(candidate.value);
  if (!reference) {
    return;
  }
  const key = safeDecodeUriComponent(reference).toLowerCase();
  const normalized = { ...candidate, value: reference };
  const existing = candidates.get(key);
  if (!existing || compareReferenceCandidates(normalized, existing) < 0) {
    candidates.set(key, normalized);
  }
}

function compareTextCandidates(left: IwaTextCandidate, right: IwaTextCandidate): number {
  const confidence = right.confidence - left.confidence;
  if (confidence !== 0) {
    return confidence;
  }
  const source = sourceRank(right.source) - sourceRank(left.source);
  if (source !== 0) {
    return source;
  }
  if (left.order !== right.order) {
    return left.order - right.order;
  }
  if (left.fieldPath && right.fieldPath) {
    const fieldPath = compareFieldPaths(left.fieldPath, right.fieldPath);
    if (fieldPath !== 0) return fieldPath;
  } else if (left.fieldPath) {
    return -1;
  } else if (right.fieldPath) {
    return 1;
  }
  return left.text.localeCompare(right.text);
}

function compareReferenceCandidates(left: IwaReferenceCandidate, right: IwaReferenceCandidate): number {
  const confidence = right.confidence - left.confidence;
  if (confidence !== 0) {
    return confidence;
  }
  const source = sourceRank(right.source) - sourceRank(left.source);
  if (source !== 0) {
    return source;
  }
  if (left.order !== right.order) {
    return left.order - right.order;
  }
  if (left.fieldPath && right.fieldPath) {
    const fieldPath = compareFieldPaths(left.fieldPath, right.fieldPath);
    if (fieldPath !== 0) return fieldPath;
  } else if (left.fieldPath) {
    return -1;
  } else if (right.fieldPath) {
    return 1;
  }
  return left.value.localeCompare(right.value);
}

function compareFieldSummaries(left: NativeIwaFieldSummary, right: NativeIwaFieldSummary): number {
  const leftScore = left.textCandidateCount * 4 + left.referenceCandidateCount * 4 + left.nestedMessageCount * 2 + left.occurrences;
  const rightScore = right.textCandidateCount * 4 + right.referenceCandidateCount * 4 + right.nestedMessageCount * 2 + right.occurrences;
  if (leftScore !== rightScore) {
    return rightScore - leftScore;
  }
  return compareFieldPaths(left.fieldPath, right.fieldPath);
}

function compareFieldPaths(left: string, right: string): number {
  const leftParts = left.split(".").map((part) => Number(part));
  const rightParts = right.split(".").map((part) => Number(part));
  const length = Math.max(leftParts.length, rightParts.length);
  for (let index = 0; index < length; index += 1) {
    const leftValue = leftParts[index] ?? -1;
    const rightValue = rightParts[index] ?? -1;
    if (leftValue !== rightValue) {
      return leftValue - rightValue;
    }
  }
  return left.localeCompare(right);
}

function sourceRank(source: "protobuf" | "raw"): number {
  return source === "protobuf" ? 2 : 1;
}

function normalizeAssetStem(value: string): string {
  return safeDecodeUriComponent(value).toLowerCase().replace(/[^a-z0-9]+/g, "");
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
