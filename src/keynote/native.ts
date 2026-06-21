import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { readdir, readFile, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { inflateRawSync } from "node:zlib";

import {
  IR_VERSION,
  type AnimationEvent,
  type Asset,
  type DeckIR,
  type IRObject,
  type JSONRecord,
  type KeyframeTrack,
  type Slide,
  type TimingDependencyEdge
} from "../ir/index.ts";

const DEFAULT_DECK_SIZE = { width: 1280, height: 720 };
const MAX_TEXT_CANDIDATES_PER_STREAM = 80;
const MAX_TEXT_OBJECTS_PER_SLIDE = 24;
const MAX_REFERENCE_CANDIDATES_PER_STREAM = 160;
const MAX_NUMERIC_CANDIDATES_PER_STREAM = 160;
const MAX_GEOMETRY_CANDIDATES_PER_STREAM = 24;
const MAX_GROUPING_HINTS_PER_STREAM = 24;
const MAX_ANIMATION_HINTS_PER_STREAM = 24;
const MAX_PROTOBUF_FIELDS_PER_PAYLOAD = 4000;
const MAX_NESTED_PROTOBUF_DEPTH = 3;
const MAX_NESTED_PROTOBUF_BYTES = 256 * 1024;
const MAX_FIELD_SUMMARIES_PER_STREAM = 48;
const MAX_ARCHIVE_RECORD_EVIDENCE_PER_STREAM = 48;
const MAX_TYPED_ARCHIVE_MESSAGE_EVIDENCE_PER_STREAM = 512;
const MAX_SNAPPY_OUTPUT_BYTES = 64 * 1024 * 1024;
const MAX_VISUAL_NUMERIC_VALUE = 100000;
const MAX_EMBEDDED_ASSET_BYTES = 2 * 1024 * 1024;
const MAX_RENDERABLE_NATIVE_ASSET_BYTES = 32 * 1024 * 1024;
const LOW_CONFIDENCE_TEXT_FALLBACK_THRESHOLD = 0.45;

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
  quickLookPreviews: NativeQuickLookPreviewMetadata[];
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
  quickLookPreviews: NativeQuickLookPreviewMetadata[];
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
  numericCandidateCount: number;
  geometryCandidateCount: number;
  groupingHintCount: number;
  animationHintCount: number;
  magicMoveHintCount: number;
  morphHintCount: number;
  assetReferenceCount: number;
  archiveRecordCount: number;
  archivePayloadCount: number;
  archiveRecords: NativeIwaArchiveRecordEvidence[];
  typedArchiveMessageCount: number;
  typedArchiveMessages: NativeIwaArchiveMessageEvidence[];
  buildRecordCount: number;
  buildTimingRecordCount: number;
  protobufFieldCount: number;
  protobufFieldPathCount: number;
  nestedMessageCount: number;
  rawStringCount: number;
  numericCandidates: NativeIwaNumericCandidate[];
  geometryCandidates: NativeIwaGeometryCandidate[];
  groupingHints: NativeIwaGroupingHint[];
  animationHints: NativeIwaAnimationHint[];
  fieldSummaries: NativeIwaFieldSummary[];
}

export type NativeIwaStreamRole = "slide" | "document" | "theme" | "master" | "layout" | "style" | "unknown";
export type NativeIwaCompression = "none" | "snappy-framed" | "snappy-block" | "length-prefixed-snappy" | "iwa-snappy-chunk";

export interface NativeIwaArchiveRecordEvidence {
  archiveOffset: number;
  archiveInfoLength: number;
  messageCount: number;
  payloadLength: number;
  messageTypes: number[];
  dataReferences: string[];
  objectReferences: string[];
  identifier?: string;
  shouldMerge?: boolean;
}

export interface NativeIwaArchiveMessageEvidence {
  archiveOffset: number;
  archiveIdentifier?: string;
  messageIndex: number;
  type?: number;
  payloadOffset: number;
  payloadLength: number;
  version: number[];
  objectReferences: string[];
  dataReferences: string[];
  textCandidates: string[];
  geometryCandidates: NativeIwaGeometryCandidate[];
  fieldSummaries: NativeIwaFieldSummary[];
  build?: NativeIwaBuildEvidence;
  buildTiming?: NativeIwaBuildTimingEvidence;
  textContent?: NativeIwaTextContentEvidence;
  textDrawable?: NativeIwaTextDrawableEvidence;
}

export interface NativeTypedVisualRecordEvidence {
  archiveIdentifier?: string;
  messageIndex: number;
  type: number;
  payloadOffset: number;
  payloadLength: number;
  objectReferences: string[];
  dataReferences: string[];
  geometryCandidates: NativeIwaGeometryCandidate[];
  fieldSummaries: NativeIwaFieldSummary[];
  numericCandidates: NativeIwaNumericCandidate[];
  confidence: number;
}

export interface NativeIwaBuildEvidence {
  kind: "build";
  buildId?: string;
  targetNativeId?: string;
  objectReferences: string[];
  delivery?: string;
  direction?: string;
  effect?: string;
  durationMs?: number;
  delayMs?: number;
  timingBase?: number;
  motionPath?: NativeIwaMotionPathEvidence;
  confidence: number;
  sourceFieldPaths: string[];
}

export interface NativeIwaBuildTimingEvidence {
  kind: "buildTiming";
  timingId?: string;
  buildId?: string;
  durationMs?: number;
  delayMs?: number;
  triggerGroupRaw?: number;
  group?: number;
  rawField5?: number;
  rawField6?: number;
  rawField72?: number;
  startRelation?: "withPrevious" | "afterPrevious" | "unknown";
  startsWithPrevious?: boolean;
  afterPrevious?: boolean;
  confidence: number;
  sourceFieldPaths: string[];
}

export interface NativeIwaMotionPathEvidence {
  kind: "motionPath";
  relative: boolean;
  points: Array<{ x: number; y: number }>;
  extentPoints?: Array<{ x: number; y: number }>;
  confidence: number;
  sourceFieldPaths: string[];
}

export interface NativeIwaTextContentEvidence {
  kind: "textContent";
  text: string;
  confidence: number;
  sourceFieldPaths: string[];
}

export interface NativeIwaTextDrawableEvidence {
  kind: "textDrawable";
  textArchiveIds: string[];
  slideArchiveId?: string;
  bounds?: { x: number; y: number; width: number; height: number };
  confidence: number;
  sourceFieldPaths: string[];
}

export interface NativeIwaFieldSummary {
  fieldPath: string;
  fieldNumber: number;
  wireType: number;
  occurrences: number;
  textCandidateCount: number;
  referenceCandidateCount: number;
  numericCandidateCount: number;
  nestedMessageCount: number;
  minLength?: number;
  maxLength?: number;
  minNumericValue?: number;
  maxNumericValue?: number;
  sampleNumericValue?: number;
  sampleNumericEncoding?: NativeIwaNumericEncoding;
  sampleText?: string;
  sampleReference?: string;
}

export type NativeIwaNumericEncoding = "varint" | "fixed32-float" | "fixed32-uint" | "fixed64-double";

export interface NativeIwaNumericCandidate {
  fieldPath: string;
  fieldNumber: number;
  wireType: number;
  value: number;
  encoding: NativeIwaNumericEncoding;
  confidence: number;
}

export interface NativeIwaGeometryCandidate {
  bounds: { x: number; y: number; width: number; height: number };
  fieldPaths: string[];
  values: number[];
  source: "protobuf";
  confidence: number;
  groupPath?: string;
  reason: string;
}

export interface NativeIwaGroupingHint {
  groupPath: string;
  titleCandidate?: string;
  textCandidateCount: number;
  referenceCandidateCount: number;
  geometryCandidateCount: number;
  animationHintCount: number;
  fieldPaths: string[];
  confidence: number;
}

export interface NativeIwaAnimationHint {
  kind: "magicMove" | "morph" | "transition" | "build";
  evidence: string;
  source: "protobuf" | "raw";
  confidence: number;
  fieldPath?: string;
}

export interface NativeQuickLookPreviewMetadata {
  path: string;
  byteLength: number;
  role: "thumbnail" | "preview" | "unknown";
  mimeType?: string;
  width?: number;
  height?: number;
}

interface NativeAssetReference {
  assetId: string;
  path: string;
  name: string;
  kind: Asset["kind"];
  mimeType?: string;
  sourceMimeType?: string;
  width?: number;
  height?: number;
  dataId?: string;
  embedded?: boolean;
  renderProxy?: NativeAssetRenderProxy;
}

interface NativeAssetRenderProxy {
  mimeType: string;
  format: "png";
  byteLength: number;
  checksum: string;
  sourceMimeType?: string;
}

interface NativeRenderableAssetData {
  data: Uint8Array;
  mimeType: string;
  format: "source" | "png";
  proxy: boolean;
  dimensions?: { width: number; height: number; source: string };
}

interface NativeAssetCatalog {
  assets: Asset[];
  previewAssets: Asset[];
  allAssets: Asset[];
  byPath: Map<string, NativeAssetReference>;
  byName: Map<string, NativeAssetReference[]>;
  byStem: Map<string, NativeAssetReference[]>;
  byDataId: Map<string, NativeAssetReference[]>;
  previewReferences: NativePreviewAssetReference[];
}

interface NativeAssetMatch {
  asset: NativeAssetReference;
  evidence: string;
  confidence: number;
  fieldPath?: string;
  source: "protobuf" | "raw" | "archive";
  geometryCandidate?: NativeIwaGeometryCandidate;
  archiveMessage?: NativeIwaArchiveMessageEvidence;
  suppressedAssets?: NativeAssetReference[];
}

interface NativeSlideOrderEvidence {
  sourcePath: string;
  method: "document-reference" | "document-archive-identifier";
  orderedPaths: string[];
  matchedPathCount: number;
  hintCount: number;
  confidence: number;
}

interface NativePreviewAssetReference extends NativeAssetReference {
  previewSource: "quicklook" | "package-asset";
  previewRole: NativeQuickLookPreviewMetadata["role"] | "snapshot";
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
  sizeSource?: string;
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

interface IwaNumericCandidate extends NativeIwaNumericCandidate {
  source: "protobuf";
  order: number;
  endOrder: number;
}

interface IwaScanResult {
  textCandidates: IwaTextCandidate[];
  referenceCandidates: IwaReferenceCandidate[];
  numericCandidates: IwaNumericCandidate[];
  geometryCandidates: NativeIwaGeometryCandidate[];
  groupingHints: NativeIwaGroupingHint[];
  animationHints: NativeIwaAnimationHint[];
  fieldSummaries: NativeIwaFieldSummary[];
  archiveRecords: NativeIwaArchiveRecordEvidence[];
  archiveMessages: NativeIwaArchiveMessageEvidence[];
  typedArchiveMessageCount: number;
  protobufFieldCount: number;
  nestedMessageCount: number;
  rawStringCount: number;
  expandedByteLength: number;
}

interface IwaArchiveMessageInfo {
  length: number;
  payloadOffset: number;
  payloadLength: number;
  type?: number;
  version: number[];
  objectReferences: string[];
  dataReferences: string[];
}

interface IwaArchiveRecord {
  archiveOffset: number;
  archiveInfoLength: number;
  payloadOffset: number;
  totalPayloadLength: number;
  messageInfos: IwaArchiveMessageInfo[];
  identifier?: string;
  shouldMerge?: boolean;
}

export async function detectNativeKeynotePackage(keynotePath: string): Promise<NativeKeynoteDetection> {
  const keynotePackage = await readKeynotePackage(keynotePath);
  const parts = readNativeKeynoteParts(keynotePackage.entries);
  const assets = createNativeAssetCatalog(parts.assetPaths, parts.entries, parts.quickLookPreviews);
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
    quickLookPreviews: parts.quickLookPreviews,
    iwaStreams: classifyIwaStreams(parts.iwaEntries, assets)
  };
}

export async function parseNativeKeynoteToIr(keynotePath: string): Promise<DeckIR> {
  const keynotePackage = await readKeynotePackage(keynotePath);
  const parts = readNativeKeynoteParts(keynotePackage.entries);
  if (!parts.hasIndexZip && parts.iwaEntries.length === 0) {
    throw new Error("Native Keynote package was not detected. Expected Index.zip or Index/*.iwa streams.");
  }

  const metadata = readNativeMetadata(parts.entries);
  const deckSize = metadata.size ?? DEFAULT_DECK_SIZE;
  const assets = createNativeAssetCatalog(parts.assetPaths, parts.entries, parts.quickLookPreviews);
  const iwaStreams = classifyIwaStreams(parts.iwaEntries, assets);
  const slideOrderEvidence = inferSlideOrderEvidence(parts.iwaEntries);
  const slideEntries = chooseSlideIwaEntries(parts.iwaEntries, slideOrderEvidence);
  const slides = buildNativeSlides(slideEntries, parts.iwaEntries, deckSize, assets, slideOrderEvidence);
  const objectCount = slides.reduce((total, slide) => total + slide.objects.length, 0);
  const recoveredTextObjectCount = slides.reduce(
    (total, slide) => total + slide.objects.filter((object) => object.type === "text").length,
    0
  );
  const recoveredAssetObjectCount = slides.reduce(
    (total, slide) =>
      total +
      slide.objects.filter(
        (object) =>
          (object.type === "image" || object.type === "media") &&
          object.metadata?.nativeFallback !== "full-slide-preview"
      ).length,
    0
  );
  const previewFallbackObjectCount = slides.reduce(
    (total, slide) => total + slide.objects.filter((object) => object.metadata?.nativeFallback === "full-slide-preview").length,
    0
  );
  const unrecoveredAssetCount = Math.max(0, assets.assets.length - countUniqueReferencedAssets(slides));
  const hasDetectedAssets = assets.assets.length > 0;
  const packageFormat = classifyPackageFormat(keynotePackage.container, parts);
  const totalProtobufFieldCount = iwaStreams.reduce((total, stream) => total + stream.protobufFieldCount, 0);
  const totalNestedMessageCount = iwaStreams.reduce((total, stream) => total + stream.nestedMessageCount, 0);
  const totalReferenceCandidateCount = iwaStreams.reduce((total, stream) => total + stream.referenceCandidateCount, 0);
  const totalNumericCandidateCount = iwaStreams.reduce((total, stream) => total + stream.numericCandidateCount, 0);
  const totalGeometryCandidateCount = iwaStreams.reduce((total, stream) => total + stream.geometryCandidateCount, 0);
  const totalGroupingHintCount = iwaStreams.reduce((total, stream) => total + stream.groupingHintCount, 0);
  const totalAnimationHintCount = iwaStreams.reduce((total, stream) => total + stream.animationHintCount, 0);
  const totalMagicMoveHintCount = iwaStreams.reduce((total, stream) => total + stream.magicMoveHintCount, 0);
  const totalMorphHintCount = iwaStreams.reduce((total, stream) => total + stream.morphHintCount, 0);
  const totalArchiveRecordCount = iwaStreams.reduce((total, stream) => total + stream.archiveRecordCount, 0);
  const totalArchivePayloadCount = iwaStreams.reduce((total, stream) => total + stream.archivePayloadCount, 0);
  const totalTypedArchiveMessageCount = iwaStreams.reduce((total, stream) => total + stream.typedArchiveMessageCount, 0);
  const totalNativeBuildRecordCount = iwaStreams.reduce((total, stream) => total + stream.buildRecordCount, 0);
  const totalNativeBuildTimingRecordCount = iwaStreams.reduce((total, stream) => total + stream.buildTimingRecordCount, 0);
  const recoveredNativeBuildAnimationCount = slides.reduce(
    (total, slide) => total + (slide.timeline?.events ?? []).filter((event) => event.metadata?.nativeSource === "keynote-iwa-build").length,
    0
  );
  const recoveredCharacterBuildAnimationCount = slides.reduce(
    (total, slide) => total + (slide.timeline?.events ?? []).filter((event) => event.metadata?.nativeBuildGranularity === "character").length,
    0
  );
  const totalNativeTypedVisualRecordCount = slides.reduce(
    (total, slide) => total + Number(slide.metadata?.nativeTypedVisualRecordCount ?? 0),
    0
  );
  const unresolvedNativeBuildRecordCount = slides.reduce(
    (total, slide) => total + Number(slide.metadata?.nativeBuildAnimationUnresolvedCount ?? 0),
    0
  );
  const recoveredNativeBuildRecordCount = Math.max(0, totalNativeBuildRecordCount - unresolvedNativeBuildRecordCount);
  const quickLookPreviewCount = parts.quickLookPreviews.length;
  const quickLookPreviewWithDimensionsCount = parts.quickLookPreviews.filter((preview) => preview.width !== undefined && preview.height !== undefined).length;
  const imageAssetCount = assets.assets.filter((asset) => asset.kind === "image").length;
  const imageAssetWithDimensionsCount = assets.assets.filter((asset) => asset.kind === "image" && asset.width !== undefined && asset.height !== undefined).length;
  const previewAssetCount = assets.previewReferences.length;
  const previewImageAssetCount = assets.previewReferences.filter((asset) => asset.kind === "image").length;
  const lossReportMetadata = {
    schema: "keymorph.keynote.native.loss.v1",
    nativeContainer: keynotePackage.container,
    packageFormat,
    automationUsed: false,
    parser: "static-iwa-protobuf-field-scan",
    limitations: [
      "private-iwa-schema",
      "approximate-object-placement",
      "no-keynote-gui-render",
      recoveredNativeBuildAnimationCount > 0 ? "partial-native-build-animation-reconstruction" : "native-animation-not-reconstructed"
    ],
    evidenceCounts: {
      iwaStreamCount: parts.iwaEntries.length,
      archiveRecordCount: totalArchiveRecordCount,
      archivePayloadCount: totalArchivePayloadCount,
      typedArchiveMessageCount: totalTypedArchiveMessageCount,
      buildRecordCount: totalNativeBuildRecordCount,
      buildTimingRecordCount: totalNativeBuildTimingRecordCount,
      typedVisualRecordCount: totalNativeTypedVisualRecordCount,
      recoveredBuildRecordCount: recoveredNativeBuildRecordCount,
      recoveredBuildAnimationCount: recoveredNativeBuildAnimationCount,
      recoveredCharacterBuildAnimationCount,
      unresolvedBuildRecordCount: unresolvedNativeBuildRecordCount,
      protobufFieldCount: totalProtobufFieldCount,
      nestedMessageCount: totalNestedMessageCount,
      numericCandidateCount: totalNumericCandidateCount,
      geometryCandidateCount: totalGeometryCandidateCount,
      groupingHintCount: totalGroupingHintCount,
      animationHintCount: totalAnimationHintCount,
      assetReferenceCandidateCount: totalReferenceCandidateCount,
      quickLookPreviewCount,
      quickLookPreviewWithDimensionsCount,
      imageAssetCount,
      imageAssetWithDimensionsCount,
      previewAssetCount,
      previewImageAssetCount,
      previewFallbackObjectCount
    },
    recommendedFallback: "Use the Keynote PPTX bridge with explicit automation opt-in for visual layout and animation downgrade."
  };
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
      description:
        totalNativeBuildRecordCount > 0
          ? `Decoded ${totalNativeBuildRecordCount} native Keynote build record(s) and ${totalNativeBuildTimingRecordCount} timing wrapper(s); ${recoveredNativeBuildAnimationCount} resolved target(s) were degraded to IR playback events while unresolved/private effects remain unsupported.`
          : totalAnimationHintCount > 0
          ? `Detected ${totalAnimationHintCount} native animation hint(s), including ${totalMagicMoveHintCount} Magic Move and ${totalMorphHintCount} morph-like hint(s), but most Keynote timing data remains private.`
          : "Keynote builds, transitions, Magic Move, and timing data are not mapped by the native fallback.",
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
      : []),
    ...(unresolvedNativeBuildRecordCount > 0
      ? [
          {
            code: "keynote-native-build-target-loss",
            severity: "warning" as const,
            area: "animation" as const,
            description:
              `Decoded ${totalNativeBuildRecordCount} Keynote build record(s), but ${unresolvedNativeBuildRecordCount} could not be attached to a recovered IR object target.`,
            fallback: "Preserve build/timing evidence in slide metadata and recover only effects whose target archive id maps to a known object."
          }
        ]
      : []),
    ...(recoveredCharacterBuildAnimationCount > 0
      ? [
          {
            code: "keynote-native-character-build-degraded",
            severity: "warning" as const,
            area: "animation" as const,
            description:
              `Recovered ${recoveredCharacterBuildAnimationCount} character-level Keynote build event(s), but the current IR runtime approximates them at object level while preserving character granularity metadata.`,
            fallback: "Preserve nativeBuildGranularity=character and upgrade the runtime to grapheme-level split text playback."
          }
        ]
      : []),
    ...(previewFallbackObjectCount > 0
      ? [
          {
            code: "keynote-native-preview-fallback",
            severity: "warning" as const,
            area: "layout" as const,
            description:
              "Low-confidence native text recovery was suppressed in favor of a full-slide preview or snapshot image from the Keynote package.",
            fallback: "Preserve the preview image as the slide visual and expose the source asset for downstream inspection."
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
      : []),
    ...(totalGeometryCandidateCount > 0
      ? [
          {
            code: "keynote-native-geometry-candidate-scan",
            severity: "warning" as const,
            description:
              `Detected ${totalGeometryCandidateCount} numeric geometry candidate(s) from protobuf field groups; object bounds remain approximate until the private schema is mapped.`,
            confidence: 0.38
          }
        ]
      : []),
    ...(totalGroupingHintCount > 0
      ? [
          {
            code: "keynote-native-object-grouping-hints",
            severity: "warning" as const,
            description:
              `Detected ${totalGroupingHintCount} slide/object grouping hint(s) from nested field paths and mixed text/reference/geometry evidence.`,
            confidence: 0.41
          }
        ]
      : []),
    ...(totalNativeBuildRecordCount > 0
      ? [
          {
            code: "keynote-native-build-timing-scan",
            severity: recoveredNativeBuildAnimationCount > 0 ? ("info" as const) : ("warning" as const),
            description:
              recoveredNativeBuildAnimationCount > 0
                ? `Decoded native Keynote build/timing records and recovered ${recoveredNativeBuildRecordCount} build record(s) into ${recoveredNativeBuildAnimationCount} conservative opacity/media event(s).`
                : `Decoded ${totalNativeBuildRecordCount} native Keynote build record(s), but no target object could be mapped safely.`,
            confidence: recoveredNativeBuildAnimationCount > 0 ? 0.62 : 0.46
          }
        ]
      : []),
    ...(totalAnimationHintCount > 0
      ? [
          {
            code: "keynote-native-magic-move-morph-hints",
            severity: "warning" as const,
            description:
              `Detected native animation terms in IWA payloads (${totalMagicMoveHintCount} Magic Move, ${totalMorphHintCount} morph-like), but target pairs and timing are uncertain.`,
            confidence: 0.33
          }
        ]
      : [])
  ];
  const deckTitle = metadata.title ?? (path.basename(keynotePath, path.extname(keynotePath)) || "Imported Keynote");
  const deckAssets = assets.allAssets;

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
        nativeIwaArchiveRecordCount: totalArchiveRecordCount,
        nativeIwaArchivePayloadCount: totalArchivePayloadCount,
        nativeIwaTypedArchiveMessageCount: totalTypedArchiveMessageCount,
        ...(slideOrderEvidence
          ? {
              nativeSlideOrderEvidence: {
                sourcePath: slideOrderEvidence.sourcePath,
                method: slideOrderEvidence.method,
                orderedPaths: slideOrderEvidence.orderedPaths,
                matchedPathCount: slideOrderEvidence.matchedPathCount,
                hintCount: slideOrderEvidence.hintCount,
                confidence: slideOrderEvidence.confidence
              }
            }
          : {}),
        nativeAssetCount: assets.assets.length,
        nativeImageAssetCount: imageAssetCount,
        nativeImageAssetWithDimensionsCount: imageAssetWithDimensionsCount,
        nativePreviewAssetCount: previewAssetCount,
        nativePreviewImageAssetCount: previewImageAssetCount,
        nativeQuickLookPreviews: parts.quickLookPreviews,
        nativeMetadataPaths: metadata.paths,
        ...(metadata.sizeSource ? { nativeDeckSizeSource: metadata.sizeSource } : {}),
        ...(metadata.documentIdentifier ? { nativeDocumentIdentifier: metadata.documentIdentifier } : {})
      }
    },
    deck: {
      id: "keynote-native-import",
      title: deckTitle,
      size: { width: deckSize.width, height: deckSize.height, unit: "px" },
      ...(deckAssets.length > 0 ? { assets: deckAssets } : {}),
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
        },
        {
          severity: previewAssetCount > 0 ? "info" : "info",
          code:
            previewAssetCount > 0
              ? previewFallbackObjectCount > 0
                ? "keynote-native-preview-fallback-used"
                : "keynote-native-preview-assets-exposed"
              : "keynote-native-no-preview-assets",
          message:
            previewAssetCount > 0
              ? previewFallbackObjectCount > 0
                ? `Exposed ${previewAssetCount} preview/snapshot asset(s) and used ${previewFallbackObjectCount} as full-slide fallback object(s).`
                : `Exposed ${previewAssetCount} preview/snapshot asset(s) for downstream inspection.`
              : "No QuickLook preview or snapshot image assets were detected."
        },
        {
          severity: totalGeometryCandidateCount > 0 ? "info" : "warning",
          code: totalGeometryCandidateCount > 0 ? "keynote-native-geometry-candidates-detected" : "keynote-native-geometry-candidates-not-detected",
          message:
            totalGeometryCandidateCount > 0
              ? `Detected ${totalGeometryCandidateCount} numeric geometry candidate(s) from protobuf-like fields.`
              : "No numeric geometry candidates were detected from protobuf-like fields."
        },
        {
          severity: totalAnimationHintCount > 0 ? "warning" : "info",
          code: totalAnimationHintCount > 0 ? "keynote-native-animation-hints-detected" : "keynote-native-animation-hints-not-detected",
          message:
            totalAnimationHintCount > 0
              ? `Detected ${totalAnimationHintCount} native animation hint(s), including ${totalMagicMoveHintCount} Magic Move and ${totalMorphHintCount} morph-like hint(s); mappings remain uncertain.`
              : "No Magic Move or morph-like native animation hints were detected in visible IWA payloads."
        },
        {
          severity: totalNativeBuildRecordCount > 0 ? (recoveredNativeBuildAnimationCount > 0 ? "info" : "warning") : "info",
          code:
            totalNativeBuildRecordCount > 0
              ? recoveredNativeBuildAnimationCount > 0
                ? "keynote-native-build-animations-recovered"
                : "keynote-native-build-animations-unresolved"
              : "keynote-native-build-animations-not-detected",
          message:
            totalNativeBuildRecordCount > 0
              ? recoveredNativeBuildAnimationCount > 0
                ? `Decoded ${totalNativeBuildRecordCount} native Keynote build record(s), recovered ${recoveredNativeBuildRecordCount} build target(s), and generated ${recoveredNativeBuildAnimationCount} conservative animation event(s).`
                : `Decoded ${totalNativeBuildRecordCount} native Keynote build record(s), but none mapped to recovered IR objects safely.`
              : "No native Keynote build records were decoded from typed IWA payloads."
        }
      ],
      unsupportedFeatures,
      degradedFeatures,
      uncertainMappings,
      statistics: {
        slideCount: slides.length,
        objectCount,
        animationCount: recoveredNativeBuildAnimationCount,
        assetCount: deckAssets.length,
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
        quickLookPreviews: parts.quickLookPreviews,
        totalProtobufFieldCount,
        totalNestedMessageCount,
        totalReferenceCandidateCount,
        totalNumericCandidateCount,
        totalGeometryCandidateCount,
        totalGroupingHintCount,
        totalAnimationHintCount,
        totalMagicMoveHintCount,
        totalMorphHintCount,
        totalArchiveRecordCount,
        totalArchivePayloadCount,
        totalTypedArchiveMessageCount,
        totalNativeBuildRecordCount,
        totalNativeBuildTimingRecordCount,
        totalNativeTypedVisualRecordCount,
        recoveredNativeBuildRecordCount,
        recoveredNativeBuildAnimationCount,
        recoveredCharacterBuildAnimationCount,
        unresolvedNativeBuildRecordCount,
        slideOrderEvidence: slideOrderEvidence
          ? {
              sourcePath: slideOrderEvidence.sourcePath,
              method: slideOrderEvidence.method,
              orderedPaths: slideOrderEvidence.orderedPaths,
              matchedPathCount: slideOrderEvidence.matchedPathCount,
              hintCount: slideOrderEvidence.hintCount,
              confidence: slideOrderEvidence.confidence
            }
          : undefined,
        recoveredTextObjectCount,
        recoveredAssetObjectCount,
        previewFallbackObjectCount,
        unrecoveredAssetCount,
        imageAssetCount,
        imageAssetWithDimensionsCount,
        previewAssetCount,
        previewImageAssetCount,
        quickLookPreviewWithDimensionsCount,
        lossReport: lossReportMetadata,
        iwaStreams: iwaStreams.map((stream) => ({
          path: stream.path,
          role: stream.role,
          compression: stream.compression,
          byteLength: stream.byteLength,
          expandedByteLength: stream.expandedByteLength,
          textCandidateCount: stream.textCandidateCount,
          referenceCandidateCount: stream.referenceCandidateCount,
          numericCandidateCount: stream.numericCandidateCount,
          geometryCandidateCount: stream.geometryCandidateCount,
          groupingHintCount: stream.groupingHintCount,
          animationHintCount: stream.animationHintCount,
          magicMoveHintCount: stream.magicMoveHintCount,
          morphHintCount: stream.morphHintCount,
          assetReferenceCount: stream.assetReferenceCount,
          archiveRecordCount: stream.archiveRecordCount,
          archivePayloadCount: stream.archivePayloadCount,
          archiveRecords: stream.archiveRecords,
          typedArchiveMessageCount: stream.typedArchiveMessageCount,
          typedArchiveMessages: stream.typedArchiveMessages,
          buildRecordCount: stream.buildRecordCount,
          buildTimingRecordCount: stream.buildTimingRecordCount,
          protobufFieldCount: stream.protobufFieldCount,
          protobufFieldPathCount: stream.protobufFieldPathCount,
          nestedMessageCount: stream.nestedMessageCount,
          rawStringCount: stream.rawStringCount,
          numericCandidates: stream.numericCandidates,
          geometryCandidates: stream.geometryCandidates,
          groupingHints: stream.groupingHints,
          animationHints: stream.animationHints,
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
  const quickLookPaths = allPaths.filter((entryPath) => isPreviewPath(entryPath));
  const quickLookPreviews = quickLookPaths.map((entryPath) =>
    createQuickLookPreviewMetadata(entryPath, combinedEntries.get(entryPath) ?? new Uint8Array())
  );
  return {
    hasIndexZip,
    entries: combinedEntries,
    iwaEntries: allPaths
      .filter((entryPath) => entryPath.toLowerCase().endsWith(".iwa"))
      .map((entryPath) => ({ path: entryPath, data: combinedEntries.get(entryPath)! })),
    metadataPaths: allPaths.filter((entryPath) => isMetadataPath(entryPath)),
    assetPaths: allPaths.filter((entryPath) => entryPath.startsWith("Data/")),
    quickLookPaths,
    quickLookPreviews,
    indexZipEntryCount,
    hasLooseIndexDirectory,
    hasDataDirectory: allPaths.some((entryPath) => entryPath.startsWith("Data/")),
    hasQuickLookPreview: quickLookPaths.length > 0
  };
}

function isPreviewPath(entryPath: string): boolean {
  return /^QuickLook\//i.test(entryPath) || /^preview(?:[-_][a-z0-9]+)?\.(?:jpe?g|png|gif|webp)$/i.test(entryPath);
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

function createNativeAssetCatalog(
  assetPaths: string[],
  entries: Map<string, Uint8Array>,
  quickLookPreviews: NativeQuickLookPreviewMetadata[]
): NativeAssetCatalog {
  const assets: Asset[] = [];
  const previewAssets: Asset[] = [];
  const byPath = new Map<string, NativeAssetReference>();
  const byName = new Map<string, NativeAssetReference[]>();
  const byStem = new Map<string, NativeAssetReference[]>();
  const byDataId = new Map<string, NativeAssetReference[]>();
  const previewReferences: NativePreviewAssetReference[] = [];

  for (const assetPath of assetPaths) {
    const data = entries.get(assetPath);
    const name = path.posix.basename(assetPath);
    if (!name || !data) {
      continue;
    }

    const classification = classifyAssetPath(assetPath);
    const renderable = createNativeRenderableAssetData(assetPath, data, classification.mimeType);
    const imageDimensions =
      renderable?.dimensions ?? (classification.kind === "image" ? readImageDimensions(data, classification.mimeType) : undefined);
    const assetId = stableAssetId(assetPath);
    const dataId = dataIdFromAssetPath(assetPath);
    const embedded = renderable !== undefined;
    const asset: Asset = {
      id: assetId,
      kind: classification.kind,
      uri: embedded && renderable ? dataUri(renderable.mimeType, renderable.data) : undefined,
      name,
      mimeType: renderable?.mimeType ?? classification.mimeType,
      ...(imageDimensions ? { width: imageDimensions.width, height: imageDimensions.height } : {}),
      checksum: `sha256:${sha256Hex(data)}`,
      metadata: {
        nativeSourcePath: assetPath,
        byteLength: data.byteLength,
        nativeDataEmbedded: embedded,
        ...(renderable && renderable.data !== data
          ? {
              nativeRenderableProxy: true,
              nativeRenderableProxyFormat: renderable.format,
              nativeRenderableProxyMimeType: renderable.mimeType,
              nativeRenderableProxyByteLength: renderable.data.byteLength,
              nativeRenderableProxyChecksum: `sha256:${sha256Hex(renderable.data)}`,
              nativeSourceMimeType: classification.mimeType ?? ""
            }
          : {}),
        ...(dataId ? { nativeDataId: dataId } : {}),
        ...(isSnapshotAssetPath(assetPath)
          ? {
              nativePreviewAsset: true,
              nativePreviewRole: "snapshot",
              nativePreviewSource: "package-asset"
            }
          : {}),
        ...(imageDimensions
          ? {
              nativeImageWidth: imageDimensions.width,
              nativeImageHeight: imageDimensions.height,
              nativeImageDimensionSource: imageDimensions.source
            }
          : {})
      }
    };
    const reference: NativeAssetReference = {
      assetId,
      path: assetPath,
      name,
      kind: asset.kind,
      mimeType: asset.mimeType,
      sourceMimeType: classification.mimeType,
      width: asset.width,
      height: asset.height,
      embedded,
      ...(renderable?.proxy
        ? {
            renderProxy: {
              mimeType: renderable.mimeType,
              format: renderable.format,
              byteLength: renderable.data.byteLength,
              checksum: `sha256:${sha256Hex(renderable.data)}`,
              sourceMimeType: classification.mimeType
            }
          }
        : {}),
      ...(dataId ? { dataId } : {})
    };

    assets.push(asset);
    byPath.set(assetPath.toLowerCase(), reference);
    addMultiMapValue(byName, name.toLowerCase(), reference);
    addMultiMapValue(byStem, normalizeAssetStem(path.posix.basename(name, path.posix.extname(name))), reference);
    if (dataId) {
      addMultiMapValue(byDataId, dataId, reference);
    }
    if (isSnapshotAssetPath(assetPath) && asset.kind === "image") {
      previewReferences.push({
        ...reference,
        previewSource: "package-asset",
        previewRole: "snapshot"
      });
    }
  }

  for (const preview of quickLookPreviews) {
    const data = entries.get(preview.path);
    const classification = classifyAssetPath(preview.path);
    if (!data || classification.kind !== "image") {
      continue;
    }

    const imageDimensions = readImageDimensions(data, classification.mimeType);
    const assetId = stableAssetId(preview.path);
    const name = path.posix.basename(preview.path);
    const asset: Asset = {
      id: assetId,
      kind: "image",
      uri: dataUri(classification.mimeType, data),
      name,
      mimeType: classification.mimeType,
      ...(imageDimensions ? { width: imageDimensions.width, height: imageDimensions.height } : {}),
      checksum: `sha256:${sha256Hex(data)}`,
      metadata: {
        nativeSourcePath: preview.path,
        byteLength: data.byteLength,
        nativePreviewAsset: true,
        nativeDataEmbedded: true,
        nativePreviewRole: preview.role,
        nativePreviewSource: "quicklook",
        ...(imageDimensions
          ? {
              nativeImageWidth: imageDimensions.width,
              nativeImageHeight: imageDimensions.height,
              nativeImageDimensionSource: imageDimensions.source
            }
          : {})
      }
    };
    const reference: NativePreviewAssetReference = {
      assetId,
      path: preview.path,
      name,
      kind: "image",
      mimeType: asset.mimeType,
      width: asset.width,
      height: asset.height,
      previewSource: "quicklook",
      previewRole: preview.role
    };

    previewAssets.push(asset);
    previewReferences.push(reference);
  }

  assets.sort((left, right) => comparePartPaths(String(left.metadata?.nativeSourcePath ?? left.name ?? left.id), String(right.metadata?.nativeSourcePath ?? right.name ?? right.id)));
  previewAssets.sort((left, right) => comparePartPaths(String(left.metadata?.nativeSourcePath ?? left.name ?? left.id), String(right.metadata?.nativeSourcePath ?? right.name ?? right.id)));
  previewReferences.sort(comparePreviewAssetReferences);
  return { assets, previewAssets, allAssets: [...assets, ...previewAssets], byPath, byName, byStem, byDataId, previewReferences };
}

function shouldEmbedNativeAsset(data: Uint8Array, mimeType: string | undefined): boolean {
  return Boolean(mimeType?.startsWith("image/")) && mimeType !== "image/tiff" && data.byteLength <= MAX_EMBEDDED_ASSET_BYTES;
}

function createNativeRenderableAssetData(
  assetPath: string,
  data: Uint8Array,
  mimeType: string | undefined
): NativeRenderableAssetData | undefined {
  const sourceDimensions = readImageDimensions(data, mimeType);
  if (isBrowserRenderableImageMimeType(mimeType)) {
    if (shouldEmbedNativeAsset(data, mimeType)) {
      return { data, mimeType: mimeType!, format: "source", proxy: false, dimensions: sourceDimensions };
    }
    return undefined;
  }

  if (mimeType === "image/tiff" || mimeType === "image/heic") {
    const converted = convertNativeImageToPng(assetPath, data);
    if (!converted || converted.byteLength > MAX_RENDERABLE_NATIVE_ASSET_BYTES) {
      return undefined;
    }
    return {
      data: converted,
      mimeType: "image/png",
      format: "png",
      proxy: true,
      dimensions: readPngDimensions(converted) ?? sourceDimensions
    };
  }
  return undefined;
}

function isBrowserRenderableImageMimeType(mimeType: string | undefined): boolean {
  return ["image/png", "image/jpeg", "image/gif", "image/webp", "image/svg+xml"].includes(mimeType ?? "");
}

function convertNativeImageToPng(assetPath: string, data: Uint8Array): Uint8Array | undefined {
  if (!existsSync("/usr/bin/sips")) {
    return undefined;
  }
  const dir = mkdtempSync(path.join(tmpdir(), "keymorph-native-image-"));
  const inputPath = path.join(dir, `input${path.extname(assetPath) || ".image"}`);
  const outputPath = path.join(dir, "output.png");
  try {
    writeFileSync(inputPath, data);
    execFileSync("/usr/bin/sips", ["-s", "format", "png", inputPath, "--out", outputPath], { stdio: "ignore" });
    return new Uint8Array(readFileSync(outputPath));
  } catch {
    return undefined;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function dataUri(mimeType: string | undefined, data: Uint8Array): string {
  return `data:${mimeType ?? "application/octet-stream"};base64,${Buffer.from(data).toString("base64")}`;
}

function isSnapshotAssetPath(assetPath: string): boolean {
  const name = path.posix.basename(assetPath).toLowerCase();
  return /^st-[a-f0-9-]+\.(?:png|jpe?g|webp)$/i.test(name) || /^preview(?:[-_][a-z0-9]+)?\.(?:png|jpe?g|webp)$/i.test(name);
}

function comparePreviewAssetReferences(left: NativePreviewAssetReference, right: NativePreviewAssetReference): number {
  const role = previewRoleRank(right.previewRole) - previewRoleRank(left.previewRole);
  if (role !== 0) return role;
  if (left.previewRole === "snapshot" && right.previewRole === "snapshot") {
    return comparePartPaths(left.path, right.path);
  }
  const pixels = (right.width ?? 0) * (right.height ?? 0) - (left.width ?? 0) * (left.height ?? 0);
  if (pixels !== 0) return pixels;
  return comparePartPaths(left.path, right.path);
}

function previewRoleRank(role: NativePreviewAssetReference["previewRole"]): number {
  if (role === "snapshot") return 3;
  if (role === "preview") return 2;
  if (role === "thumbnail") return 1;
  return 0;
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
      numericCandidateCount: scan.numericCandidates.length,
      geometryCandidateCount: scan.geometryCandidates.length,
      groupingHintCount: scan.groupingHints.length,
      animationHintCount: scan.animationHints.length,
      magicMoveHintCount: scan.animationHints.filter((hint) => hint.kind === "magicMove").length,
      morphHintCount: scan.animationHints.filter((hint) => hint.kind === "morph").length,
      assetReferenceCount: findIwaAssetMatchesFromScan(scan, assets).length,
      archiveRecordCount: scan.archiveRecords.length,
      archivePayloadCount: scan.archiveRecords.reduce((total, record) => total + record.messageCount, 0),
      archiveRecords: scan.archiveRecords,
      typedArchiveMessageCount: scan.typedArchiveMessageCount,
      typedArchiveMessages: scan.archiveMessages,
      buildRecordCount: scan.archiveMessages.filter((message) => message.build).length,
      buildTimingRecordCount: scan.archiveMessages.filter((message) => message.buildTiming).length,
      protobufFieldCount: scan.protobufFieldCount,
      protobufFieldPathCount: scan.fieldSummaries.length,
      nestedMessageCount: scan.nestedMessageCount,
      rawStringCount: scan.rawStringCount,
      numericCandidates: scan.numericCandidates.map(stripNumericCandidateInternalFields),
      geometryCandidates: scan.geometryCandidates,
      groupingHints: scan.groupingHints,
      animationHints: scan.animationHints,
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

  const plistSize = readDeckSize(values);
  const documentSize = plistSize ? undefined : readNativeDocumentDeckSize(entries);

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
    size: plistSize ?? documentSize,
    sizeSource: plistSize ? "metadata-plist" : documentSize ? "document-iwa" : undefined,
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
  assets: NativeAssetCatalog,
  slideOrderEvidence?: NativeSlideOrderEvidence
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
    const typedVisualRecords = createNativeTypedVisualRecords(scan.archiveMessages);
    const previewFallback = selectPreviewFallbackAsset(assets.previewReferences, index, entries.length);
    const usePreviewFallback = previewFallback
      ? shouldUsePreviewFallback(textCandidates, assetMatches, previewFallback)
      : false;
    const textObjects = textCandidates.map((candidate, objectIndex) =>
      createTextObject(slideId, candidate, objectIndex, entry.path, deckSize, scan.geometryCandidates[objectIndex])
    );
    const assetObjects = assetMatches.map((match, assetIndex) =>
      createAssetObject(
        slideId,
        match,
        textObjects.length + assetIndex,
        entry.path,
        deckSize,
        match.geometryCandidate ?? scan.geometryCandidates[textObjects.length + assetIndex]
      )
    );
    applyNativePlacementGroupMetadata(assetObjects);
    const nativeTextObjects = createNativeTextDrawableObjects(slideId, scan.archiveMessages, entry.path, deckSize);
    const textObjectsToKeep = suppressDuplicateFallbackTextObjects(textObjects, nativeTextObjects);
    const objects =
      usePreviewFallback && previewFallback
        ? [createPreviewImageObject(slideId, previewFallback, entry.path, deckSize, textCandidates)]
      : textObjects.length > 0 || nativeTextObjects.length > 0 || assetObjects.length > 0
        ? [...textObjectsToKeep, ...nativeTextObjects, ...assetObjects]
        : [createPlaceholderObject(slideId, entry.path, deckSize)];
    const buildAnimationRecovery = recoverNativeBuildAnimations(slideId, objects, scan.archiveMessages);

    return {
      id: slideId,
      index,
      name: readSlideName(entry.path, index),
      background: { type: "solid" as const, color: "#ffffff" },
      objects,
      timeline: {
        durationMs: buildAnimationRecovery.durationMs,
        events: buildAnimationRecovery.events,
        dependencyGraph: { edges: buildAnimationRecovery.dependencyEdges },
        ...(buildAnimationRecovery.buildRecords.length > 0 || buildAnimationRecovery.timingRecords.length > 0
          ? {
              metadata: {
                nativeBuildRecordCount: buildAnimationRecovery.buildRecords.length,
                nativeBuildTimingRecordCount: buildAnimationRecovery.timingRecords.length,
                nativeBuildAnimationRecoveredCount: buildAnimationRecovery.events.length,
                nativeBuildAnimationUnresolvedCount: buildAnimationRecovery.unresolvedBuildCount,
                nativeBuildTimingDependencyCount: buildAnimationRecovery.dependencyEdges.length
              }
            }
          : {})
      },
      metadata: {
        nativeSourcePath: entry.path,
        nativeTextCandidateCount: textCandidates.length,
        nativeReferenceCandidateCount: scan.referenceCandidates.length,
        nativeNumericCandidateCount: scan.numericCandidates.length,
        nativeGeometryCandidateCount: scan.geometryCandidates.length,
        nativeGroupingHintCount: scan.groupingHints.length,
        nativeAnimationHintCount: scan.animationHints.length,
        nativeMagicMoveHintCount: scan.animationHints.filter((hint) => hint.kind === "magicMove").length,
        nativeMorphHintCount: scan.animationHints.filter((hint) => hint.kind === "morph").length,
        nativeAssetReferenceCount: assetMatches.length,
        nativeArchiveRecordCount: scan.archiveRecords.length,
        nativeArchivePayloadCount: scan.archiveRecords.reduce((total, record) => total + record.messageCount, 0),
        ...(scan.archiveRecords.length > 0 ? { nativeArchiveRecords: scan.archiveRecords.slice(0, 12) } : {}),
        nativeTypedArchiveMessageCount: scan.typedArchiveMessageCount,
        ...(scan.archiveMessages.length > 0 ? { nativeTypedArchiveMessages: scan.archiveMessages.slice(0, 24) } : {}),
        nativeTypedVisualRecordCount: typedVisualRecords.length,
        ...(typedVisualRecords.length > 0 ? { nativeTypedVisualRecords: typedVisualRecords.slice(0, 24) } : {}),
        nativeBuildRecordCount: buildAnimationRecovery.buildRecords.length,
        nativeBuildTimingRecordCount: buildAnimationRecovery.timingRecords.length,
        nativeBuildAnimationRecoveredCount: buildAnimationRecovery.events.length,
        nativeBuildAnimationUnresolvedCount: buildAnimationRecovery.unresolvedBuildCount,
        nativeBuildTimingDependencyCount: buildAnimationRecovery.dependencyEdges.length,
        ...(buildAnimationRecovery.buildRecords.length > 0
          ? { nativeBuildRecords: buildAnimationRecovery.buildRecords.slice(0, 24) }
          : {}),
        ...(buildAnimationRecovery.timingRecords.length > 0
          ? { nativeBuildTimingRecords: buildAnimationRecovery.timingRecords.slice(0, 24) }
          : {}),
        ...(slideOrderEvidence
          ? {
              nativeSlideOrderSourcePath: slideOrderEvidence.sourcePath,
              nativeSlideOrderMethod: slideOrderEvidence.method,
              nativeSlideOrderConfidence: slideOrderEvidence.confidence
            }
          : {}),
        ...(previewFallback
          ? {
              nativePreviewFallbackAvailable: true,
              nativePreviewFallbackPath: previewFallback.path,
              nativePreviewFallbackUsed: usePreviewFallback
            }
          : {}),
        nativeProtobufFieldCount: scan.protobufFieldCount,
        nativeNestedMessageCount: scan.nestedMessageCount,
        nativeNumericCandidates: scan.numericCandidates.slice(0, 24).map(stripNumericCandidateInternalFields),
        nativeGeometryCandidates: scan.geometryCandidates.slice(0, 12),
        nativeGroupingHints: scan.groupingHints.slice(0, 12),
        nativeAnimationHints: scan.animationHints.slice(0, 12),
        nativeFieldSummaries: scan.fieldSummaries.slice(0, 12),
        nativeParser: "iwa-protobuf-field-scan"
      }
    };
  });

  return slides.length > 0 ? slides : [createPlaceholderSlide(0, "Native Keynote placeholder", undefined, deckSize)];
}

function chooseSlideIwaEntries(
  iwaEntries: Array<{ path: string; data: Uint8Array }>,
  slideOrderEvidence?: NativeSlideOrderEvidence
): Array<{ path: string; data: Uint8Array }> {
  const slides = iwaEntries.filter((entry) => isSlideIwaPath(entry.path));
  if (slideOrderEvidence) {
    const order = new Map(slideOrderEvidence.orderedPaths.map((entryPath, index) => [entryPath.toLowerCase(), index]));
    return slides.sort((left, right) => {
      const leftOrder = order.get(left.path.toLowerCase());
      const rightOrder = order.get(right.path.toLowerCase());
      if (leftOrder !== undefined && rightOrder !== undefined && leftOrder !== rightOrder) {
        return leftOrder - rightOrder;
      }
      if (leftOrder !== undefined) return -1;
      if (rightOrder !== undefined) return 1;
      return comparePartPaths(left.path, right.path);
    });
  }
  return slides.sort((left, right) => comparePartPaths(left.path, right.path));
}

function inferSlideOrderEvidence(iwaEntries: Array<{ path: string; data: Uint8Array }>): NativeSlideOrderEvidence | undefined {
  const documentEntry = iwaEntries.find((entry) => classifyIwaRole(entry.path) === "document");
  if (!documentEntry) {
    return undefined;
  }

  const slideEntries = iwaEntries.filter((entry) => isSlideIwaPath(entry.path));
  if (slideEntries.length < 2) {
    return undefined;
  }

  const slidePathLookup = createSlidePathLookup(slideEntries);
  const hints: Array<{ value: string; order: number; method: NativeSlideOrderEvidence["method"] }> = [];
  const payloads = expandIwaPayloads(documentEntry.data);
  let order = 0;

  for (const payload of payloads) {
    const records = parseIwaArchiveRecords(payload.data);
    for (const record of records) {
      if (record.identifier) {
        hints.push({ value: record.identifier, order, method: "document-archive-identifier" });
        order += 1;
      }
      for (const messageInfo of record.messageInfos) {
        for (const reference of messageInfo.objectReferences) {
          hints.push({ value: reference, order, method: "document-archive-identifier" });
          order += 1;
        }
      }
    }

    const rawScan = scanRawStrings(payload.data, order);
    for (const candidate of [...rawScan.referenceCandidates, ...rawScan.textCandidates.map((candidate) => ({ value: candidate.text, order: candidate.order }))]) {
      hints.push({ value: candidate.value, order: candidate.order, method: "document-reference" });
    }
    order += payload.data.byteLength + 1;
  }

  const orderedPaths: string[] = [];
  const seen = new Set<string>();
  let archiveHintCount = 0;
  for (const hint of hints.sort((left, right) => left.order - right.order)) {
    const matchedPath = matchSlidePathHint(hint.value, slidePathLookup);
    if (!matchedPath || seen.has(matchedPath.toLowerCase())) {
      continue;
    }
    if (hint.method === "document-archive-identifier") {
      archiveHintCount += 1;
    }
    seen.add(matchedPath.toLowerCase());
    orderedPaths.push(matchedPath);
  }

  if (orderedPaths.length < 2) {
    return undefined;
  }

  const fallbackPaths = slideEntries.map((entry) => entry.path).sort(comparePartPaths);
  for (const fallbackPath of fallbackPaths) {
    if (!seen.has(fallbackPath.toLowerCase())) {
      orderedPaths.push(fallbackPath);
    }
  }

  return {
    sourcePath: documentEntry.path,
    method: archiveHintCount > 0 ? "document-archive-identifier" : "document-reference",
    orderedPaths,
    matchedPathCount: seen.size,
    hintCount: hints.length,
    confidence: roundConfidence(Math.min(0.78, 0.32 + seen.size / slideEntries.length * 0.32 + (archiveHintCount > 0 ? 0.14 : 0)))
  };
}

function createSlidePathLookup(slideEntries: Array<{ path: string; data: Uint8Array }>): Map<string, string> {
  const lookup = new Map<string, string>();
  for (const entry of slideEntries) {
    const normalizedPath = normalizePartPath(entry.path).toLowerCase();
    const baseName = path.posix.basename(entry.path).toLowerCase();
    const stem = path.posix.basename(entry.path, path.posix.extname(entry.path)).toLowerCase();
    lookup.set(normalizedPath, entry.path);
    lookup.set(baseName, entry.path);
    lookup.set(stem, entry.path);
    lookup.set(normalizeAssetStem(stem), entry.path);
  }
  return lookup;
}

function matchSlidePathHint(value: string, lookup: Map<string, string>): string | undefined {
  const normalized = normalizeReferenceText(value);
  const decoded = safeDecodeUriComponent(normalized);
  for (const candidate of [normalized, decoded]) {
    const withoutQuery = candidate.split(/[?#]/, 1)[0] ?? candidate;
    const partPath = normalizePartPath(withoutQuery).toLowerCase();
    const baseName = path.posix.basename(partPath);
    const stem = path.posix.basename(baseName, path.posix.extname(baseName));
    for (const key of [partPath, baseName, stem, normalizeAssetStem(stem)]) {
      const match = lookup.get(key);
      if (match) {
        return match;
      }
    }
  }
  return undefined;
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
  deckSize: { width: number; height: number },
  geometryCandidate?: NativeIwaGeometryCandidate
): IRObject {
  const text = candidate.text;
  const marginX = Math.round(deckSize.width * 0.075);
  const top = Math.round(deckSize.height * 0.12);
  const lineHeight = objectIndex === 0 ? 76 : 54;
  const y = top + objectIndex * (lineHeight + 18);
  const bounds = geometryCandidate
    ? clampGeometryCandidateBounds(geometryCandidate.bounds, deckSize)
    : {
        x: marginX,
        y,
        width: deckSize.width - marginX * 2,
        height: Math.max(44, lineHeight)
      };

  return {
    id: `${slideId}-text-${objectIndex + 1}`,
    type: "text",
    name: objectIndex === 0 ? "Native text" : `Native text ${objectIndex + 1}`,
    bounds,
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
      nativeTextConfidence: candidate.confidence,
      ...(geometryCandidate
        ? {
            nativeGeometryCandidate: geometryCandidate,
            nativeGeometryConfidence: geometryCandidate.confidence
          }
        : {})
    }
  };
}

function createNativeTextDrawableObjects(
  slideId: string,
  messages: NativeIwaArchiveMessageEvidence[],
  sourcePath: string,
  deckSize: { width: number; height: number }
): IRObject[] {
  const targetIds = nativeBuildTargetIds(messages);
  if (targetIds.size === 0) {
    return [];
  }

  const textByArchiveId = new Map<string, NativeIwaTextContentEvidence>();
  for (const message of messages) {
    if (message.archiveIdentifier && message.textContent) {
      textByArchiveId.set(message.archiveIdentifier, message.textContent);
    }
  }

  const objects: IRObject[] = [];
  const seenDrawableIds = new Set<string>();
  for (const message of messages) {
    const drawable = message.textDrawable;
    const drawableId = message.archiveIdentifier;
    if (!drawable || !drawableId || !targetIds.has(drawableId) || seenDrawableIds.has(drawableId)) {
      continue;
    }
    const textEntry = drawable.textArchiveIds.map((id) => textByArchiveId.get(id)).find((entry) => entry?.text);
    if (!textEntry) {
      continue;
    }
    const object = createNativeTextDrawableObject(slideId, message, drawable, textEntry, objects.length, sourcePath, deckSize);
    if (object) {
      seenDrawableIds.add(drawableId);
      objects.push(object);
    }
  }
  return objects;
}

function nativeBuildTargetIds(messages: NativeIwaArchiveMessageEvidence[]): Set<string> {
  return new Set(
    messages
      .map((message) => message.build?.targetNativeId)
      .filter((targetId): targetId is string => Boolean(targetId))
  );
}

function createNativeTextDrawableObject(
  slideId: string,
  message: NativeIwaArchiveMessageEvidence,
  drawable: NativeIwaTextDrawableEvidence,
  textEntry: NativeIwaTextContentEvidence,
  objectIndex: number,
  sourcePath: string,
  deckSize: { width: number; height: number }
): IRObject | undefined {
  const text = cleanTextCandidate(textEntry.text);
  if (!text || !message.archiveIdentifier) {
    return undefined;
  }
  const bounds = drawable.bounds
    ? normalizeNativeTextDrawableBounds(drawable.bounds, deckSize)
    : {
        x: Math.round(deckSize.width * 0.1),
        y: Math.round(deckSize.height * 0.1) + objectIndex * 64,
        width: Math.round(deckSize.width * 0.72),
        height: 64
      };
  const fontSize = Math.max(14, Math.min(48, Math.round((bounds.height || 64) * 0.42)));
  return {
    id: `${slideId}-native-text-${objectIndex + 1}`,
    type: "text",
    name: `Native text drawable ${objectIndex + 1}`,
    bounds,
    opacity: 1,
    text: {
      plainText: text,
      runs: [
        {
          text,
          style: {
            fontFamily: "Helvetica Neue",
            fontSize,
            color: "#111827"
          }
        }
      ]
    },
    metadata: {
      nativeSourcePath: sourcePath,
      nativeExtraction: "typed-iwa-text-drawable",
      nativeArchiveMessageType: message.type,
      nativeArchiveIdentifier: message.archiveIdentifier,
      nativeArchiveMessageIndex: message.messageIndex,
      nativeArchivePayloadOffset: message.payloadOffset,
      nativeArchivePayloadLength: message.payloadLength,
      nativeArchiveObjectReferences: message.objectReferences,
      nativeTextArchiveIdentifier: drawable.textArchiveIds[0],
      nativeTextArchiveIdentifiers: drawable.textArchiveIds,
      nativeTextDrawableConfidence: drawable.confidence,
      nativeTextContentConfidence: textEntry.confidence,
      nativeTextDrawableFieldPaths: drawable.sourceFieldPaths,
      nativeTextContentFieldPaths: textEntry.sourceFieldPaths,
      ...(drawable.slideArchiveId ? { nativeTextDrawableSlideArchiveId: drawable.slideArchiveId } : {}),
      ...(drawable.bounds ? { nativeTextDrawableRawBounds: drawable.bounds } : {})
    }
  };
}

function normalizeNativeTextDrawableBounds(
  bounds: { x: number; y: number; width: number; height: number },
  deckSize: { width: number; height: number }
): { x: number; y: number; width: number; height: number } {
  const width = Math.max(8, Math.min(MAX_VISUAL_NUMERIC_VALUE, bounds.width));
  const height = Math.max(8, Math.min(MAX_VISUAL_NUMERIC_VALUE, bounds.height));
  return {
    x: roundGeometryNumber(Math.max(-deckSize.width, Math.min(deckSize.width * 2, bounds.x))),
    y: roundGeometryNumber(Math.max(-deckSize.height, Math.min(deckSize.height * 2, bounds.y))),
    width: roundGeometryNumber(width),
    height: roundGeometryNumber(height)
  };
}

function suppressDuplicateFallbackTextObjects(textObjects: IRObject[], nativeTextObjects: IRObject[]): IRObject[] {
  const nativeText = new Set(
    nativeTextObjects
      .filter((object): object is Extract<IRObject, { type: "text" }> => object.type === "text")
      .map((object) => normalizeTextForDuplicateCheck(object.text.plainText))
      .filter((text): text is string => Boolean(text))
  );
  if (nativeText.size === 0) {
    return textObjects;
  }
  return textObjects.filter((object) => {
    if (object.type !== "text") {
      return true;
    }
    const normalized = normalizeTextForDuplicateCheck(object.text.plainText);
    return !normalized || !nativeText.has(normalized);
  });
}

function normalizeTextForDuplicateCheck(text: string | undefined): string | undefined {
  const cleaned = text?.replace(/\s+/g, "").trim().toLowerCase();
  return cleaned || undefined;
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

function createPreviewImageObject(
  slideId: string,
  preview: NativePreviewAssetReference,
  sourcePath: string,
  deckSize: { width: number; height: number },
  suppressedTextCandidates: IwaTextCandidate[]
): IRObject {
  return {
    id: `${slideId}-preview-fallback`,
    type: "image",
    name: preview.name,
    bounds: {
      x: 0,
      y: 0,
      width: deckSize.width,
      height: deckSize.height
    },
    opacity: 1,
    source: {
      assetId: preview.assetId,
      metadata: {
        nativeAssetPath: preview.path,
        mimeType: preview.mimeType,
        ...(preview.width !== undefined && preview.height !== undefined
          ? {
              width: preview.width,
              height: preview.height
            }
          : {})
      }
    },
    metadata: {
      nativeSourcePath: sourcePath,
      nativeAssetPath: preview.path,
      nativeFallback: "full-slide-preview",
      nativePreviewSource: preview.previewSource,
      nativePreviewRole: preview.previewRole,
      nativeSuppressedTextCandidateCount: suppressedTextCandidates.length,
      nativeSuppressedTextCandidates: suppressedTextCandidates.slice(0, 8).map((candidate) => ({
        text: candidate.text,
        confidence: candidate.confidence,
        source: candidate.source,
        ...(candidate.fieldPath ? { fieldPath: candidate.fieldPath } : {})
      }))
    }
  };
}

function shouldUsePreviewFallback(
  textCandidates: IwaTextCandidate[],
  assetMatches: NativeAssetMatch[],
  preview: NativePreviewAssetReference
): boolean {
  if (preview.kind !== "image") {
    return false;
  }
  if (assetMatches.length > 0) {
    return false;
  }
  if (preview.previewRole === "snapshot") {
    return true;
  }
  if (textCandidates.length === 0) {
    return true;
  }
  return textCandidates.every((candidate) => candidate.confidence <= LOW_CONFIDENCE_TEXT_FALLBACK_THRESHOLD);
}

function selectPreviewFallbackAsset(
  previews: NativePreviewAssetReference[],
  slideIndex: number,
  slideCount: number
): NativePreviewAssetReference | undefined {
  if (previews.length === 0) {
    return undefined;
  }

  const slideNumber = slideIndex + 1;
  const numbered = previews.find((preview) => pathLooksLikeSlidePreview(preview.path, slideNumber));
  if (numbered) {
    return numbered;
  }
  if (slideCount === 1) {
    return previews.find((preview) => preview.previewRole === "preview") ?? previews[0];
  }
  const snapshots = previews.filter((preview) => preview.previewRole === "snapshot");
  if (slideIndex < snapshots.length) {
    return snapshots[slideIndex];
  }
  return undefined;
}

function pathLooksLikeSlidePreview(entryPath: string, slideNumber: number): boolean {
  const baseName = path.posix.basename(entryPath, path.posix.extname(entryPath)).toLowerCase();
  const escaped = String(slideNumber).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`(?:^|[-_\\s])(?:slide|page|preview|snapshot|st)?[-_\\s]*0*${escaped}(?:$|[-_\\s])`, "i").test(baseName);
}

function createAssetObject(
  slideId: string,
  match: NativeAssetMatch,
  objectIndex: number,
  sourcePath: string,
  deckSize: { width: number; height: number },
  geometryCandidate?: NativeIwaGeometryCandidate
): IRObject {
  const asset = match.asset;
  const marginX = Math.round(deckSize.width * 0.075);
  const aspect = asset.width && asset.height ? asset.width / asset.height : undefined;
  const width = Math.round(deckSize.width * 0.38);
  const height = aspect && aspect > 0 ? Math.round(width / aspect) : Math.round(deckSize.height * 0.38);
  const column = objectIndex % 2;
  const row = Math.floor(objectIndex / 2);
  const bounds = geometryCandidate
    ? match.geometryCandidate === geometryCandidate
      ? geometryCandidate.bounds
      : clampGeometryCandidateBounds(geometryCandidate.bounds, deckSize)
    : {
        x: marginX + column * Math.round(deckSize.width * 0.42),
        y: Math.round(deckSize.height * 0.48) + row * Math.round(deckSize.height * 0.12),
        width,
        height
      };
  const metadata = {
    nativeSourcePath: sourcePath,
    nativeAssetPath: asset.path,
    ...(asset.dataId ? { nativeAssetDataId: asset.dataId } : {}),
    ...(asset.width !== undefined && asset.height !== undefined
      ? {
          nativeAssetWidth: asset.width,
          nativeAssetHeight: asset.height
        }
      : {}),
    ...(asset.renderProxy
      ? {
          nativeAssetRenderableProxy: true,
          nativeAssetRenderableProxyFormat: asset.renderProxy.format,
          nativeAssetRenderableProxyMimeType: asset.renderProxy.mimeType,
          nativeAssetRenderableProxyByteLength: asset.renderProxy.byteLength,
          nativeAssetRenderableProxyChecksum: asset.renderProxy.checksum,
          nativeAssetSourceMimeType: asset.renderProxy.sourceMimeType ?? ""
        }
      : {}),
    nativeAssetEvidence: match.evidence,
    nativeAssetMatchConfidence: match.confidence,
    nativeAssetFieldPath: match.fieldPath,
    ...(match.suppressedAssets && match.suppressedAssets.length > 0
      ? {
          nativeSuppressedAssetVariantCount: match.suppressedAssets.length,
          nativeSuppressedAssetVariants: match.suppressedAssets.map((asset) => ({
            path: asset.path,
            mimeType: asset.mimeType,
            ...(asset.sourceMimeType && asset.sourceMimeType !== asset.mimeType ? { sourceMimeType: asset.sourceMimeType } : {}),
            ...(asset.width !== undefined && asset.height !== undefined ? { width: asset.width, height: asset.height } : {}),
            embedded: asset.embedded === true,
            ...(asset.renderProxy ? { renderProxy: true } : {})
          }))
        }
      : {}),
    ...(match.archiveMessage
      ? {
          nativeArchiveMessageType: match.archiveMessage.type,
          nativeArchiveIdentifier: match.archiveMessage.archiveIdentifier,
          nativeArchiveMessageIndex: match.archiveMessage.messageIndex,
          nativeArchivePayloadOffset: match.archiveMessage.payloadOffset,
          nativeArchivePayloadLength: match.archiveMessage.payloadLength,
          nativeArchiveObjectReferences: match.archiveMessage.objectReferences,
          nativeArchiveDataReferences: match.archiveMessage.dataReferences
        }
      : {}),
    nativeExtraction:
      match.source === "protobuf"
        ? "asset-protobuf-field-string-scan"
        : match.source === "archive"
        ? "asset-archive-info-data-reference"
        : "asset-raw-string-scan",
    ...(geometryCandidate
      ? {
          nativeGeometryCandidate: geometryCandidate,
          nativeGeometryConfidence: geometryCandidate.confidence
        }
      : {})
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
          mimeType: asset.mimeType,
          ...(asset.sourceMimeType && asset.sourceMimeType !== asset.mimeType ? { sourceMimeType: asset.sourceMimeType } : {}),
          ...(asset.renderProxy ? { renderProxy: true, renderProxyFormat: asset.renderProxy.format } : {}),
          ...(asset.width !== undefined && asset.height !== undefined
            ? {
                width: asset.width,
                height: asset.height
              }
            : {})
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
          mimeType: asset.mimeType,
          ...(asset.sourceMimeType && asset.sourceMimeType !== asset.mimeType ? { sourceMimeType: asset.sourceMimeType } : {}),
          ...(asset.renderProxy ? { renderProxy: true, renderProxyFormat: asset.renderProxy.format } : {}),
          ...(asset.width !== undefined && asset.height !== undefined
            ? {
                width: asset.width,
                height: asset.height
              }
            : {})
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

function applyNativePlacementGroupMetadata(objects: IRObject[]): void {
  const groups = new Map<string, IRObject[]>();
  for (const object of objects) {
    const key = nativePlacementGroupKey(object);
    if (!key) {
      continue;
    }
    const list = groups.get(key) ?? [];
    list.push(object);
    groups.set(key, list);
  }

  for (const [key, group] of groups) {
    if (group.length < 2) {
      continue;
    }
    const memberIds = group.map((object) => object.id);
    for (const [index, object] of group.entries()) {
      object.metadata = {
        ...(object.metadata ?? {}),
        nativePlacementGroupKey: key,
        nativePlacementGroupIndex: index,
        nativePlacementGroupSize: group.length,
        nativePlacementGroupObjectIds: memberIds
      };
    }
  }
}

function nativePlacementGroupKey(object: IRObject): string | undefined {
  if (object.type !== "image" && object.type !== "media") {
    return undefined;
  }
  const dataReferences = Array.isArray(object.metadata?.nativeArchiveDataReferences)
    ? object.metadata.nativeArchiveDataReferences.map(String)
    : object.metadata?.nativeAssetDataId
      ? [String(object.metadata.nativeAssetDataId)]
      : [];
  if (dataReferences.length === 0 || !object.bounds) {
    return undefined;
  }
  const bounds = object.bounds;
  return [
    "placement",
    uniqueStrings(dataReferences).sort(comparePartPaths).join("+"),
    Math.round(bounds.x),
    Math.round(bounds.y),
    Math.round(bounds.width),
    Math.round(bounds.height)
  ].join(":");
}

interface NativeBuildAnimationRecovery {
  events: AnimationEvent[];
  dependencyEdges: TimingDependencyEdge[];
  durationMs: number;
  buildRecords: NativeIwaBuildEvidence[];
  timingRecords: NativeIwaBuildTimingEvidence[];
  unresolvedBuildCount: number;
}

interface NativeBuildTimingAnchor {
  primaryEventId: string;
  eventIds: string[];
  startMs: number;
  durationMs: number;
}

function recoverNativeBuildAnimations(
  slideId: string,
  objects: IRObject[],
  messages: NativeIwaArchiveMessageEvidence[]
): NativeBuildAnimationRecovery {
  const buildRecords = messages.map((message) => message.build).filter((build): build is NativeIwaBuildEvidence => Boolean(build));
  const timingRecords = messages
    .map((message) => message.buildTiming)
    .filter((timing): timing is NativeIwaBuildTimingEvidence => Boolean(timing));
  if (buildRecords.length === 0) {
    return { events: [], dependencyEdges: [], durationMs: 2500, buildRecords, timingRecords, unresolvedBuildCount: 0 };
  }

  const objectLookup = createNativeBuildObjectLookup(objects);
  const timingByBuildId = new Map<string, NativeIwaBuildTimingEvidence>();
  for (const timing of timingRecords) {
    if (timing.buildId && !timingByBuildId.has(timing.buildId)) {
      timingByBuildId.set(timing.buildId, timing);
    }
  }
  const buildById = new Map<string, NativeIwaBuildEvidence>();
  for (const build of buildRecords) {
    if (build.buildId && !buildById.has(build.buildId)) {
      buildById.set(build.buildId, build);
    }
  }
  const sequencedBuilds = sequenceNativeBuildRecords(buildRecords, timingRecords, buildById);

  const events: AnimationEvent[] = [];
  const dependencyEdges: TimingDependencyEdge[] = [];
  let cursorMs = 0;
  let cursorAnchor: NativeBuildTimingAnchor | undefined;
  let groupStartMs = 0;
  let groupAnchor: NativeBuildTimingAnchor | undefined;
  let unresolvedBuildCount = 0;
  for (const [index, build] of sequencedBuilds.entries()) {
    const timing = build.buildId ? timingByBuildId.get(build.buildId) : undefined;
    const resolutions = resolveNativeBuildTargets(build, objectLookup);
    const durationMs = clampNativeBuildDuration(timing?.durationMs ?? build.durationMs);
    const delayMs = clampNativeBuildDelay(timing?.delayMs ?? build.delayMs);
    const startsWithPrevious = isNativeTimingWithPrevious(timing);
    const baseStartMs = startsWithPrevious ? groupStartMs : cursorMs;
    const startMs = baseStartMs + delayMs;
    const sequenceReference = cursorAnchor;
    const groupReference = groupAnchor ?? cursorAnchor;
    const generatedEvents: AnimationEvent[] = [];
    for (const [targetIndex, resolution] of resolutions.entries()) {
      const event = nativeBuildEventFromEvidence(slideId, index, targetIndex, build, timing, resolution, startMs, durationMs);
      if (event) {
        events.push(event);
        generatedEvents.push(event);
      }
    }

    const generatedAnchor = createNativeBuildTimingAnchor(generatedEvents, startMs, durationMs);
    if (generatedAnchor) {
      addNativeBuildSiblingTimingEdges(dependencyEdges, generatedAnchor);
      if (startsWithPrevious) {
        addNativeBuildTimingDependencyEdge(
          dependencyEdges,
          groupReference?.primaryEventId,
          generatedAnchor.primaryEventId,
          "with",
          startMs - (groupReference?.startMs ?? startMs)
        );
      } else if (sequenceReference) {
        addNativeBuildTimingDependencyEdge(
          dependencyEdges,
          sequenceReference.primaryEventId,
          generatedAnchor.primaryEventId,
          "after",
          startMs - (sequenceReference.startMs + sequenceReference.durationMs)
        );
      }
    }

    const generatedEventCount = generatedEvents.length;
    if (generatedEventCount > 0 || timing || build.durationMs !== undefined) {
      if (!startsWithPrevious) {
        groupStartMs = startMs;
        groupAnchor = generatedAnchor;
      }
      const endMs = startMs + durationMs;
      if (!cursorAnchor || endMs > cursorMs) {
        cursorMs = endMs;
        cursorAnchor = generatedAnchor;
      }
    }
    if (generatedEventCount === 0) {
      unresolvedBuildCount += 1;
    }
  }

  return {
    events,
    dependencyEdges,
    durationMs: Math.max(2500, events.reduce((max, event) => Math.max(max, eventEndMs(event)), 0) + 500),
    buildRecords,
    timingRecords,
    unresolvedBuildCount
  };
}

function isNativeTimingWithPrevious(timing: NativeIwaBuildTimingEvidence | undefined): boolean {
  return timing?.startRelation === "withPrevious" || timing?.startsWithPrevious === true;
}

function createNativeBuildTimingAnchor(events: AnimationEvent[], startMs: number, durationMs: number): NativeBuildTimingAnchor | undefined {
  if (events.length === 0) {
    return undefined;
  }
  return {
    primaryEventId: events[0]!.id,
    eventIds: events.map((event) => event.id),
    startMs,
    durationMs
  };
}

function addNativeBuildSiblingTimingEdges(edges: TimingDependencyEdge[], anchor: NativeBuildTimingAnchor): void {
  for (const eventId of anchor.eventIds.slice(1)) {
    addNativeBuildTimingDependencyEdge(edges, anchor.primaryEventId, eventId, "with", 0);
  }
}

function addNativeBuildTimingDependencyEdge(
  edges: TimingDependencyEdge[],
  from: string | undefined,
  to: string | undefined,
  relation: TimingDependencyEdge["relation"],
  offsetMs: number
): void {
  if (!from || !to || from === to) {
    return;
  }
  const roundedOffsetMs = Number.isFinite(offsetMs) ? Math.round(offsetMs) : 0;
  const key = `${from}|${to}|${relation}|${roundedOffsetMs}`;
  if (edges.some((edge) => `${edge.from}|${edge.to}|${edge.relation}|${Math.round(Number(edge.offsetMs ?? 0))}` === key)) {
    return;
  }
  edges.push({
    id: `native-timing-${nativeBuildSlug(from)}-${nativeBuildSlug(to)}-${relation}`,
    from,
    to,
    relation,
    ...(roundedOffsetMs !== 0 ? { offsetMs: roundedOffsetMs } : {})
  });
}

function createNativeBuildObjectLookup(objects: IRObject[]): Map<string, Array<{ object: IRObject; method: string }>> {
  const lookup = new Map<string, Array<{ object: IRObject; method: string }>>();
  for (const object of flattenNativeObjects(objects)) {
    addNativeBuildObjectLookupValue(lookup, object.id, object, "object-id");
    if (object.morphKey) {
      addNativeBuildObjectLookupValue(lookup, object.morphKey, object, "morph-key");
    }
    addNativeBuildObjectLookupValue(lookup, object.metadata?.nativeArchiveIdentifier, object, "archive-identifier");
    addNativeBuildObjectLookupValues(lookup, object.metadata?.nativeArchiveObjectReferences, object, "archive-object-reference");
  }
  return lookup;
}

function createNativeTypedVisualRecords(messages: NativeIwaArchiveMessageEvidence[]): NativeTypedVisualRecordEvidence[] {
  return messages
    .filter((message) => isTypedVisualGeometryMessage(message.type))
    .map((message) => {
      const numericCandidates = message.fieldSummaries
        .filter((summary) => summary.sampleNumericValue !== undefined)
        .map((summary) => ({
          fieldPath: summary.fieldPath,
          fieldNumber: summary.fieldNumber,
          wireType: summary.wireType,
          value: roundGeometryNumber(summary.sampleNumericValue!),
          encoding: summary.sampleNumericEncoding ?? "varint",
          confidence: 0.58
        }));
      const confidence =
        0.42 +
        (message.archiveIdentifier ? 0.08 : 0) +
        (message.geometryCandidates.length > 0 ? 0.26 : 0) +
        (message.dataReferences.length > 0 ? 0.12 : 0) +
        (message.type === 3006 ? 0.04 : 0);
      return {
        ...(message.archiveIdentifier ? { archiveIdentifier: message.archiveIdentifier } : {}),
        messageIndex: message.messageIndex,
        type: message.type!,
        payloadOffset: message.payloadOffset,
        payloadLength: message.payloadLength,
        objectReferences: message.objectReferences,
        dataReferences: message.dataReferences,
        geometryCandidates: message.geometryCandidates.slice(0, 4),
        fieldSummaries: message.fieldSummaries.slice(0, 16),
        numericCandidates: numericCandidates.slice(0, 16),
        confidence: Math.min(0.92, confidence)
      };
    })
    .sort((left, right) => {
      const type = left.type - right.type;
      if (type !== 0) return type;
      return Number(left.archiveIdentifier ?? 0) - Number(right.archiveIdentifier ?? 0) || left.payloadOffset - right.payloadOffset;
    });
}

function sequenceNativeBuildRecords(
  buildRecords: NativeIwaBuildEvidence[],
  timingRecords: NativeIwaBuildTimingEvidence[],
  buildById: Map<string, NativeIwaBuildEvidence>
): NativeIwaBuildEvidence[] {
  const ordered: NativeIwaBuildEvidence[] = [];
  const seen = new Set<NativeIwaBuildEvidence>();
  for (const timing of timingRecords) {
    const build = timing.buildId ? buildById.get(timing.buildId) : undefined;
    if (!build || seen.has(build)) {
      continue;
    }
    seen.add(build);
    ordered.push(build);
  }
  for (const build of buildRecords) {
    if (!seen.has(build)) {
      ordered.push(build);
    }
  }
  return ordered;
}

function flattenNativeObjects(objects: IRObject[]): IRObject[] {
  const out: IRObject[] = [];
  for (const object of objects) {
    out.push(object);
    if (object.type === "group") {
      out.push(...flattenNativeObjects(object.children));
    }
  }
  return out;
}

function addNativeBuildObjectLookupValues(
  lookup: Map<string, Array<{ object: IRObject; method: string }>>,
  values: unknown,
  object: IRObject,
  method: string
): void {
  if (!Array.isArray(values)) {
    return;
  }
  for (const value of values) {
    addNativeBuildObjectLookupValue(lookup, value, object, method);
  }
}

function addNativeBuildObjectLookupValue(
  lookup: Map<string, Array<{ object: IRObject; method: string }>>,
  value: unknown,
  object: IRObject,
  method: string
): void {
  const key = nativeBuildLookupKey(value);
  if (!key) {
    return;
  }
  const list = lookup.get(key) ?? [];
  if (!list.some((entry) => entry.object.id === object.id && entry.method === method)) {
    list.push({ object, method });
  }
  lookup.set(key, list);
}

function resolveNativeBuildTargets(
  build: NativeIwaBuildEvidence,
  lookup: Map<string, Array<{ object: IRObject; method: string }>>
): Array<{ object: IRObject; method: string; nativeTargetId?: string }> {
  const targetIds = uniqueStrings([build.targetNativeId, ...build.objectReferences].filter((value): value is string => Boolean(value)));
  for (const targetId of targetIds) {
    const key = nativeBuildLookupKey(targetId);
    const matches = key ? lookup.get(key) ?? [] : [];
    const unique = uniqueObjectMatches(matches);
    if (unique.length > 0) {
      return unique.map((match) => ({
        object: match.object,
        method: match.method,
        nativeTargetId: targetId
      }));
    }
  }
  return [];
}

function uniqueObjectMatches(matches: Array<{ object: IRObject; method: string }>): Array<{ object: IRObject; method: string }> {
  const out: Array<{ object: IRObject; method: string }> = [];
  const seen = new Set<string>();
  for (const match of matches) {
    if (seen.has(match.object.id)) {
      continue;
    }
    seen.add(match.object.id);
    out.push(match);
  }
  return out;
}

function nativeBuildEventFromEvidence(
  slideId: string,
  index: number,
  targetIndex: number,
  build: NativeIwaBuildEvidence,
  timing: NativeIwaBuildTimingEvidence | undefined,
  resolution: { object: IRObject; method: string; nativeTargetId?: string },
  startMs: number,
  durationMs: number
): AnimationEvent | undefined {
  const direction = normalizeNativeBuildDirection(build.direction, build.effect);
  if (build.effect && /movie-start/i.test(build.effect) && resolution.object.type === "media") {
    return {
      id: nativeBuildEventId(slideId, build, timing, index, targetIndex, resolution.object.id, "media-play"),
      kind: "media",
      label: "Keynote movie start",
      targetId: resolution.object.id,
      action: "play",
      start: { type: "absolute", atMs: startMs },
      durationMs: 0,
      fill: "forwards",
      metadata: nativeBuildEventMetadata(build, timing, resolution, "media-play")
    };
  }
  if (build.effect && /motion-path/i.test(build.effect) && build.motionPath) {
    const startPoint = build.motionPath.points[0];
    const endPoint = build.motionPath.points[build.motionPath.points.length - 1];
    if (startPoint && endPoint) {
      return {
        id: nativeBuildEventId(slideId, build, timing, index, targetIndex, resolution.object.id, "motion-path"),
        kind: "keyframes",
        label: "Keynote motion path",
        targetId: resolution.object.id,
        start: { type: "absolute", atMs: startMs },
        durationMs,
        fill: "both",
        easing: "easeInOut",
        tracks: [
          {
            property: "transform.translateX",
            interpolation: "number",
            keyframes: [
              { offset: 0, value: startPoint.x },
              { offset: 1, value: endPoint.x }
            ]
          },
          {
            property: "transform.translateY",
            interpolation: "number",
            keyframes: [
              { offset: 0, value: startPoint.y },
              { offset: 1, value: endPoint.y }
            ]
          }
        ],
        metadata: nativeBuildEventMetadata(build, timing, resolution, "motion-path")
      };
    }
  }
  if (direction !== "in" && direction !== "out") {
    return undefined;
  }
  if (build.effect && /motion-path|movie-start/i.test(build.effect)) {
    return undefined;
  }
  if (isNativeAppearEffect(build.effect)) {
    return nativeAppearBuildEventFromEvidence(slideId, index, targetIndex, build, timing, resolution, startMs, durationMs, direction);
  }
  if (build.effect && /wipe/i.test(build.effect) && resolution.object.bounds) {
    return nativeWipeBuildEventFromEvidence(slideId, index, targetIndex, build, timing, resolution, startMs, durationMs, direction);
  }
  if (build.effect && /blur/i.test(build.effect)) {
    return nativeBlurBuildEventFromEvidence(slideId, index, targetIndex, build, timing, resolution, startMs, durationMs, direction);
  }
  if (build.effect && /dissolve/i.test(build.effect)) {
    return nativeDissolveBuildEventFromEvidence(slideId, index, targetIndex, build, timing, resolution, startMs, durationMs, direction);
  }
  if (build.effect && /crumble/i.test(build.effect)) {
    return nativeCrumbleBuildEventFromEvidence(slideId, index, targetIndex, build, timing, resolution, startMs, durationMs, direction);
  }
  if (build.effect && /anvil/i.test(build.effect)) {
    return nativeAnvilBuildEventFromEvidence(slideId, index, targetIndex, build, timing, resolution, startMs, durationMs, direction);
  }
  const from = direction === "out" ? 1 : 0;
  const to = direction === "out" ? 0 : 1;
  return {
    id: nativeBuildEventId(slideId, build, timing, index, targetIndex, resolution.object.id, `opacity-${direction}`),
    kind: "keyframes",
    label: `Keynote ${direction === "out" ? "build out" : "build in"}`,
    targetId: resolution.object.id,
    start: { type: "absolute", atMs: startMs },
    durationMs,
    fill: "both",
    easing: "easeInOut",
    tracks: [
      {
        property: "opacity",
        interpolation: "number",
        keyframes: [
          { offset: 0, value: from },
          { offset: 1, value: to }
        ]
      }
    ],
    metadata: nativeBuildEventMetadata(build, timing, resolution, `opacity-${direction}`)
  };
}

function nativeAppearBuildEventFromEvidence(
  slideId: string,
  index: number,
  targetIndex: number,
  build: NativeIwaBuildEvidence,
  timing: NativeIwaBuildTimingEvidence | undefined,
  resolution: { object: IRObject; method: string; nativeTargetId?: string },
  startMs: number,
  durationMs: number,
  direction: "in" | "out"
): AnimationEvent {
  const from = direction === "out" ? 1 : 0;
  const to = direction === "out" ? 0 : 1;
  return {
    id: nativeBuildEventId(slideId, build, timing, index, targetIndex, resolution.object.id, `appear-${direction}`),
    kind: "keyframes",
    label: `Keynote appear ${direction}`,
    targetId: resolution.object.id,
    start: { type: "absolute", atMs: startMs },
    durationMs,
    fill: "both",
    easing: { type: "steps", count: 1, position: "end" },
    tracks: [
      {
        property: "opacity",
        interpolation: "discrete",
        keyframes: [
          { offset: 0, value: from },
          { offset: 1, value: to }
        ]
      }
    ],
    metadata: {
      ...nativeBuildEventMetadata(build, timing, resolution, `appear-${direction}`),
      nativeBuildDegradation: "Keynote Appear is represented as a discrete opacity step instead of a fade."
    }
  };
}

function nativeBlurBuildEventFromEvidence(
  slideId: string,
  index: number,
  targetIndex: number,
  build: NativeIwaBuildEvidence,
  timing: NativeIwaBuildTimingEvidence | undefined,
  resolution: { object: IRObject; method: string; nativeTargetId?: string },
  startMs: number,
  durationMs: number,
  direction: "in" | "out"
): AnimationEvent {
  const fromOpacity = direction === "out" ? 1 : 0;
  const toOpacity = direction === "out" ? 0 : 1;
  const fromBlur = direction === "out" ? 0 : 18;
  const toBlur = direction === "out" ? 18 : 0;
  return {
    id: nativeBuildEventId(slideId, build, timing, index, targetIndex, resolution.object.id, `blur-${direction}`),
    kind: "keyframes",
    label: `Keynote blur ${direction}`,
    targetId: resolution.object.id,
    start: { type: "absolute", atMs: startMs },
    durationMs,
    fill: "both",
    easing: "easeInOut",
    tracks: [
      twoPointTrack("opacity", fromOpacity, toOpacity),
      twoPointTrack("filter.blurPx", fromBlur, toBlur)
    ],
    metadata: {
      ...nativeBuildEventMetadata(build, timing, resolution, `blur-${direction}`),
      nativeBuildDegradation: "Keynote blur radius/curve is approximated with a fixed CSS blur track."
    }
  };
}

function nativeDissolveBuildEventFromEvidence(
  slideId: string,
  index: number,
  targetIndex: number,
  build: NativeIwaBuildEvidence,
  timing: NativeIwaBuildTimingEvidence | undefined,
  resolution: { object: IRObject; method: string; nativeTargetId?: string },
  startMs: number,
  durationMs: number,
  direction: "in" | "out"
): AnimationEvent {
  const from = direction === "out" ? 1 : 0;
  const to = direction === "out" ? 0 : 1;
  return {
    id: nativeBuildEventId(slideId, build, timing, index, targetIndex, resolution.object.id, `dissolve-${direction}`),
    kind: "keyframes",
    label: `Keynote dissolve ${direction}`,
    targetId: resolution.object.id,
    start: { type: "absolute", atMs: startMs },
    durationMs,
    fill: "both",
    easing: "easeInOut",
    tracks: [twoPointTrack("opacity", from, to)],
    metadata: {
      ...nativeBuildEventMetadata(build, timing, resolution, `dissolve-${direction}`),
      ...(nativeBuildGranularity(build.effect)
        ? {
            nativeBuildGranularity: nativeBuildGranularity(build.effect)!
          }
        : {}),
      nativeBuildDegradation: build.effect && /character/i.test(build.effect) ? "Per-character dissolve is approximated as object-level opacity." : "Dissolve is approximated as object-level opacity."
    }
  };
}

function nativeCrumbleBuildEventFromEvidence(
  slideId: string,
  index: number,
  targetIndex: number,
  build: NativeIwaBuildEvidence,
  timing: NativeIwaBuildTimingEvidence | undefined,
  resolution: { object: IRObject; method: string; nativeTargetId?: string },
  startMs: number,
  durationMs: number,
  direction: "in" | "out"
): AnimationEvent {
  const entering = direction === "in";
  return {
    id: nativeBuildEventId(slideId, build, timing, index, targetIndex, resolution.object.id, `crumble-${direction}`),
    kind: "keyframes",
    label: `Keynote crumble ${direction}`,
    targetId: resolution.object.id,
    start: { type: "absolute", atMs: startMs },
    durationMs,
    fill: "both",
    easing: "easeInOut",
    tracks: [
      twoPointTrack("opacity", entering ? 0 : 1, entering ? 1 : 0),
      twoPointTrack("transform.translateY", entering ? 18 : 0, entering ? 0 : 36),
      twoPointTrack("transform.rotateDeg", entering ? -4 : 0, entering ? 0 : 8),
      twoPointTrack("transform.scale", entering ? 0.96 : 1, entering ? 1 : 0.9)
    ],
    metadata: {
      ...nativeBuildEventMetadata(build, timing, resolution, `crumble-${direction}`),
      nativeBuildDegradation: "Keynote crumble particle/shatter behavior is approximated with object-level fall, rotation, scale, and opacity."
    }
  };
}

function nativeAnvilBuildEventFromEvidence(
  slideId: string,
  index: number,
  targetIndex: number,
  build: NativeIwaBuildEvidence,
  timing: NativeIwaBuildTimingEvidence | undefined,
  resolution: { object: IRObject; method: string; nativeTargetId?: string },
  startMs: number,
  durationMs: number,
  direction: "in" | "out"
): AnimationEvent {
  const entering = direction === "in";
  return {
    id: nativeBuildEventId(slideId, build, timing, index, targetIndex, resolution.object.id, `anvil-${direction}`),
    kind: "keyframes",
    label: `Keynote anvil ${direction}`,
    targetId: resolution.object.id,
    start: { type: "absolute", atMs: startMs },
    durationMs,
    fill: "both",
    easing: entering ? "backOut" : "easeInOut",
    tracks: [
      twoPointTrack("opacity", entering ? 0 : 1, entering ? 1 : 0),
      twoPointTrack("transform.translateY", entering ? -48 : 0, entering ? 0 : 48),
      twoPointTrack("transform.scale", entering ? 1.16 : 1, entering ? 1 : 0.86),
      twoPointTrack("transform.rotateDeg", entering ? -2 : 0, entering ? 0 : 4)
    ],
    metadata: {
      ...nativeBuildEventMetadata(build, timing, resolution, `anvil-${direction}`),
      nativeBuildDegradation: "Keynote anvil impact/physics behavior is approximated with object-level drop, scale, rotation, and opacity."
    }
  };
}

function twoPointTrack(property: string, from: number, to: number): KeyframeTrack {
  return {
    property,
    interpolation: "number",
    keyframes: [
      { offset: 0, value: from },
      { offset: 1, value: to }
    ]
  };
}

function nativeWipeBuildEventFromEvidence(
  slideId: string,
  index: number,
  targetIndex: number,
  build: NativeIwaBuildEvidence,
  timing: NativeIwaBuildTimingEvidence | undefined,
  resolution: { object: IRObject; method: string; nativeTargetId?: string },
  startMs: number,
  durationMs: number,
  direction: "in" | "out"
): AnimationEvent {
  const bounds = resolution.object.bounds!;
  const collapsed = { ...bounds, width: Math.max(1, Math.min(bounds.width, 1)) };
  return {
    id: nativeBuildEventId(slideId, build, timing, index, targetIndex, resolution.object.id, `wipe-${direction}`),
    kind: "keyframes",
    label: `Keynote wipe ${direction}`,
    targetId: resolution.object.id,
    start: { type: "absolute", atMs: startMs },
    durationMs,
    fill: "both",
    easing: "easeInOut",
    tracks: [
      {
        property: "bounds",
        interpolation: "matrix",
        keyframes:
          direction === "out"
            ? [
                { offset: 0, value: bounds },
                { offset: 1, value: collapsed }
              ]
            : [
                { offset: 0, value: collapsed },
                { offset: 1, value: bounds }
              ]
      },
      {
        property: "opacity",
        interpolation: "number",
        keyframes:
          direction === "out"
            ? [
                { offset: 0, value: 1 },
                { offset: 1, value: 0 }
              ]
            : [
                { offset: 0, value: 0 },
                { offset: 0.05, value: 1 },
                { offset: 1, value: 1 }
              ]
      }
    ],
    metadata: {
      ...nativeBuildEventMetadata(build, timing, resolution, `wipe-${direction}`),
      nativeBuildDegradation: "wipe approximated as left-to-right bounds reveal; Keynote mask direction/easing is not fully decoded"
    }
  };
}

function nativeBuildEventMetadata(
  build: NativeIwaBuildEvidence,
  timing: NativeIwaBuildTimingEvidence | undefined,
  resolution: { method: string; nativeTargetId?: string },
  fallback: string
): JSONRecord {
  return {
    nativeSource: "keynote-iwa-build",
    nativeBuildFallback: fallback,
    ...(build.buildId ? { nativeBuildId: build.buildId } : {}),
    ...(timing?.timingId ? { nativeBuildTimingId: timing.timingId } : {}),
    ...(timing?.buildId ? { nativeBuildTimingBuildId: timing.buildId } : {}),
    ...(build.targetNativeId ? { nativeBuildTargetId: build.targetNativeId } : {}),
    ...(resolution.nativeTargetId ? { nativeResolvedTargetId: resolution.nativeTargetId } : {}),
    nativeResolvedTargetBy: resolution.method,
    nativeBuildObjectReferences: build.objectReferences,
    ...(build.delivery ? { nativeBuildDelivery: build.delivery } : {}),
    ...(build.direction ? { nativeBuildDirection: build.direction } : {}),
    ...(build.effect ? { nativeBuildEffect: build.effect } : {}),
    nativeBuildConfidence: build.confidence,
    ...(build.motionPath
      ? {
          nativeMotionPathRelative: build.motionPath.relative,
          nativeMotionPathPoints: build.motionPath.points,
          ...(build.motionPath.extentPoints ? { nativeMotionPathExtentPoints: build.motionPath.extentPoints } : {}),
          nativeMotionPathConfidence: build.motionPath.confidence,
          nativeMotionPathFieldPaths: build.motionPath.sourceFieldPaths
        }
      : {}),
    ...(timing ? { nativeBuildTimingConfidence: timing.confidence } : {}),
    ...(timing?.group !== undefined ? { nativeBuildTimingGroup: timing.group } : {}),
    ...(timing?.triggerGroupRaw !== undefined ? { nativeBuildTimingTriggerGroupRaw: timing.triggerGroupRaw } : {}),
    ...(timing?.rawField5 !== undefined ? { nativeBuildTimingRawField5: timing.rawField5 } : {}),
    ...(timing?.rawField6 !== undefined ? { nativeBuildTimingRawField6: timing.rawField6 } : {}),
    ...(timing?.rawField72 !== undefined ? { nativeBuildTimingRawField72: timing.rawField72 } : {}),
    ...(timing?.startRelation ? { nativeBuildStartRelation: timing.startRelation } : {}),
    ...(timing?.startsWithPrevious !== undefined ? { nativeBuildStartsWithPrevious: timing.startsWithPrevious } : {}),
    ...(timing?.afterPrevious !== undefined ? { nativeBuildAfterPrevious: timing.afterPrevious } : {})
  };
}

function nativeBuildGranularity(effect: string | undefined): "character" | "word" | "line" | "object" | undefined {
  const normalized = effect?.toLowerCase() ?? "";
  if (/\bcharacter\b/.test(normalized)) return "character";
  if (/\bword\b/.test(normalized)) return "word";
  if (/\bline\b/.test(normalized)) return "line";
  if (normalized) return "object";
  return undefined;
}

function isNativeAppearEffect(effect: string | undefined): boolean {
  return /\b(?:bc-)?appear\b/i.test(effect ?? "");
}

function normalizeNativeBuildDirection(direction: string | undefined, effect: string | undefined): "in" | "out" | "action" | undefined {
  const normalized = `${direction ?? ""} ${effect ?? ""}`.toLowerCase();
  if (/\bout\b/.test(normalized)) return "out";
  if (/\baction\b|motion-path|movie-start/.test(normalized)) return "action";
  if (/\bin\b/.test(normalized)) return "in";
  if (/appear|dissolve|blur|fade|build/i.test(normalized)) return "in";
  return undefined;
}

function nativeBuildEventId(
  slideId: string,
  build: NativeIwaBuildEvidence,
  timing: NativeIwaBuildTimingEvidence | undefined,
  index: number,
  targetIndex: number,
  targetId: string,
  suffix: string
): string {
  const stableId = build.buildId ?? timing?.timingId ?? build.targetNativeId ?? String(index + 1);
  const targetSuffix = targetIndex > 0 ? `-${nativeBuildSlug(targetId)}` : "";
  return `${slideId}-native-build-${nativeBuildSlug(stableId)}${targetSuffix}-${suffix}`;
}

function nativeBuildSlug(value: string): string {
  const slug = value.replace(/[^A-Za-z0-9_-]+/g, "-").replace(/^-+|-+$/g, "");
  return slug || "unknown";
}

function eventEndMs(event: AnimationEvent): number {
  const start = event.start?.type === "absolute" ? event.start.atMs : event.delayMs ?? 0;
  return Math.max(0, start) + Math.max(0, event.durationMs ?? 0);
}

function clampNativeBuildDuration(value: number | undefined): number {
  if (value === undefined || !Number.isFinite(value) || value <= 0) {
    return 500;
  }
  return Math.max(1, Math.min(120000, Math.round(value)));
}

function clampNativeBuildDelay(value: number | undefined): number {
  if (value === undefined || !Number.isFinite(value) || value <= 0) {
    return 0;
  }
  return Math.max(0, Math.min(120000, Math.round(value)));
}

function nativeBuildLookupKey(value: unknown): string | undefined {
  if (typeof value !== "string" && typeof value !== "number") {
    return undefined;
  }
  const normalized = String(value).trim();
  return normalized ? normalized.toLowerCase() : undefined;
}

function scanIwaEntry(data: Uint8Array): IwaScanResult {
  return scanIwaPayloads(expandIwaPayloads(data));
}

function scanIwaPayloads(payloads: ExpandedIwaPayload[]): IwaScanResult {
  const textCandidates = new Map<string, IwaTextCandidate>();
  const referenceCandidates = new Map<string, IwaReferenceCandidate>();
  const numericCandidates = new Map<string, IwaNumericCandidate>();
  const animationHints = new Map<string, NativeIwaAnimationHint>();
  const fieldSummaries = new Map<string, NativeIwaFieldSummary>();
  const archiveRecords = new Map<string, NativeIwaArchiveRecordEvidence>();
  const archiveMessages = new Map<string, NativeIwaArchiveMessageEvidence>();
  let protobufFieldCount = 0;
  let nestedMessageCount = 0;
  let rawStringCount = 0;
  let typedArchiveMessageCount = 0;
  let orderBase = 0;

  for (const payload of payloads) {
    const framedRecords = parseIwaArchiveRecords(payload.data);
    for (const record of framedRecords) {
      addBestArchiveRecordEvidence(archiveRecords, archiveRecordEvidence(record));
    }
    const scanData =
      framedRecords.length > 0
        ? archiveMessagePayloadsFromRecords(payload.data, framedRecords)
        : [{ data: payload.data, archiveMessage: undefined }];
    typedArchiveMessageCount += scanData.filter((item) => item.archiveMessage).length;

    for (const item of scanData) {
      const data = item.data;
      const protobufScan = scanProtobufPayload(data, orderBase);
      const suppressPresentationText = isNativeInternalTextMessageType(item.archiveMessage?.messageInfo.type);
      protobufFieldCount += protobufScan.protobufFieldCount;
      nestedMessageCount += protobufScan.nestedMessageCount;
      for (const summary of protobufScan.fieldSummaries) {
        mergeFieldSummary(fieldSummaries, summary);
      }
      if (!suppressPresentationText) {
        for (const candidate of protobufScan.textCandidates) {
          addBestTextCandidate(textCandidates, candidate);
        }
      }
      for (const candidate of protobufScan.referenceCandidates) {
        addBestReferenceCandidate(referenceCandidates, candidate);
      }
      for (const candidate of protobufScan.numericCandidates) {
        addBestNumericCandidate(numericCandidates, candidate);
      }
      for (const hint of protobufScan.animationHints) {
        addBestAnimationHint(animationHints, hint);
      }
      if (item.archiveMessage) {
        const typedMessage = createArchiveMessageEvidence(
          item.archiveMessage.record,
          item.archiveMessage.messageInfo,
          item.archiveMessage.messageIndex,
          protobufScan,
          data
        );
        addBestArchiveMessageEvidence(archiveMessages, typedMessage);
      }

      const rawScan = scanRawStrings(data, orderBase + data.byteLength + 1);
      rawStringCount += rawScan.rawStringCount;
      if (!suppressPresentationText) {
        for (const candidate of rawScan.textCandidates) {
          addBestTextCandidate(textCandidates, candidate);
        }
      }
      for (const candidate of rawScan.referenceCandidates) {
        addBestReferenceCandidate(referenceCandidates, candidate);
      }
      for (const hint of rawScan.animationHints) {
        addBestAnimationHint(animationHints, hint);
      }
      orderBase += data.byteLength * 2 + 2;
    }
  }

  const fieldSummaryList = Array.from(fieldSummaries.values()).sort(compareFieldSummaries).slice(0, MAX_FIELD_SUMMARIES_PER_STREAM);
  const numericCandidateList = Array.from(numericCandidates.values())
    .sort(compareNumericCandidates)
    .slice(0, MAX_NUMERIC_CANDIDATES_PER_STREAM);
  const geometryCandidates = inferGeometryCandidates(numericCandidateList);
  const animationHintList = Array.from(animationHints.values()).sort(compareAnimationHints).slice(0, MAX_ANIMATION_HINTS_PER_STREAM);
  return {
    textCandidates: Array.from(textCandidates.values()).sort(compareTextCandidates).slice(0, MAX_TEXT_CANDIDATES_PER_STREAM),
    referenceCandidates: Array.from(referenceCandidates.values())
      .sort(compareReferenceCandidates)
      .slice(0, MAX_REFERENCE_CANDIDATES_PER_STREAM),
    numericCandidates: numericCandidateList,
    geometryCandidates,
    groupingHints: inferGroupingHints(
      Array.from(textCandidates.values()),
      Array.from(referenceCandidates.values()),
      geometryCandidates,
      animationHintList
    ),
    animationHints: animationHintList,
    fieldSummaries: fieldSummaryList,
    archiveRecords: Array.from(archiveRecords.values())
      .sort(compareArchiveRecordEvidence)
      .slice(0, MAX_ARCHIVE_RECORD_EVIDENCE_PER_STREAM),
    archiveMessages: Array.from(archiveMessages.values())
      .sort(compareArchiveMessageEvidence)
      .slice(0, MAX_TYPED_ARCHIVE_MESSAGE_EVIDENCE_PER_STREAM),
    typedArchiveMessageCount,
    protobufFieldCount,
    nestedMessageCount,
    rawStringCount,
    expandedByteLength: payloads.reduce((max, payload) => Math.max(max, payload.data.byteLength), 0)
  };
}

function isNativeInternalTextMessageType(type: number | undefined): boolean {
  return type === 8 || type === 153 || type === 3097;
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
  for (const record of scan.archiveRecords) {
    for (const dataId of record.dataReferences) {
      for (const asset of assets.byDataId.get(normalizeDataId(dataId)) ?? []) {
        const existing = matches.get(asset.assetId);
        const match: NativeAssetMatch = {
          asset,
          evidence: `data-id:${dataId}`,
          confidence: 0.84,
          source: "archive"
        };
        if (!existing || match.confidence > existing.confidence) {
          matches.set(asset.assetId, match);
        }
      }
    }
  }
  for (const message of scan.archiveMessages) {
    for (const dataId of message.dataReferences) {
      for (const asset of assets.byDataId.get(normalizeDataId(dataId)) ?? []) {
        const key = archiveAssetMatchKey(asset.assetId, message);
        const existing = matches.get(key);
        const match: NativeAssetMatch = {
          asset,
          evidence: `typed-message:${message.type ?? "unknown"}:data-id:${dataId}`,
          confidence: typedArchiveAssetMatchConfidence(message),
          source: "archive",
          geometryCandidate: message.geometryCandidates[0],
          archiveMessage: message
        };
        const generic = matches.get(asset.assetId);
        if (generic && !generic.archiveMessage) {
          matches.delete(asset.assetId);
        }
        if (!existing || match.confidence > existing.confidence || (!existing.geometryCandidate && match.geometryCandidate)) {
          matches.set(key, match);
        }
      }
    }
  }
  for (const match of findIwaMovieStartMediaAssetMatchesFromScan(scan, assets)) {
    const key = match.archiveMessage ? archiveAssetMatchKey(match.asset.assetId, match.archiveMessage) : match.asset.assetId;
    const existing = matches.get(key);
    const generic = matches.get(match.asset.assetId);
    if (generic && !generic.archiveMessage) {
      matches.delete(match.asset.assetId);
    }
    if (!existing || match.confidence > existing.confidence || !existing.archiveMessage) {
      matches.set(key, match);
    }
  }

  return coalesceNativeAssetVariantMatches(Array.from(matches.values())).sort((left, right) => comparePartPaths(left.asset.path, right.asset.path));
}

function coalesceNativeAssetVariantMatches(matches: NativeAssetMatch[]): NativeAssetMatch[] {
  const grouped = new Map<string, NativeAssetMatch[]>();
  const passthrough: NativeAssetMatch[] = [];
  for (const match of matches) {
    const key = nativeAssetVariantGroupKey(match);
    if (!key) {
      passthrough.push(match);
      continue;
    }
    const list = grouped.get(key) ?? [];
    list.push(match);
    grouped.set(key, list);
  }

  const coalesced: NativeAssetMatch[] = [];
  for (const list of grouped.values()) {
    if (list.length <= 1) {
      coalesced.push(...list);
      continue;
    }
    const sorted = [...list].sort(compareNativeAssetVariantMatches);
    const best = sorted[0]!;
    coalesced.push({
      ...best,
      suppressedAssets: sorted.slice(1).map((match) => match.asset)
    });
  }
  return [...passthrough, ...coalesced];
}

function nativeAssetVariantGroupKey(match: NativeAssetMatch): string | undefined {
  const message = match.archiveMessage;
  if (!message?.archiveIdentifier || !isTypedVisualGeometryMessage(message.type)) {
    return undefined;
  }
  return [
    "archive",
    message.archiveIdentifier,
    message.type ?? "unknown",
    message.messageIndex,
    message.payloadOffset,
    match.geometryCandidate?.fieldPaths.join(",") ?? ""
  ].join(":");
}

function compareNativeAssetVariantMatches(left: NativeAssetMatch, right: NativeAssetMatch): number {
  const displayable = nativeAssetDisplayRank(right.asset) - nativeAssetDisplayRank(left.asset);
  if (displayable !== 0) return displayable;
  const pixels = nativeAssetPixelArea(right.asset) - nativeAssetPixelArea(left.asset);
  if (pixels !== 0) return pixels;
  const role = nativeAssetVariantPenalty(left.asset) - nativeAssetVariantPenalty(right.asset);
  if (role !== 0) return role;
  return comparePartPaths(left.asset.path, right.asset.path);
}

function nativeAssetDisplayRank(asset: NativeAssetReference): number {
  if (asset.kind !== "image") return 3;
  if (asset.embedded || asset.renderProxy) return 3;
  if (asset.mimeType === "image/tiff" || asset.mimeType === "image/heic") return 1;
  return 2;
}

function nativeAssetPixelArea(asset: NativeAssetReference): number {
  return Math.max(0, asset.width ?? 0) * Math.max(0, asset.height ?? 0);
}

function nativeAssetVariantPenalty(asset: NativeAssetReference): number {
  const name = asset.name.toLowerCase();
  let penalty = 0;
  if (/(?:^|[-_])small(?:[-_.]|$)/.test(name)) penalty += 20;
  if (/(?:^|[-_])filtered(?:[-_.]|$)/.test(name)) penalty += 8;
  if ((asset.sourceMimeType ?? asset.mimeType) === "image/tiff" || (asset.sourceMimeType ?? asset.mimeType) === "image/heic") {
    penalty += asset.renderProxy ? 0 : 4;
  }
  return penalty;
}

function archiveAssetMatchKey(assetId: string, message: NativeIwaArchiveMessageEvidence): string {
  return [
    assetId,
    "archive",
    message.archiveIdentifier ?? "unknown",
    message.type ?? "unknown",
    message.messageIndex,
    message.payloadOffset
  ].join(":");
}

function findIwaMovieStartMediaAssetMatchesFromScan(scan: IwaScanResult, assets: NativeAssetCatalog): NativeAssetMatch[] {
  const mediaAssets = Array.from(
    new Map(
      Array.from(assets.byPath.values())
        .filter((asset) => asset.kind === "video" || asset.kind === "audio")
        .map((asset) => [asset.assetId, asset])
    ).values()
  );
  if (mediaAssets.length !== 1) {
    return [];
  }
  const mediaAsset = mediaAssets[0];
  if (!mediaAsset) {
    return [];
  }
  const movieStartTargets = new Set(
    scan.archiveMessages
      .map((message) => message.build)
      .filter((build): build is NativeIwaBuildEvidence => Boolean(build?.targetNativeId && /movie-start/i.test(build.effect ?? "")))
      .map((build) => build.targetNativeId!)
  );
  if (movieStartTargets.size === 0) {
    return [];
  }
  return scan.archiveMessages
    .filter(
      (message) =>
        message.type === 3007 &&
        message.archiveIdentifier !== undefined &&
        movieStartTargets.has(message.archiveIdentifier)
    )
    .map((message) => ({
      asset: mediaAsset,
      evidence: `typed-message:3007:movie-start-target:${message.archiveIdentifier}`,
      confidence: 0.58,
      source: "archive" as const,
      geometryCandidate: message.geometryCandidates[0],
      archiveMessage: message
    }));
}

function typedArchiveAssetMatchConfidence(message: NativeIwaArchiveMessageEvidence): number {
  const typeScore = message.type === 3005 || message.type === 3006 ? 0.08 : 0;
  const geometryScore = message.geometryCandidates.length > 0 ? 0.06 : 0;
  return roundConfidence(Math.min(0.96, 0.84 + typeScore + geometryScore));
}

function parseIwaArchiveRecords(data: Uint8Array): IwaArchiveRecord[] {
  const records: IwaArchiveRecord[] = [];
  let offset = 0;

  while (offset < data.byteLength) {
    const archiveOffset = offset;
    const lengthInfo = readVarint(data, offset);
    if (!lengthInfo || lengthInfo.value <= 0 || lengthInfo.value > MAX_NESTED_PROTOBUF_BYTES) {
      return [];
    }
    const archiveInfoStart = lengthInfo.nextOffset;
    const archiveInfoEnd = archiveInfoStart + lengthInfo.value;
    if (archiveInfoEnd > data.byteLength) {
      return [];
    }

    const archiveInfo = parseArchiveInfo(data.subarray(archiveInfoStart, archiveInfoEnd));
    if (!archiveInfo || archiveInfo.messageInfos.length === 0) {
      return [];
    }

    let payloadOffset = archiveInfoEnd;
    for (const messageInfo of archiveInfo.messageInfos) {
      if (messageInfo.length < 0 || payloadOffset + messageInfo.length > data.byteLength) {
        return [];
      }
      messageInfo.payloadOffset = payloadOffset;
      messageInfo.payloadLength = messageInfo.length;
      payloadOffset += messageInfo.length;
    }

    records.push({
      archiveOffset,
      archiveInfoLength: lengthInfo.value,
      payloadOffset: archiveInfoEnd,
      totalPayloadLength: payloadOffset - archiveInfoEnd,
      messageInfos: archiveInfo.messageInfos,
      identifier: archiveInfo.identifier,
      shouldMerge: archiveInfo.shouldMerge
    });
    offset = payloadOffset;
  }

  return records.length > 0 && offset === data.byteLength ? records : [];
}

function parseArchiveInfo(data: Uint8Array): { messageInfos: IwaArchiveMessageInfo[]; identifier?: string; shouldMerge?: boolean } | undefined {
  const fields = readSequentialProtobufFields(data);
  if (!fields || fields.length === 0) {
    return undefined;
  }

  const messageInfos: IwaArchiveMessageInfo[] = [];
  let identifier: string | undefined;
  let shouldMerge: boolean | undefined;

  for (const field of fields) {
    if (field.fieldNumber === 1 && field.wireType === 0 && field.numericValue !== undefined) {
      identifier = String(field.numericValue);
      continue;
    }
    if (field.fieldNumber === 2 && field.wireType === 2 && field.value) {
      const messageInfo = parseArchiveMessageInfo(field.value);
      if (messageInfo) {
        messageInfos.push(messageInfo);
      }
      continue;
    }
    if (field.fieldNumber === 3 && field.wireType === 0 && field.numericValue !== undefined) {
      shouldMerge = field.numericValue !== 0;
    }
  }

  return messageInfos.length > 0 ? { messageInfos, identifier, shouldMerge } : undefined;
}

function parseArchiveMessageInfo(data: Uint8Array): IwaArchiveMessageInfo | undefined {
  const fields = readSequentialProtobufFields(data);
  if (!fields || fields.length === 0) {
    return undefined;
  }

  let length: number | undefined;
  let type: number | undefined;
  const version: number[] = [];
  const objectReferences: string[] = [];
  const dataReferences: string[] = [];

  for (const field of fields) {
    if (field.fieldNumber === 1 && field.wireType === 0 && field.numericValue !== undefined) {
      type = field.numericValue;
      continue;
    }
    if (field.fieldNumber === 2) {
      if (field.wireType === 0 && field.numericValue !== undefined) {
        version.push(field.numericValue);
      } else if (field.wireType === 2 && field.value) {
        version.push(...readPackedVarints(field.value));
      }
      continue;
    }
    if (field.fieldNumber === 3 && field.wireType === 0 && field.numericValue !== undefined) {
      length = field.numericValue;
      continue;
    }
    if (field.fieldNumber === 4) {
      continue;
    }
    if (field.fieldNumber === 5) {
      if (field.wireType === 0 && field.numericValue !== undefined) {
        objectReferences.push(String(field.numericValue));
      } else if (field.wireType === 2 && field.value) {
        objectReferences.push(...readReferenceIdentifiers(field.value));
      }
      continue;
    }
    if (field.fieldNumber === 6) {
      if (field.wireType === 0 && field.numericValue !== undefined) {
        dataReferences.push(normalizeDataId(String(field.numericValue)));
      } else if (field.wireType === 2 && field.value) {
        dataReferences.push(...readReferenceIdentifiers(field.value).map(normalizeDataId));
      }
    }
  }

  if (length === undefined || length < 0 || length > MAX_NESTED_PROTOBUF_BYTES) {
    return undefined;
  }
  return {
    length,
    payloadOffset: 0,
    payloadLength: 0,
    ...(type !== undefined ? { type } : {}),
    version,
    objectReferences: Array.from(new Set(objectReferences)),
    dataReferences: Array.from(new Set(dataReferences))
  };
}

function readSequentialProtobufFields(
  data: Uint8Array
): Array<{
  fieldNumber: number;
  wireType: number;
  nextOffset: number;
  value?: Uint8Array;
  numericValue?: number;
  rawValue?: Uint8Array;
}> | undefined {
  const fields = [];
  let offset = 0;
  while (offset < data.byteLength && fields.length < MAX_PROTOBUF_FIELDS_PER_PAYLOAD) {
    const field = readProtobufFieldAt(data, offset);
    if (!field || field.nextOffset <= offset) {
      return undefined;
    }
    fields.push(field);
    offset = field.nextOffset;
  }
  return offset === data.byteLength ? fields : undefined;
}

function readPackedVarints(data: Uint8Array): number[] {
  const values: number[] = [];
  let offset = 0;
  while (offset < data.byteLength) {
    const value = readVarint(data, offset);
    if (!value || value.nextOffset <= offset) {
      return [];
    }
    values.push(value.value);
    offset = value.nextOffset;
  }
  return values;
}

function readReferenceIdentifiers(data: Uint8Array): string[] {
  const fields = readSequentialProtobufFields(data);
  if (!fields) {
    return [];
  }
  const ids: string[] = [];
  for (const field of fields) {
    if (field.fieldNumber === 1 && field.wireType === 0 && field.numericValue !== undefined) {
      ids.push(String(field.numericValue));
    }
    if (field.fieldNumber === 1 && field.wireType === 2 && field.value) {
      const text = decodeUtf8(field.value)?.replace(/\u0000/g, "").trim();
      if (text && !text.includes("\ufffd")) {
        ids.push(text);
      }
    }
  }
  return ids;
}

function numericValueAtFieldPath(data: Uint8Array, fieldPath: number[]): number | undefined {
  return numericValuesAtFieldPath(data, fieldPath)[0];
}

function numericValuesAtFieldPath(data: Uint8Array, fieldPath: number[]): number[] {
  if (fieldPath.length === 0) {
    return [];
  }
  const fields = readSequentialProtobufFields(data);
  if (!fields) {
    return [];
  }
  const [fieldNumber, ...rest] = fieldPath;
  const matches = fields.filter((field) => field.fieldNumber === fieldNumber);
  const values: number[] = [];
  for (const field of matches) {
    if (rest.length === 0) {
      const value = numericValueFromSequentialField(field);
      if (value !== undefined) {
        values.push(value);
      }
      continue;
    }
    if (field.value) {
      values.push(...numericValuesAtFieldPath(field.value, rest));
    }
  }
  return values;
}

function numericValueFromSequentialField(field: {
  wireType: number;
  numericValue?: number;
  rawValue?: Uint8Array;
}): number | undefined {
  if (field.wireType === 0 && field.numericValue !== undefined) {
    return field.numericValue;
  }
  if (field.wireType === 5 && field.rawValue?.byteLength === 4) {
    const view = new DataView(field.rawValue.buffer, field.rawValue.byteOffset, field.rawValue.byteLength);
    const floatValue = view.getFloat32(0, true);
    if (Number.isFinite(floatValue) && Math.abs(floatValue) <= MAX_VISUAL_NUMERIC_VALUE) {
      return roundGeometryNumber(floatValue);
    }
    const intValue = view.getUint32(0, true);
    return intValue <= MAX_VISUAL_NUMERIC_VALUE ? intValue : undefined;
  }
  if (field.wireType === 1 && field.rawValue?.byteLength === 8) {
    const view = new DataView(field.rawValue.buffer, field.rawValue.byteOffset, field.rawValue.byteLength);
    const doubleValue = view.getFloat64(0, true);
    return Number.isFinite(doubleValue) && Math.abs(doubleValue) <= MAX_VISUAL_NUMERIC_VALUE ? roundGeometryNumber(doubleValue) : undefined;
  }
  return undefined;
}

function stringValueAtFieldPath(data: Uint8Array, fieldPath: number[]): string | undefined {
  return stringValuesAtFieldPath(data, fieldPath)[0];
}

function stringValuesAtFieldPath(data: Uint8Array, fieldPath: number[]): string[] {
  if (fieldPath.length === 0) {
    return [];
  }
  const fields = readSequentialProtobufFields(data);
  if (!fields) {
    return [];
  }
  const [fieldNumber, ...rest] = fieldPath;
  const matches = fields.filter((field) => field.fieldNumber === fieldNumber);
  const values: string[] = [];
  for (const field of matches) {
    if (!field.value) {
      continue;
    }
    if (rest.length === 0) {
      const text = cleanNativeBuildString(field.value);
      if (text) {
        values.push(text);
      }
      continue;
    }
    values.push(...stringValuesAtFieldPath(field.value, rest));
  }
  return uniqueStrings(values);
}

function bytesValueAtFieldPath(data: Uint8Array, fieldPath: number[]): Uint8Array | undefined {
  return bytesValuesAtFieldPath(data, fieldPath)[0];
}

function bytesValuesAtFieldPath(data: Uint8Array, fieldPath: number[]): Uint8Array[] {
  if (fieldPath.length === 0) {
    return [];
  }
  const fields = readSequentialProtobufFields(data);
  if (!fields) {
    return [];
  }
  const [fieldNumber, ...rest] = fieldPath;
  const values: Uint8Array[] = [];
  for (const field of fields.filter((candidate) => candidate.fieldNumber === fieldNumber)) {
    if (!field.value) {
      continue;
    }
    if (rest.length === 0) {
      values.push(field.value);
      continue;
    }
    values.push(...bytesValuesAtFieldPath(field.value, rest));
  }
  return values;
}

function cleanNativeBuildString(value: Uint8Array): string | undefined {
  const text = decodeUtf8(value)?.replace(/\u0000/g, "").replace(/\s+/g, " ").trim();
  if (!text || text.includes("\ufffd") || text.length > 256) {
    return undefined;
  }
  return text;
}

function nativeBuildEvidenceFromPayload(
  record: IwaArchiveRecord,
  messageInfo: IwaArchiveMessageInfo,
  payload: Uint8Array
): NativeIwaBuildEvidence | undefined {
  if (messageInfo.type !== 8) {
    return undefined;
  }
  const targetNativeId = nativeIdFromNumber(numericValueAtFieldPath(payload, [1, 1]));
  const delivery = stringValueAtFieldPath(payload, [2]);
  const direction = stringValueAtFieldPath(payload, [4, 18, 1]);
  const effect = stringValueAtFieldPath(payload, [4, 18, 2]) ?? nativeBuildEffectFromStrings(extractArchiveEvidenceStrings(payload));
  const durationMs = nativeSecondsToMs(numericValueAtFieldPath(payload, [4, 18, 3]));
  const delayMs = nativeSecondsToMs(numericValueAtFieldPath(payload, [4, 18, 5]) ?? numericValueAtFieldPath(payload, [3]));
  const timingBase = numericValueAtFieldPath(payload, [4, 17]);
  const motionPath = nativeMotionPathEvidenceFromPayload(payload);
  const sourceFieldPaths = [
    ...(targetNativeId ? ["1.1"] : []),
    ...(delivery ? ["2"] : []),
    ...(delayMs !== undefined ? ["3", "4.18.5"] : []),
    ...(timingBase !== undefined ? ["4.17"] : []),
    ...(direction ? ["4.18.1"] : []),
    ...(effect ? ["4.18.2"] : []),
    ...(durationMs !== undefined ? ["4.18.3"] : []),
    ...(motionPath ? motionPath.sourceFieldPaths : [])
  ];
  const confidence = roundConfidence(
    0.28 +
      (targetNativeId ? 0.2 : 0) +
      (effect ? 0.2 : 0) +
      (direction ? 0.12 : 0) +
      (durationMs !== undefined ? 0.1 : 0) +
      (motionPath ? 0.08 : 0) +
      (record.identifier ? 0.06 : 0) +
      (messageInfo.objectReferences.length > 0 ? 0.04 : 0)
  );
  return {
    kind: "build",
    ...(record.identifier ? { buildId: record.identifier } : {}),
    ...(targetNativeId ? { targetNativeId } : {}),
    objectReferences: messageInfo.objectReferences.slice(0, 24),
    ...(delivery ? { delivery } : {}),
    ...(direction ? { direction } : {}),
    ...(effect ? { effect } : {}),
    ...(durationMs !== undefined ? { durationMs } : {}),
    ...(delayMs !== undefined ? { delayMs } : {}),
    ...(timingBase !== undefined ? { timingBase } : {}),
    ...(motionPath ? { motionPath } : {}),
    confidence,
    sourceFieldPaths: uniqueStrings(sourceFieldPaths)
  };
}

function nativeBuildTimingEvidenceFromPayload(
  record: IwaArchiveRecord,
  messageInfo: IwaArchiveMessageInfo,
  payload: Uint8Array
): NativeIwaBuildTimingEvidence | undefined {
  if (messageInfo.type !== 153) {
    return undefined;
  }
  const buildId = nativeIdFromNumber(numericValueAtFieldPath(payload, [1, 1]));
  const delayMs = nativeSecondsToMs(numericValueAtFieldPath(payload, [3]));
  const durationMs = nativeSecondsToMs(numericValueAtFieldPath(payload, [4]));
  const startsWithPreviousRaw = numericValueAtFieldPath(payload, [5]);
  const afterPreviousRaw = numericValueAtFieldPath(payload, [6]);
  const triggerGroupRaw = numericValueAtFieldPath(payload, [7, 2]);
  const startsWithPrevious = startsWithPreviousRaw !== undefined ? startsWithPreviousRaw !== 0 : undefined;
  const afterPreviousRawFlag = afterPreviousRaw !== undefined ? afterPreviousRaw !== 0 : undefined;
  const startRelation =
    startsWithPrevious === true
      ? "withPrevious"
      : afterPreviousRawFlag === true
        ? "afterPrevious"
        : startsWithPrevious !== undefined || afterPreviousRawFlag !== undefined
          ? "unknown"
          : undefined;
  const afterPrevious = startRelation === "afterPrevious" ? true : startRelation === "withPrevious" ? false : afterPreviousRawFlag;
  const sourceFieldPaths = [
    ...(buildId ? ["1.1"] : []),
    ...(delayMs !== undefined ? ["3"] : []),
    ...(durationMs !== undefined ? ["4"] : []),
    ...(startsWithPreviousRaw !== undefined ? ["5"] : []),
    ...(afterPreviousRaw !== undefined ? ["6"] : []),
    ...(triggerGroupRaw !== undefined ? ["7.2"] : [])
  ];
  const confidence = roundConfidence(
    0.32 +
      (buildId ? 0.28 : 0) +
      (durationMs !== undefined ? 0.16 : 0) +
      (record.identifier ? 0.08 : 0) +
      (triggerGroupRaw !== undefined ? 0.04 : 0)
  );
  return {
    kind: "buildTiming",
    ...(record.identifier ? { timingId: record.identifier } : {}),
    ...(buildId ? { buildId } : {}),
    ...(durationMs !== undefined ? { durationMs } : {}),
    ...(delayMs !== undefined ? { delayMs } : {}),
    ...(triggerGroupRaw !== undefined ? { triggerGroupRaw, group: triggerGroupRaw, rawField72: triggerGroupRaw } : {}),
    ...(startsWithPreviousRaw !== undefined ? { rawField5: startsWithPreviousRaw, startsWithPrevious } : {}),
    ...(afterPreviousRaw !== undefined ? { rawField6: afterPreviousRaw, afterPrevious } : {}),
    ...(startRelation ? { startRelation } : {}),
    confidence,
    sourceFieldPaths: uniqueStrings(sourceFieldPaths)
  };
}

function nativeMotionPathEvidenceFromPayload(payload: Uint8Array): NativeIwaMotionPathEvidence | undefined {
  const pathPayload = bytesValueAtFieldPath(payload, [4, 22]);
  if (!pathPayload) {
    return undefined;
  }
  const pointClusters = nativeMotionPathPointClusters(pathPayload);
  if (pointClusters.length < 2) {
    return undefined;
  }
  const first = pointClusters[0];
  const last = pointClusters[pointClusters.length - 1];
  if (!first || !last || !isZeroishPoint(first) || !isFiniteMotionPoint(last)) {
    return undefined;
  }
  const points = [
    { x: 0, y: 0 },
    { x: roundGeometryNumber(last.x), y: roundGeometryNumber(last.y) }
  ];
  if (Math.abs(points[1]!.x) < 0.001 && Math.abs(points[1]!.y) < 0.001) {
    return undefined;
  }
  const extentPoints = nativeMotionPathExtentPoints(pathPayload);
  return {
    kind: "motionPath",
    relative: true,
    points,
    ...(extentPoints.length > 0 ? { extentPoints } : {}),
    confidence: pointClusters.length === 2 ? 0.86 : 0.78,
    sourceFieldPaths: [
      "4.22",
      "4.22.8.1.1.1.1",
      "4.22.8.1.1.1.2",
      "4.22.8.1.1.2.1",
      "4.22.8.1.1.2.2",
      "4.22.8.1.1.3.1",
      "4.22.8.1.1.3.2",
      ...(extentPoints.length > 0 ? ["4.22.8.2.1", "4.22.8.2.2"] : [])
    ]
  };
}

function nativeMotionPathExtentPoints(payload: Uint8Array): Array<{ x: number; y: number }> {
  const extentPayloads = bytesValuesAtFieldPath(payload, [8, 2]);
  return extentPayloads
    .map((extentPayload) => {
      const x = numericValueAtFieldPath(extentPayload, [1]);
      const y = numericValueAtFieldPath(extentPayload, [2]);
      if (x === undefined || y === undefined || !Number.isFinite(x) || !Number.isFinite(y)) {
        return undefined;
      }
      return { x: roundGeometryNumber(x), y: roundGeometryNumber(y) };
    })
    .filter((point): point is { x: number; y: number } => Boolean(point));
}

function nativeMotionPathPointClusters(payload: Uint8Array): Array<{ x: number; y: number }> {
  const containers = bytesValuesAtFieldPath(payload, [8, 1, 1]);
  const clusters: Array<{ x: number; y: number }> = [];
  for (const container of containers) {
    const controls = [
      nativeMotionPathPointAt(container, [1]),
      nativeMotionPathPointAt(container, [2]),
      nativeMotionPathPointAt(container, [3])
    ].filter((point): point is { x: number; y: number } => Boolean(point));
    if (controls.length === 0) {
      continue;
    }
    const stable = controls.every((point) => pointsNearlyEqual(point, controls[0]!));
    if (!stable) {
      continue;
    }
    clusters.push(controls[0]!);
  }
  return clusters;
}

function nativeMotionPathPointAt(payload: Uint8Array, fieldPath: number[]): { x: number; y: number } | undefined {
  const pointPayload = bytesValueAtFieldPath(payload, fieldPath);
  if (!pointPayload) {
    return undefined;
  }
  const x = numericValueAtFieldPath(pointPayload, [1]);
  const y = numericValueAtFieldPath(pointPayload, [2]);
  if (x === undefined || y === undefined || !Number.isFinite(x) || !Number.isFinite(y)) {
    return undefined;
  }
  return { x, y };
}

function isZeroishPoint(point: { x: number; y: number }): boolean {
  return Math.abs(point.x) <= 0.001 && Math.abs(point.y) <= 0.001;
}

function isFiniteMotionPoint(point: { x: number; y: number }): boolean {
  return Number.isFinite(point.x) && Number.isFinite(point.y) && Math.abs(point.x) <= MAX_VISUAL_NUMERIC_VALUE && Math.abs(point.y) <= MAX_VISUAL_NUMERIC_VALUE;
}

function pointsNearlyEqual(left: { x: number; y: number }, right: { x: number; y: number }): boolean {
  return Math.abs(left.x - right.x) <= 0.001 && Math.abs(left.y - right.y) <= 0.001;
}

function nativeBuildEffectFromStrings(values: string[]): string | undefined {
  return values.find((value) => /^(?:apple:|com\.apple\.iWork\.Keynote\.)/.test(value));
}

function nativeSecondsToMs(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value) || value < 0 || value > 3600) {
    return undefined;
  }
  return Math.round(value * 1000);
}

function nativeIdFromNumber(value: number | undefined): string | undefined {
  if (value === undefined || !Number.isInteger(value) || value <= 0) {
    return undefined;
  }
  return String(value);
}

function archiveMessagePayloadsFromRecords(
  data: Uint8Array,
  records: IwaArchiveRecord[]
): Array<{
  data: Uint8Array;
  archiveMessage: {
    record: IwaArchiveRecord;
    messageInfo: IwaArchiveMessageInfo;
    messageIndex: number;
  };
}> {
  const payloads: Array<{
    data: Uint8Array;
    archiveMessage: {
      record: IwaArchiveRecord;
      messageInfo: IwaArchiveMessageInfo;
      messageIndex: number;
    };
  }> = [];
  for (const record of records) {
    for (const [messageIndex, messageInfo] of record.messageInfos.entries()) {
      payloads.push({
        data: data.subarray(messageInfo.payloadOffset, messageInfo.payloadOffset + messageInfo.payloadLength),
        archiveMessage: { record, messageInfo, messageIndex }
      });
    }
  }
  return payloads;
}

function createArchiveMessageEvidence(
  record: IwaArchiveRecord,
  messageInfo: IwaArchiveMessageInfo,
  messageIndex: number,
  scan: Pick<IwaScanResult, "textCandidates" | "numericCandidates" | "fieldSummaries">,
  payload: Uint8Array
): NativeIwaArchiveMessageEvidence {
  const geometryCandidates = inferArchiveMessageGeometryCandidates(messageInfo, scan.numericCandidates, payload);
  const dataReferences = uniqueStrings([
    ...messageInfo.dataReferences.map(normalizeDataId),
    ...typedDataReferencesFromPayload(messageInfo, payload),
    ...typedDataReferencesFromNumericFields(messageInfo, scan.numericCandidates)
  ]).slice(0, 24);
  const textCandidates = uniqueStrings([
    ...scan.textCandidates.map((candidate) => candidate.text),
    ...extractArchiveEvidenceStrings(payload)
  ]).slice(0, 24);
  const build = nativeBuildEvidenceFromPayload(record, messageInfo, payload);
  const buildTiming = nativeBuildTimingEvidenceFromPayload(record, messageInfo, payload);
  const textContent = nativeTextContentEvidenceFromPayload(messageInfo, payload);
  const textDrawable = nativeTextDrawableEvidenceFromPayload(messageInfo, payload);
  return {
    archiveOffset: record.archiveOffset,
    ...(record.identifier ? { archiveIdentifier: record.identifier } : {}),
    messageIndex,
    ...(messageInfo.type !== undefined ? { type: messageInfo.type } : {}),
    payloadOffset: messageInfo.payloadOffset,
    payloadLength: messageInfo.payloadLength,
    version: messageInfo.version.slice(0, 8),
    objectReferences: messageInfo.objectReferences.slice(0, 24),
    dataReferences,
    textCandidates,
    geometryCandidates,
    fieldSummaries: scan.fieldSummaries.slice(0, 24),
    ...(build ? { build } : {}),
    ...(buildTiming ? { buildTiming } : {}),
    ...(textContent ? { textContent } : {}),
    ...(textDrawable ? { textDrawable } : {})
  };
}

function typedDataReferencesFromNumericFields(messageInfo: IwaArchiveMessageInfo, numericCandidates: IwaNumericCandidate[]): string[] {
  if (messageInfo.type !== 3005 && messageInfo.type !== 3006) {
    return [];
  }
  const dataReferenceFieldPaths = new Set(["11.1", "12.1", "13.1", "15.1"]);
  return uniqueStrings(
    numericCandidates
      .filter(
        (candidate) =>
          dataReferenceFieldPaths.has(candidate.fieldPath) &&
          candidate.encoding === "varint" &&
          Number.isInteger(candidate.value) &&
          candidate.value > 0
      )
      .map((candidate) => normalizeDataId(String(candidate.value)))
  );
}

function typedDataReferencesFromPayload(messageInfo: IwaArchiveMessageInfo, payload: Uint8Array): string[] {
  if (messageInfo.type !== 3005 && messageInfo.type !== 3006) {
    return [];
  }
  const dataReferenceFieldPaths = [
    [11, 1],
    [12, 1],
    [13, 1],
    [15, 1]
  ];
  return uniqueStrings(
    dataReferenceFieldPaths
      .map((fieldPath) => numericValueAtFieldPath(payload, fieldPath))
      .filter((value): value is number => value !== undefined && Number.isInteger(value) && value > 0)
      .map((value) => normalizeDataId(String(value)))
  );
}

function inferArchiveMessageGeometryCandidates(
  messageInfo: IwaArchiveMessageInfo,
  numericCandidates: IwaNumericCandidate[],
  payload: Uint8Array
): NativeIwaGeometryCandidate[] {
  const candidates = [
    ...inferTypedImageGeometryCandidatesFromPayload(messageInfo, payload),
    ...inferTypedImageGeometryCandidates(messageInfo, numericCandidates),
    ...inferGeometryCandidates(numericCandidates)
  ];
  return dedupeGeometryCandidates(candidates)
    .sort((left, right) => right.confidence - left.confidence || compareFieldPaths(left.fieldPaths[0] ?? "", right.fieldPaths[0] ?? ""))
    .slice(0, MAX_GEOMETRY_CANDIDATES_PER_STREAM);
}

function inferTypedImageGeometryCandidatesFromPayload(
  messageInfo: IwaArchiveMessageInfo,
  payload: Uint8Array
): NativeIwaGeometryCandidate[] {
  if (!isTypedVisualGeometryMessage(messageInfo.type)) {
    return [];
  }

  const xValues = numericValuesAtFieldPath(payload, [1, 1, 1, 1]);
  const yValues = numericValuesAtFieldPath(payload, [1, 1, 1, 2]);
  const widthValues = numericValuesAtFieldPath(payload, [1, 1, 2, 1]);
  const heightValues = numericValuesAtFieldPath(payload, [1, 1, 2, 2]);
  const tupleCount = Math.min(xValues.length, yValues.length, widthValues.length, heightValues.length);
  const tuples = Array.from({ length: tupleCount }, (_, index) => ({
    x: xValues[index]!,
    y: yValues[index]!,
    width: widthValues[index]!,
    height: heightValues[index]!
  }))
    .filter((tuple) => isPlausibleTypedImageGeometryTuple(tuple.x, tuple.y, tuple.width, tuple.height))
    .sort((left, right) => right.width * right.height - left.width * left.height);
  const best = tuples[0];
  if (!best) {
    return [];
  }

  return [
    {
      bounds: {
        x: roundGeometryNumber(best.x),
        y: roundGeometryNumber(best.y),
        width: roundGeometryNumber(best.width),
        height: roundGeometryNumber(best.height)
      },
      fieldPaths: ["1.1.1.1", "1.1.1.2", "1.1.2.1", "1.1.2.2"],
      values: [best.x, best.y, best.width, best.height].map(roundGeometryNumber),
      source: "protobuf",
      confidence: 0.97,
      groupPath: "1.1",
      reason: `typed Keynote ${messageInfo.type === 3007 ? "media" : "image"} geometry from archive message type ${messageInfo.type}`
    }
  ];
}

function inferTypedImageGeometryCandidates(
  messageInfo: IwaArchiveMessageInfo,
  numericCandidates: IwaNumericCandidate[]
): NativeIwaGeometryCandidate[] {
  if (!isTypedVisualGeometryMessage(messageInfo.type)) {
    return [];
  }

  const byPath = new Map(numericCandidates.map((candidate) => [candidate.fieldPath, candidate]));
  const x = byPath.get("1.1.1.1");
  const y = byPath.get("1.1.1.2");
  const width = byPath.get("1.1.2.1");
  const height = byPath.get("1.1.2.2");
  if (!x || !y || !width || !height || !isPlausibleGeometryTuple(x.value, y.value, width.value, height.value)) {
    return [];
  }

  const tuple = [x, y, width, height];
  return [
    {
      bounds: {
        x: roundGeometryNumber(x.value),
        y: roundGeometryNumber(y.value),
        width: roundGeometryNumber(width.value),
        height: roundGeometryNumber(height.value)
      },
      fieldPaths: tuple.map((candidate) => candidate.fieldPath),
      values: tuple.map((candidate) => roundGeometryNumber(candidate.value)),
      source: "protobuf",
      confidence: 0.94,
      groupPath: "1.1",
      reason: `typed Keynote ${messageInfo.type === 3007 ? "media" : "image"} geometry from archive message type ${messageInfo.type}`
    }
  ];
}

function isTypedVisualGeometryMessage(type: number | undefined): boolean {
  return type === 3005 || type === 3006 || type === 3007;
}

function nativeTextContentEvidenceFromPayload(
  messageInfo: IwaArchiveMessageInfo,
  payload: Uint8Array
): NativeIwaTextContentEvidence | undefined {
  if (messageInfo.type !== 2001) {
    return undefined;
  }
  const cleaned = [
    ...stringValuesAtFieldPath(payload, [3]),
    ...extractArchiveEvidenceStrings(payload)
  ]
    .map((text) => cleanTextCandidate(text))
    .find((text): text is string => Boolean(text));
  if (!cleaned) {
    return undefined;
  }
  return {
    kind: "textContent",
    text: cleaned,
    confidence: 0.9,
    sourceFieldPaths: ["3"]
  };
}

function nativeTextDrawableEvidenceFromPayload(
  messageInfo: IwaArchiveMessageInfo,
  payload: Uint8Array
): NativeIwaTextDrawableEvidence | undefined {
  if (messageInfo.type !== 2011) {
    return undefined;
  }

  const textArchiveIds = uniqueStrings(
    [
      ...numericValuesAtFieldPath(payload, [2, 1]),
      ...numericValuesAtFieldPath(payload, [4, 1])
    ]
      .map(nativeIdFromNumber)
      .filter((value): value is string => Boolean(value))
  );
  const slideArchiveId = nativeIdFromNumber(numericValueAtFieldPath(payload, [1, 1, 2, 1]));
  const x = numericValueAtFieldPath(payload, [1, 1, 1, 1, 1]);
  const y = numericValueAtFieldPath(payload, [1, 1, 1, 1, 2]);
  const width = numericValueAtFieldPath(payload, [1, 3, 5, 2, 1]);
  const height = numericValueAtFieldPath(payload, [1, 3, 5, 2, 2]);
  const bounds =
    x !== undefined &&
    y !== undefined &&
    width !== undefined &&
    height !== undefined &&
    isPlausibleTypedTextDrawableGeometryTuple(x, y, width, height)
      ? {
          x: roundGeometryNumber(x),
          y: roundGeometryNumber(y),
          width: roundGeometryNumber(width),
          height: roundGeometryNumber(height)
        }
      : undefined;
  if (textArchiveIds.length === 0 && !bounds) {
    return undefined;
  }
  const sourceFieldPaths = [
    ...(slideArchiveId ? ["1.1.2.1"] : []),
    ...(x !== undefined ? ["1.1.1.1.1"] : []),
    ...(y !== undefined ? ["1.1.1.1.2"] : []),
    ...(width !== undefined ? ["1.3.5.2.1"] : []),
    ...(height !== undefined ? ["1.3.5.2.2"] : []),
    ...(textArchiveIds.length > 0 ? ["2.1", "4.1"] : [])
  ];
  return {
    kind: "textDrawable",
    textArchiveIds,
    ...(slideArchiveId ? { slideArchiveId } : {}),
    ...(bounds ? { bounds } : {}),
    confidence: roundConfidence(0.36 + (textArchiveIds.length > 0 ? 0.3 : 0) + (bounds ? 0.22 : 0) + (slideArchiveId ? 0.04 : 0)),
    sourceFieldPaths: uniqueStrings(sourceFieldPaths)
  };
}

function isPlausibleTypedTextDrawableGeometryTuple(x: number, y: number, width: number, height: number): boolean {
  return (
    Number.isFinite(x) &&
    Number.isFinite(y) &&
    Number.isFinite(width) &&
    Number.isFinite(height) &&
    Math.abs(x) <= MAX_VISUAL_NUMERIC_VALUE &&
    Math.abs(y) <= MAX_VISUAL_NUMERIC_VALUE &&
    width >= 1 &&
    height >= 1 &&
    width <= MAX_VISUAL_NUMERIC_VALUE &&
    height <= MAX_VISUAL_NUMERIC_VALUE
  );
}

function extractArchiveEvidenceStrings(data: Uint8Array): string[] {
  const strings = new Set<string>();
  let start = -1;
  for (let index = 0; index <= data.length; index += 1) {
    const byte = data[index];
    const printable = byte !== undefined && (byte === 0x09 || byte === 0x0a || byte === 0x0d || byte >= 0x20);
    if (printable) {
      if (start < 0) start = index;
      continue;
    }
    if (start >= 0 && index - start >= 3) {
      const decoded = decodeUtf8(data.subarray(start, index))?.replace(/\u0000/g, "").replace(/\s+/g, " ").trim();
      if (decoded && !decoded.includes("\ufffd") && decoded.length >= 3 && decoded.length <= 200) {
        strings.add(decoded);
      }
    }
    start = -1;
  }
  return Array.from(strings).slice(0, 24);
}

function archiveRecordEvidence(record: IwaArchiveRecord): NativeIwaArchiveRecordEvidence {
  const dataReferences = Array.from(new Set(record.messageInfos.flatMap((messageInfo) => messageInfo.dataReferences).map(normalizeDataId)));
  const objectReferences = Array.from(new Set(record.messageInfos.flatMap((messageInfo) => messageInfo.objectReferences)));
  return {
    archiveOffset: record.archiveOffset,
    archiveInfoLength: record.archiveInfoLength,
    messageCount: record.messageInfos.length,
    payloadLength: record.totalPayloadLength,
    messageTypes: Array.from(
      new Set(record.messageInfos.map((messageInfo) => messageInfo.type).filter((type): type is number => type !== undefined))
    ).slice(0, 24),
    dataReferences: dataReferences.slice(0, 24),
    objectReferences: objectReferences.slice(0, 24),
    ...(record.identifier ? { identifier: record.identifier } : {}),
    ...(record.shouldMerge !== undefined ? { shouldMerge: record.shouldMerge } : {})
  };
}

function addBestArchiveRecordEvidence(
  records: Map<string, NativeIwaArchiveRecordEvidence>,
  record: NativeIwaArchiveRecordEvidence
): void {
  const key = `${record.archiveOffset}:${record.archiveInfoLength}:${record.payloadLength}`;
  const existing = records.get(key);
  if (!existing || record.messageCount > existing.messageCount) {
    records.set(key, record);
  }
}

function addBestArchiveMessageEvidence(
  messages: Map<string, NativeIwaArchiveMessageEvidence>,
  message: NativeIwaArchiveMessageEvidence
): void {
  const key = `${message.archiveOffset}:${message.messageIndex}:${message.type ?? "unknown"}:${message.payloadOffset}`;
  const existing = messages.get(key);
  if (!existing || messageEvidenceScore(message) > messageEvidenceScore(existing)) {
    messages.set(key, message);
  }
}

function compareArchiveRecordEvidence(left: NativeIwaArchiveRecordEvidence, right: NativeIwaArchiveRecordEvidence): number {
  if (left.archiveOffset !== right.archiveOffset) {
    return left.archiveOffset - right.archiveOffset;
  }
  return right.messageCount - left.messageCount;
}

function compareArchiveMessageEvidence(left: NativeIwaArchiveMessageEvidence, right: NativeIwaArchiveMessageEvidence): number {
  if (left.archiveOffset !== right.archiveOffset) {
    return left.archiveOffset - right.archiveOffset;
  }
  if (left.messageIndex !== right.messageIndex) {
    return left.messageIndex - right.messageIndex;
  }
  return (left.type ?? 0) - (right.type ?? 0);
}

function messageEvidenceScore(message: NativeIwaArchiveMessageEvidence): number {
  return message.dataReferences.length * 4 + message.geometryCandidates.length * 3 + message.textCandidates.length * 2 + message.fieldSummaries.length;
}

function scanProtobufPayload(
  data: Uint8Array,
  orderBase: number
): Pick<
  IwaScanResult,
  | "textCandidates"
  | "referenceCandidates"
  | "numericCandidates"
  | "animationHints"
  | "fieldSummaries"
  | "protobufFieldCount"
  | "nestedMessageCount"
> {
  const state = {
    textCandidates: new Map<string, IwaTextCandidate>(),
    referenceCandidates: new Map<string, IwaReferenceCandidate>(),
    numericCandidates: new Map<string, IwaNumericCandidate>(),
    animationHints: new Map<string, NativeIwaAnimationHint>(),
    fieldSummaries: new Map<string, NativeIwaFieldSummary>(),
    protobufFieldCount: 0,
    nestedMessageCount: 0,
    orderBase
  };
  scanProtobufFields(data, "", 0, state);
  return {
    textCandidates: Array.from(state.textCandidates.values()),
    referenceCandidates: Array.from(state.referenceCandidates.values()),
    numericCandidates: Array.from(state.numericCandidates.values()),
    animationHints: Array.from(state.animationHints.values()),
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
    numericCandidates: Map<string, IwaNumericCandidate>;
    animationHints: Map<string, NativeIwaAnimationHint>;
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

    const numeric = decodeNumericCandidate(field, fieldPath, state.orderBase + offset);
    if (numeric) {
      summary.numericCandidateCount += 1;
      summary.minNumericValue =
        summary.minNumericValue === undefined ? numeric.value : Math.min(summary.minNumericValue, numeric.value);
      summary.maxNumericValue =
        summary.maxNumericValue === undefined ? numeric.value : Math.max(summary.maxNumericValue, numeric.value);
      summary.sampleNumericValue ??= numeric.value;
      summary.sampleNumericEncoding ??= numeric.encoding;
      addBestNumericCandidate(state.numericCandidates, numeric);
    }

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
      for (const hint of detectAnimationHints(text, "protobuf", depth > 0 ? 0.6 : 0.54, fieldPath)) {
        addBestAnimationHint(state.animationHints, hint);
      }
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
): { fieldNumber: number; wireType: number; nextOffset: number; value?: Uint8Array; numericValue?: number; rawValue?: Uint8Array } | undefined {
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
    return {
      fieldNumber,
      wireType,
      nextOffset: value.nextOffset,
      numericValue: value.value,
      rawValue: data.subarray(key.nextOffset, value.nextOffset)
    };
  }

  if (wireType === 1) {
    const nextOffset = key.nextOffset + 8;
    return nextOffset <= data.length
      ? { fieldNumber, wireType, nextOffset, rawValue: data.subarray(key.nextOffset, nextOffset) }
      : undefined;
  }

  if (wireType === 5) {
    const nextOffset = key.nextOffset + 4;
    return nextOffset <= data.length
      ? { fieldNumber, wireType, nextOffset, rawValue: data.subarray(key.nextOffset, nextOffset) }
      : undefined;
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

function decodeNumericCandidate(
  field: {
    fieldNumber: number;
    wireType: number;
    nextOffset: number;
    numericValue?: number;
    rawValue?: Uint8Array;
  },
  fieldPath: string,
  order: number
): IwaNumericCandidate | undefined {
  if (field.wireType === 0 && field.numericValue !== undefined) {
    const value = field.numericValue;
    if (!Number.isFinite(value) || value < 0 || value > MAX_VISUAL_NUMERIC_VALUE) {
      return undefined;
    }
    return {
      fieldPath,
      fieldNumber: field.fieldNumber,
      wireType: field.wireType,
      value,
      encoding: "varint",
      source: "protobuf",
      confidence: value <= 10000 ? 0.52 : 0.36,
      order,
      endOrder: order + (field.rawValue?.byteLength ?? 1)
    };
  }

  if (field.wireType === 5 && field.rawValue?.byteLength === 4) {
    const view = new DataView(field.rawValue.buffer, field.rawValue.byteOffset, field.rawValue.byteLength);
    const floatValue = view.getFloat32(0, true);
    if (Number.isFinite(floatValue) && Math.abs(floatValue) <= MAX_VISUAL_NUMERIC_VALUE && Math.abs(floatValue) >= 0.0001) {
      return {
        fieldPath,
        fieldNumber: field.fieldNumber,
        wireType: field.wireType,
        value: roundGeometryNumber(floatValue),
        encoding: "fixed32-float",
        source: "protobuf",
        confidence: 0.7,
        order,
        endOrder: order + 4
      };
    }

    const intValue = view.getUint32(0, true);
    if (intValue <= MAX_VISUAL_NUMERIC_VALUE) {
      return {
        fieldPath,
        fieldNumber: field.fieldNumber,
        wireType: field.wireType,
        value: intValue,
        encoding: "fixed32-uint",
        source: "protobuf",
        confidence: intValue <= 10000 ? 0.5 : 0.34,
        order,
        endOrder: order + 4
      };
    }
  }

  if (field.wireType === 1 && field.rawValue?.byteLength === 8) {
    const view = new DataView(field.rawValue.buffer, field.rawValue.byteOffset, field.rawValue.byteLength);
    const doubleValue = view.getFloat64(0, true);
    if (Number.isFinite(doubleValue) && Math.abs(doubleValue) <= MAX_VISUAL_NUMERIC_VALUE && Math.abs(doubleValue) >= 0.0001) {
      return {
        fieldPath,
        fieldNumber: field.fieldNumber,
        wireType: field.wireType,
        value: roundGeometryNumber(doubleValue),
        encoding: "fixed64-double",
        source: "protobuf",
        confidence: 0.72,
        order,
        endOrder: order + 8
      };
    }
  }

  return undefined;
}

function inferGeometryCandidates(numericCandidates: IwaNumericCandidate[]): NativeIwaGeometryCandidate[] {
  const candidates: NativeIwaGeometryCandidate[] = [];
  candidates.push(...inferGroupedGeometryCandidates(numericCandidates));
  const windows = consecutiveNumericWindows(numericCandidates);

  for (const window of windows) {
    if (window.length < 4) {
      continue;
    }
    for (let index = 0; index <= window.length - 4; index += 1) {
      const group = window.slice(index, index + 4);
      const values = group.map((candidate) => candidate.value);
      const [x, y, width, height] = values;
      if (!isPlausibleGeometryTuple(x, y, width, height)) {
        continue;
      }
      if (!hasSameParentFieldPath(group) || !hasSequentialFieldNumbers(group)) {
        continue;
      }
      const fieldPaths = group.map((candidate) => candidate.fieldPath);
      const groupPath = commonFieldPathPrefix(fieldPaths);
      candidates.push({
        bounds: {
          x: roundGeometryNumber(x),
          y: roundGeometryNumber(y),
          width: roundGeometryNumber(width),
          height: roundGeometryNumber(height)
        },
        fieldPaths,
        values: values.map(roundGeometryNumber),
        source: "protobuf",
        confidence: geometryTupleConfidence(group, groupPath),
        groupPath,
        reason: groupPath
          ? "four nearby visual numeric fields under a common protobuf field path"
          : "four nearby visual numeric fields in protobuf order"
      });
    }
  }

  return dedupeGeometryCandidates(candidates)
    .sort((left, right) => right.confidence - left.confidence || compareFieldPaths(left.fieldPaths[0] ?? "", right.fieldPaths[0] ?? ""))
    .slice(0, MAX_GEOMETRY_CANDIDATES_PER_STREAM);
}

function inferGroupedGeometryCandidates(numericCandidates: IwaNumericCandidate[]): NativeIwaGeometryCandidate[] {
  const groups = new Map<string, IwaNumericCandidate[]>();
  for (const candidate of numericCandidates) {
    const parent = parentFieldPath(candidate.fieldPath);
    if (!parent) {
      continue;
    }
    const group = groups.get(parent);
    if (group) {
      group.push(candidate);
    } else {
      groups.set(parent, [candidate]);
    }
  }

  const candidates: NativeIwaGeometryCandidate[] = [];
  for (const [groupPath, group] of groups) {
    const byFieldNumber = new Map<number, IwaNumericCandidate[]>();
    for (const candidate of group) {
      const existing = byFieldNumber.get(candidate.fieldNumber);
      if (existing) {
        existing.push(candidate);
      } else {
        byFieldNumber.set(candidate.fieldNumber, [candidate]);
      }
    }

    const fieldNumbers = Array.from(byFieldNumber.keys()).sort((left, right) => left - right);
    for (let index = 0; index <= fieldNumbers.length - 4; index += 1) {
      const first = fieldNumbers[index]!;
      const tupleFieldNumbers = [first, first + 1, first + 2, first + 3];
      if (!tupleFieldNumbers.every((fieldNumber) => byFieldNumber.has(fieldNumber))) {
        continue;
      }

      const tuple = tupleFieldNumbers.map((fieldNumber) =>
        byFieldNumber.get(fieldNumber)!.sort(compareNumericCandidates)[0]!
      );
      const orderedTuple = tuple.slice().sort((left, right) => left.order - right.order);
      if (!tuple.every((candidate, tupleIndex) => candidate === orderedTuple[tupleIndex])) {
        continue;
      }
      const values = tuple.map((candidate) => candidate.value);
      const [x, y, width, height] = values;
      if (!isPlausibleGeometryTuple(x, y, width, height)) {
        continue;
      }
      candidates.push({
        bounds: {
          x: roundGeometryNumber(x),
          y: roundGeometryNumber(y),
          width: roundGeometryNumber(width),
          height: roundGeometryNumber(height)
        },
        fieldPaths: tuple.map((candidate) => candidate.fieldPath),
        values: values.map(roundGeometryNumber),
        source: "protobuf",
        confidence: geometryTupleConfidence(tuple, groupPath),
        groupPath,
        reason: "sequential visual numeric fields under a common protobuf parent path"
      });
    }
  }
  return candidates;
}

function consecutiveNumericWindows(numericCandidates: IwaNumericCandidate[]): IwaNumericCandidate[][] {
  const sorted = numericCandidates
    .filter((candidate) => candidate.value >= 0 && candidate.value <= MAX_VISUAL_NUMERIC_VALUE)
    .sort((left, right) => left.order - right.order);
  const windows: IwaNumericCandidate[][] = [];
  let current: IwaNumericCandidate[] = [];

  for (const candidate of sorted) {
    const previous = current[current.length - 1];
    if (!previous || candidate.order - previous.endOrder <= 8) {
      current.push(candidate);
      continue;
    }
    if (current.length >= 4) {
      windows.push(current);
    }
    current = [candidate];
  }
  if (current.length >= 4) {
    windows.push(current);
  }
  return windows;
}

function isPlausibleGeometryTuple(x: number, y: number, width: number, height: number): boolean {
  if (![x, y, width, height].every((value) => Number.isFinite(value) && value >= 0)) {
    return false;
  }
  if (width < 1 || height < 1 || width > MAX_VISUAL_NUMERIC_VALUE || height > MAX_VISUAL_NUMERIC_VALUE) {
    return false;
  }
  if (x > MAX_VISUAL_NUMERIC_VALUE || y > MAX_VISUAL_NUMERIC_VALUE) {
    return false;
  }
  return width >= 2 || height >= 2;
}

function isPlausibleTypedImageGeometryTuple(x: number, y: number, width: number, height: number): boolean {
  if (![x, y, width, height].every((value) => Number.isFinite(value))) {
    return false;
  }
  if (width < 1 || height < 1 || width > MAX_VISUAL_NUMERIC_VALUE || height > MAX_VISUAL_NUMERIC_VALUE) {
    return false;
  }
  return Math.abs(x) <= MAX_VISUAL_NUMERIC_VALUE && Math.abs(y) <= MAX_VISUAL_NUMERIC_VALUE;
}

function geometryTupleConfidence(group: IwaNumericCandidate[], groupPath: string | undefined): number {
  const encodingScore = group.reduce((total, candidate) => total + candidate.confidence, 0) / group.length;
  const pathScore = groupPath ? 0.12 : 0;
  const fieldScore = hasSequentialFieldNumbers(group) ? 0.1 : 0;
  return roundConfidence(Math.min(0.88, encodingScore + pathScore + fieldScore));
}

function hasSequentialFieldNumbers(group: IwaNumericCandidate[]): boolean {
  for (let index = 1; index < group.length; index += 1) {
    if (group[index]!.fieldNumber !== group[index - 1]!.fieldNumber + 1) {
      return false;
    }
  }
  return true;
}

function hasSameParentFieldPath(group: IwaNumericCandidate[]): boolean {
  const parent = parentFieldPath(group[0]?.fieldPath ?? "");
  return parent !== undefined && group.every((candidate) => parentFieldPath(candidate.fieldPath) === parent);
}

function dedupeGeometryCandidates(candidates: NativeIwaGeometryCandidate[]): NativeIwaGeometryCandidate[] {
  const seen = new Map<string, NativeIwaGeometryCandidate>();
  for (const candidate of candidates) {
    const key = `${candidate.bounds.x}:${candidate.bounds.y}:${candidate.bounds.width}:${candidate.bounds.height}:${candidate.fieldPaths.join(",")}`;
    const existing = seen.get(key);
    if (!existing || candidate.confidence > existing.confidence) {
      seen.set(key, candidate);
    }
  }
  return Array.from(seen.values());
}

function inferGroupingHints(
  textCandidates: IwaTextCandidate[],
  referenceCandidates: IwaReferenceCandidate[],
  geometryCandidates: NativeIwaGeometryCandidate[],
  animationHints: NativeIwaAnimationHint[]
): NativeIwaGroupingHint[] {
  const groups = new Map<
    string,
    {
      textCandidates: IwaTextCandidate[];
      referenceCandidates: IwaReferenceCandidate[];
      geometryCandidates: NativeIwaGeometryCandidate[];
      animationHints: NativeIwaAnimationHint[];
      fieldPaths: Set<string>;
    }
  >();

  for (const candidate of textCandidates) {
    addGroupingFieldPath(groups, candidate.fieldPath, "textCandidates", candidate);
  }
  for (const candidate of referenceCandidates) {
    addGroupingFieldPath(groups, candidate.fieldPath, "referenceCandidates", candidate);
  }
  for (const candidate of geometryCandidates) {
    const groupPath = candidate.groupPath ?? commonFieldPathPrefix(candidate.fieldPaths);
    addGroupingFieldPath(groups, groupPath, "geometryCandidates", candidate);
    for (const fieldPath of candidate.fieldPaths) {
      const group = groupPath ? groups.get(groupPath) : undefined;
      group?.fieldPaths.add(fieldPath);
    }
  }
  for (const hint of animationHints) {
    addGroupingFieldPath(groups, hint.fieldPath, "animationHints", hint);
  }

  return Array.from(groups.entries())
    .map(([groupPath, group]) => {
      const titleCandidate = group.textCandidates.sort(compareTextCandidates)[0]?.text;
      const signalKinds = [
        group.textCandidates.length > 0,
        group.referenceCandidates.length > 0,
        group.geometryCandidates.length > 0,
        group.animationHints.length > 0
      ].filter(Boolean).length;
      return {
        groupPath,
        ...(titleCandidate ? { titleCandidate } : {}),
        textCandidateCount: group.textCandidates.length,
        referenceCandidateCount: group.referenceCandidates.length,
        geometryCandidateCount: group.geometryCandidates.length,
        animationHintCount: group.animationHints.length,
        fieldPaths: Array.from(group.fieldPaths).sort(compareFieldPaths).slice(0, 12),
        confidence: roundConfidence(Math.min(0.82, 0.24 + signalKinds * 0.12 + Math.min(0.18, group.fieldPaths.size * 0.015)))
      };
    })
    .filter((hint) => hint.fieldPaths.length >= 2 || hint.geometryCandidateCount > 0 || hint.animationHintCount > 0)
    .sort((left, right) => right.confidence - left.confidence || compareFieldPaths(left.groupPath, right.groupPath))
    .slice(0, MAX_GROUPING_HINTS_PER_STREAM);
}

function addGroupingFieldPath<
  K extends "textCandidates" | "referenceCandidates" | "geometryCandidates" | "animationHints"
>(
  groups: Map<
    string,
    {
      textCandidates: IwaTextCandidate[];
      referenceCandidates: IwaReferenceCandidate[];
      geometryCandidates: NativeIwaGeometryCandidate[];
      animationHints: NativeIwaAnimationHint[];
      fieldPaths: Set<string>;
    }
  >,
  fieldPath: string | undefined,
  key: K,
  value: {
    textCandidates: IwaTextCandidate;
    referenceCandidates: IwaReferenceCandidate;
    geometryCandidates: NativeIwaGeometryCandidate;
    animationHints: NativeIwaAnimationHint;
  }[K]
): void {
  if (!fieldPath) {
    return;
  }
  const groupPath = parentFieldPath(fieldPath);
  if (!groupPath) {
    return;
  }
  let group = groups.get(groupPath);
  if (!group) {
    group = {
      textCandidates: [],
      referenceCandidates: [],
      geometryCandidates: [],
      animationHints: [],
      fieldPaths: new Set<string>()
    };
    groups.set(groupPath, group);
  }
  (group[key] as Array<typeof value>).push(value);
  group.fieldPaths.add(fieldPath);
}

function detectAnimationHints(
  text: string,
  source: "protobuf" | "raw",
  confidence: number,
  fieldPath?: string
): NativeIwaAnimationHint[] {
  const normalized = text.replace(/\s+/g, " ").trim();
  const lower = normalized.toLowerCase();
  const hints: NativeIwaAnimationHint[] = [];
  if (/\bmagic\s*move\b/.test(lower) || /\bmagicmove\b/.test(lower)) {
    hints.push({
      kind: "magicMove",
      evidence: normalized.slice(0, 120),
      source,
      confidence: roundConfidence(confidence + 0.18),
      fieldPath
    });
  }
  if (/\bmorph(?:ing)?\b/.test(lower)) {
    hints.push({
      kind: "morph",
      evidence: normalized.slice(0, 120),
      source,
      confidence: roundConfidence(confidence + 0.12),
      fieldPath
    });
  }
  if (/\btransition\b/.test(lower)) {
    hints.push({
      kind: "transition",
      evidence: normalized.slice(0, 120),
      source,
      confidence: roundConfidence(confidence + 0.06),
      fieldPath
    });
  }
  if (/\b(build|appear|dissolve|move in|fly in)\b/.test(lower)) {
    hints.push({
      kind: "build",
      evidence: normalized.slice(0, 120),
      source,
      confidence: roundConfidence(confidence),
      fieldPath
    });
  }
  return hints;
}

function scanRawStrings(data: Uint8Array, orderBase: number): {
  textCandidates: IwaTextCandidate[];
  referenceCandidates: IwaReferenceCandidate[];
  animationHints: NativeIwaAnimationHint[];
  rawStringCount: number;
} {
  const textCandidates = new Map<string, IwaTextCandidate>();
  const referenceCandidates = new Map<string, IwaReferenceCandidate>();
  const animationHints = new Map<string, NativeIwaAnimationHint>();
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
        for (const hint of detectAnimationHints(text, "raw", 0.36)) {
          addBestAnimationHint(animationHints, hint);
        }
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
    animationHints: Array.from(animationHints.values()),
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

  const dataId = dataIdFromReference(candidate.value);
  if (dataId) {
    const dataIdMatches = assets.byDataId.get(dataId) ?? [];
    if (dataIdMatches.length === 1) {
      return [
        {
          asset: dataIdMatches[0]!,
          evidence: candidate.value,
          confidence: Math.min(0.86, candidate.confidence + 0.12),
          fieldPath: candidate.fieldPath,
          source: candidate.source
        }
      ];
    }
    if (dataIdMatches.length > 1) {
      return dataIdMatches.map((asset) => ({
        asset,
        evidence: candidate.value,
        confidence: Math.min(0.7, candidate.confidence),
        fieldPath: candidate.fieldPath,
        source: candidate.source
      }));
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

function dataIdFromAssetPath(assetPath: string): string | undefined {
  const normalized = normalizePartPath(assetPath);
  if (!normalized.startsWith("Data/")) {
    return undefined;
  }
  const stem = path.posix.basename(normalized, path.posix.extname(normalized));
  const match = stem.match(/(?:^|[-_])(\d{1,12})$/);
  return match?.[1] ? normalizeDataId(match[1]) : undefined;
}

function dataIdFromReference(value: string): string | undefined {
  const normalized = safeDecodeUriComponent(normalizeReferenceText(value));
  const explicit = normalized.match(/(?:^|[^A-Za-z0-9])(?:data[-_ ]?id|sfdata|dataref|data-ref)[:=#\s-]*(\d{1,12})(?:$|[^A-Za-z0-9])/i);
  if (explicit?.[1]) {
    return normalizeDataId(explicit[1]);
  }
  const dataPathId = dataIdFromAssetPath(normalized);
  if (dataPathId) {
    return dataPathId;
  }
  if (/^\d{1,12}$/.test(normalized)) {
    return normalizeDataId(normalized);
  }
  return undefined;
}

function normalizeDataId(value: string): string {
  const digits = value.trim().replace(/^0+(?=\d)/, "");
  return digits || "0";
}

function uniqueStrings(values: string[]): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  for (const value of values) {
    const normalized = value.trim();
    if (!normalized) {
      continue;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    out.push(normalized);
  }
  return out;
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
    numericCandidateCount: 0,
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
  summary.numericCandidateCount += incoming.numericCandidateCount;
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
  summary.minNumericValue =
    summary.minNumericValue === undefined
      ? incoming.minNumericValue
      : incoming.minNumericValue === undefined
        ? summary.minNumericValue
        : Math.min(summary.minNumericValue, incoming.minNumericValue);
  summary.maxNumericValue =
    summary.maxNumericValue === undefined
      ? incoming.maxNumericValue
      : incoming.maxNumericValue === undefined
        ? summary.maxNumericValue
        : Math.max(summary.maxNumericValue, incoming.maxNumericValue);
  summary.sampleNumericValue ??= incoming.sampleNumericValue;
  summary.sampleNumericEncoding ??= incoming.sampleNumericEncoding;
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

function addBestNumericCandidate(candidates: Map<string, IwaNumericCandidate>, candidate: IwaNumericCandidate): void {
  const key = `${candidate.fieldPath}:${candidate.encoding}:${candidate.value}:${candidate.order}`;
  const existing = candidates.get(key);
  if (!existing || compareNumericCandidates(candidate, existing) < 0) {
    candidates.set(key, candidate);
  }
}

function addBestAnimationHint(candidates: Map<string, NativeIwaAnimationHint>, candidate: NativeIwaAnimationHint): void {
  const key = `${candidate.kind}:${candidate.evidence.toLowerCase()}`;
  const existing = candidates.get(key);
  if (!existing || compareAnimationHints(candidate, existing) < 0) {
    candidates.set(key, candidate);
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

function compareNumericCandidates(left: IwaNumericCandidate, right: IwaNumericCandidate): number {
  const confidence = right.confidence - left.confidence;
  if (confidence !== 0) {
    return confidence;
  }
  if (left.order !== right.order) {
    return left.order - right.order;
  }
  const fieldPath = compareFieldPaths(left.fieldPath, right.fieldPath);
  if (fieldPath !== 0) return fieldPath;
  return left.value - right.value;
}

function compareAnimationHints(left: NativeIwaAnimationHint, right: NativeIwaAnimationHint): number {
  const kind = animationKindRank(right.kind) - animationKindRank(left.kind);
  if (kind !== 0) {
    return kind;
  }
  const confidence = (right.confidence ?? 0) - (left.confidence ?? 0);
  if (confidence !== 0) {
    return confidence;
  }
  if (left.fieldPath && right.fieldPath) {
    const fieldPath = compareFieldPaths(left.fieldPath, right.fieldPath);
    if (fieldPath !== 0) return fieldPath;
  } else if (left.fieldPath) {
    return -1;
  } else if (right.fieldPath) {
    return 1;
  }
  return left.evidence.localeCompare(right.evidence);
}

function compareFieldSummaries(left: NativeIwaFieldSummary, right: NativeIwaFieldSummary): number {
  const leftScore =
    left.textCandidateCount * 4 +
    left.referenceCandidateCount * 4 +
    left.numericCandidateCount * 2 +
    left.nestedMessageCount * 2 +
    left.occurrences;
  const rightScore =
    right.textCandidateCount * 4 +
    right.referenceCandidateCount * 4 +
    right.numericCandidateCount * 2 +
    right.nestedMessageCount * 2 +
    right.occurrences;
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

function animationKindRank(kind: NativeIwaAnimationHint["kind"]): number {
  switch (kind) {
    case "magicMove":
      return 4;
    case "morph":
      return 3;
    case "transition":
      return 2;
    case "build":
      return 1;
  }
}

function stripNumericCandidateInternalFields(candidate: IwaNumericCandidate): NativeIwaNumericCandidate {
  return {
    fieldPath: candidate.fieldPath,
    fieldNumber: candidate.fieldNumber,
    wireType: candidate.wireType,
    value: candidate.value,
    encoding: candidate.encoding,
    confidence: candidate.confidence
  };
}

function parentFieldPath(fieldPath: string): string | undefined {
  const parts = fieldPath.split(".");
  if (parts.length < 2) {
    return undefined;
  }
  return parts.slice(0, -1).join(".");
}

function commonFieldPathPrefix(fieldPaths: string[]): string | undefined {
  if (fieldPaths.length === 0) {
    return undefined;
  }
  const first = fieldPaths[0]?.split(".") ?? [];
  const prefix: string[] = [];
  for (let index = 0; index < first.length; index += 1) {
    const value = first[index];
    if (value === undefined || !fieldPaths.every((fieldPath) => fieldPath.split(".")[index] === value)) {
      break;
    }
    prefix.push(value);
  }
  return prefix.length > 0 ? prefix.join(".") : undefined;
}

function roundGeometryNumber(value: number): number {
  return Number(value.toFixed(3));
}

function roundConfidence(value: number): number {
  return Number(Math.max(0, Math.min(1, value)).toFixed(3));
}

function clampGeometryCandidateBounds(
  bounds: { x: number; y: number; width: number; height: number },
  deckSize: { width: number; height: number }
): { x: number; y: number; width: number; height: number } {
  const maxWidth = Math.max(1, deckSize.width * 2);
  const maxHeight = Math.max(1, deckSize.height * 2);
  const x = Math.max(0, Math.min(bounds.x, maxWidth));
  const y = Math.max(0, Math.min(bounds.y, maxHeight));
  const width = Math.max(1, Math.min(bounds.width, Math.max(1, maxWidth - x)));
  const height = Math.max(1, Math.min(bounds.height, Math.max(1, maxHeight - y)));
  return {
    x: roundGeometryNumber(x),
    y: roundGeometryNumber(y),
    width: roundGeometryNumber(width),
    height: roundGeometryNumber(height)
  };
}

function normalizeAssetStem(value: string): string {
  return safeDecodeUriComponent(value).toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function expandIwaPayloads(data: Uint8Array): ExpandedIwaPayload[] {
  const iwaChunks = decodeIwaSnappyChunks(data);
  if (iwaChunks.length > 0) {
    return dedupeExpandedPayloads(iwaChunks.map((chunk) => ({ data: chunk, compression: "iwa-snappy-chunk" })));
  }

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

function decodeIwaSnappyChunks(data: Uint8Array): Uint8Array[] {
  const chunks: Uint8Array[] = [];
  let offset = 0;

  while (offset + 4 <= data.byteLength) {
    const chunkType = data[offset];
    const chunkLength = readUint24(data, offset + 1);
    const chunkStart = offset + 4;
    const chunkEnd = chunkStart + chunkLength;
    if (chunkType === undefined || chunkLength <= 0 || chunkEnd > data.byteLength) {
      return [];
    }

    const decoded = decodeSnappyBlock(data.subarray(chunkStart, chunkEnd));
    if (!decoded) {
      return [];
    }
    chunks.push(decoded);
    offset = chunkEnd;
  }

  return chunks.length > 0 && offset === data.byteLength ? chunks : [];
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
  let text = stripInvalidTextChars(value).replace(/\u0000/g, "").replace(/\s+/g, " ").trim();
  for (const marker of CONTAINER_MARKERS) {
    text = text.replaceAll(marker, " ");
  }
  text = cleanHumanTextResidue(text);
  text = text.replace(/^[|:;,\-_.\s]+|[|:;,\-_.\s]+$/g, "").replace(/\s+/g, " ").trim();
  if (!text || isContainerMarker(text)) {
    return undefined;
  }
  if (!looksLikePresentationText(text)) {
    return undefined;
  }
  return text;
}

function cleanHumanTextResidue(value: string): string {
  let text = value.trim();
  text = text.replace(/[\s$*\\]+$/g, "").trim();
  if (/[\p{Script=Han}]/u.test(text)) {
    text = text.replace(/\s+["'“”‘’][A-Za-z0-9+&()]{1,4}$/u, "").trim();
    text = text.replace(/[$*\\]+[A-Za-z0-9+&()]{1,4}$/u, "").trim();
    text = text.replace(/[\s$*\\]+$/g, "").trim();
  }
  text = cleanDuplicatedAsciiTail(text);
  return text;
}

function looksLikePresentationText(text: string): boolean {
  if (isLikelyNativeNoiseText(text) || hasBinaryResidue(text)) {
    return false;
  }
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
  if (/^(apple|com\.apple|x-apple):/i.test(text) || /^apple[-.:]/i.test(text)) {
    return false;
  }
  if (/^(decimal|double|float|string|boolean|integer|uint|sint|fixed|sfixed)$/i.test(text)) {
    return false;
  }
  if (/^(path|transition|build|animation|effect)$/i.test(text) && text.length < 14) {
    return false;
  }
  if (/^(?:[A-Za-z0-9]{1,3}[\\/\s-]+)?transition$/i.test(text)) {
    return false;
  }
  if (/^(xbo\s+transition|image\s+\d+\p{L}?|picture\s+\d+\p{L}?|图片\s*\d+\p{L}?|圖像\s*\d+\p{L}?|影像\s*\d+\p{L}?)$/iu.test(text)) {
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
  if (text.length <= 4 && !/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(text)) {
    return false;
  }
  if (text.length <= 8 && !/[\p{Script=Han}]/u.test(text) && /[^A-Za-z0-9\s\p{L}\p{N}]/u.test(text)) {
    return false;
  }
  if (text.length <= 10 && hasMixedLowSignalScripts(text)) {
    return false;
  }
  if (hasSuspiciousShortAsciiTail(text)) {
    return false;
  }
  if (!hasHumanTextSignal(text)) {
    return false;
  }
  return /[\p{L}\p{N}]/u.test(text);
}

function isLikelyNativeNoiseText(text: string): boolean {
  const lower = text.toLowerCase();
  if (hasInternalKeynoteToken(text)) {
    return true;
  }
  if (hasNativeLocaleResidue(text)) {
    return true;
  }
  if (hasInternalKeynoteTextResidue(text)) {
    return true;
  }
  if (hasConcatenatedNativeTextResidue(text)) {
    return true;
  }
  if (isLowSignalNativeAsciiResidue(text)) {
    return true;
  }
  if (/\b(decimal|double|float|string|boolean|integer|uint|sint|fixed|sfixed)\b/.test(lower) && /[.$*@\\]/.test(text)) {
    return true;
  }
  if (/\bmagic-move-implied-motion-path\b/.test(lower)) {
    return true;
  }
  if (/\bxbo transition\b/i.test(text) && /[$*@\\]|\bdecimal\b/i.test(text)) {
    return true;
  }
  if (/\btransition\b/i.test(text) && /\bmagic[- ]move\b/i.test(text) && /[$*@\\]/.test(text)) {
    return true;
  }
  if (/[\p{Script=Han}]/u.test(text) && /[A-Za-z0-9][*\\]$/.test(text)) {
    return true;
  }
  if (/[\p{Script=Han}]/u.test(text) && /["'“”‘’]?[A-Za-z0-9+&()]{1,4}$/.test(text)) {
    return true;
  }
  return false;
}

function hasNativeLocaleResidue(text: string): boolean {
  const tokens = text.trim().split(/\s+/).filter(Boolean);
  if (tokens.length < 2) {
    return false;
  }
  const localeTokens = tokens.filter(isNativeLocaleToken);
  if (localeTokens.length >= 2 && localeTokens.length / tokens.length >= 0.5) {
    return true;
  }
  return localeTokens.length > 0 && tokens.some((token) => /[=+&]/.test(token) || isLowSignalNativeAsciiResidue(token));
}

function isNativeLocaleToken(token: string): boolean {
  return /^[a-z]{2,3}(?:-[a-z0-9]{2,8})+$/i.test(token);
}

function hasInternalKeynoteTextResidue(text: string): boolean {
  const lower = text.toLowerCase();
  const hasTransitionResidue = /(?:^|\s)(?:xbo\s+)?(?:t|r)?ansition\b/i.test(text);
  if (
    hasTransitionResidue &&
    (hasNativeLocaleResidue(text) || /\b\d+[A-Z]?\/\s*transition\b/i.test(text) || /[$*@\\]/.test(text) || /\bmagic[- ]move\b/i.test(lower))
  ) {
    return true;
  }
  return /(?:^|\s|[a-z])nsition\$[a-z]\b/i.test(text);
}

function hasConcatenatedNativeTextResidue(text: string): boolean {
  if (!/[\p{Script=Han}]/u.test(text)) {
    return false;
  }
  return /\s+\d+[A-Z][A-Za-z]/.test(text) || /(?:^|\s|[a-z])nsition\$[a-z]\b/i.test(text);
}

function isLowSignalNativeAsciiResidue(text: string): boolean {
  if (text.length > 18 || /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(text)) {
    return false;
  }
  const chars = Array.from(text.replace(/\s+/g, ""));
  if (chars.length < 3) {
    return false;
  }
  const letters = chars.filter((char) => /[A-Za-z]/.test(char)).length;
  const digits = chars.filter((char) => /\d/.test(char)).length;
  const symbols = chars.filter((char) => /[^A-Za-z0-9]/.test(char)).length;
  const vowels = chars.filter((char) => /[AEIOUaeiou]/.test(char)).length;
  if (letters >= 4 && vowels === 0 && /\s/.test(text) && /^[A-Z0-9\s]+$/.test(text)) {
    return true;
  }
  if (symbols >= 2 && digits + symbols >= letters) {
    return true;
  }
  if (symbols >= 1 && digits >= 1 && vowels <= 1 && letters <= 8) {
    return true;
  }
  if (symbols >= 1 && vowels === 0 && letters >= 4) {
    return true;
  }
  return false;
}

function stripInvalidTextChars(value: string): string {
  return Array.from(value)
    .filter((char) => {
      const code = char.codePointAt(0) ?? 0;
      return code === 0x09 || code === 0x0a || code === 0x0d || (code >= 0x20 && code <= 0xd7ff) || (code >= 0xe000 && code <= 0xfffd) || (code >= 0x10000 && code <= 0x10ffff);
    })
    .join("");
}

function hasInternalKeynoteToken(text: string): boolean {
  return /(^|[\s$*])apple:[a-z0-9._:-]+/i.test(text) || /\bmagic-move-implied-motion-path\b/i.test(text);
}

function hasBinaryResidue(text: string): boolean {
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(text)) {
    return true;
  }
  if (/^[A-Za-z0-9@*.$\\/_-]{1,8}$/.test(text) && /[.@*$\\/_-]/.test(text)) {
    return true;
  }
  if (/^[\p{L}\p{N}\s"'.,:;!?%&()[\]{}<>/@#$*+\\_-]+$/u.test(text)) {
    const asciiSymbols = Array.from(text).filter((char) => /["'.,:;!?%&()[\]{}<>/@#$*+\\_-]/.test(char)).length;
    const lettersAndNumbers = Array.from(text).filter((char) => /[\p{L}\p{N}]/u.test(char)).length;
    const threshold = /[\p{Script=Han}]/u.test(text) ? 0.45 : 0.28;
    if (lettersAndNumbers > 0 && asciiSymbols / lettersAndNumbers > threshold) {
      return true;
    }
  }
  return false;
}

function hasMixedLowSignalScripts(text: string): boolean {
  const hasCyrillic = /[\p{Script=Cyrillic}]/u.test(text);
  const hasLatin = /[A-Za-z]/.test(text);
  const hasDigit = /\d/.test(text);
  const hasHan = /[\p{Script=Han}]/u.test(text);
  return !hasHan && hasCyrillic && (hasLatin || hasDigit);
}

function cleanDuplicatedAsciiTail(text: string): string {
  const match = text.match(/^([A-Za-z][A-Za-z0-9]*(?: [A-Za-z][A-Za-z0-9]*){1,5})([A-Za-z]{2,3})$/);
  if (!match) return text;
  const base = match[1] ?? text;
  const tail = match[2] ?? "";
  const lastWord = base.split(/\s+/).at(-1) ?? "";
  return lastWord.toLowerCase().startsWith(tail.toLowerCase()) ? base : text;
}

function hasSuspiciousShortAsciiTail(text: string): boolean {
  if (!/\s/.test(text)) return false;
  if (!/^[A-Za-z0-9 ?!.,:;'"/&()+_-]+$/.test(text)) return false;
  const words = text.trim().split(/\s+/);
  const last = words.at(-1) ?? "";
  const previous = words.at(-2) ?? "";
  if (/^[A-Za-z]{1,3}$/.test(last) && previous.length >= 4 && previous.toLowerCase().startsWith(last.toLowerCase())) {
    return true;
  }
  return false;
}

function hasHumanTextSignal(text: string): boolean {
  if (/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(text)) {
    return true;
  }
  if (/\s/.test(text) && /[A-Za-z]{2,}/.test(text)) {
    return true;
  }
  if (/^[A-Z][A-Za-z0-9]+(?:[ -][A-Za-z0-9]+)+$/.test(text)) {
    return true;
  }
  return false;
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

function readNativeDocumentDeckSize(entries: Map<string, Uint8Array>): { width: number; height: number } | undefined {
  const documentEntry = Array.from(entries.entries()).find(([entryPath]) => /^Index\/Document\.iwa$/i.test(normalizePartPath(entryPath)));
  if (!documentEntry) {
    return undefined;
  }
  const scan = scanIwaEntry(documentEntry[1]);
  const candidates = scan.archiveMessages
    .filter((message) => message.type === 2)
    .map((message) => nativeDocumentDeckSizeCandidate(message))
    .filter((size): size is { width: number; height: number; confidence: number } => Boolean(size))
    .sort((left, right) => right.confidence - left.confidence || right.width * right.height - left.width * left.height);
  const best = candidates[0];
  return best ? { width: best.width, height: best.height } : undefined;
}

function nativeDocumentDeckSizeCandidate(
  message: NativeIwaArchiveMessageEvidence
): { width: number; height: number; confidence: number } | undefined {
  const byPath = new Map(message.fieldSummaries.map((summary) => [summary.fieldPath, summary]));
  const width = byPath.get("1")?.sampleNumericValue;
  const height = byPath.get("2")?.sampleNumericValue;
  if (!isPlausibleNativeDeckSize(width, height)) {
    return undefined;
  }
  const confidence = roundConfidence(
    0.72 +
      (message.archiveIdentifier ? 0.08 : 0) +
      (byPath.get("1")?.occurrences === 1 ? 0.04 : 0) +
      (byPath.get("2")?.occurrences === 1 ? 0.04 : 0)
  );
  return {
    width: Math.round(width),
    height: Math.round(height),
    confidence
  };
}

function isPlausibleNativeDeckSize(width: number | undefined, height: number | undefined): boolean {
  return (
    width !== undefined &&
    height !== undefined &&
    Number.isFinite(width) &&
    Number.isFinite(height) &&
    width >= 320 &&
    height >= 240 &&
    width <= 10000 &&
    height <= 10000
  );
}

function createQuickLookPreviewMetadata(entryPath: string, data: Uint8Array): NativeQuickLookPreviewMetadata {
  const classification = classifyAssetPath(entryPath);
  const dimensions = readImageDimensions(data, classification.mimeType);
  return {
    path: entryPath,
    byteLength: data.byteLength,
    role: classifyQuickLookRole(entryPath),
    mimeType: classification.mimeType,
    ...(dimensions ? { width: dimensions.width, height: dimensions.height } : {})
  };
}

function classifyQuickLookRole(entryPath: string): NativeQuickLookPreviewMetadata["role"] {
  const baseName = path.posix.basename(entryPath).toLowerCase();
  if (/thumb|thumbnail/.test(baseName)) {
    return "thumbnail";
  }
  if (/preview|slide|page/.test(baseName)) {
    return "preview";
  }
  return "unknown";
}

function readImageDimensions(
  data: Uint8Array,
  mimeType: string | undefined
): { width: number; height: number; source: string } | undefined {
  return (
    readPngDimensions(data) ??
    readJpegDimensions(data) ??
    readGifDimensions(data) ??
    readWebpDimensions(data) ??
    (mimeType === "image/svg+xml" ? readSvgDimensions(data) : undefined) ??
    (mimeType === "image/tiff" ? readTiffDimensions(data) : undefined)
  );
}

function readPngDimensions(data: Uint8Array): { width: number; height: number; source: string } | undefined {
  if (
    data.byteLength < 24 ||
    data[0] !== 0x89 ||
    data[1] !== 0x50 ||
    data[2] !== 0x4e ||
    data[3] !== 0x47 ||
    decodeAscii(data.subarray(12, 16)) !== "IHDR"
  ) {
    return undefined;
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  return validateImageDimensions(view.getUint32(16, false), view.getUint32(20, false), "png-ihdr");
}

function readJpegDimensions(data: Uint8Array): { width: number; height: number; source: string } | undefined {
  if (data.byteLength < 4 || data[0] !== 0xff || data[1] !== 0xd8) {
    return undefined;
  }
  let offset = 2;
  while (offset + 9 <= data.byteLength) {
    if (data[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    let markerOffset = offset + 1;
    while (data[markerOffset] === 0xff) {
      markerOffset += 1;
    }
    const marker = data[markerOffset];
    if (marker === undefined || marker === 0xd9 || marker === 0xda) {
      break;
    }
    const lengthOffset = markerOffset + 1;
    if (lengthOffset + 2 > data.byteLength) {
      break;
    }
    const segmentLength = ((data[lengthOffset] ?? 0) << 8) | (data[lengthOffset + 1] ?? 0);
    if (segmentLength < 2 || lengthOffset + segmentLength > data.byteLength) {
      break;
    }
    if (isJpegStartOfFrameMarker(marker)) {
      const height = ((data[lengthOffset + 3] ?? 0) << 8) | (data[lengthOffset + 4] ?? 0);
      const width = ((data[lengthOffset + 5] ?? 0) << 8) | (data[lengthOffset + 6] ?? 0);
      return validateImageDimensions(width, height, "jpeg-sof");
    }
    offset = lengthOffset + segmentLength;
  }
  return undefined;
}

function isJpegStartOfFrameMarker(marker: number | undefined): boolean {
  return marker !== undefined && marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker);
}

function readGifDimensions(data: Uint8Array): { width: number; height: number; source: string } | undefined {
  if (data.byteLength < 10) {
    return undefined;
  }
  const signature = decodeAscii(data.subarray(0, 6));
  if (signature !== "GIF87a" && signature !== "GIF89a") {
    return undefined;
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  return validateImageDimensions(view.getUint16(6, true), view.getUint16(8, true), "gif-logical-screen");
}

function readWebpDimensions(data: Uint8Array): { width: number; height: number; source: string } | undefined {
  if (data.byteLength < 30 || decodeAscii(data.subarray(0, 4)) !== "RIFF" || decodeAscii(data.subarray(8, 12)) !== "WEBP") {
    return undefined;
  }
  const chunkType = decodeAscii(data.subarray(12, 16));
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (chunkType === "VP8X" && data.byteLength >= 30) {
    const width = 1 + (data[24] ?? 0) + ((data[25] ?? 0) << 8) + ((data[26] ?? 0) << 16);
    const height = 1 + (data[27] ?? 0) + ((data[28] ?? 0) << 8) + ((data[29] ?? 0) << 16);
    return validateImageDimensions(width, height, "webp-vp8x");
  }
  if (chunkType === "VP8 " && data.byteLength >= 30) {
    const width = view.getUint16(26, true) & 0x3fff;
    const height = view.getUint16(28, true) & 0x3fff;
    return validateImageDimensions(width, height, "webp-vp8");
  }
  if (chunkType === "VP8L" && data.byteLength >= 25) {
    const b0 = data[21] ?? 0;
    const b1 = data[22] ?? 0;
    const b2 = data[23] ?? 0;
    const b3 = data[24] ?? 0;
    const width = 1 + (((b1 & 0x3f) << 8) | b0);
    const height = 1 + ((b3 << 6) | (b2 << 2) | ((b1 & 0xc0) >>> 6));
    return validateImageDimensions(width, height, "webp-vp8l");
  }
  return undefined;
}

function readSvgDimensions(data: Uint8Array): { width: number; height: number; source: string } | undefined {
  const text = decodeUtf8(data);
  if (!text || !/<svg[\s>]/i.test(text)) {
    return undefined;
  }
  const svgTag = text.match(/<svg\b[^>]*>/i)?.[0];
  if (!svgTag) {
    return undefined;
  }
  const width = parseSvgLength(readXmlAttribute(svgTag, "width"));
  const height = parseSvgLength(readXmlAttribute(svgTag, "height"));
  if (width !== undefined && height !== undefined) {
    return validateImageDimensions(width, height, "svg-size");
  }
  const viewBox = readXmlAttribute(svgTag, "viewBox");
  if (viewBox) {
    const values = viewBox.trim().split(/[\s,]+/).map(Number);
    if (values.length >= 4) {
      return validateImageDimensions(values[2]!, values[3]!, "svg-viewbox");
    }
  }
  return undefined;
}

function readXmlAttribute(tag: string, name: string): string | undefined {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = tag.match(new RegExp(`\\s${escaped}\\s*=\\s*(['"])(.*?)\\1`, "i"));
  return match?.[2] ? unescapeXml(match[2]) : undefined;
}

function parseSvgLength(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = value.trim().match(/^([+-]?(?:\d+\.?\d*|\.\d+))(?:px)?$/i);
  if (!match) {
    return undefined;
  }
  const parsed = Number(match[1]);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function readTiffDimensions(data: Uint8Array): { width: number; height: number; source: string } | undefined {
  if (data.byteLength < 16) {
    return undefined;
  }
  const littleEndian = data[0] === 0x49 && data[1] === 0x49;
  const bigEndian = data[0] === 0x4d && data[1] === 0x4d;
  if (!littleEndian && !bigEndian) {
    return undefined;
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const magic = view.getUint16(2, littleEndian);
  if (magic !== 42) {
    return undefined;
  }
  const ifdOffset = view.getUint32(4, littleEndian);
  if (ifdOffset + 2 > data.byteLength) {
    return undefined;
  }
  const entryCount = view.getUint16(ifdOffset, littleEndian);
  let width: number | undefined;
  let height: number | undefined;
  for (let index = 0; index < entryCount; index += 1) {
    const entryOffset = ifdOffset + 2 + index * 12;
    if (entryOffset + 12 > data.byteLength) {
      break;
    }
    const tag = view.getUint16(entryOffset, littleEndian);
    const type = view.getUint16(entryOffset + 2, littleEndian);
    const count = view.getUint32(entryOffset + 4, littleEndian);
    if (count !== 1 || (type !== 3 && type !== 4)) {
      continue;
    }
    const value = type === 3 ? view.getUint16(entryOffset + 8, littleEndian) : view.getUint32(entryOffset + 8, littleEndian);
    if (tag === 256) width = value;
    if (tag === 257) height = value;
  }
  return width !== undefined && height !== undefined ? validateImageDimensions(width, height, "tiff-ifd") : undefined;
}

function validateImageDimensions(
  width: number,
  height: number,
  source: string
): { width: number; height: number; source: string } | undefined {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0 || width > 100000 || height > 100000) {
    return undefined;
  }
  return { width, height, source };
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
    case ".svg":
      return { kind: "image", mimeType: "image/svg+xml" };
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
