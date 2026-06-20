import { readFile, stat } from "node:fs/promises";
import path from "node:path";

import {
  IR_VERSION,
  type AnimationEvent,
  type Asset,
  type DeckIR,
  type Easing,
  type IRObject,
  type JSONRecord,
  type Keyframe,
  type KeyframeTrack,
  type ObjectStateProperties
} from "../ir/index.ts";

export interface KeynoteHtmlImportOptions {
  sourceKeynotePath?: string;
  sourceName?: string;
}

export type KeynoteHtmlParseOptions = KeynoteHtmlImportOptions;

type KeynoteHtmlSlideRef = string | { id?: string; slideId?: string; name?: string; title?: string };

interface KeynoteHtmlHeader {
  title?: string;
  filename?: string;
  creator?: string;
  timestamp?: string;
  slideWidth?: number;
  slideHeight?: number;
  slideList?: KeynoteHtmlSlideRef[];
  fonts?: unknown[];
  autoplayTransitionDelay?: number;
  autoplayBuildDelay?: number;
}

interface KeynoteHtmlSlide {
  assets?: Record<string, KeynoteHtmlAsset> | KeynoteHtmlAsset[];
  events?: Record<string, KeynoteHtmlEvent | KeynoteHtmlEvent[]>;
  baseLayer?: KeynoteHtmlLayer;
  layers?: KeynoteHtmlLayer[];
}

interface KeynoteHtmlAsset {
  id?: string;
  name?: string;
  kind?: string;
  type?: string;
  mimeType?: string;
  index?: number;
  path?: string;
  src?: string;
  url?:
    | string
    | {
    native?: string;
    web?: string;
  };
  width?: number;
  height?: number;
  duration?: number;
  durationMs?: number;
}

interface KeynoteHtmlEvent {
  beginTime?: number;
  duration?: number;
  baseLayer?: KeynoteHtmlLayer;
  effects?: KeynoteHtmlEvent[];
  movie?: string;
  type?: string;
  name?: string;
  objectID?: string;
}

interface KeynoteHtmlLayer {
  objectID?: string;
  id?: string;
  name?: string;
  initialState?: KeynoteHtmlLayerState;
  layers?: KeynoteHtmlLayer[];
  sublayers?: KeynoteHtmlLayer[];
  children?: KeynoteHtmlLayer[];
  animations?: KeynoteHtmlAnimation[];
  texture?: string;
  contents?: string;
  asset?: string;
  assetId?: string;
  texturedRectangle?: JSONRecord;
}

interface KeynoteHtmlLayerState {
  position?: { pointX?: number; pointY?: number; x?: number; y?: number } | number[];
  width?: number;
  height?: number;
  opacity?: number;
  hidden?: boolean;
  rotation?: number;
  scale?: number | { pointX?: number; pointY?: number; x?: number; y?: number } | number[];
  anchorPoint?: { pointX?: number; pointY?: number; x?: number; y?: number } | number[];
  contents?: string;
  src?: string;
  asset?: string;
  assetId?: string;
  contentsRect?: { x?: number; y?: number; width?: number; height?: number } | number[];
}

interface KeynoteHtmlAnimation {
  property?: string;
  beginTime?: number;
  duration?: number;
  fillMode?: string;
  timingFunction?: string;
  from?: KeynoteHtmlAnimationValue;
  to?: KeynoteHtmlAnimationValue;
  values?: KeynoteHtmlAnimationValue[];
  keyTimes?: number[];
  animations?: KeynoteHtmlAnimation[];
  removedOnCompletion?: boolean;
  autoreverses?: boolean;
  additive?: boolean;
  repeatCount?: number;
  timeOffset?: number;
}

type KeynoteHtmlAnimationValue =
  | { scalar?: number | boolean; pointX?: number; pointY?: number; x?: number; y?: number; width?: number; height?: number; texture?: string }
  | number[]
  | number
  | boolean
  | string;

interface SlideParseContext {
  exportDir: string;
  assetsDir: string;
  slideId: string;
  slideIndex: number;
  slideWidth: number;
  slideHeight: number;
  slideAssets: Record<string, KeynoteHtmlAsset>;
  deckAssets: Map<string, Asset>;
  objectIdsByLayerKey: Map<string, string[]>;
  events: AnimationEvent[];
  messages: string[];
  unsupportedProperties: Set<string>;
  recoveredLayerCount: number;
  recoveredTextureLayerCount: number;
  recoveredAnimationCount: number;
}

export async function parseKeynoteHtmlToIr(exportDir: string, options: KeynoteHtmlImportOptions = {}): Promise<DeckIR> {
  const resolvedExportDir = path.resolve(exportDir);
  const assetsDir = path.join(resolvedExportDir, "assets");
  const header = await readJson<KeynoteHtmlHeader>(path.join(assetsDir, "header.json"));
  const slideRefs = readSlideRefs(header.slideList);
  const width = finiteNumber(header.slideWidth) ?? 1920;
  const height = finiteNumber(header.slideHeight) ?? 1080;
  const deckAssets = new Map<string, Asset>();
  const unsupportedProperties = new Set<string>();
  let recoveredLayerCount = 0;
  let recoveredTextureLayerCount = 0;
  let recoveredAnimationCount = 0;
  const messages: string[] = [];

  const slides = [];
  for (const [index, slideRef] of slideRefs.entries()) {
    const slideId = slideRef.id;
    const slideJson = await readJson<KeynoteHtmlSlide>(path.join(assetsDir, slideId, `${slideId}.json`));
    const ctx: SlideParseContext = {
      exportDir: resolvedExportDir,
      assetsDir,
      slideId,
      slideIndex: index,
      slideWidth: width,
      slideHeight: height,
      slideAssets: normalizeSlideAssets(slideJson.assets),
      deckAssets,
      objectIdsByLayerKey: new Map(),
      events: [],
      messages,
      unsupportedProperties,
      recoveredLayerCount: 0,
      recoveredTextureLayerCount: 0,
      recoveredAnimationCount: 0
    };

    const baseLayers = extractSlideLayers(slideJson);
    const objects = baseLayers.flatMap((layer, baseIndex) =>
      layerToObjects(layer, ctx, {
        path: `events.baseLayer[${baseIndex}]`,
        parentBounds: { x: 0, y: 0, width, height },
        root: true
      })
    );
    appendEffectAnimations(slideJson.events, ctx);

    recoveredLayerCount += ctx.recoveredLayerCount;
    recoveredTextureLayerCount += ctx.recoveredTextureLayerCount;
    recoveredAnimationCount += ctx.recoveredAnimationCount;

    slides.push({
      id: `keynote-html-slide-${index + 1}`,
      index,
      name: slideRef.title ?? slideId,
      background: { type: "solid" as const, color: "#ffffff" },
      objects,
      timeline: {
        id: `keynote-html-slide-${index + 1}-timeline`,
        durationMs: inferSlideDurationMs(ctx.events, finiteNumber(header.autoplayTransitionDelay) ?? 5),
        events: ctx.events,
        metadata: {
          sourceSlideId: slideId,
          source: "keynote-html-export"
        }
      },
      metadata: {
        sourceSlideId: slideId,
        keynoteHtmlAssetCount: Object.keys(ctx.slideAssets).length,
        keynoteHtmlBaseLayerCount: baseLayers.length
      }
    });
  }

  const deck: DeckIR = {
    irVersion: IR_VERSION,
    deck: {
      id: "keynote-html-deck",
      title: header.title ?? options.sourceName ?? header.filename ?? "Keynote HTML Export",
      size: { width, height, unit: "px" },
      slides,
      assets: [...deckAssets.values()],
      metadata: {
        source: "keynote-html-export",
        creator: header.creator,
        timestamp: header.timestamp,
        fonts: header.fonts ?? []
      }
    },
    metadata: {
      title: header.title ?? options.sourceName ?? header.filename,
      sourceApplication: header.creator ?? "Apple Keynote",
      custom: {
        sourceFormat: "keynote-html-export",
        sourceKeynotePath: options.sourceKeynotePath,
        exportDir: resolvedExportDir
      }
    },
    conversion: {
      status: unsupportedProperties.size > 0 ? "partial" : "success",
      tool: "keymorph-keynote-html-parser",
      source: {
        kind: "html",
        uri: resolvedExportDir,
        application: "Apple Keynote HTML Export"
      },
      messages: [
        {
          severity: "info",
          code: "keynote-html-layer-json-import",
          message: "Imported Keynote HTML export layer JSON into KeyMorph IR."
        },
        {
          severity: "info",
          code: "keynote-html-static-import",
          message: `Parsed ${slides.length} slide(s), ${recoveredLayerCount} layer object(s), ${deckAssets.size} asset(s), and ${recoveredAnimationCount} animation event(s) from Keynote HTML export JSON.`
        },
        ...messages.map((message) => ({
          severity: "warning" as const,
          code: "keynote-html-layer-import-warning",
          message
        }))
      ],
      unsupportedFeatures:
        unsupportedProperties.size > 0
          ? [
              {
                code: "keynote-html-animation-property-unsupported",
                description: `Unsupported Keynote HTML animation properties were preserved in metadata: ${[...unsupportedProperties].sort().join(", ")}.`,
                severity: "warning",
                fallback: "Render supported layer geometry and basic keyframe properties; use Keynote movie runtime for exact playback."
              }
            ]
          : [],
      uncertainMappings: [
        {
          code: "keynote-html-textures-not-editable-text",
          description: "Keynote HTML export represents most text and shapes as rendered texture layers, so IR preserves visual layers rather than editable text runs.",
          severity: "warning",
          confidence: 0.86
        },
        ...(unsupportedProperties.size > 0
          ? [
              {
                code: "keynote-html-animation-custom-property",
                description: "One or more Keynote HTML animation properties were preserved as custom keyframe tracks.",
                severity: "warning" as const,
                confidence: 0.5
              }
            ]
          : [])
      ],
      statistics: {
        slideCount: slides.length,
        objectCount: recoveredLayerCount,
        animationCount: recoveredAnimationCount,
        assetCount: deckAssets.size,
        unsupportedFeatureCount: unsupportedProperties.size > 0 ? 1 : 0,
        uncertainMappingCount: unsupportedProperties.size > 0 ? 2 : 1
      },
      metadata: {
        recoveredLayerCount,
        recoveredTextureLayerCount,
        recoveredAnimationCount,
        unsupportedAnimationProperties: [...unsupportedProperties].sort()
      }
    }
  };

  return deck;
}

export async function parseKeynoteHtmlExportToIr(exportDir: string, options: KeynoteHtmlParseOptions = {}): Promise<DeckIR> {
  return parseKeynoteHtmlToIr(exportDir, options);
}

function readSlideRefs(slideList: KeynoteHtmlHeader["slideList"]): Array<{ id: string; title?: string }> {
  if (!Array.isArray(slideList)) return [];
  return slideList
    .map((entry, index) => {
      if (typeof entry === "string") return { id: entry, title: entry };
      if (!entry || typeof entry !== "object") return undefined;
      const id = stringValue(entry.id) ?? stringValue(entry.slideId) ?? stringValue(entry.name) ?? String(index + 1);
      return { id, title: stringValue(entry.title) ?? stringValue(entry.name) };
    })
    .filter((entry): entry is { id: string; title?: string } => Boolean(entry?.id));
}

function normalizeSlideAssets(assets: KeynoteHtmlSlide["assets"]): Record<string, KeynoteHtmlAsset> {
  if (!assets) return {};
  if (Array.isArray(assets)) {
    return Object.fromEntries(assets.map((asset, index) => [asset.id ?? asset.name ?? String(index), asset]));
  }
  return assets;
}

function extractSlideLayers(slideJson: KeynoteHtmlSlide): KeynoteHtmlLayer[] {
  if (slideJson.baseLayer) return extractLayerArray(slideJson.baseLayer);
  if (Array.isArray(slideJson.layers)) return slideJson.layers;
  return extractBaseLayers(slideJson.events);
}

function extractLayerArray(layer: KeynoteHtmlLayer): KeynoteHtmlLayer[] {
  const children = layer.layers ?? layer.sublayers ?? layer.children;
  return children && children.length > 0 ? children : [layer];
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function extractBaseLayers(events: KeynoteHtmlSlide["events"]): KeynoteHtmlLayer[] {
  const layers: KeynoteHtmlLayer[] = [];
  for (const value of Object.values(events ?? {})) {
    if (Array.isArray(value)) {
      for (const item of value) {
        if (item?.baseLayer) layers.push(item.baseLayer);
      }
    } else if (value?.baseLayer) {
      layers.push(value.baseLayer);
    }
  }
  return layers;
}

function layerToObjects(
  layer: KeynoteHtmlLayer,
  ctx: SlideParseContext,
  options: { path: string; parentBounds: { x: number; y: number; width: number; height: number }; root?: boolean }
): IRObject[] {
  const state = layerStateToObjectState(layer.initialState, options.parentBounds);
  const bounds = state.bounds ?? { x: 0, y: 0, width: ctx.slideWidth, height: ctx.slideHeight };
  const children: IRObject[] = [];

  for (const [childIndex, child] of (layer.layers ?? layer.sublayers ?? layer.children ?? []).entries()) {
    children.push(
      ...layerToObjects(child, ctx, {
        path: `${options.path}.layers[${childIndex}]`,
        parentBounds: bounds
      })
    );
  }

  const sourceLayerId = layer.objectID ?? layer.id ?? layer.name;
  const objectId = stableObjectId(ctx.slideIndex, options.path, sourceLayerId);
  const assetRef = layer.texture ?? layer.initialState?.contents ?? layer.initialState?.asset ?? layer.initialState?.assetId ?? layer.initialState?.src ?? layer.contents ?? layer.asset ?? layer.assetId;
  if (assetRef) {
    const asset = ensureLayerAsset(ctx, assetRef);
    if (asset) {
      ctx.recoveredLayerCount += 1;
      ctx.recoveredTextureLayerCount += 1;
      const object: IRObject =
        asset.kind === "video"
          ? {
              id: objectId,
              type: "media",
              name: layer.name ?? asset.name,
              mediaType: "video",
              source: { assetId: asset.id },
              playback: { autoplay: false, muted: true },
              bounds,
              opacity: state.opacity,
              visible: state.visible,
              transform: state.transform,
              metadata: layerMetadata(ctx, layer, options.path)
            }
          : {
              id: objectId,
              type: "image",
              name: layer.name ?? asset.name,
              source: { assetId: asset.id },
              crop: contentsRectToCrop(layer.initialState?.contentsRect),
              bounds,
              opacity: state.opacity,
              visible: state.visible,
              transform: state.transform,
              metadata: layerMetadata(ctx, layer, options.path)
            };
      appendLayerAnimations(ctx, object.id, layer.animations, options.path);
      registerLayerObject(ctx, layer, object.id);
      if (children.length > 0) {
        const group: IRObject = {
          id: `${objectId}-group`,
          type: "group",
          name: layer.name ?? `Keynote layer ${objectId}`,
          bounds,
          opacity: state.opacity,
          visible: state.visible,
          transform: state.transform,
          children: [{ ...object, metadata: { ...(object.metadata ?? {}), keynoteHtmlChildLayerCount: children.length } }, ...children],
          metadata: layerMetadata(ctx, layer, options.path)
        };
        ctx.recoveredLayerCount += 1;
        return [group];
      }
      return [object];
    }
  }

  if (children.length > 0) {
    ctx.recoveredLayerCount += 1;
    const group: IRObject = {
      id: objectId,
      type: "group",
      name: layer.name ?? `Keynote layer ${objectId}`,
      bounds,
      opacity: state.opacity,
      visible: state.visible,
      transform: state.transform,
      children,
      metadata: layerMetadata(ctx, layer, options.path)
    };
    appendLayerAnimations(ctx, group.id, layer.animations, options.path);
    registerLayerObject(ctx, layer, group.id);
    return [group];
  }

  appendLayerAnimations(ctx, objectId, layer.animations, options.path);
  return children;
}

function registerLayerObject(ctx: SlideParseContext, layer: KeynoteHtmlLayer, objectId: string): void {
  const keys = layerKeys(layer);
  for (const key of keys) {
    const ids = ctx.objectIdsByLayerKey.get(key) ?? [];
    if (!ids.includes(objectId)) ids.push(objectId);
    ctx.objectIdsByLayerKey.set(key, ids);
  }
}

function layerKeys(layer: KeynoteHtmlLayer): string[] {
  const keys = [];
  const sourceLayerId = layer.objectID ?? layer.id ?? layer.name;
  if (sourceLayerId) keys.push(`id:${sourceLayerId}`);
  const assetRef = layer.texture ?? layer.initialState?.contents ?? layer.initialState?.asset ?? layer.initialState?.assetId ?? layer.initialState?.src ?? layer.contents ?? layer.asset ?? layer.assetId;
  if (assetRef) keys.push(`asset:${assetRef}`);
  return keys;
}

function appendEffectAnimations(events: KeynoteHtmlSlide["events"], ctx: SlideParseContext): void {
  const visitEvent = (event: KeynoteHtmlEvent | undefined, sourcePath: string): void => {
    if (!event || typeof event !== "object") return;
    for (const [effectIndex, effect] of (event.effects ?? []).entries()) {
      if (effect.baseLayer) appendEffectLayerAnimations(effect.baseLayer, ctx, `${sourcePath}.effects[${effectIndex}].baseLayer`);
      visitEvent(effect, `${sourcePath}.effects[${effectIndex}]`);
    }
  };
  for (const [key, value] of Object.entries(events ?? {})) {
    if (Array.isArray(value)) {
      value.forEach((event, index) => visitEvent(event, `events.${key}[${index}]`));
    } else {
      visitEvent(value, `events.${key}`);
    }
  }
}

function appendEffectLayerAnimations(layer: KeynoteHtmlLayer, ctx: SlideParseContext, sourcePath: string): void {
  const targetId = findRegisteredObjectId(ctx, layer);
  if (targetId) appendLayerAnimations(ctx, targetId, layer.animations, sourcePath);
  for (const [childIndex, child] of (layer.layers ?? layer.sublayers ?? layer.children ?? []).entries()) {
    appendEffectLayerAnimations(child, ctx, `${sourcePath}.layers[${childIndex}]`);
  }
}

function findRegisteredObjectId(ctx: SlideParseContext, layer: KeynoteHtmlLayer): string | undefined {
  for (const key of layerKeys(layer)) {
    const ids = ctx.objectIdsByLayerKey.get(key);
    if (ids?.length) return ids[0];
  }
  return undefined;
}

function layerStateToObjectState(
  state: KeynoteHtmlLayerState | undefined,
  parentBounds: { x: number; y: number; width: number; height: number }
): ObjectStateProperties {
  const width = finiteNumber(state?.width) ?? parentBounds.width;
  const height = finiteNumber(state?.height) ?? parentBounds.height;
  const anchor = pointValue(state?.anchorPoint, 0.5, 0.5);
  const position = pointValue(state?.position, parentBounds.width * anchor.x, parentBounds.height * anchor.y);
  const scale = pointValue(state?.scale, 1, 1);
  const rotation = finiteNumber(state?.rotation) ?? 0;
  return {
    visible: state?.hidden === true ? false : true,
    opacity: finiteNumber(state?.opacity) ?? 1,
    bounds: {
      x: position.x - width * anchor.x,
      y: position.y - height * anchor.y,
      width,
      height
    },
    transform: {
      scaleX: scale.x,
      scaleY: scale.y,
      rotateDeg: radiansToDegrees(rotation),
      origin: { x: anchor.x, y: anchor.y }
    }
  };
}

function ensureLayerAsset(ctx: SlideParseContext, textureId: string): Asset | undefined {
  const source = ctx.slideAssets[textureId];
  if (!source) {
    ctx.messages.push(`Layer ${textureId} referenced an asset missing from slide ${ctx.slideId}.`);
    return undefined;
  }
  const url = assetUrl(source);
  if (!url) {
    ctx.messages.push(`Layer ${textureId} referenced a Keynote asset without a URL.`);
    return undefined;
  }
  const normalizedUrl = normalizeAssetUri(ctx, url);
  const assetId = `keynote-html-asset-${ctx.slideId}-${sanitizeId(textureId)}`;
  const existing = ctx.deckAssets.get(assetId);
  if (existing) return existing;
  const kind = assetKind(source.kind ?? source.type, normalizedUrl, source.mimeType);
  const asset: Asset = {
    id: assetId,
    kind,
    uri: normalizedUrl,
    name: path.posix.basename(url),
    width: finiteNumber(source.width),
    height: finiteNumber(source.height),
    durationMs: source.durationMs !== undefined ? finiteNumber(source.durationMs) : secondsToMs(source.duration),
    mimeType: mimeTypeForUri(normalizedUrl, kind),
    metadata: {
      source: "keynote-html-export",
      sourceSlideId: ctx.slideId,
      sourceTextureId: textureId,
      sourceType: source.kind ?? source.type,
      sourceIndex: source.index,
      sourceUrl: url
    }
  };
  ctx.deckAssets.set(assetId, asset);
  return asset;
}

function assetUrl(source: KeynoteHtmlAsset): string | undefined {
  if (typeof source.url === "string") return source.url;
  return source.url?.web ?? source.url?.native ?? source.path ?? source.src;
}

function normalizeAssetUri(ctx: SlideParseContext, rawUrl: string): string {
  if (/^(?:https?:|file:|data:)/i.test(rawUrl)) return rawUrl;
  const normalized = rawUrl.replaceAll("\\", "/");
  const slidePrefix = `assets/${ctx.slideId}/`;
  if (normalized.startsWith(slidePrefix)) return normalized;
  return pathToPosix(path.join("assets", ctx.slideId, normalized));
}

function appendLayerAnimations(ctx: SlideParseContext, targetId: string, animations: KeynoteHtmlAnimation[] | undefined, sourcePath: string): void {
  const flattened = flattenAnimations(animations);
  for (const [index, animation] of flattened.entries()) {
    const tracks = animationToTracks(animation);
    if (tracks.length === 0) {
      if (animation.property && animation.property !== "zPosition") ctx.unsupportedProperties.add(animation.property);
      continue;
    }
    if (animation.property && tracks.some((track) => track.property.startsWith("custom.keynote."))) {
      ctx.unsupportedProperties.add(animation.property);
    }
    const beginMs = secondsToMs(animation.beginTime);
    const durationMs = Math.max(0, secondsToMs(animation.duration) ?? 0);
    ctx.events.push({
      id: `${targetId}-keynote-html-animation-${ctx.events.length + 1}`,
      kind: "keyframes",
      targetId,
      start: { type: "absolute", atMs: beginMs ?? 0 },
      durationMs,
      easing: easingName(animation.timingFunction),
      fill: fillMode(animation.fillMode),
      tracks,
      metadata: {
        source: "keynote-html-export",
        sourcePath,
        sourceAnimationIndex: index,
        sourceProperty: animation.property
      }
    });
    ctx.recoveredAnimationCount += 1;
  }
}

function flattenAnimations(animations: KeynoteHtmlAnimation[] | undefined): KeynoteHtmlAnimation[] {
  const out: KeynoteHtmlAnimation[] = [];
  const visit = (animation: KeynoteHtmlAnimation): void => {
    if (animation.property) out.push(animation);
    for (const child of animation.animations ?? []) visit(child);
  };
  for (const animation of animations ?? []) visit(animation);
  return out;
}

function animationToTracks(animation: KeynoteHtmlAnimation): KeyframeTrack[] {
  const rawValues = animation.values && animation.values.length > 0 ? animation.values : animation.from !== undefined && animation.to !== undefined ? [animation.from, animation.to] : [];
  if (rawValues.length === 0) return [];
  const keyTimes = animation.keyTimes?.length === rawValues.length ? animation.keyTimes : undefined;
  const keyframes = rawValues.map((value, index) => ({
    offset: clamp01(finiteNumber(keyTimes?.[index]) ?? index / Math.max(1, rawValues.length - 1)),
    value
  }));

  switch (animation.property) {
    case "opacity":
      return [numberTrack("opacity", keyframes)];
    case "hidden":
      return [
        {
          property: "visible",
          interpolation: "discrete",
          keyframes: keyframes.map(({ offset, value }) => ({ offset, value: !Boolean(unwrapAnimationValue(value)) }))
        }
      ];
    case "position":
    case "transform.translation":
      return pointTracks(animation.property === "position" ? ["bounds.x", "bounds.y"] : ["transform.translateX", "transform.translateY"], keyframes);
    case "scale":
      return pointTracks(["transform.scaleX", "transform.scaleY"], keyframes, 1);
    case "transform.scale.x":
      return [numberTrack("transform.scaleX", keyframes)];
    case "transform.scale.y":
      return [numberTrack("transform.scaleY", keyframes)];
    case "rotation":
    case "transform.rotation":
      return [numberTrack("transform.rotateDeg", keyframes, (value) => radiansToDegrees(Number(value) || 0))];
    case "bounds":
      return boundsTracks(keyframes);
    case "transform":
      return transformTracks(keyframes);
    default:
      return animation.property
        ? [
            {
              property: `custom.keynote.${animation.property}`,
              interpolation: "linear",
              keyframes: keyframes.map(({ offset, value }) => ({ offset, value: normalizeCustomAnimationValue(value) }))
            }
          ]
        : [];
  }
}

function numberTrack(
  property: string,
  keyframes: Array<{ offset: number; value: KeynoteHtmlAnimationValue }>,
  convert: (value: number | boolean | string) => number = (value) => Number(value) || 0
): KeyframeTrack {
  return {
    property,
    interpolation: "number",
    keyframes: keyframes.map(({ offset, value }) => ({ offset, value: convert(unwrapAnimationValue(value)) }))
  };
}

function pointTracks(
  properties: [string, string],
  keyframes: Array<{ offset: number; value: KeynoteHtmlAnimationValue }>,
  fallback = 0
): KeyframeTrack[] {
  return properties.map((property, axis) => ({
    property,
    interpolation: "number" as const,
    keyframes: keyframes.map(({ offset, value }) => {
      const point = animationPointValue(value, fallback, fallback);
      return { offset, value: axis === 0 ? point.x : point.y };
    })
  }));
}

function boundsTracks(keyframes: Array<{ offset: number; value: KeynoteHtmlAnimationValue }>): KeyframeTrack[] {
  const keys = ["bounds.x", "bounds.y", "bounds.width", "bounds.height"] as const;
  return keys.map((property) => ({
    property,
    interpolation: "number" as const,
    keyframes: keyframes.map(({ offset, value }) => {
      const bounds = animationRectValue(value);
      return { offset, value: bounds[property.slice("bounds.".length) as keyof typeof bounds] };
    })
  }));
}

function transformTracks(keyframes: Array<{ offset: number; value: KeynoteHtmlAnimationValue }>): KeyframeTrack[] {
  const scaleTracks = pointTracks(["transform.scaleX", "transform.scaleY"], keyframes, 1);
  const rotation = numberTrack("transform.rotateDeg", keyframes, (value) => radiansToDegrees(Number(value) || 0));
  return [...scaleTracks, rotation];
}

function normalizeCustomAnimationValue(value: KeynoteHtmlAnimationValue): number | boolean | string | JSONRecord {
  if (typeof value === "object" && value !== null) return value as JSONRecord;
  return value;
}

function unwrapAnimationValue(value: KeynoteHtmlAnimationValue): number | boolean | string {
  if (typeof value !== "object" || value === null) return value;
  if ("scalar" in value && value.scalar !== undefined) return value.scalar;
  if ("pointX" in value || "pointY" in value) return 0;
  if ("texture" in value && value.texture) return value.texture;
  return 0;
}

function animationPointValue(value: KeynoteHtmlAnimationValue, fallbackX: number, fallbackY: number): { x: number; y: number } {
  if (isPoint(value)) {
    return {
      x: finiteNumber(value.pointX ?? value.x) ?? fallbackX,
      y: finiteNumber(value.pointY ?? value.y) ?? fallbackY
    };
  }
  if (Array.isArray(value)) return { x: finiteNumber(value[0]) ?? fallbackX, y: finiteNumber(value[1]) ?? fallbackY };
  if (typeof value === "object" && value !== null) {
    const record = value as JSONRecord;
    return { x: finiteNumber(record.x) ?? fallbackX, y: finiteNumber(record.y) ?? fallbackY };
  }
  const scalar = finiteNumber(value) ?? fallbackX;
  return { x: scalar, y: scalar };
}

function animationRectValue(value: KeynoteHtmlAnimationValue): { x: number; y: number; width: number; height: number } {
  if (Array.isArray(value)) {
    return {
      x: finiteNumber(value[0]) ?? 0,
      y: finiteNumber(value[1]) ?? 0,
      width: finiteNumber(value[2]) ?? 0,
      height: finiteNumber(value[3]) ?? 0
    };
  }
  if (typeof value === "object" && value !== null) {
    const record = value as JSONRecord;
    return {
      x: finiteNumber(record.x) ?? 0,
      y: finiteNumber(record.y) ?? 0,
      width: finiteNumber(record.width) ?? 0,
      height: finiteNumber(record.height) ?? 0
    };
  }
  return { x: 0, y: 0, width: 0, height: 0 };
}

function isPoint(value: KeynoteHtmlAnimationValue): value is { pointX?: number; pointY?: number; x?: number; y?: number } {
  return typeof value === "object" && value !== null && ("pointX" in value || "pointY" in value || "x" in value || "y" in value);
}

function contentsRectToCrop(contentsRect: KeynoteHtmlLayerState["contentsRect"] | undefined): { x: number; y: number; width: number; height: number; unit: "ratio" } | undefined {
  if (!contentsRect) return undefined;
  const x = Array.isArray(contentsRect) ? finiteNumber(contentsRect[0]) ?? 0 : finiteNumber(contentsRect.x) ?? 0;
  const y = Array.isArray(contentsRect) ? finiteNumber(contentsRect[1]) ?? 0 : finiteNumber(contentsRect.y) ?? 0;
  const width = Array.isArray(contentsRect) ? finiteNumber(contentsRect[2]) ?? 1 : finiteNumber(contentsRect.width) ?? 1;
  const height = Array.isArray(contentsRect) ? finiteNumber(contentsRect[3]) ?? 1 : finiteNumber(contentsRect.height) ?? 1;
  if (x === 0 && y === 0 && width === 1 && height === 1) return undefined;
  return { x, y, width, height, unit: "ratio" };
}

function layerMetadata(ctx: SlideParseContext, layer: KeynoteHtmlLayer, sourcePath: string): JSONRecord {
  return {
    source: "keynote-html-export",
    sourceSlideId: ctx.slideId,
    sourcePath,
    sourceObjectID: layer.objectID,
    sourceTextureId: layer.texture,
    texturedRectangle: layer.texturedRectangle
  };
}

function inferSlideDurationMs(events: AnimationEvent[], autoplayTransitionDelaySeconds: number): number {
  const animationEnd = events.reduce((max, event) => {
    const start = event.start?.type === "absolute" ? event.start.atMs : 0;
    return Math.max(max, start + (event.durationMs ?? 0));
  }, 0);
  return Math.max(animationEnd, autoplayTransitionDelaySeconds * 1000, 1);
}

function assetKind(type: string | undefined, uri: string, mimeType?: string): Asset["kind"] {
  const descriptor = `${type ?? ""} ${mimeType ?? ""}`.toLowerCase();
  if (descriptor.includes("video") || /\.(?:mp4|m4v|mov)$/i.test(uri)) return "video";
  if (descriptor.includes("audio") || /\.(?:m4a|mp3|wav|aac)$/i.test(uri)) return "audio";
  if (descriptor.includes("texture") || descriptor.includes("image") || /\.(?:png|jpe?g|gif|webp|svg|pdf)$/i.test(uri)) return "image";
  return "other";
}

function mimeTypeForUri(uri: string, kind: Asset["kind"]): string | undefined {
  const ext = path.posix.extname(uri).toLowerCase();
  if (ext === ".png") return "image/png";
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".gif") return "image/gif";
  if (ext === ".webp") return "image/webp";
  if (ext === ".svg") return "image/svg+xml";
  if (ext === ".pdf") return "application/pdf";
  if (ext === ".mp4" || ext === ".m4v") return "video/mp4";
  if (ext === ".mov") return "video/quicktime";
  if (kind === "image") return "image/*";
  if (kind === "video") return "video/*";
  return undefined;
}

function easingName(value: string | undefined): Easing | undefined {
  switch (value) {
    case "EaseIn":
      return "easeIn";
    case "EaseOut":
      return "easeOut";
    case "EaseInEaseOut":
      return "easeInOut";
    case "Linear":
      return "linear";
    default:
      return value ? { type: "custom", name: value } : undefined;
  }
}

function fillMode(value: string | undefined): AnimationEvent["fill"] {
  if (value === "forwards") return "forwards";
  if (value === "backwards") return "backwards";
  if (value === "both") return "both";
  return "none";
}

async function readJson<T>(filePath: string): Promise<T> {
  const source = await readFile(filePath, "utf8");
  return JSON.parse(source) as T;
}

async function fileExists(filePath: string): Promise<boolean> {
  try {
    return (await stat(filePath)).isFile();
  } catch {
    return false;
  }
}

function stableObjectId(slideIndex: number, sourcePath: string, objectId: string | undefined): string {
  const source = objectId ? `${objectId}-${sourcePath}` : sourcePath;
  return `keynote-html-s${slideIndex + 1}-${sanitizeId(source)}`;
}

function pointValue(value: unknown, fallbackX: number, fallbackY: number): { x: number; y: number } {
  if (Array.isArray(value)) return { x: finiteNumber(value[0]) ?? fallbackX, y: finiteNumber(value[1]) ?? fallbackY };
  if (typeof value === "object" && value !== null) {
    const record = value as JSONRecord;
    return {
      x: finiteNumber(record.pointX) ?? finiteNumber(record.x) ?? fallbackX,
      y: finiteNumber(record.pointY) ?? finiteNumber(record.y) ?? fallbackY
    };
  }
  const scalar = finiteNumber(value);
  return scalar === undefined ? { x: fallbackX, y: fallbackY } : { x: scalar, y: scalar };
}

function sanitizeId(value: string): string {
  const cleaned = value.replace(/[^a-zA-Z0-9_-]+/g, "-").replace(/^-+|-+$/g, "");
  return cleaned || "layer";
}

function finiteNumber(value: unknown): number | undefined {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : undefined;
}

function secondsToMs(value: unknown): number | undefined {
  const numeric = finiteNumber(value);
  return numeric === undefined ? undefined : numeric * 1000;
}

function radiansToDegrees(value: number): number {
  return Math.abs(value) <= Math.PI * 2 + 0.001 ? (value * 180) / Math.PI : value;
}

function clamp01(value: number): number {
  return Math.max(0, Math.min(1, value));
}

function pathToPosix(value: string): string {
  return value.split(path.sep).join(path.posix.sep);
}

export async function isKeynoteHtmlExportDir(exportDir: string): Promise<boolean> {
  return fileExists(path.join(path.resolve(exportDir), "assets", "header.json"));
}
