import { readFile, writeFile } from "node:fs/promises";
import { deflateRawSync, inflateRawSync } from "node:zlib";

import {
  IR_VERSION,
  type AnimationEvent,
  type ConversionReport,
  type DeckIR,
  type DegradedFeature,
  type IRObject,
  type Slide,
  type TimingDependencyEdge,
  type TimingNode,
  type TimelineTrigger,
  type UnsupportedFeature
} from "../ir/index.ts";

const EMU_PER_INCH = 914400;
const WIDE_LAYOUT = { width: 13.333333, height: 7.5 };
const OPENXML_ANGLE_UNITS_PER_DEGREE = 60000;

export async function parsePptxToIr(filePath: string): Promise<DeckIR> {
  const data = await readFile(filePath);
  const entries = readZipEntries(data);
  const presentationXml = getTextEntry(entries, "ppt/presentation.xml");
  const slidePaths = presentationXml ? readSlidePaths(entries, presentationXml) : [];
  const size = presentationXml ? readSlideSize(presentationXml) : { width: 1280, height: 720 };
  const slideResults = slidePaths.map((slidePath, index) => parseSlideXml(getTextEntry(entries, slidePath) ?? "", index, size));
  const slides = slideResults.map((result) => result.slide);
  const parsedSlides = slides.length
    ? slides
    : [
        {
          id: "slide-1",
          index: 0,
          name: "Import placeholder",
          background: { type: "solid" as const, color: "#ffffff" },
          objects: [
            {
              id: "import-placeholder",
              type: "text" as const,
              name: "Import placeholder",
              bounds: { x: 96, y: 96, width: 960, height: 96 },
              opacity: 1,
              text: {
                plainText: "PPTX import placeholder",
                runs: [
                  {
                    text: "PPTX import placeholder",
                    style: { fontFamily: "Arial", fontSize: 36, color: "#111827" }
                  }
                ]
              }
            }
          ],
          timeline: { durationMs: 2500, events: [], dependencyGraph: { edges: [] } }
        }
      ];
  const unsupportedFeatures = slideResults.flatMap((result) => result.unsupportedFeatures);
  const degradedFeatures = slideResults.flatMap((result) => result.degradedFeatures);
  const animationCount = parsedSlides.reduce((total, slide) => total + (slide.timeline?.events.length ?? 0), 0);

  const report: ConversionReport = {
    source: { kind: "pptx", uri: filePath },
    status: slidePaths.length ? "partial" : "failed",
    generatedAt: new Date().toISOString(),
    tool: "keymorph-pptx-parser-mvp",
    messages: [
      {
        severity: slidePaths.length ? "info" : "warning",
        code: slidePaths.length ? "pptx-static-import" : "pptx-import-placeholder",
        message: slidePaths.length
          ? "Extracted static slide objects and supported timing animations from PresentationML."
          : "The PPTX package was readable, but no slide parts were discovered."
      }
    ],
    unsupportedFeatures,
    degradedFeatures,
    uncertainMappings: [],
    statistics: {
      slideCount: parsedSlides.length,
      objectCount: parsedSlides.reduce((total, slide) => total + slide.objects.length, 0),
      animationCount,
      unsupportedFeatureCount: unsupportedFeatures.length,
      degradedFeatureCount: degradedFeatures.length
    },
    metadata: {
      byteLength: data.byteLength
    }
  };

  return {
    irVersion: IR_VERSION,
    metadata: { title: "Imported PPTX", sourceApplication: "PowerPoint" },
    deck: {
      id: "pptx-import",
      title: "Imported PPTX",
      size: { width: size.width, height: size.height, unit: "px" },
      slides: parsedSlides
    },
    conversion: report
  };
}

export async function exportIrToPptx(deck: DeckIR, outputPath: string): Promise<void> {
  const files = createPptxFiles(deck);
  await writeFile(outputPath, createZip(files));
}

function createPptxFiles(deck: DeckIR): Map<string, string | Uint8Array> {
  const files = new Map<string, string | Uint8Array>();
  const slideCount = deck.deck.slides.length;

  files.set("[Content_Types].xml", contentTypes(slideCount));
  files.set("_rels/.rels", rootRels());
  files.set("docProps/core.xml", coreProps(deck));
  files.set("docProps/app.xml", appProps(slideCount));
  files.set("ppt/presentation.xml", presentationXml(deck));
  files.set("ppt/_rels/presentation.xml.rels", presentationRels(slideCount));
  files.set("ppt/theme/theme1.xml", themeXml());
  files.set("ppt/slideMasters/slideMaster1.xml", slideMasterXml());
  files.set("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRels());
  files.set("ppt/slideLayouts/slideLayout1.xml", slideLayoutXml());
  files.set("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayoutRels());

  deck.deck.slides.forEach((slide, index) => {
    const slideNumber = index + 1;
    files.set(`ppt/slides/slide${slideNumber}.xml`, slideXml(deck, slide));
    files.set(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, slideRels());
  });

  return files;
}

interface ParsedSlide {
  slide: Slide;
  unsupportedFeatures: UnsupportedFeature[];
  degradedFeatures: DegradedFeature[];
}

function parseSlideXml(source: string, index: number, slideSize: { width: number; height: number }): ParsedSlide {
  const objects: IRObject[] = [];
  const shapeIdToObjectId = new Map<string, string>();
  const shapeSources = source.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) ?? [];

  shapeSources.forEach((shapeXml, objectIndex) => {
    const id = `slide-${index + 1}-object-${objectIndex + 1}`;
    const sourceShapeId = readShapeSourceId(shapeXml);
    const text = readShapeText(shapeXml);
    const bounds = readShapeBounds(shapeXml, { x: 96 + objectIndex * 28, y: 96 + objectIndex * 28, width: 400, height: 100 });
    if (text) {
      objects.push({
        id,
        type: "text",
        name: readShapeName(shapeXml) ?? "Text",
        bounds,
        opacity: 1,
        metadata: sourceShapeId ? { pptxShapeId: sourceShapeId } : undefined,
        text: {
          plainText: text,
          runs: [{ text, style: { fontFamily: "Arial", fontSize: 28, color: "#111827" } }]
        }
      });
    } else if (!shapeXml.includes("<p:nvGrpSpPr")) {
      objects.push({
        id,
        type: "shape",
        name: readShapeName(shapeXml) ?? "Shape",
        shape: readShapePreset(shapeXml),
        bounds,
        opacity: 1,
        metadata: sourceShapeId ? { pptxShapeId: sourceShapeId } : undefined,
        style: {
          fill: { type: "solid", color: readShapeFill(shapeXml) ?? "#e2e8f0" },
          stroke: { color: "#94a3b8", width: 1 }
        }
      });
    }
    if (sourceShapeId && objects.some((object) => object.id === id)) {
      shapeIdToObjectId.set(sourceShapeId, id);
    }
  });

  const timing = parseSlideTiming(source, index, slideSize, shapeIdToObjectId);
  const maxEventEnd = timing.events.reduce((max, event) => Math.max(max, eventStartMs(event) + (event.durationMs ?? 0)), 0);

  return {
    slide: {
      id: `slide-${index + 1}`,
      index,
      name: `Slide ${index + 1}`,
      background: { type: "solid", color: readSlideBackground(source) ?? "#ffffff" },
      objects,
      timeline: {
        durationMs: Math.max(2500, Math.ceil(maxEventEnd)),
        triggers: timing.triggers.length ? timing.triggers : undefined,
        events: timing.events,
        dependencyGraph: { nodes: timing.dependencyNodes.length ? timing.dependencyNodes : undefined, edges: timing.dependencyEdges }
      }
    },
    unsupportedFeatures: timing.unsupportedFeatures,
    degradedFeatures: timing.degradedFeatures
  };
}

function readSlidePaths(entries: Map<string, Uint8Array>, presentationXml: string): string[] {
  const rels = parseRelationships(getTextEntry(entries, "ppt/_rels/presentation.xml.rels") ?? "");
  const slideRefs = Array.from(presentationXml.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)).map((match) => match[1]);
  return slideRefs
    .map((relationshipId) => rels.get(relationshipId))
    .filter((target): target is string => Boolean(target))
    .map((target) => normalizePartPath(`ppt/${target}`));
}

function readSlideSize(presentationXml: string): { width: number; height: number } {
  const match = presentationXml.match(/<p:sldSz\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
  if (!match) return { width: 1280, height: 720 };
  return {
    width: Math.round((Number(match[1]) / EMU_PER_INCH) * 96),
    height: Math.round((Number(match[2]) / EMU_PER_INCH) * 96)
  };
}

function parseRelationships(source: string): Map<string, string> {
  const relationships = new Map<string, string>();
  for (const match of source.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relationships.set(match[1], match[2]);
  }
  return relationships;
}

function readShapeText(source: string): string {
  return Array.from(source.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g))
    .map((match) => unescapeXml(match[1]))
    .join("");
}

function readShapeName(source: string): string | undefined {
  const match = source.match(/<p:cNvPr\b[^>]*name="([^"]*)"/);
  return match ? unescapeXml(match[1]) : undefined;
}

function readShapeSourceId(source: string): string | undefined {
  const match = source.match(/<p:cNvPr\b[^>]*id="([^"]*)"/);
  return match ? unescapeXml(match[1]) : undefined;
}

function readShapeBounds(source: string, fallback: { x: number; y: number; width: number; height: number }) {
  const offset = source.match(/<a:off\b[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/);
  const extent = source.match(/<a:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
  if (!offset || !extent) return fallback;
  return {
    x: emuToWidePx(Number(offset[1]), WIDE_LAYOUT.width),
    y: emuToWidePx(Number(offset[2]), WIDE_LAYOUT.height),
    width: emuToWidePx(Number(extent[1]), WIDE_LAYOUT.width),
    height: emuToWidePx(Number(extent[2]), WIDE_LAYOUT.height)
  };
}

function readShapePreset(source: string): "rect" | "roundRect" | "ellipse" | "triangle" {
  const match = source.match(/<a:prstGeom\b[^>]*prst="([^"]+)"/);
  if (match?.[1] === "ellipse") return "ellipse";
  if (match?.[1] === "roundRect") return "roundRect";
  if (match?.[1] === "triangle") return "triangle";
  return "rect";
}

function readShapeFill(source: string): string | undefined {
  const match = source.match(/<a:solidFill>[\s\S]*?<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
  return match ? `#${match[1]}` : undefined;
}

function readSlideBackground(source: string): string | undefined {
  const match = source.match(/<p:bg>[\s\S]*?<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
  return match ? `#${match[1]}` : undefined;
}

interface TimingParseResult {
  events: AnimationEvent[];
  triggers: TimelineTrigger[];
  dependencyNodes: TimingNode[];
  dependencyEdges: TimingDependencyEdge[];
  unsupportedFeatures: UnsupportedFeature[];
  degradedFeatures: DegradedFeature[];
}

interface ParsedTimingNode {
  tag: SupportedTimingTag;
  xml: string;
  range: XmlElementRange;
  path: string[];
}

type SupportedTimingTag = "animEffect" | "set" | "animMotion" | "animScale" | "animRot";

interface TimingSequenceContext {
  kind: "absolute" | "withPrevious" | "afterPrevious" | "onClick";
  delayMs: number;
  triggerId?: string;
  triggerTargetId?: string;
  clickIndex?: number;
}

interface ParsedTimingEvent {
  event: AnimationEvent;
  timingContext: TimingSequenceContext;
}

interface TriggerBuildSequence {
  targetIds: string[];
  clickIndex: number;
}

interface XmlElementRange {
  localName: string;
  start: number;
  openEnd: number;
  closeStart: number;
  end: number;
  parent?: XmlElementRange;
  children: XmlElementRange[];
}

function parseSlideTiming(
  source: string,
  slideIndex: number,
  slideSize: { width: number; height: number },
  shapeIdToObjectId: Map<string, string>
): TimingParseResult {
  const timingXml = extractXmlElements(source, "timing")[0];
  const result: TimingParseResult = {
    events: [],
    triggers: [],
    dependencyNodes: [],
    dependencyEdges: [],
    unsupportedFeatures: [],
    degradedFeatures: []
  };
  if (!timingXml) return result;

  const timingTree = parseXmlElementRanges(timingXml);
  const timingNodes = extractSupportedTimingNodes(timingXml, timingTree);
  const parsedEvents: ParsedTimingEvent[] = [];

  for (const node of timingNodes) {
    const event = parseSupportedTimingEvent(node, slideIndex, parsedEvents.length, slideSize, shapeIdToObjectId, result.unsupportedFeatures);
    if (!event) continue;
    const timingContext = readTimingSequenceContext(
      timingXml,
      timingTree,
      node,
      slideIndex,
      parsedEvents.length + 1,
      shapeIdToObjectId,
      result.degradedFeatures
    );
    applyTimingContext(event, timingContext);
    parsedEvents.push({ event, timingContext });
    result.events.push(event);
  }

  for (const tag of ["anim", "animClr", "cmd"]) {
    for (const _node of extractXmlElements(timingXml, tag)) {
      result.unsupportedFeatures.push({
        code: `presentationml-${tag}`,
        severity: "warning",
        area: "animation",
        description: `PowerPoint p:${tag} animation nodes are not mapped to the IR timeline yet.`,
        sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/p:${tag}`,
        fallback: "Preserve the static object and omit this animation."
      });
    }
  }

  if (/<p:(?:nextCondLst|endCondLst)\b/.test(timingXml)) {
    result.degradedFeatures.push({
      code: "presentationml-timing-conditions",
      severity: "warning",
      area: "animation",
      description: "PowerPoint sequence next/end conditions are not fully modeled by the IR timeline.",
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing`,
      fallback: "Import supported effect order and explicit start conditions; omit next/end condition behavior."
    });
  }

  const triggerBuildSequences = parseTriggerBuildSequences(timingXml, shapeIdToObjectId);
  addTimingDependencies(slideIndex, parsedEvents, triggerBuildSequences, result);
  const buildList = parseBuildListSequencing(timingXml, slideIndex, shapeIdToObjectId, result.events);
  result.degradedFeatures.push(...buildList.degradedFeatures);
  addDependencyEdges(result.dependencyEdges, buildList.dependencyEdges);
  return result;
}

function parseSupportedTimingEvent(
  node: ParsedTimingNode,
  slideIndex: number,
  eventIndex: number,
  slideSize: { width: number; height: number },
  shapeIdToObjectId: Map<string, string>,
  unsupportedFeatures: UnsupportedFeature[]
): AnimationEvent | undefined {
  if (node.tag === "animEffect") {
    return parseFadeEffect(node.xml, slideIndex, eventIndex, shapeIdToObjectId, unsupportedFeatures);
  }
  if (node.tag === "set") {
    return parseVisibilitySet(node.xml, slideIndex, eventIndex, shapeIdToObjectId, unsupportedFeatures);
  }
  if (node.tag === "animMotion") {
    return parseMotionEffect(node.xml, slideIndex, eventIndex, slideSize, shapeIdToObjectId, unsupportedFeatures);
  }
  if (node.tag === "animScale") {
    return parseScaleEffect(node.xml, slideIndex, eventIndex, shapeIdToObjectId, unsupportedFeatures);
  }
  return parseRotationEffect(node.xml, slideIndex, eventIndex, shapeIdToObjectId, unsupportedFeatures);
}

function applyTimingContext(event: AnimationEvent, context: TimingSequenceContext): void {
  const childDelayMs = eventStartMs(event);
  const offsetMs = context.delayMs + childDelayMs;
  if (context.kind === "withPrevious") {
    event.start = { type: "withPrevious", offsetMs };
  } else if (context.kind === "afterPrevious") {
    event.start = { type: "afterPrevious", offsetMs };
  } else if (context.kind === "onClick" && context.triggerId) {
    event.start = { type: "trigger", triggerId: context.triggerId, offsetMs };
  } else {
    event.start = { type: "absolute", atMs: offsetMs };
  }
}

function parseFadeEffect(
  node: string,
  slideIndex: number,
  eventIndex: number,
  shapeIdToObjectId: Map<string, string>,
  unsupportedFeatures: UnsupportedFeature[]
): AnimationEvent | undefined {
  const attrs = readXmlAttributes(node, "animEffect");
  const filter = String(attrs.get("filter") ?? "").toLowerCase();
  if (!filter.includes("fade")) {
    unsupportedFeatures.push({
      code: "presentationml-animation-effect",
      severity: "warning",
      area: "animation",
      description: `PowerPoint animEffect filter "${attrs.get("filter") ?? "unknown"}" is not mapped to IR keyframes.`,
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/p:animEffect`,
      fallback: "Preserve the static object and omit this effect."
    });
    return undefined;
  }

  const targetId = resolveTimingTarget(node, slideIndex, shapeIdToObjectId, unsupportedFeatures);
  if (!targetId) return undefined;

  const transition = String(attrs.get("transition") ?? "in").toLowerCase();
  const fadesOut = transition === "out";
  return {
    id: `slide-${slideIndex + 1}-anim-${eventIndex + 1}`,
    kind: "keyframes",
    label: fadesOut ? "Fade out" : "Fade in",
    targetId,
    start: { type: "absolute", atMs: readTimingDelay(node) },
    durationMs: readTimingDuration(node, 500),
    easing: "linear",
    fill: "both",
    tracks: [
      {
        property: "opacity",
        keyframes: [
          { offset: 0, value: fadesOut ? 1 : 0 },
          { offset: 1, value: fadesOut ? 0 : 1 }
        ]
      }
    ],
    metadata: { pptxEffect: "fade", pptxTransition: fadesOut ? "out" : "in" }
  };
}

function parseVisibilitySet(
  node: string,
  slideIndex: number,
  eventIndex: number,
  shapeIdToObjectId: Map<string, string>,
  unsupportedFeatures: UnsupportedFeature[]
): AnimationEvent | undefined {
  const attrNames = Array.from(node.matchAll(/<p:attrName>([\s\S]*?)<\/p:attrName>/g)).map((match) => unescapeXml(match[1]));
  if (!attrNames.some((name) => name.toLowerCase() === "style.visibility")) {
    unsupportedFeatures.push({
      code: "presentationml-set-animation",
      severity: "warning",
      area: "animation",
      description: `PowerPoint set animation attributes "${attrNames.join(", ") || "unknown"}" are not mapped to IR visibility events.`,
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/p:set`,
      fallback: "Preserve the static object and omit this set animation."
    });
    return undefined;
  }

  const targetId = resolveTimingTarget(node, slideIndex, shapeIdToObjectId, unsupportedFeatures);
  if (!targetId) return undefined;

  const visible = readVisibilityValue(node);
  if (visible === undefined) {
    unsupportedFeatures.push({
      code: "presentationml-visibility-value",
      severity: "warning",
      area: "animation",
      description: "PowerPoint visibility set animation has an unsupported target value.",
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/p:set/p:to`,
      fallback: "Preserve the static object and omit this visibility toggle."
    });
    return undefined;
  }

  return {
    id: `slide-${slideIndex + 1}-anim-${eventIndex + 1}`,
    kind: "visibility",
    label: visible ? "Appear" : "Disappear",
    targetId,
    start: { type: "absolute", atMs: readTimingDelay(node) },
    durationMs: 0,
    fill: "both",
    visible,
    metadata: { pptxEffect: "visibility" }
  };
}

function parseMotionEffect(
  node: string,
  slideIndex: number,
  eventIndex: number,
  slideSize: { width: number; height: number },
  shapeIdToObjectId: Map<string, string>,
  unsupportedFeatures: UnsupportedFeature[]
): AnimationEvent | undefined {
  const targetId = resolveTimingTarget(node, slideIndex, shapeIdToObjectId, unsupportedFeatures);
  if (!targetId) return undefined;

  const attrs = readXmlAttributes(node, "animMotion");
  const delta =
    parseMotionPath(attrs.get("path"), slideSize) ??
    parseMotionBy(attrs.get("by"), slideSize) ??
    parseMotionFromTo(attrs.get("from"), attrs.get("to"), slideSize);
  if (!delta) {
    unsupportedFeatures.push({
      code: "presentationml-motion-path",
      severity: "warning",
      area: "animation",
      description: "PowerPoint motion path is not a simple line/by/from-to path that can be mapped to IR translate keyframes.",
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/p:animMotion`,
      fallback: "Preserve the static object and omit this motion path."
    });
    return undefined;
  }

  const tracks = [];
  if (delta.x !== 0) {
    tracks.push({
      property: "transform.translateX",
      keyframes: [
        { offset: 0, value: 0 },
        { offset: 1, value: delta.x }
      ]
    });
  }
  if (delta.y !== 0) {
    tracks.push({
      property: "transform.translateY",
      keyframes: [
        { offset: 0, value: 0 },
        { offset: 1, value: delta.y }
      ]
    });
  }
  if (tracks.length === 0) return undefined;

  return {
    id: `slide-${slideIndex + 1}-anim-${eventIndex + 1}`,
    kind: "keyframes",
    label: "Motion path",
    targetId,
    start: { type: "absolute", atMs: readTimingDelay(node) },
    durationMs: readTimingDuration(node, 500),
    easing: "linear",
    fill: "both",
    tracks,
    metadata: { pptxEffect: "motionPath" }
  };
}

function parseScaleEffect(
  node: string,
  slideIndex: number,
  eventIndex: number,
  shapeIdToObjectId: Map<string, string>,
  unsupportedFeatures: UnsupportedFeature[]
): AnimationEvent | undefined {
  const targetId = resolveTimingTarget(node, slideIndex, shapeIdToObjectId, unsupportedFeatures);
  if (!targetId) return undefined;

  const attrs = readXmlAttributes(node, "animScale");
  const scale =
    parseScalePair(attrs.get("by")) ??
    parseScaleFromTo(attrs.get("from"), attrs.get("to")) ??
    parseScaleValuesFromChildNodes(node);
  if (!scale) {
    unsupportedFeatures.push({
      code: "presentationml-scale-animation",
      severity: "warning",
      area: "animation",
      description: "PowerPoint scale animation is not a simple by/from-to pair that can be mapped to IR scale keyframes.",
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/p:animScale`,
      fallback: "Preserve the static object and omit this scale animation."
    });
    return undefined;
  }

  const tracks = [];
  if (scale.fromX !== scale.toX) {
    tracks.push({
      property: "transform.scaleX",
      keyframes: [
        { offset: 0, value: scale.fromX },
        { offset: 1, value: scale.toX }
      ]
    });
  }
  if (scale.fromY !== scale.toY) {
    tracks.push({
      property: "transform.scaleY",
      keyframes: [
        { offset: 0, value: scale.fromY },
        { offset: 1, value: scale.toY }
      ]
    });
  }
  if (tracks.length === 0) return undefined;

  return {
    id: `slide-${slideIndex + 1}-anim-${eventIndex + 1}`,
    kind: "keyframes",
    label: "Scale",
    targetId,
    start: { type: "absolute", atMs: readTimingDelay(node) },
    durationMs: readTimingDuration(node, 500),
    easing: "linear",
    fill: "both",
    tracks,
    metadata: { pptxEffect: "scale" }
  };
}

function parseRotationEffect(
  node: string,
  slideIndex: number,
  eventIndex: number,
  shapeIdToObjectId: Map<string, string>,
  unsupportedFeatures: UnsupportedFeature[]
): AnimationEvent | undefined {
  const targetId = resolveTimingTarget(node, slideIndex, shapeIdToObjectId, unsupportedFeatures);
  if (!targetId) return undefined;

  const attrs = readXmlAttributes(node, "animRot");
  const rotation = parseRotationValues(attrs.get("by"), attrs.get("from"), attrs.get("to")) ?? parseRotationValuesFromChildNodes(node);
  if (!rotation || rotation.from === rotation.to) {
    unsupportedFeatures.push({
      code: "presentationml-rotation-animation",
      severity: "warning",
      area: "animation",
      description: "PowerPoint rotation animation is not a simple by/from-to value that can be mapped to IR rotation keyframes.",
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/p:animRot`,
      fallback: "Preserve the static object and omit this rotation animation."
    });
    return undefined;
  }

  return {
    id: `slide-${slideIndex + 1}-anim-${eventIndex + 1}`,
    kind: "keyframes",
    label: "Rotate",
    targetId,
    start: { type: "absolute", atMs: readTimingDelay(node) },
    durationMs: readTimingDuration(node, 500),
    easing: "linear",
    fill: "both",
    tracks: [
      {
        property: "transform.rotateDeg",
        keyframes: [
          { offset: 0, value: rotation.from },
          { offset: 1, value: rotation.to }
        ]
      }
    ],
    metadata: { pptxEffect: "rotation" }
  };
}

function parseBuildListSequencing(
  timingXml: string,
  slideIndex: number,
  shapeIdToObjectId: Map<string, string>,
  events: AnimationEvent[]
): { dependencyEdges: TimingDependencyEdge[]; degradedFeatures: DegradedFeature[] } {
  const buildListXml = extractXmlElements(timingXml, "bldLst")[0];
  if (!buildListXml) return { dependencyEdges: [], degradedFeatures: [] };

  const objectIds = uniqueStrings(Array.from(buildListXml.matchAll(/spid="([^"]+)"/g))
    .map((match) => shapeIdToObjectId.get(match[1]))
    .filter((targetId): targetId is string => Boolean(targetId)));
  const sequencedEvents = objectIds
    .map((targetId) => events.find((event) => "targetId" in event && event.targetId === targetId))
    .filter((event): event is AnimationEvent => Boolean(event));
  const dependencyEdges: TimingDependencyEdge[] = [];

  for (let index = 1; index < sequencedEvents.length; index += 1) {
    dependencyEdges.push({
      id: `slide-${slideIndex + 1}-build-edge-${index}`,
      from: sequencedEvents[index - 1].id,
      to: sequencedEvents[index].id,
      relation: "after"
    });
  }

  return {
    dependencyEdges,
    degradedFeatures: [
      {
        code: "presentationml-build-list",
        severity: "warning",
        area: "animation",
        description:
          dependencyEdges.length > 0
            ? "PowerPoint build-list sequencing was reduced to simple after-dependencies between supported object timing events."
            : "PowerPoint build-list sequencing was detected, but no supported object timing events could be associated with it.",
        sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/p:bldLst`,
        fallback: "Import supported object-level effects and preserve build order metadata where possible."
      }
    ]
  };
}

function extractSupportedTimingNodes(timingXml: string, ranges: XmlElementRange[]): ParsedTimingNode[] {
  const supported = new Set<SupportedTimingTag>(["animEffect", "set", "animMotion", "animScale", "animRot"]);
  return ranges
    .filter((range): range is XmlElementRange & { localName: SupportedTimingTag } => supported.has(range.localName as SupportedTimingTag))
    .sort((a, b) => a.start - b.start)
    .map((range) => ({
      tag: range.localName,
      xml: timingXml.slice(range.start, range.end),
      range,
      path: xmlAncestorPath(range)
    }));
}

function readTimingSequenceContext(
  timingXml: string,
  ranges: XmlElementRange[],
  node: ParsedTimingNode,
  slideIndex: number,
  eventOrdinal: number,
  shapeIdToObjectId: Map<string, string>,
  degradedFeatures: DegradedFeature[]
): TimingSequenceContext {
  const ancestors = xmlAncestors(node.range);
  const timingContainer = ancestors.find((ancestor) => ancestor.localName === "par" || ancestor.localName === "seq");
  const containerCTn = timingContainer ? immediateChildRange(ranges, timingContainer, "cTn") : undefined;
  const containerCTnXml = containerCTn ? timingXml.slice(containerCTn.start, containerCTn.end) : undefined;
  const attrs = containerCTnXml ? readXmlAttributes(containerCTnXml, "cTn") : new Map();
  const nodeType = String(attrs.get("nodeType") ?? "").toLowerCase();
  const condSource = containerCTn ? immediateChildXml(timingXml, ranges, containerCTn, "stCondLst") : undefined;
  const cond = readStartCondition(condSource ?? node.xml);
  const delayMs = parseTimingMs(cond.delay, 0);

  if (cond.delay === "indefinite" || cond.event === "onClick" || nodeType === "clickeffect" || nodeType === "clickpar") {
    const triggerId = `slide-${slideIndex + 1}-click-${eventOrdinal}`;
    const triggerTargetId = cond.targetSpid ? shapeIdToObjectId.get(cond.targetSpid) : undefined;
    return { kind: "onClick", delayMs, triggerId, triggerTargetId, clickIndex: eventOrdinal };
  }

  if (nodeType === "aftereffect" || cond.event === "onEnd") {
    return { kind: "afterPrevious", delayMs };
  }

  if (nodeType === "witheffect" || cond.event === "begin") {
    return { kind: "withPrevious", delayMs };
  }

  if (cond.event && !["begin", "onEnd", "onClick"].includes(cond.event)) {
    degradedFeatures.push({
      code: "presentationml-trigger-timing",
      severity: "warning",
      area: "animation",
      description: `PowerPoint timing trigger event "${cond.event}" is not represented exactly in the IR timeline.`,
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/${node.path.join("/")}`,
      fallback: "Import the supported effect with its delay as absolute timing."
    });
  }

  if (cond.delay === "indefinite") {
    degradedFeatures.push({
      code: "presentationml-trigger-timing",
      severity: "warning",
      area: "animation",
      description: "PowerPoint indefinite trigger timing could not be associated with a click sequence.",
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing/${node.path.join("/")}`,
      fallback: "Import the supported effect at the beginning of the slide timeline."
    });
  }

  return { kind: "absolute", delayMs };
}

function readStartCondition(source: string): { delay?: string; event?: string; targetSpid?: string } {
  const stCondList = extractXmlElements(source, "stCondLst")[0];
  const condSource = stCondList ?? source;
  const cond = extractXmlElements(condSource, "cond")[0] ?? condSource.match(/<p:cond\b[^>]*\/>/)?.[0];
  if (!cond) return {};
  const attrs = readXmlAttributes(cond, "cond");
  return {
    delay: attrs.get("delay"),
    event: attrs.get("evt"),
    targetSpid: cond.match(/<p:spTgt\b[^>]*spid="([^"]+)"/)?.[1]
  };
}

function addTimingDependencies(
  slideIndex: number,
  parsedEvents: ParsedTimingEvent[],
  triggerBuildSequences: TriggerBuildSequence[],
  result: TimingParseResult
): void {
  const triggers = new Map<string, TimelineTrigger>();
  const dependencyNodes: TimingNode[] = [];
  const dependencyEdges: TimingDependencyEdge[] = [];

  parsedEvents.forEach(({ event, timingContext }, index) => {
    dependencyNodes.push({ id: event.id, eventId: event.id, label: event.label, kind: "event" });
    if (timingContext.kind === "onClick" && timingContext.triggerId) {
      const trigger = {
        id: timingContext.triggerId,
        type: "onClick" as const,
        targetId: timingContext.triggerTargetId,
        clickIndex: timingContext.clickIndex ?? index + 1
      };
      triggers.set(trigger.id, trigger);
      dependencyNodes.push({ id: trigger.id, label: `Click ${trigger.clickIndex}`, kind: "trigger" });
      dependencyEdges.push({
        id: `${event.id}-trigger-edge`,
        from: trigger.id,
        to: event.id,
        relation: "triggers",
        offsetMs: timingContext.delayMs
      });
    }

    const previous = parsedEvents[index - 1]?.event;
    if (!previous) return;
    if (timingContext.kind === "withPrevious") {
      dependencyEdges.push({
        id: `${event.id}-with-previous`,
        from: previous.id,
        to: event.id,
        relation: "with",
        offsetMs: timingContext.delayMs
      });
    } else if (timingContext.kind === "afterPrevious") {
      dependencyEdges.push({
        id: `${event.id}-after-previous`,
        from: previous.id,
        to: event.id,
        relation: "after",
        offsetMs: timingContext.delayMs
      });
    }
  });

  for (const buildSequence of triggerBuildSequences) {
    const triggerId = `slide-${slideIndex + 1}-build-click-${buildSequence.clickIndex}`;
    triggers.set(triggerId, { id: triggerId, type: "onClick", clickIndex: buildSequence.clickIndex });
    dependencyNodes.push({ id: triggerId, label: `Build click ${buildSequence.clickIndex}`, kind: "trigger" });
    const eventsByTarget = buildSequence.targetIds
      .map((targetId) => parsedEvents.find(({ event }) => "targetId" in event && event.targetId === targetId)?.event)
      .filter((event): event is AnimationEvent => Boolean(event));
    eventsByTarget.forEach((event, index) => {
      dependencyEdges.push({
        id: `${event.id}-build-click-${buildSequence.clickIndex}`,
        from: index === 0 ? triggerId : eventsByTarget[index - 1].id,
        to: event.id,
        relation: index === 0 ? "triggers" : "after"
      });
    });
  }

  result.triggers = uniqueById([...result.triggers, ...triggers.values()]);
  result.dependencyNodes = uniqueById([...result.dependencyNodes, ...dependencyNodes]);
  addDependencyEdges(result.dependencyEdges, dependencyEdges);
}

function parseTriggerBuildSequences(timingXml: string, shapeIdToObjectId: Map<string, string>): TriggerBuildSequence[] {
  return extractXmlElements(timingXml, "bldLst")
    .flatMap((buildListXml) => extractXmlElements(buildListXml, "bldP"))
    .map((build, index) => {
      const attrs = readXmlAttributes(build, "bldP");
      const targetId = attrs.get("spid") ? shapeIdToObjectId.get(attrs.get("spid") ?? "") : undefined;
      const childTargetIds = Array.from(build.matchAll(/<p:bld\b[^>]*spid="([^"]+)"/g))
        .map((match) => match[1])
        .map((shapeId) => (shapeId ? shapeIdToObjectId.get(shapeId) : undefined))
        .filter((id): id is string => Boolean(id));
      return {
        targetIds: uniqueStrings([...(targetId ? [targetId] : []), ...childTargetIds]),
        clickIndex: Number(attrs.get("grpId")) || index + 1
      };
    })
    .filter((sequence) => sequence.targetIds.length > 1);
}

function uniqueStrings(items: string[]): string[] {
  return Array.from(new Set(items));
}

function addDependencyEdges(target: TimingDependencyEdge[], edges: TimingDependencyEdge[]): void {
  const seen = new Set(target.map((edge) => `${edge.from}\u0000${edge.to}\u0000${edge.relation}\u0000${edge.offsetMs ?? ""}`));
  for (const edge of edges) {
    const key = `${edge.from}\u0000${edge.to}\u0000${edge.relation}\u0000${edge.offsetMs ?? ""}`;
    if (seen.has(key)) continue;
    seen.add(key);
    target.push(edge);
  }
}

function uniqueById<T extends { id: string }>(items: T[]): T[] {
  const seen = new Set<string>();
  const result: T[] = [];
  for (const item of items) {
    if (seen.has(item.id)) continue;
    seen.add(item.id);
    result.push(item);
  }
  return result;
}

function resolveTimingTarget(
  node: string,
  slideIndex: number,
  shapeIdToObjectId: Map<string, string>,
  unsupportedFeatures: UnsupportedFeature[]
): string | undefined {
  const shapeId = node.match(/<p:spTgt\b[^>]*spid="([^"]+)"/)?.[1];
  if (!shapeId) {
    unsupportedFeatures.push({
      code: "presentationml-animation-target",
      severity: "warning",
      area: "animation",
      description: "PowerPoint timing node does not target a slide shape.",
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing`,
      fallback: "Preserve the static object and omit this animation."
    });
    return undefined;
  }

  const targetId = shapeIdToObjectId.get(shapeId);
  if (!targetId) {
    unsupportedFeatures.push({
      code: "presentationml-animation-target",
      severity: "warning",
      area: "animation",
      description: `PowerPoint timing node targets shape id "${shapeId}", which was not imported as an IR object.`,
      sourcePath: `ppt/slides/slide${slideIndex + 1}.xml#/p:timing`,
      fallback: "Preserve the static object and omit this animation."
    });
  }
  return targetId;
}

function readTimingDuration(node: string, fallbackMs: number): number {
  const cBhvr = extractXmlElements(node, "cBhvr")[0] ?? node;
  const cTnAttrs = readXmlAttributes(cBhvr, "cTn");
  return parseTimingMs(cTnAttrs.get("dur"), fallbackMs);
}

function readTimingDelay(node: string): number {
  const match = node.match(/<p:cond\b[^>]*delay="([^"]+)"/);
  return parseTimingMs(match?.[1], 0);
}

function readVisibilityValue(node: string): boolean | undefined {
  const strValue = node.match(/<p:strVal\b[^>]*val="([^"]+)"/)?.[1]?.toLowerCase();
  if (strValue === "visible" || strValue === "true" || strValue === "1") return true;
  if (strValue === "hidden" || strValue === "false" || strValue === "0") return false;
  const boolValue = node.match(/<p:boolVal\b[^>]*val="([^"]+)"/)?.[1]?.toLowerCase();
  if (boolValue === "true" || boolValue === "1") return true;
  if (boolValue === "false" || boolValue === "0") return false;
  return undefined;
}

function parseMotionPath(path: string | undefined, slideSize: { width: number; height: number }): { x: number; y: number } | undefined {
  if (!path) return undefined;
  const points = Array.from(path.matchAll(/[ML]\s*(-?\d*\.?\d+)[,\s]+(-?\d*\.?\d+)/gi)).map((match) => ({
    x: Number(match[1]),
    y: Number(match[2])
  }));
  if (points.length < 2 || points.some((point) => !Number.isFinite(point.x) || !Number.isFinite(point.y))) return undefined;
  const first = points[0];
  const last = points[points.length - 1];
  return {
    x: motionCoordinateToPx(last.x - first.x, slideSize.width),
    y: motionCoordinateToPx(last.y - first.y, slideSize.height)
  };
}

function parseMotionBy(value: string | undefined, slideSize: { width: number; height: number }): { x: number; y: number } | undefined {
  const point = parseMotionPoint(value);
  if (!point) return undefined;
  return {
    x: motionCoordinateToPx(point.x, slideSize.width),
    y: motionCoordinateToPx(point.y, slideSize.height)
  };
}

function parseMotionFromTo(
  from: string | undefined,
  to: string | undefined,
  slideSize: { width: number; height: number }
): { x: number; y: number } | undefined {
  const fromPoint = parseMotionPoint(from);
  const toPoint = parseMotionPoint(to);
  if (!fromPoint || !toPoint) return undefined;
  return {
    x: motionCoordinateToPx(toPoint.x - fromPoint.x, slideSize.width),
    y: motionCoordinateToPx(toPoint.y - fromPoint.y, slideSize.height)
  };
}

function parseMotionPoint(value: string | undefined): { x: number; y: number } | undefined {
  if (!value) return undefined;
  const parts = value
    .trim()
    .split(/[,\s]+/)
    .filter(Boolean)
    .map(Number);
  if (parts.length < 2 || !Number.isFinite(parts[0]) || !Number.isFinite(parts[1])) return undefined;
  return { x: parts[0], y: parts[1] };
}

function parseScalePair(value: string | undefined): { fromX: number; fromY: number; toX: number; toY: number } | undefined {
  const point = parseMotionPoint(value);
  if (!point) return undefined;
  return {
    fromX: 1,
    fromY: 1,
    toX: normalizeScaleValue(point.x),
    toY: normalizeScaleValue(point.y)
  };
}

function parseScaleFromTo(
  from: string | undefined,
  to: string | undefined
): { fromX: number; fromY: number; toX: number; toY: number } | undefined {
  const fromPoint = parseMotionPoint(from);
  const toPoint = parseMotionPoint(to);
  if (!fromPoint || !toPoint) return undefined;
  return {
    fromX: normalizeScaleValue(fromPoint.x),
    fromY: normalizeScaleValue(fromPoint.y),
    toX: normalizeScaleValue(toPoint.x),
    toY: normalizeScaleValue(toPoint.y)
  };
}

function parseScaleValuesFromChildNodes(node: string): { fromX: number; fromY: number; toX: number; toY: number } | undefined {
  const from = node.match(/<p:from>[\s\S]*?<p:pt\b[^>]*x="([^"]+)"[^>]*y="([^"]+)"/);
  const to = node.match(/<p:to>[\s\S]*?<p:pt\b[^>]*x="([^"]+)"[^>]*y="([^"]+)"/);
  if (!from || !to) return undefined;
  return {
    fromX: normalizeScaleValue(Number(from[1])),
    fromY: normalizeScaleValue(Number(from[2])),
    toX: normalizeScaleValue(Number(to[1])),
    toY: normalizeScaleValue(Number(to[2]))
  };
}

function parseRotationValues(
  by: string | undefined,
  from: string | undefined,
  to: string | undefined
): { from: number; to: number } | undefined {
  if (from !== undefined && to !== undefined) {
    return { from: openXmlAngleToDeg(Number(from)), to: openXmlAngleToDeg(Number(to)) };
  }
  if (by !== undefined) {
    return { from: 0, to: openXmlAngleToDeg(Number(by)) };
  }
  return undefined;
}

function parseRotationValuesFromChildNodes(node: string): { from: number; to: number } | undefined {
  const from = node.match(/<p:from>[\s\S]*?<p:fltVal\b[^>]*val="([^"]+)"/)?.[1];
  const to = node.match(/<p:to>[\s\S]*?<p:fltVal\b[^>]*val="([^"]+)"/)?.[1];
  if (from === undefined || to === undefined) return undefined;
  return { from: openXmlAngleToDeg(Number(from)), to: openXmlAngleToDeg(Number(to)) };
}

function motionCoordinateToPx(value: number, slidePixels: number): number {
  const px = Math.abs(value) <= 1.5 ? value * slidePixels : value;
  return Math.round(px * 1000) / 1000;
}

function normalizeScaleValue(value: number): number {
  if (!Number.isFinite(value)) return 1;
  const normalized = Math.abs(value) > 10 ? value / 100000 : value;
  return Math.round(normalized * 10000) / 10000;
}

function formatScalePair(x: number, y: number): string {
  return `${Math.round(x * 100000)} ${Math.round(y * 100000)}`;
}

function openXmlAngleToDeg(value: number): number {
  if (!Number.isFinite(value)) return 0;
  const deg = Math.abs(value) > 360 ? value / OPENXML_ANGLE_UNITS_PER_DEGREE : value;
  return Math.round(deg * 1000) / 1000;
}

function degToOpenXmlAngle(value: number): number {
  return Math.round(value * OPENXML_ANGLE_UNITS_PER_DEGREE);
}

function parseTimingMs(value: string | undefined, fallbackMs: number): number {
  if (!value || value === "indefinite") return fallbackMs;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? Math.max(0, numeric) : fallbackMs;
}

function eventStartMs(event: AnimationEvent): number {
  if (event.start?.type === "absolute") return event.start.atMs;
  if (event.start && "offsetMs" in event.start && typeof event.start.offsetMs === "number") return event.start.offsetMs;
  return event.delayMs ?? 0;
}

function extractXmlElements(source: string, localName: string): string[] {
  const elementPattern = new RegExp(`<p:${localName}\\b[\\s\\S]*?<\\/p:${localName}>`, "g");
  return source.match(elementPattern) ?? [];
}

function parseXmlElementRanges(source: string): XmlElementRange[] {
  const ranges: XmlElementRange[] = [];
  const stack: XmlElementRange[] = [];
  const tagPattern = /<(?<closing>\/)?(?<prefix>[A-Za-z_][\w.-]*):(?<localName>[A-Za-z_][\w.-]*)(?<attrs>[^<>]*?)(?<selfClosing>\/)?>/g;
  for (const match of source.matchAll(tagPattern)) {
    const localName = match.groups?.localName;
    if (!localName) continue;
    const start = match.index ?? 0;
    const end = start + match[0].length;
    if (match.groups?.closing) {
      const stackIndex = findLastIndex(stack, (range) => range.localName === localName);
      if (stackIndex === -1) continue;
      const range = stack[stackIndex];
      stack.splice(stackIndex);
      range.closeStart = start;
      range.end = end;
      ranges.push(range);
      continue;
    }

    const parent = stack[stack.length - 1];
    const range: XmlElementRange = { localName, start, openEnd: end, closeStart: end, end, parent, children: [] };
    if (parent) parent.children.push(range);
    if (match.groups?.selfClosing || /\/\s*>$/.test(match[0])) {
      ranges.push(range);
    } else {
      stack.push(range);
    }
  }
  return ranges.sort((a, b) => a.start - b.start);
}

function findLastIndex<T>(items: T[], predicate: (item: T) => boolean): number {
  for (let index = items.length - 1; index >= 0; index -= 1) {
    if (predicate(items[index])) return index;
  }
  return -1;
}

function xmlAncestors(range: XmlElementRange): XmlElementRange[] {
  const ancestors: XmlElementRange[] = [];
  for (let current = range.parent; current; current = current.parent) {
    ancestors.push(current);
  }
  return ancestors;
}

function xmlAncestorPath(range: XmlElementRange): string[] {
  return [...xmlAncestors(range).reverse().map((ancestor) => `p:${ancestor.localName}`), `p:${range.localName}`];
}

function immediateChildXml(source: string, ranges: XmlElementRange[], parent: XmlElementRange, localName: string): string | undefined {
  const child = immediateChildRange(ranges, parent, localName);
  return child ? source.slice(child.start, child.end) : undefined;
}

function immediateChildRange(ranges: XmlElementRange[], parent: XmlElementRange, localName: string): XmlElementRange | undefined {
  return (
    parent.children.find((candidate) => candidate.localName === localName) ??
    ranges.find((candidate) => candidate.localName === localName && candidate.parent === parent)
  );
}

function readXmlAttributes(source: string, localName: string): Map<string, string> {
  const attrs = new Map<string, string>();
  const match = source.match(new RegExp(`<p:${localName}\\b([^>]*)>`));
  if (!match) return attrs;
  for (const attr of match[1].matchAll(/([A-Za-z_:][\w:.-]*)="([^"]*)"/g)) {
    attrs.set(attr[1], unescapeXml(attr[2]));
  }
  return attrs;
}

function presentationXml(deck: DeckIR): string {
  const ids = deck.deck.slides
    .map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`)
    .join("");
  return xml(`\
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${ids}</p:sldIdLst>
  <p:sldSz cx="${inchesToEmu(WIDE_LAYOUT.width)}" cy="${inchesToEmu(WIDE_LAYOUT.height)}" type="wide"/>
  <p:notesSz cx="${inchesToEmu(10)}" cy="${inchesToEmu(7.5)}"/>
</p:presentation>`);
}

function presentationRels(slideCount: number): string {
  const slideRelsXml = Array.from({ length: slideCount }, (_, index) => {
    const id = index + 2;
    return `<Relationship Id="rId${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`;
  }).join("");
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${slideRelsXml}
  <Relationship Id="rId${slideCount + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`);
}

function slideXml(deck: DeckIR, slide: Slide): string {
  const bg = solidFill(slide.background) ?? "#ffffff";
  const shapeIds = assignShapeIds(slide.objects);
  const shapes = slide.objects.map((object) => objectToShape(deck, object, shapeIds)).join("");
  const timing = timingXml(deck, slide);
  return xml(`\
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="${stripHash(bg)}"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      ${shapes}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  ${timing}
</p:sld>`);
}

function timingXml(deck: DeckIR, slide: Slide): string {
  const shapeIds = assignShapeIds(slide.objects);
  const entries: string[] = [];
  let nextNodeIndex = 0;
  for (const event of slide.timeline?.events ?? []) {
    const nodes = eventToTimingNodes(deck, slide, event, nextNodeIndex, shapeIds);
    entries.push(...nodes);
    nextNodeIndex += nodes.length;
  }
  if (entries.length === 0) return "";

  const mainSequence = entries.join("");
  const buildList = buildListXml(slide, shapeIds);
  return `\
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                ${mainSequence}
              </p:childTnLst>
            </p:cTn>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
  ${buildList}
</p:timing>`;
}

function eventToTimingNodes(
  deck: DeckIR,
  slide: Slide,
  event: AnimationEvent,
  startIndex: number,
  shapeIds: Map<string, string>
): string[] {
  if (event.kind === "visibility") {
    const shapeId = shapeIds.get(event.targetId);
    if (!shapeId) return [];
    return [timingWrapper(slide, event, startIndex, visibilitySetXml(startIndex, shapeId, event.visible, 0), shapeIds)];
  }

  if (event.kind !== "keyframes") return [];
  const shapeId = shapeIds.get(event.targetId);
  if (!shapeId) return [];

  const timingNodes: string[] = [];
  const opacity = opacityFade(event);
  if (opacity) {
    const nodeIndex = startIndex + timingNodes.length;
    timingNodes.push(timingWrapper(slide, event, nodeIndex, fadeEffectXml(nodeIndex, shapeId, opacity, event.durationMs ?? 500, 0), shapeIds));
  }

  const motion = motionOffsets(deck, slide, event);
  if (motion && (motion.fromX !== motion.toX || motion.fromY !== motion.toY)) {
    const nodeIndex = startIndex + timingNodes.length;
    timingNodes.push(timingWrapper(slide, event, nodeIndex, motionPathXml(deck, nodeIndex, shapeId, motion, event.durationMs ?? 500, 0), shapeIds));
  }

  const scale = scaleValues(event);
  if (scale && (scale.fromX !== scale.toX || scale.fromY !== scale.toY)) {
    const nodeIndex = startIndex + timingNodes.length;
    timingNodes.push(timingWrapper(slide, event, nodeIndex, scaleXml(nodeIndex, shapeId, scale, event.durationMs ?? 500, 0), shapeIds));
  }

  const rotation = rotationValues(event);
  if (rotation && rotation.from !== rotation.to) {
    const nodeIndex = startIndex + timingNodes.length;
    timingNodes.push(timingWrapper(slide, event, nodeIndex, rotationXml(nodeIndex, shapeId, rotation, event.durationMs ?? 500, 0), shapeIds));
  }

  return timingNodes;
}

function timingWrapper(slide: Slide, event: AnimationEvent, index: number, childXml: string, shapeIds: Map<string, string>): string {
  const parId = 3 + index * 3;
  const context = exportTimingContext(slide, event, index, shapeIds);
  const stCond = startConditionXml(context);
  const nodeType = context.nodeType ? ` nodeType="${context.nodeType}"` : "";
  const presetClass = context.presetClass ? ` presetClass="${context.presetClass}"` : "";
  return `\
<p:par>
  <p:cTn id="${parId}" fill="hold"${nodeType}${presetClass}>
    ${stCond}
    <p:childTnLst>
      ${childXml}
    </p:childTnLst>
  </p:cTn>
</p:par>`;
}

function exportTimingContext(
  slide: Slide,
  event: AnimationEvent,
  index: number,
  shapeIds: Map<string, string>
): { kind: "absolute" | "withPrevious" | "afterPrevious" | "onClick"; delayMs: number; nodeType?: string; presetClass?: string; triggerShapeId?: string } {
  const start = event.start;
  if (start?.type === "withPrevious" || start?.type === "with") {
    return { kind: "withPrevious", delayMs: start.offsetMs ?? 0, nodeType: "withEffect" };
  }
  if (start?.type === "afterPrevious" || start?.type === "after") {
    return { kind: "afterPrevious", delayMs: start.offsetMs ?? 0, nodeType: "afterEffect" };
  }
  if (start?.type === "onClick") {
    const triggerShapeId = start.targetId ? shapeIds.get(start.targetId) : targetShapeId(event, shapeIds);
    return { kind: "onClick", delayMs: 0, nodeType: "clickEffect", presetClass: "entr", triggerShapeId };
  }
  if (start?.type === "trigger") {
    const trigger = slide.timeline?.triggers?.find((candidate) => candidate.id === start.triggerId);
    if (trigger?.type === "onClick") {
      const triggerShapeId = trigger.targetId ? shapeIds.get(trigger.targetId) : targetShapeId(event, shapeIds);
      return { kind: "onClick", delayMs: start.offsetMs ?? 0, nodeType: "clickEffect", presetClass: "entr", triggerShapeId };
    }
    if (trigger?.type === "afterPrevious") {
      return { kind: "afterPrevious", delayMs: trigger.offsetMs ?? start.offsetMs ?? 0, nodeType: "afterEffect" };
    }
  }
  return { kind: "absolute", delayMs: eventStartMs(event) };
}

function startConditionXml(context: { kind: string; delayMs: number; triggerShapeId?: string }): string {
  const delay = Math.max(0, Math.round(context.delayMs));
  if (context.kind === "onClick") {
    const target = context.triggerShapeId ? `<p:tgtEl><p:spTgt spid="${context.triggerShapeId}"/></p:tgtEl>` : "";
    const delayValue = delay > 0 ? String(delay) : "indefinite";
    return `<p:stCondLst><p:cond delay="${delayValue}" evt="onClick">${target}</p:cond></p:stCondLst>`;
  }
  if (context.kind === "withPrevious") {
    return `<p:stCondLst><p:cond delay="${delay}" evt="begin"/></p:stCondLst>`;
  }
  if (context.kind === "afterPrevious") {
    return `<p:stCondLst><p:cond delay="${delay}" evt="onEnd"/></p:stCondLst>`;
  }
  return `<p:stCondLst><p:cond delay="${delay}"/></p:stCondLst>`;
}

function targetShapeId(event: AnimationEvent, shapeIds: Map<string, string>): string | undefined {
  return "targetId" in event ? shapeIds.get(event.targetId) : undefined;
}

function buildListXml(slide: Slide, shapeIds: Map<string, string>): string {
  const orderedShapeIds = orderedBuildShapeIds(slide, shapeIds);
  if (orderedShapeIds.length === 0) return "";
  const items = orderedShapeIds
    .map((shapeId, index) => `<p:bldP spid="${shapeId}" grpId="${index + 1}"><p:bld spid="${shapeId}"/></p:bldP>`)
    .join("");
  return `<p:bldLst>${items}</p:bldLst>`;
}

function orderedBuildShapeIds(slide: Slide, shapeIds: Map<string, string>): string[] {
  const events = slide.timeline?.events ?? [];
  const eventIds = new Set(events.map((event) => event.id));
  const edges = (slide.timeline?.dependencyGraph?.edges ?? []).filter(
    (edge) => eventIds.has(edge.to) && (eventIds.has(edge.from) || edge.relation === "triggers")
  );
  const linkedEventIds = new Set(edges.flatMap((edge) => [edge.from, edge.to]).filter((id) => eventIds.has(id)));
  events.forEach((event) => {
    if (["with", "withPrevious", "after", "afterPrevious", "onClick", "trigger"].includes(String(event.start?.type))) {
      linkedEventIds.add(event.id);
    }
  });
  const triggeredEventIds = new Set(
    events
      .filter((event) => event.start?.type === "onClick" || event.start?.type === "trigger" || linkedEventIds.has(event.id))
      .map((event) => event.id)
  );
  if (triggeredEventIds.size === 0) return [];
  const shapeIdsInOrder: string[] = [];
  for (const event of events) {
    if (!triggeredEventIds.has(event.id)) continue;
    const shapeId = targetShapeId(event, shapeIds);
    if (!shapeId || shapeIdsInOrder.includes(shapeId)) continue;
    shapeIdsInOrder.push(shapeId);
  }
  return shapeIdsInOrder;
}

function fadeEffectXml(index: number, shapeId: string, fade: "in" | "out", durationMs: number, delayMs: number): string {
  return `\
<p:animEffect transition="${fade}" filter="fade">
  <p:cBhvr>
    <p:cTn id="${4 + index * 3}" dur="${Math.max(1, Math.round(durationMs))}" fill="hold">
      <p:stCondLst><p:cond delay="${Math.max(0, Math.round(delayMs))}"/></p:stCondLst>
    </p:cTn>
    <p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>
  </p:cBhvr>
</p:animEffect>`;
}

function visibilitySetXml(index: number, shapeId: string, visible: boolean, delayMs: number): string {
  return `\
<p:set>
  <p:cBhvr>
    <p:cTn id="${4 + index * 3}" dur="1" fill="hold">
      <p:stCondLst><p:cond delay="${Math.max(0, Math.round(delayMs))}"/></p:stCondLst>
    </p:cTn>
    <p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>
    <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
  </p:cBhvr>
  <p:to><p:strVal val="${visible ? "visible" : "hidden"}"/></p:to>
</p:set>`;
}

function motionPathXml(
  deck: DeckIR,
  index: number,
  shapeId: string,
  motion: { fromX: number; fromY: number; toX: number; toY: number },
  durationMs: number,
  delayMs: number
): string {
  const fromX = formatMotionCoordinate(motion.fromX / deck.deck.size.width);
  const fromY = formatMotionCoordinate(motion.fromY / deck.deck.size.height);
  const toX = formatMotionCoordinate(motion.toX / deck.deck.size.width);
  const toY = formatMotionCoordinate(motion.toY / deck.deck.size.height);
  return `\
<p:animMotion origin="layout" path="M ${fromX} ${fromY} L ${toX} ${toY} E" pathEditMode="relative">
  <p:cBhvr>
    <p:cTn id="${4 + index * 3}" dur="${Math.max(1, Math.round(durationMs))}" fill="hold">
      <p:stCondLst><p:cond delay="${Math.max(0, Math.round(delayMs))}"/></p:stCondLst>
    </p:cTn>
    <p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>
  </p:cBhvr>
</p:animMotion>`;
}

function scaleXml(
  index: number,
  shapeId: string,
  scale: { fromX: number; fromY: number; toX: number; toY: number },
  durationMs: number,
  delayMs: number
): string {
  return `\
<p:animScale from="${formatScalePair(scale.fromX, scale.fromY)}" to="${formatScalePair(scale.toX, scale.toY)}">
  <p:cBhvr>
    <p:cTn id="${4 + index * 3}" dur="${Math.max(1, Math.round(durationMs))}" fill="hold">
      <p:stCondLst><p:cond delay="${Math.max(0, Math.round(delayMs))}"/></p:stCondLst>
    </p:cTn>
    <p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>
  </p:cBhvr>
</p:animScale>`;
}

function rotationXml(
  index: number,
  shapeId: string,
  rotation: { from: number; to: number },
  durationMs: number,
  delayMs: number
): string {
  return `\
<p:animRot from="${degToOpenXmlAngle(rotation.from)}" to="${degToOpenXmlAngle(rotation.to)}">
  <p:cBhvr>
    <p:cTn id="${4 + index * 3}" dur="${Math.max(1, Math.round(durationMs))}" fill="hold">
      <p:stCondLst><p:cond delay="${Math.max(0, Math.round(delayMs))}"/></p:stCondLst>
    </p:cTn>
    <p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>
  </p:cBhvr>
</p:animRot>`;
}

function opacityFade(event: AnimationEvent): "in" | "out" | undefined {
  if (event.kind !== "keyframes") return undefined;
  const track = event.tracks.find((candidate) => candidate.property === "opacity");
  const values = twoPointNumericValues(track?.keyframes);
  if (!values) return undefined;
  if (values.from === 0 && values.to === 1) return "in";
  if (values.from === 1 && values.to === 0) return "out";
  return undefined;
}

function motionOffsets(deck: DeckIR, slide: Slide, event: AnimationEvent): { fromX: number; fromY: number; toX: number; toY: number } | undefined {
  if (event.kind !== "keyframes") return undefined;
  const object = slide.objects.find((candidate) => candidate.id === event.targetId);
  const translateX = twoPointNumericValues(event.tracks.find((track) => track.property === "transform.translateX")?.keyframes);
  const translateY = twoPointNumericValues(event.tracks.find((track) => track.property === "transform.translateY")?.keyframes);
  const boundsX = twoPointNumericValues(event.tracks.find((track) => track.property === "bounds.x")?.keyframes);
  const boundsY = twoPointNumericValues(event.tracks.find((track) => track.property === "bounds.y")?.keyframes);
  const fromX = translateX ? translateX.from : boundsX ? boundsX.from - (object?.bounds?.x ?? 0) : 0;
  const toX = translateX ? translateX.to : boundsX ? boundsX.to - (object?.bounds?.x ?? 0) : 0;
  const fromY = translateY ? translateY.from : boundsY ? boundsY.from - (object?.bounds?.y ?? 0) : 0;
  const toY = translateY ? translateY.to : boundsY ? boundsY.to - (object?.bounds?.y ?? 0) : 0;
  if (fromX === toX && fromY === toY) return undefined;
  if (!Number.isFinite(deck.deck.size.width) || !Number.isFinite(deck.deck.size.height)) return undefined;
  const deltaX = toX - fromX;
  const deltaY = toY - fromY;
  if (object?.bounds && Math.abs(deltaX) > deck.deck.size.width * 2 && Math.abs(deltaY) > deck.deck.size.height * 2) return undefined;
  return { fromX, fromY, toX, toY };
}

function scaleValues(event: AnimationEvent): { fromX: number; fromY: number; toX: number; toY: number } | undefined {
  if (event.kind !== "keyframes") return undefined;
  const uniform = twoPointNumericValues(event.tracks.find((track) => track.property === "transform.scale")?.keyframes);
  const scaleX = twoPointNumericValues(event.tracks.find((track) => track.property === "transform.scaleX")?.keyframes);
  const scaleY = twoPointNumericValues(event.tracks.find((track) => track.property === "transform.scaleY")?.keyframes);
  if (!uniform && !scaleX && !scaleY) return undefined;
  return {
    fromX: scaleX?.from ?? uniform?.from ?? 1,
    fromY: scaleY?.from ?? uniform?.from ?? 1,
    toX: scaleX?.to ?? uniform?.to ?? 1,
    toY: scaleY?.to ?? uniform?.to ?? 1
  };
}

function rotationValues(event: AnimationEvent): { from: number; to: number } | undefined {
  if (event.kind !== "keyframes") return undefined;
  return twoPointNumericValues(event.tracks.find((track) => track.property === "transform.rotateDeg")?.keyframes);
}

function twoPointNumericValues(keyframes: { offset: number; value: unknown }[] | undefined): { from: number; to: number } | undefined {
  if (!keyframes || keyframes.length < 2) return undefined;
  const first = keyframes[0];
  const last = keyframes[keyframes.length - 1];
  if (first.offset !== 0 || last.offset !== 1) return undefined;
  if (typeof first.value !== "number" || typeof last.value !== "number") return undefined;
  return { from: first.value, to: last.value };
}

function formatMotionCoordinate(value: number): string {
  if (!Number.isFinite(value)) return "0";
  const rounded = Math.round(value * 1000000) / 1000000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

function assignShapeIds(objects: IRObject[]): Map<string, string> {
  const ids = new Map<string, string>();
  let nextId = 2;
  const visit = (object: IRObject) => {
    if (object.type !== "group") {
      ids.set(object.id, String(nextId));
      nextId += 1;
      return;
    }
    for (const child of object.children) visit(child);
  };
  for (const object of objects) visit(object);
  return ids;
}

function objectToShape(deck: DeckIR, object: IRObject, shapeIds: Map<string, string>): string {
  if (object.type === "group") {
    return object.children.map((child) => objectToShape(deck, child, shapeIds)).join("");
  }

  const id = Number(shapeIds.get(object.id) ?? 2);
  const bounds = object.bounds ?? { x: 0, y: 0, width: 100, height: 100 };
  const x = pxToEmu(bounds.x, deck.deck.size.width, WIDE_LAYOUT.width);
  const y = pxToEmu(bounds.y, deck.deck.size.height, WIDE_LAYOUT.height);
  const cx = pxToEmu(bounds.width, deck.deck.size.width, WIDE_LAYOUT.width);
  const cy = pxToEmu(bounds.height, deck.deck.size.height, WIDE_LAYOUT.height);
  const name = escapeXml(object.name ?? object.id);

  if (object.type === "text") {
    const text = escapeXml(object.text.plainText ?? object.text.runs?.map((run) => run.text).join("") ?? "");
    const style = object.text.runs?.[0]?.style ?? object.style?.textStyle ?? {};
    return `\
<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square"/>
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="${fontSizeToOpenXml(style.fontSize ?? 28)}"><a:solidFill><a:srgbClr val="${stripHash(colorToString(style.color) ?? "#111827")}"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>
  </p:txBody>
</p:sp>`;
  }

  const fill = object.type === "shape" ? solidFill(object.style?.fill) ?? "#e2e8f0" : "#e2e8f0";
  const shape = object.type === "shape" ? shapePreset(object.shape) : "rect";
  return `\
<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="${shape}"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="${stripHash(fill)}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
</p:sp>`;
}

function contentTypes(slideCount: number): string {
  const slideOverrides = Array.from({ length: slideCount }, (_, index) => {
    return `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }).join("");
  return xml(`\
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  ${slideOverrides}
</Types>`);
}

function rootRels(): string {
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
}

function slideRels(): string {
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);
}

function slideMasterRels(): string {
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`);
}

function slideLayoutRels(): string {
  return xml(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);
}

function slideMasterXml(): string {
  return xml(`\
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`);
}

function slideLayoutXml(): string {
  return xml(`\
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`);
}

function themeXml(): string {
  return xml(`\
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="KeyMorph">
  <a:themeElements>
    <a:clrScheme name="KeyMorph">
      <a:dk1><a:srgbClr val="111827"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="0F766E"/></a:accent1><a:accent2><a:srgbClr val="2563EB"/></a:accent2><a:accent3><a:srgbClr val="DC2626"/></a:accent3>
      <a:accent4><a:srgbClr val="CA8A04"/></a:accent4><a:accent5><a:srgbClr val="7C3AED"/></a:accent5><a:accent6><a:srgbClr val="0891B2"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="KeyMorph"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="KeyMorph"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`);
}

function coreProps(deck: DeckIR): string {
  return xml(`\
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(deck.deck.title ?? deck.metadata?.title ?? "KeyMorph deck")}</dc:title>
  <dc:creator>${escapeXml(deck.metadata?.author ?? "KeyMorph")}</dc:creator>
  <cp:lastModifiedBy>KeyMorph</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`);
}

function appProps(slideCount: number): string {
  return xml(`\
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>KeyMorph</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>${slideCount}</Slides>
</Properties>`);
}

function createZip(files: Map<string, string | Uint8Array>): Uint8Array {
  const chunks: Uint8Array[] = [];
  const centralDirectory: Uint8Array[] = [];
  let offset = 0;

  for (const [name, content] of files) {
    const nameBytes = textBytes(name);
    const data = typeof content === "string" ? textBytes(content) : content;
    const compressed = deflateRawSync(data);
    const crc = crc32(data);
    const local = concat([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(8),
      u16(0),
      u16(0),
      u32(crc),
      u32(compressed.length),
      u32(data.length),
      u16(nameBytes.length),
      u16(0),
      nameBytes,
      compressed
    ]);
    chunks.push(local);
    centralDirectory.push(
      concat([
        u32(0x02014b50),
        u16(20),
        u16(20),
        u16(0),
        u16(8),
        u16(0),
        u16(0),
        u32(crc),
        u32(compressed.length),
        u32(data.length),
        u16(nameBytes.length),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(0),
        u32(offset),
        nameBytes
      ])
    );
    offset += local.length;
  }

  const centralStart = offset;
  const central = concat(centralDirectory);
  const end = concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(files.size),
    u16(files.size),
    u32(central.length),
    u32(centralStart),
    u16(0)
  ]);

  return concat([...chunks, central, end]);
}

function xml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body.replaceAll("\r\n", "\n")}`;
}

function inchesToEmu(inches: number): number {
  return Math.round(inches * EMU_PER_INCH);
}

function pxToEmu(px: number, deckPixels: number, inches: number): number {
  return Math.round((px / deckPixels) * inches * EMU_PER_INCH);
}

function emuToWidePx(emu: number, inches: number): number {
  const pixels = inches === WIDE_LAYOUT.width ? 1280 : 720;
  return Math.round((emu / (inches * EMU_PER_INCH)) * pixels);
}

function fontSizeToOpenXml(px: number): number {
  return Math.round(px * 0.75 * 100);
}

function shapePreset(shape: string): string {
  if (shape === "ellipse") return "ellipse";
  if (shape === "roundRect") return "roundRect";
  if (shape === "triangle") return "triangle";
  return "rect";
}

function solidFill(fill: unknown): string | undefined {
  if (typeof fill === "string") return fill;
  if (fill && typeof fill === "object" && "type" in fill && (fill as { type: string }).type === "solid") {
    return colorToString((fill as { color?: unknown }).color);
  }
  return undefined;
}

function colorToString(color: unknown): string | undefined {
  if (typeof color === "string") return color;
  if (color && typeof color === "object" && "value" in color) return String((color as { value: unknown }).value);
  return undefined;
}

function stripHash(color: string): string {
  return color.replace(/^#/, "").slice(0, 6);
}

function escapeXml(value: unknown): string {
  return String(value ?? "").replace(/[<>&'"]/g, (char) => {
    switch (char) {
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case "&":
        return "&amp;";
      case "'":
        return "&apos;";
      default:
        return "&quot;";
    }
  });
}

function unescapeXml(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function textBytes(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function getTextEntry(entries: Map<string, Uint8Array>, path: string): string | undefined {
  const value = entries.get(normalizePartPath(path));
  return value ? new TextDecoder().decode(value) : undefined;
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
    const name = new TextDecoder().decode(data.slice(nameStart, nameStart + nameLength));

    if ((flags & 0x08) !== 0) {
      throw new Error("PPTX ZIP entries using data descriptors are not supported by the MVP reader.");
    }
    if (dataEnd > data.length) {
      throw new Error(`Invalid PPTX ZIP entry size for ${name}.`);
    }

    const compressed = data.slice(dataStart, dataEnd);
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

  return entries;
}

function normalizePartPath(path: string): string {
  const parts: string[] = [];
  for (const part of path.replace(/^\/+/, "").split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }
  return parts.join("/");
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

function u16(value: number): Uint8Array {
  const out = new Uint8Array(2);
  new DataView(out.buffer).setUint16(0, value, true);
  return out;
}

function u32(value: number): Uint8Array {
  const out = new Uint8Array(4);
  new DataView(out.buffer).setUint32(0, value >>> 0, true);
  return out;
}

const CRC_TABLE = new Uint32Array(256).map((_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit += 1) {
    value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  }
  return value >>> 0;
});

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
