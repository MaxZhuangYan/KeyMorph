import { IR_VERSION, type KeyMorphIR, type TimingDependencyGraph } from "./types.ts";

export type ValidationSeverity = "error" | "warning";

export interface ValidationIssue {
  severity: ValidationSeverity;
  code: string;
  path: string;
  message: string;
}

export interface ValidationResult<T = KeyMorphIR> {
  valid: boolean;
  value?: T;
  errors: ValidationIssue[];
  warnings: ValidationIssue[];
}

interface ValidationContext {
  errors: ValidationIssue[];
  warnings: ValidationIssue[];
  objectIds: Set<string>;
  stateIds: Set<string>;
  slideIds: Set<string>;
  assetIds: Set<string>;
  eventIds: Set<string>;
  triggerIds: Set<string>;
}

const OBJECT_TYPES = new Set(["shape", "text", "image", "media", "path", "group", "embed", "placeholder"]);
const EVENT_KINDS = new Set(["setState", "property", "keyframes", "morph", "transition", "media", "visibility"]);
const EASING_NAMES = new Set([
  "linear",
  "ease",
  "easeIn",
  "easeOut",
  "easeInOut",
  "easeInCubic",
  "easeOutCubic",
  "easeInOutCubic",
  "backIn",
  "backOut",
  "backInOut",
  "bounceIn",
  "bounceOut",
  "bounceInOut",
  "elasticIn",
  "elasticOut",
  "elasticInOut",
]);
const TRANSITION_TYPES = new Set(["none", "cut", "fade", "push", "wipe", "zoom", "dissolve", "morph", "magicMove"]);
const MORPH_STRATEGIES = new Set(["morph", "magicMove"]);
const TIMING_START_TYPES = new Set([
  "absolute",
  "with",
  "after",
  "before",
  "onClick",
  "withPrevious",
  "afterPrevious",
  "trigger",
]);

export function validateIR(input: unknown): ValidationResult<KeyMorphIR> {
  const ctx: ValidationContext = {
    errors: [],
    warnings: [],
    objectIds: new Set(),
    stateIds: new Set(),
    slideIds: new Set(),
    assetIds: new Set(),
    eventIds: new Set(),
    triggerIds: new Set(),
  };

  if (!isRecord(input)) {
    addIssue(ctx, "error", "IR_NOT_OBJECT", "$", "IR must be an object.");
    return toResult(input, ctx);
  }

  if (input.irVersion !== IR_VERSION) {
    addIssue(ctx, "error", "INVALID_IR_VERSION", "$.irVersion", `IR version must be ${IR_VERSION}.`);
  }

  if (!isRecord(input.deck)) {
    addIssue(ctx, "error", "DECK_REQUIRED", "$.deck", "IR must include a deck object.");
    return toResult(input, ctx);
  }

  validateDeck(input.deck, "$.deck", ctx);
  validateConversionReport(input.conversion, "$.conversion", ctx);

  return toResult(input, ctx);
}

export function isValidIR(input: unknown): input is KeyMorphIR {
  return validateIR(input).valid;
}

function validateDeck(deck: Record<string, unknown>, path: string, ctx: ValidationContext): void {
  requireId(deck.id, `${path}.id`, ctx, undefined, "deck");
  validateDeckSize(deck.size, `${path}.size`, ctx);

  if (deck.assets !== undefined) {
    if (!Array.isArray(deck.assets)) {
      addIssue(ctx, "error", "ASSETS_NOT_ARRAY", `${path}.assets`, "Deck assets must be an array.");
    } else {
      deck.assets.forEach((asset, index) => validateAsset(asset, `${path}.assets[${index}]`, ctx));
    }
  }

  if (!Array.isArray(deck.slides) || deck.slides.length === 0) {
    addIssue(ctx, "error", "SLIDES_REQUIRED", `${path}.slides`, "Deck must include at least one slide.");
    return;
  }

  deck.slides.forEach((slide, index) => {
    if (!isRecord(slide)) {
      addIssue(ctx, "error", "SLIDE_NOT_OBJECT", `${path}.slides[${index}]`, "Slide must be an object.");
      return;
    }

    const slideId = requireId(slide.id, `${path}.slides[${index}].id`, ctx, ctx.slideIds, "slide");
    if (slideId !== undefined && slide.index !== undefined && slide.index !== index) {
      addIssue(ctx, "warning", "SLIDE_INDEX_MISMATCH", `${path}.slides[${index}].index`, "Slide index does not match its array position.");
    }

    validateSlide(slide, `${path}.slides[${index}]`, ctx);
  });
}

function validateDeckSize(size: unknown, path: string, ctx: ValidationContext): void {
  if (!isRecord(size)) {
    addIssue(ctx, "error", "DECK_SIZE_REQUIRED", path, "Deck size must be an object.");
    return;
  }

  requireFiniteNumber(size.width, `${path}.width`, ctx);
  requireFiniteNumber(size.height, `${path}.height`, ctx);

  if (!["px", "pt", "in", "cm"].includes(String(size.unit))) {
    addIssue(ctx, "error", "INVALID_DECK_SIZE_UNIT", `${path}.unit`, "Deck size unit must be one of px, pt, in, or cm.");
  }
}

function validateAsset(asset: unknown, path: string, ctx: ValidationContext): void {
  if (!isRecord(asset)) {
    addIssue(ctx, "error", "ASSET_NOT_OBJECT", path, "Asset must be an object.");
    return;
  }

  requireId(asset.id, `${path}.id`, ctx, ctx.assetIds, "asset");
  if (!["image", "video", "audio", "font", "data", "other"].includes(String(asset.kind))) {
    addIssue(ctx, "error", "INVALID_ASSET_KIND", `${path}.kind`, "Asset kind is not supported.");
  }

  validateOptionalNonNegativeNumber(asset.width, `${path}.width`, ctx);
  validateOptionalNonNegativeNumber(asset.height, `${path}.height`, ctx);
  validateOptionalNonNegativeNumber(asset.durationMs, `${path}.durationMs`, ctx);
}

function validateSlide(slide: Record<string, unknown>, path: string, ctx: ValidationContext): void {
  if (!Array.isArray(slide.objects)) {
    addIssue(ctx, "error", "OBJECTS_NOT_ARRAY", `${path}.objects`, "Slide objects must be an array.");
  } else {
    slide.objects.forEach((object, index) => validateObject(object, `${path}.objects[${index}]`, ctx, undefined));
  }

  if (slide.states !== undefined) {
    if (!Array.isArray(slide.states)) {
      addIssue(ctx, "error", "STATES_NOT_ARRAY", `${path}.states`, "Slide states must be an array.");
    } else {
      slide.states.forEach((state, index) => validateState(state, `${path}.states[${index}]`, ctx));
    }
  }

  validateTransition(slide.transition, `${path}.transition`, ctx);
  validateTimeline(slide.timeline, `${path}.timeline`, ctx);
}

function validateObject(object: unknown, path: string, ctx: ValidationContext, parentId: string | undefined): void {
  if (!isRecord(object)) {
    addIssue(ctx, "error", "OBJECT_NOT_OBJECT", path, "IR object must be an object.");
    return;
  }

  const objectId = requireId(object.id, `${path}.id`, ctx, ctx.objectIds, "object");
  if (!OBJECT_TYPES.has(String(object.type))) {
    addIssue(ctx, "error", "INVALID_OBJECT_TYPE", `${path}.type`, "Object type is not supported.");
  }

  validateRect(object.bounds, `${path}.bounds`, ctx, false);
  validateTransform(object.transform, `${path}.transform`, ctx);
  validateOptionalRatio(object.opacity, `${path}.opacity`, ctx);

  switch (object.type) {
    case "shape":
      if (typeof object.shape !== "string" || object.shape.length === 0) {
        addIssue(ctx, "error", "SHAPE_KIND_REQUIRED", `${path}.shape`, "Shape objects must include a shape kind.");
      }
      break;
    case "text":
      validateTextContent(object.text, `${path}.text`, ctx, true);
      break;
    case "image":
      validateSource(object.source, `${path}.source`, ctx);
      break;
    case "media":
      if (!["video", "audio"].includes(String(object.mediaType))) {
        addIssue(ctx, "error", "INVALID_MEDIA_TYPE", `${path}.mediaType`, "Media type must be video or audio.");
      }
      validateSource(object.source, `${path}.source`, ctx);
      break;
    case "path":
      validatePathCommands(object.path, `${path}.path`, ctx);
      break;
    case "group":
      if (!Array.isArray(object.children)) {
        addIssue(ctx, "error", "GROUP_CHILDREN_NOT_ARRAY", `${path}.children`, "Group objects must include a children array.");
      } else {
        object.children.forEach((child, index) => validateObject(child, `${path}.children[${index}]`, ctx, objectId));
      }
      break;
    case "embed":
      if (!["html", "iframe", "chart", "table", "unknown"].includes(String(object.embedType))) {
        addIssue(ctx, "error", "INVALID_EMBED_TYPE", `${path}.embedType`, "Embed type is not supported.");
      }
      break;
    case "placeholder":
      if (typeof object.placeholderType !== "string" || object.placeholderType.length === 0) {
        addIssue(ctx, "error", "PLACEHOLDER_TYPE_REQUIRED", `${path}.placeholderType`, "Placeholder type is required.");
      }
      break;
  }

  if (parentId !== undefined && objectId === parentId) {
    addIssue(ctx, "error", "GROUP_SELF_REFERENCE", `${path}.id`, "Group child cannot reuse the parent object id.");
  }
}

function validateState(state: unknown, path: string, ctx: ValidationContext): void {
  if (!isRecord(state)) {
    addIssue(ctx, "error", "STATE_NOT_OBJECT", path, "State must be an object.");
    return;
  }

  requireId(state.id, `${path}.id`, ctx, ctx.stateIds, "state");
  const targetId = requireId(state.targetId, `${path}.targetId`, ctx, undefined, "state target");
  if (targetId !== undefined && !ctx.objectIds.has(targetId)) {
    addIssue(ctx, "error", "UNKNOWN_STATE_TARGET", `${path}.targetId`, "State targetId must reference an object on this deck.");
  }

  if (!isRecord(state.properties)) {
    addIssue(ctx, "error", "STATE_PROPERTIES_REQUIRED", `${path}.properties`, "State must include properties.");
    return;
  }

  validateStateProperties(state.properties, `${path}.properties`, ctx);
}

function validateStateProperties(properties: Record<string, unknown>, path: string, ctx: ValidationContext): void {
  validateOptionalRatio(properties.opacity, `${path}.opacity`, ctx);
  validateRect(properties.bounds, `${path}.bounds`, ctx, false);
  validateTransform(properties.transform, `${path}.transform`, ctx);
  validateTextContent(properties.text, `${path}.text`, ctx, false);
}

function validateTimeline(timeline: unknown, path: string, ctx: ValidationContext): void {
  if (timeline === undefined) {
    return;
  }

  if (!isRecord(timeline)) {
    addIssue(ctx, "error", "TIMELINE_NOT_OBJECT", path, "Timeline must be an object.");
    return;
  }

  validateOptionalNonNegativeNumber(timeline.durationMs, `${path}.durationMs`, ctx);
  validateEasing(timeline.defaultEasing, `${path}.defaultEasing`, ctx);

  if (timeline.triggers !== undefined) {
    if (!Array.isArray(timeline.triggers)) {
      addIssue(ctx, "error", "TRIGGERS_NOT_ARRAY", `${path}.triggers`, "Timeline triggers must be an array.");
    } else {
      timeline.triggers.forEach((trigger, index) => validateTrigger(trigger, `${path}.triggers[${index}]`, ctx));
    }
  }

  if (!Array.isArray(timeline.events)) {
    addIssue(ctx, "error", "TIMELINE_EVENTS_NOT_ARRAY", `${path}.events`, "Timeline events must be an array.");
    return;
  }

  const beforeEventIds = new Set(ctx.eventIds);
  timeline.events.forEach((event, index) => {
    if (isRecord(event)) {
      requireId(event.id, `${path}.events[${index}].id`, ctx, ctx.eventIds, "event");
    }
  });

  timeline.events.forEach((event, index) => validateAnimationEvent(event, `${path}.events[${index}]`, ctx));
  validateTimingDependencyGraph(timeline.dependencyGraph, `${path}.dependencyGraph`, ctx, beforeEventIds);
  validateImplicitEventDependencies(timeline.events, `${path}.events`, ctx);
}

function validateTrigger(trigger: unknown, path: string, ctx: ValidationContext): void {
  if (!isRecord(trigger)) {
    addIssue(ctx, "error", "TRIGGER_NOT_OBJECT", path, "Trigger must be an object.");
    return;
  }

  requireId(trigger.id, `${path}.id`, ctx, ctx.triggerIds, "trigger");
  if (!["onEnter", "onClick", "afterPrevious", "custom"].includes(String(trigger.type))) {
    addIssue(ctx, "error", "INVALID_TRIGGER_TYPE", `${path}.type`, "Trigger type is not supported.");
  }

  if (trigger.targetId !== undefined && typeof trigger.targetId === "string" && !ctx.objectIds.has(trigger.targetId)) {
    addIssue(ctx, "error", "UNKNOWN_TRIGGER_TARGET", `${path}.targetId`, "Trigger targetId must reference an object.");
  }
}

function validateAnimationEvent(event: unknown, path: string, ctx: ValidationContext): void {
  if (!isRecord(event)) {
    addIssue(ctx, "error", "EVENT_NOT_OBJECT", path, "Animation event must be an object.");
    return;
  }

  if (!EVENT_KINDS.has(String(event.kind))) {
    addIssue(ctx, "error", "INVALID_EVENT_KIND", `${path}.kind`, "Animation event kind is not supported.");
  }

  validateTimingStart(event.start, `${path}.start`, ctx);
  validateOptionalNonNegativeNumber(event.durationMs, `${path}.durationMs`, ctx);
  validateOptionalFiniteNumber(event.delayMs, `${path}.delayMs`, ctx);
  validateEasing(event.easing, `${path}.easing`, ctx);
  validateDependencyRefs(event.dependencies, `${path}.dependencies`, ctx);

  switch (event.kind) {
    case "setState":
      validateTargetReference(event.targetId, `${path}.targetId`, ctx);
      validateStateReference(event.stateId, `${path}.stateId`, ctx);
      break;
    case "property":
      validateTargetReference(event.targetId, `${path}.targetId`, ctx);
      requireString(event.property, `${path}.property`, ctx);
      if (!Object.prototype.hasOwnProperty.call(event, "to")) {
        addIssue(ctx, "error", "PROPERTY_TO_REQUIRED", `${path}.to`, "Property animation must include a to value.");
      }
      break;
    case "keyframes":
      validateTargetReference(event.targetId, `${path}.targetId`, ctx);
      validateKeyframeTracks(event.tracks, `${path}.tracks`, ctx);
      break;
    case "morph":
      validateMorphEvent(event, path, ctx);
      break;
    case "transition":
      validateTransition(event.transition, `${path}.transition`, ctx);
      validateOptionalSlideReference(event.fromSlideId, `${path}.fromSlideId`, ctx);
      validateOptionalSlideReference(event.toSlideId, `${path}.toSlideId`, ctx);
      break;
    case "media":
      validateTargetReference(event.targetId, `${path}.targetId`, ctx);
      if (!["play", "pause", "seek", "stop", "mute", "unmute"].includes(String(event.action))) {
        addIssue(ctx, "error", "INVALID_MEDIA_ACTION", `${path}.action`, "Media action is not supported.");
      }
      validateOptionalNonNegativeNumber(event.seekMs, `${path}.seekMs`, ctx);
      break;
    case "visibility":
      validateTargetReference(event.targetId, `${path}.targetId`, ctx);
      if (typeof event.visible !== "boolean") {
        addIssue(ctx, "error", "VISIBLE_BOOLEAN_REQUIRED", `${path}.visible`, "Visibility event must include a boolean visible value.");
      }
      break;
  }
}

function validateTimingStart(start: unknown, path: string, ctx: ValidationContext): void {
  if (start === undefined) {
    return;
  }

  if (!isRecord(start)) {
    addIssue(ctx, "error", "TIMING_START_NOT_OBJECT", path, "Timing start must be an object.");
    return;
  }

  if (!TIMING_START_TYPES.has(String(start.type))) {
    addIssue(ctx, "error", "INVALID_TIMING_START", `${path}.type`, "Timing start type is not supported.");
    return;
  }

  if (start.type === "absolute") {
    validateOptionalNonNegativeNumber(start.atMs, `${path}.atMs`, ctx);
  }

  if (["with", "after", "before"].includes(String(start.type))) {
    validateEventReference(start.eventId, `${path}.eventId`, ctx);
  }

  if (start.type === "onClick" && start.targetId !== undefined) {
    validateTargetReference(start.targetId, `${path}.targetId`, ctx);
  }

  if (start.type === "trigger") {
    validateTriggerReference(start.triggerId, `${path}.triggerId`, ctx);
  }

  validateOptionalFiniteNumber(start.offsetMs, `${path}.offsetMs`, ctx);
}

function validateDependencyRefs(dependencies: unknown, path: string, ctx: ValidationContext): void {
  if (dependencies === undefined) {
    return;
  }

  if (!Array.isArray(dependencies)) {
    addIssue(ctx, "error", "DEPENDENCIES_NOT_ARRAY", path, "Event dependencies must be an array.");
    return;
  }

  dependencies.forEach((dependency, index) => {
    const dependencyPath = `${path}[${index}]`;
    if (!isRecord(dependency)) {
      addIssue(ctx, "error", "DEPENDENCY_NOT_OBJECT", dependencyPath, "Event dependency must be an object.");
      return;
    }

    validateEventReference(dependency.eventId, `${dependencyPath}.eventId`, ctx);
    if (dependency.relation !== undefined && !["after", "with", "before"].includes(String(dependency.relation))) {
      addIssue(ctx, "error", "INVALID_DEPENDENCY_RELATION", `${dependencyPath}.relation`, "Dependency relation is not supported.");
    }
    validateOptionalFiniteNumber(dependency.offsetMs, `${dependencyPath}.offsetMs`, ctx);
  });
}

function validateImplicitEventDependencies(events: unknown[], path: string, ctx: ValidationContext): void {
  const edges: Array<[string, string, string]> = [];

  events.forEach((event, index) => {
    if (!isRecord(event) || typeof event.id !== "string") {
      return;
    }

    if (isRecord(event.start) && typeof event.start.eventId === "string") {
      edges.push([event.start.eventId, event.id, `${path}[${index}].start.eventId`]);
    }

    if (Array.isArray(event.dependencies)) {
      event.dependencies.forEach((dependency, dependencyIndex) => {
        if (isRecord(dependency) && typeof dependency.eventId === "string") {
          edges.push([dependency.eventId, event.id, `${path}[${index}].dependencies[${dependencyIndex}].eventId`]);
        }
      });
    }
  });

  reportCycles(edges, "$.deck.slides[].timeline.events", ctx);
}

function validateKeyframeTracks(tracks: unknown, path: string, ctx: ValidationContext): void {
  if (!Array.isArray(tracks) || tracks.length === 0) {
    addIssue(ctx, "error", "KEYFRAME_TRACKS_REQUIRED", path, "Keyframe animation must include at least one track.");
    return;
  }

  tracks.forEach((track, trackIndex) => {
    const trackPath = `${path}[${trackIndex}]`;
    if (!isRecord(track)) {
      addIssue(ctx, "error", "KEYFRAME_TRACK_NOT_OBJECT", trackPath, "Keyframe track must be an object.");
      return;
    }

    requireString(track.property, `${trackPath}.property`, ctx);
    if (!Array.isArray(track.keyframes) || track.keyframes.length === 0) {
      addIssue(ctx, "error", "KEYFRAMES_REQUIRED", `${trackPath}.keyframes`, "Keyframe track must include keyframes.");
      return;
    }

    let previousOffset = -Infinity;
    track.keyframes.forEach((keyframe, keyframeIndex) => {
      const keyframePath = `${trackPath}.keyframes[${keyframeIndex}]`;
      if (!isRecord(keyframe)) {
        addIssue(ctx, "error", "KEYFRAME_NOT_OBJECT", keyframePath, "Keyframe must be an object.");
        return;
      }

      if (!isNumberInRange(keyframe.offset, 0, 1)) {
        addIssue(ctx, "error", "INVALID_KEYFRAME_OFFSET", `${keyframePath}.offset`, "Keyframe offset must be a number from 0 to 1.");
      } else if (keyframe.offset < previousOffset) {
        addIssue(ctx, "error", "KEYFRAME_OFFSETS_UNSORTED", `${keyframePath}.offset`, "Keyframe offsets must be sorted ascending.");
      } else {
        previousOffset = keyframe.offset;
      }

      if (!Object.prototype.hasOwnProperty.call(keyframe, "value")) {
        addIssue(ctx, "error", "KEYFRAME_VALUE_REQUIRED", `${keyframePath}.value`, "Keyframe must include a value.");
      }
      validateEasing(keyframe.easing, `${keyframePath}.easing`, ctx);
    });
  });
}

function validateMorphEvent(event: Record<string, unknown>, path: string, ctx: ValidationContext): void {
  if (!MORPH_STRATEGIES.has(String(event.strategy))) {
    addIssue(ctx, "error", "INVALID_MORPH_STRATEGY", `${path}.strategy`, "Morph strategy must be morph or magicMove.");
  }

  validateMorphEndpoint(event.from, `${path}.from`, ctx);
  validateMorphEndpoint(event.to, `${path}.to`, ctx);
  validateMorphPairs(event.pairs, `${path}.pairs`, ctx);
}

function validateMorphEndpoint(endpoint: unknown, path: string, ctx: ValidationContext): void {
  if (!isRecord(endpoint)) {
    addIssue(ctx, "error", "MORPH_ENDPOINT_REQUIRED", path, "Morph endpoint must be an object.");
    return;
  }

  validateOptionalSlideReference(endpoint.slideId, `${path}.slideId`, ctx);
  if (endpoint.objectId !== undefined) {
    validateTargetReference(endpoint.objectId, `${path}.objectId`, ctx);
  }
  if (endpoint.stateId !== undefined) {
    validateStateReference(endpoint.stateId, `${path}.stateId`, ctx);
  }

  if (endpoint.objectId === undefined && endpoint.stateId === undefined && endpoint.snapshot === undefined) {
    addIssue(ctx, "error", "MORPH_ENDPOINT_EMPTY", path, "Morph endpoint must include objectId, stateId, or snapshot.");
  }
}

function validateMorphPairs(pairs: unknown, path: string, ctx: ValidationContext): void {
  if (pairs === undefined) {
    return;
  }

  if (!Array.isArray(pairs)) {
    addIssue(ctx, "error", "MORPH_PAIRS_NOT_ARRAY", path, "Morph pairs must be an array.");
    return;
  }

  pairs.forEach((pair, index) => {
    const pairPath = `${path}[${index}]`;
    if (!isRecord(pair)) {
      addIssue(ctx, "error", "MORPH_PAIR_NOT_OBJECT", pairPath, "Morph pair must be an object.");
      return;
    }

    validateTargetReference(pair.fromObjectId, `${pairPath}.fromObjectId`, ctx);
    validateTargetReference(pair.toObjectId, `${pairPath}.toObjectId`, ctx);
  });
}

function validateTransition(transition: unknown, path: string, ctx: ValidationContext): void {
  if (transition === undefined || transition === null) {
    return;
  }

  if (!isRecord(transition)) {
    addIssue(ctx, "error", "TRANSITION_NOT_OBJECT", path, "Transition must be an object or null.");
    return;
  }

  if (!TRANSITION_TYPES.has(String(transition.type))) {
    addIssue(ctx, "error", "INVALID_TRANSITION_TYPE", `${path}.type`, "Transition type is not supported.");
  }

  validateOptionalNonNegativeNumber(transition.durationMs, `${path}.durationMs`, ctx);
  validateEasing(transition.easing, `${path}.easing`, ctx);
  validateOptionalSlideReference(transition.fromSlideId, `${path}.fromSlideId`, ctx);
  validateOptionalSlideReference(transition.toSlideId, `${path}.toSlideId`, ctx);

  if ((transition.type === "morph" || transition.type === "magicMove") && transition.morph !== undefined) {
    if (!isRecord(transition.morph)) {
      addIssue(ctx, "error", "MORPH_TRANSITION_NOT_OBJECT", `${path}.morph`, "Morph transition options must be an object.");
    } else if (!MORPH_STRATEGIES.has(String(transition.morph.strategy))) {
      addIssue(ctx, "error", "INVALID_MORPH_TRANSITION_STRATEGY", `${path}.morph.strategy`, "Morph transition strategy must be morph or magicMove.");
    }
  }
}

function validateTimingDependencyGraph(
  graph: unknown,
  path: string,
  ctx: ValidationContext,
  eventIdsBeforeTimeline: Set<string>,
): void {
  if (graph === undefined) {
    return;
  }

  if (!isRecord(graph)) {
    addIssue(ctx, "error", "DEPENDENCY_GRAPH_NOT_OBJECT", path, "Timing dependency graph must be an object.");
    return;
  }

  const typedGraph = graph as unknown as TimingDependencyGraph;
  const nodeIds = new Set<string>();
  if (typedGraph.nodes !== undefined) {
    if (!Array.isArray(typedGraph.nodes)) {
      addIssue(ctx, "error", "DEPENDENCY_GRAPH_NODES_NOT_ARRAY", `${path}.nodes`, "Dependency graph nodes must be an array.");
    } else {
      typedGraph.nodes.forEach((node, index) => {
        if (!isRecord(node)) {
          addIssue(ctx, "error", "DEPENDENCY_GRAPH_NODE_NOT_OBJECT", `${path}.nodes[${index}]`, "Dependency graph node must be an object.");
          return;
        }

        const id = requireId(node.id, `${path}.nodes[${index}].id`, ctx, nodeIds, "dependency node");
        if (typeof node.eventId === "string" && !ctx.eventIds.has(node.eventId)) {
          addIssue(ctx, "error", "UNKNOWN_DEPENDENCY_NODE_EVENT", `${path}.nodes[${index}].eventId`, "Dependency graph node eventId must reference an event.");
        }
        if (id !== undefined && eventIdsBeforeTimeline.has(id)) {
          addIssue(ctx, "warning", "DEPENDENCY_NODE_ID_COLLISION", `${path}.nodes[${index}].id`, "Dependency node id collides with an event id from another timeline.");
        }
      });
    }
  }

  if (!Array.isArray(typedGraph.edges)) {
    addIssue(ctx, "error", "DEPENDENCY_GRAPH_EDGES_NOT_ARRAY", `${path}.edges`, "Dependency graph edges must be an array.");
    return;
  }

  const edgeRefs: Array<[string, string, string]> = [];
  typedGraph.edges.forEach((edge, index) => {
    const edgePath = `${path}.edges[${index}]`;
    if (!isRecord(edge)) {
      addIssue(ctx, "error", "DEPENDENCY_GRAPH_EDGE_NOT_OBJECT", edgePath, "Dependency graph edge must be an object.");
      return;
    }

    const from = requireString(edge.from, `${edgePath}.from`, ctx);
    const to = requireString(edge.to, `${edgePath}.to`, ctx);
    if (!["after", "with", "before", "blocks", "triggers"].includes(String(edge.relation))) {
      addIssue(ctx, "error", "INVALID_DEPENDENCY_EDGE_RELATION", `${edgePath}.relation`, "Dependency graph edge relation is not supported.");
    }
    validateOptionalFiniteNumber(edge.offsetMs, `${edgePath}.offsetMs`, ctx);

    if (from !== undefined && to !== undefined) {
      const knownFrom = nodeIds.has(from) || ctx.eventIds.has(from);
      const knownTo = nodeIds.has(to) || ctx.eventIds.has(to);
      if (!knownFrom) {
        addIssue(ctx, "error", "UNKNOWN_DEPENDENCY_EDGE_FROM", `${edgePath}.from`, "Dependency edge from must reference a graph node or event id.");
      }
      if (!knownTo) {
        addIssue(ctx, "error", "UNKNOWN_DEPENDENCY_EDGE_TO", `${edgePath}.to`, "Dependency edge to must reference a graph node or event id.");
      }
      edgeRefs.push([from, to, edgePath]);
    }
  });

  reportCycles(edgeRefs, path, ctx);
}

function reportCycles(edges: Array<[string, string, string]>, path: string, ctx: ValidationContext): void {
  const graph = new Map<string, Array<{ to: string; path: string }>>();
  edges.forEach(([from, to, edgePath]) => {
    const existing = graph.get(from) ?? [];
    existing.push({ to, path: edgePath });
    graph.set(from, existing);
    if (!graph.has(to)) {
      graph.set(to, []);
    }
  });

  const visiting = new Set<string>();
  const visited = new Set<string>();
  const reported = new Set<string>();

  const visit = (node: string, stack: string[]): void => {
    if (visiting.has(node)) {
      const cycleStart = stack.indexOf(node);
      const cycle = [...stack.slice(Math.max(0, cycleStart)), node].join(" -> ");
      if (!reported.has(cycle)) {
        addIssue(ctx, "error", "TIMING_DEPENDENCY_CYCLE", path, `Timing dependencies contain a cycle: ${cycle}.`);
        reported.add(cycle);
      }
      return;
    }

    if (visited.has(node)) {
      return;
    }

    visiting.add(node);
    const nextStack = [...stack, node];
    for (const edge of graph.get(node) ?? []) {
      visit(edge.to, nextStack);
    }
    visiting.delete(node);
    visited.add(node);
  };

  for (const node of graph.keys()) {
    visit(node, []);
  }
}

function validateEasing(easing: unknown, path: string, ctx: ValidationContext): void {
  if (easing === undefined) {
    return;
  }

  if (typeof easing === "string") {
    if (!EASING_NAMES.has(easing)) {
      addIssue(ctx, "error", "INVALID_EASING_NAME", path, "Easing name is not supported.");
    }
    return;
  }

  if (!isRecord(easing)) {
    addIssue(ctx, "error", "EASING_NOT_OBJECT", path, "Easing must be a supported name or object.");
    return;
  }

  switch (easing.type) {
    case "cubicBezier":
      requireFiniteNumber(easing.x1, `${path}.x1`, ctx);
      requireFiniteNumber(easing.y1, `${path}.y1`, ctx);
      requireFiniteNumber(easing.x2, `${path}.x2`, ctx);
      requireFiniteNumber(easing.y2, `${path}.y2`, ctx);
      break;
    case "spring":
      validateOptionalPositiveNumber(easing.mass, `${path}.mass`, ctx);
      validateOptionalPositiveNumber(easing.stiffness, `${path}.stiffness`, ctx);
      validateOptionalPositiveNumber(easing.damping, `${path}.damping`, ctx);
      validateOptionalFiniteNumber(easing.velocity, `${path}.velocity`, ctx);
      break;
    case "steps":
      if (!isPositiveInteger(easing.count)) {
        addIssue(ctx, "error", "INVALID_STEPS_COUNT", `${path}.count`, "Steps easing count must be a positive integer.");
      }
      if (easing.position !== undefined && !["start", "end"].includes(String(easing.position))) {
        addIssue(ctx, "error", "INVALID_STEPS_POSITION", `${path}.position`, "Steps easing position must be start or end.");
      }
      break;
    case "custom":
      requireString(easing.name, `${path}.name`, ctx);
      break;
    default:
      addIssue(ctx, "error", "INVALID_EASING_TYPE", `${path}.type`, "Easing type is not supported.");
      break;
  }
}

function validateConversionReport(report: unknown, path: string, ctx: ValidationContext): void {
  if (report === undefined) {
    return;
  }

  if (!isRecord(report)) {
    addIssue(ctx, "error", "CONVERSION_REPORT_NOT_OBJECT", path, "Conversion report must be an object.");
    return;
  }

  if (!["success", "partial", "failed"].includes(String(report.status))) {
    addIssue(ctx, "error", "INVALID_CONVERSION_STATUS", `${path}.status`, "Conversion report status must be success, partial, or failed.");
  }

  if (!Array.isArray(report.messages)) {
    addIssue(ctx, "error", "CONVERSION_MESSAGES_NOT_ARRAY", `${path}.messages`, "Conversion report messages must be an array.");
  } else {
    report.messages.forEach((message, index) => {
      const messagePath = `${path}.messages[${index}]`;
      if (!isRecord(message)) {
        addIssue(ctx, "error", "CONVERSION_MESSAGE_NOT_OBJECT", messagePath, "Conversion message must be an object.");
        return;
      }
      if (!["info", "warning", "error"].includes(String(message.severity))) {
        addIssue(ctx, "error", "INVALID_CONVERSION_MESSAGE_SEVERITY", `${messagePath}.severity`, "Conversion message severity is invalid.");
      }
      requireString(message.code, `${messagePath}.code`, ctx);
      requireString(message.message, `${messagePath}.message`, ctx);
    });
  }
}

function validateRect(rect: unknown, path: string, ctx: ValidationContext, required: boolean): void {
  if (rect === undefined) {
    if (required) {
      addIssue(ctx, "error", "RECT_REQUIRED", path, "Rectangle is required.");
    }
    return;
  }

  if (!isRecord(rect)) {
    addIssue(ctx, "error", "RECT_NOT_OBJECT", path, "Rectangle must be an object.");
    return;
  }

  requireFiniteNumber(rect.x, `${path}.x`, ctx);
  requireFiniteNumber(rect.y, `${path}.y`, ctx);
  requireFiniteNumber(rect.width, `${path}.width`, ctx);
  requireFiniteNumber(rect.height, `${path}.height`, ctx);
}

function validateTransform(transform: unknown, path: string, ctx: ValidationContext): void {
  if (transform === undefined) {
    return;
  }

  if (!isRecord(transform)) {
    addIssue(ctx, "error", "TRANSFORM_NOT_OBJECT", path, "Transform must be an object.");
    return;
  }

  [
    "translateX",
    "translateY",
    "scaleX",
    "scaleY",
    "rotateDeg",
    "skewXDeg",
    "skewYDeg",
  ].forEach((key) => validateOptionalFiniteNumber(transform[key], `${path}.${key}`, ctx));
}

function validateSource(source: unknown, path: string, ctx: ValidationContext): void {
  if (!isRecord(source)) {
    addIssue(ctx, "error", "SOURCE_REQUIRED", path, "Object source must be an object.");
    return;
  }

  if (source.assetId === undefined && source.uri === undefined && source.dataUri === undefined) {
    addIssue(ctx, "error", "SOURCE_EMPTY", path, "Object source must include assetId, uri, or dataUri.");
  }

  if (typeof source.assetId === "string" && !ctx.assetIds.has(source.assetId)) {
    addIssue(ctx, "error", "UNKNOWN_ASSET_ID", `${path}.assetId`, "Source assetId must reference a deck asset.");
  }
}

function validateTextContent(text: unknown, path: string, ctx: ValidationContext, required: boolean): void {
  if (text === undefined) {
    if (required) {
      addIssue(ctx, "error", "TEXT_REQUIRED", path, "Text object must include text content.");
    }
    return;
  }

  if (!isRecord(text)) {
    addIssue(ctx, "error", "TEXT_NOT_OBJECT", path, "Text content must be an object.");
    return;
  }

  if (text.plainText === undefined && text.runs === undefined && text.paragraphs === undefined) {
    addIssue(ctx, "warning", "TEXT_EMPTY", path, "Text content does not include plainText, runs, or paragraphs.");
  }
}

function validatePathCommands(commands: unknown, path: string, ctx: ValidationContext): void {
  if (!Array.isArray(commands) || commands.length === 0) {
    addIssue(ctx, "error", "PATH_COMMANDS_REQUIRED", path, "Path object must include path commands.");
    return;
  }

  commands.forEach((command, index) => {
    const commandPath = `${path}[${index}]`;
    if (!isRecord(command)) {
      addIssue(ctx, "error", "PATH_COMMAND_NOT_OBJECT", commandPath, "Path command must be an object.");
      return;
    }

    if (!["moveTo", "lineTo", "curveTo", "quadTo", "close"].includes(String(command.command))) {
      addIssue(ctx, "error", "INVALID_PATH_COMMAND", `${commandPath}.command`, "Path command is not supported.");
    }
  });
}

function validateTargetReference(value: unknown, path: string, ctx: ValidationContext): string | undefined {
  const id = requireId(value, path, ctx, undefined, "target");
  if (id !== undefined && !ctx.objectIds.has(id)) {
    addIssue(ctx, "error", "UNKNOWN_TARGET_ID", path, "Target id must reference an object.");
  }
  return id;
}

function validateStateReference(value: unknown, path: string, ctx: ValidationContext): string | undefined {
  const id = requireId(value, path, ctx, undefined, "state");
  if (id !== undefined && !ctx.stateIds.has(id)) {
    addIssue(ctx, "error", "UNKNOWN_STATE_ID", path, "State id must reference a slide state.");
  }
  return id;
}

function validateEventReference(value: unknown, path: string, ctx: ValidationContext): string | undefined {
  const id = requireId(value, path, ctx, undefined, "event");
  if (id !== undefined && !ctx.eventIds.has(id)) {
    addIssue(ctx, "error", "UNKNOWN_EVENT_ID", path, "Event id must reference a timeline event.");
  }
  return id;
}

function validateTriggerReference(value: unknown, path: string, ctx: ValidationContext): string | undefined {
  const id = requireId(value, path, ctx, undefined, "trigger");
  if (id !== undefined && !ctx.triggerIds.has(id)) {
    addIssue(ctx, "error", "UNKNOWN_TRIGGER_ID", path, "Trigger id must reference a timeline trigger.");
  }
  return id;
}

function validateOptionalSlideReference(value: unknown, path: string, ctx: ValidationContext): void {
  if (value === undefined) {
    return;
  }

  const id = requireId(value, path, ctx, undefined, "slide");
  if (id !== undefined && !ctx.slideIds.has(id)) {
    addIssue(ctx, "error", "UNKNOWN_SLIDE_ID", path, "Slide id must reference a deck slide.");
  }
}

function requireId(
  value: unknown,
  path: string,
  ctx: ValidationContext,
  uniquenessSet: Set<string> | undefined,
  label: string,
): string | undefined {
  const id = requireString(value, path, ctx);
  if (id === undefined) {
    return undefined;
  }

  if (id.trim().length === 0) {
    addIssue(ctx, "error", "EMPTY_ID", path, `${label} id cannot be empty.`);
    return undefined;
  }

  if (uniquenessSet !== undefined) {
    if (uniquenessSet.has(id)) {
      addIssue(ctx, "error", "DUPLICATE_ID", path, `Duplicate ${label} id '${id}'.`);
    } else {
      uniquenessSet.add(id);
    }
  }

  return id;
}

function requireString(value: unknown, path: string, ctx: ValidationContext): string | undefined {
  if (typeof value !== "string") {
    addIssue(ctx, "error", "STRING_REQUIRED", path, "Expected a string.");
    return undefined;
  }
  return value;
}

function requireFiniteNumber(value: unknown, path: string, ctx: ValidationContext): number | undefined {
  if (!isFiniteNumber(value)) {
    addIssue(ctx, "error", "NUMBER_REQUIRED", path, "Expected a finite number.");
    return undefined;
  }
  return value;
}

function validateOptionalFiniteNumber(value: unknown, path: string, ctx: ValidationContext): void {
  if (value !== undefined && !isFiniteNumber(value)) {
    addIssue(ctx, "error", "NUMBER_REQUIRED", path, "Expected a finite number.");
  }
}

function validateOptionalNonNegativeNumber(value: unknown, path: string, ctx: ValidationContext): void {
  if (value === undefined) {
    return;
  }
  if (!isFiniteNumber(value) || value < 0) {
    addIssue(ctx, "error", "NON_NEGATIVE_NUMBER_REQUIRED", path, "Expected a non-negative finite number.");
  }
}

function validateOptionalPositiveNumber(value: unknown, path: string, ctx: ValidationContext): void {
  if (value === undefined) {
    return;
  }
  if (!isFiniteNumber(value) || value <= 0) {
    addIssue(ctx, "error", "POSITIVE_NUMBER_REQUIRED", path, "Expected a positive finite number.");
  }
}

function validateOptionalRatio(value: unknown, path: string, ctx: ValidationContext): void {
  if (value !== undefined && !isNumberInRange(value, 0, 1)) {
    addIssue(ctx, "error", "RATIO_REQUIRED", path, "Expected a number from 0 to 1.");
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function isNumberInRange(value: unknown, min: number, max: number): value is number {
  return isFiniteNumber(value) && value >= min && value <= max;
}

function isPositiveInteger(value: unknown): value is number {
  return Number.isInteger(value) && Number(value) > 0;
}

function addIssue(
  ctx: ValidationContext,
  severity: ValidationSeverity,
  code: string,
  path: string,
  message: string,
): void {
  const issue = { severity, code, path, message };
  if (severity === "error") {
    ctx.errors.push(issue);
  } else {
    ctx.warnings.push(issue);
  }
}

function toResult(input: unknown, ctx: ValidationContext): ValidationResult<KeyMorphIR> {
  return {
    valid: ctx.errors.length === 0,
    value: ctx.errors.length === 0 ? (input as KeyMorphIR) : undefined,
    errors: ctx.errors,
    warnings: ctx.warnings,
  };
}
