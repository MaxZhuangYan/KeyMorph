import type {
  AnimationEvent,
  Color,
  DeckIR,
  Easing,
  Fill,
  IRObject,
  JSONRecord,
  KeyframeAnimationEvent,
  KeyframeTrack,
  MorphProperty,
  ObjectSource,
  MorphTransitionOptions,
  ObjectStateProperties,
  ShapeObject,
  Slide,
  SlideTransition,
  Stroke,
  TextObject,
  Transform2D
} from "../ir/index.ts";

export interface HtmlRuntimeOptions {
  controls?: boolean;
  initialSlideIndex?: number;
  stageScale?: number;
}

export interface DeckSlideTimelineSpan {
  slideIndex: number;
  slideId: string;
  startMs: number;
  transitionStartMs: number;
  contentStartMs: number;
  endMs: number;
  transitionDurationMs: number;
  slideDurationMs: number;
}

export interface DeckTimelinePlan {
  durationMs: number;
  slides: DeckSlideTimelineSpan[];
}

export interface DeckTimeResolution {
  globalTimeMs: number;
  slideIndex: number;
  slideId: string;
  slideTimeMs: number;
  slideStartMs: number;
  slideEndMs: number;
  inTransition: boolean;
  transitionProgress: number;
  transitionDurationMs: number;
  previousSlideIndex?: number;
}

export interface RuntimeTimingWarning {
  code: "cycle" | "unresolved";
  eventIds: string[];
  message: string;
}

export interface SlideTimingPlan {
  starts: Record<string, number>;
  order: string[];
  warnings: RuntimeTimingWarning[];
}

export interface SlideObjectStateResolutionOptions {
  effectiveGroupStates?: boolean;
}

export interface MorphTransitionStateResolution {
  pairs: { fromObjectId: string; toObjectId: string; morphKey?: string }[];
  states: Map<string, ObjectStateProperties>;
}

export type RuntimeTimelineEventPhase = "before" | "active" | "after" | "instant";

export interface RuntimeTimelineEventDiagnostic {
  eventId: string;
  kind: AnimationEvent["kind"];
  label?: string;
  targetId?: string;
  property?: string;
  transitionType?: string;
  startMs: number;
  endMs: number;
  durationMs: number;
  phase: RuntimeTimelineEventPhase;
  active: boolean;
  applied: boolean;
  fill?: AnimationEvent["fill"];
  rawProgress?: number;
  easedProgress?: number;
  appliedProgress?: number;
}

export interface RuntimeSlideObjectStateSnapshot {
  slideIndex: number;
  slideId: string;
  phase: "current" | "previous" | "inactive";
  timeMs: number;
  states: Record<string, ObjectStateProperties>;
  events: RuntimeTimelineEventDiagnostic[];
  timingPlan: SlideTimingPlan;
}

export interface RuntimeFrameSnapshotOptions extends SlideObjectStateResolutionOptions {
  includeInactiveSlides?: boolean;
}

export interface RuntimeFrameSnapshot {
  deckId: string;
  globalTimeMs: number;
  resolution: DeckTimeResolution;
  objects: Record<string, ObjectStateProperties>;
  slides: RuntimeSlideObjectStateSnapshot[];
  transition?: {
    previousSlideIndex?: number;
    previousSlideId?: string;
    currentSlideIndex: number;
    currentSlideId: string;
    progress: number;
    pairs: { fromObjectId: string; toObjectId: string; morphKey?: string }[];
    states: Record<string, ObjectStateProperties>;
  };
  warnings: RuntimeTimingWarning[];
}

type NormalizedRuntimeProperty =
  | "bounds"
  | "transform"
  | "bounds.x"
  | "bounds.y"
  | "bounds.width"
  | "bounds.height"
  | "transform.translateX"
  | "transform.translateY"
  | "transform.scale"
  | "transform.scaleX"
  | "transform.scaleY"
  | "transform.rotateDeg"
  | "transform.skewXDeg"
  | "transform.skewYDeg"
  | "opacity"
  | "visible"
  | "filter.blurPx"
  | "style.fill"
  | "style.stroke"
  | "text"
  | "crop"
  | string;

const DEFAULT_SLIDE_DURATION_MS = 2500;
const DEFAULT_TRANSITION_DURATION_MS = 600;

export function renderHtmlDocument(deck: DeckIR, options: HtmlRuntimeOptions = {}): string {
  const payload = JSON.stringify(deck).replaceAll("<", "\\u003c");
  const controls = options.controls ?? true;
  const initialSlideIndex = options.initialSlideIndex ?? 0;
  const runtimeOptions = JSON.stringify({ stageScale: options.stageScale }).replaceAll("<", "\\u003c");
  const exportMode = typeof options.stageScale === "number" && Number.isFinite(options.stageScale);

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(deck.metadata?.title ?? deck.deck.title ?? "KeyMorph Runtime")}</title>
  <style>
    :root { color-scheme: light; font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
    body { margin: 0; min-height: 100vh; background: ${exportMode ? "transparent" : "#202124"}; color: #f8fafc; display: grid; grid-template-rows: 1fr auto; }
    #viewport { min-height: 0; display: grid; place-items: ${exportMode ? "start" : "center"}; padding: ${exportMode ? "0" : "24px"}; overflow: hidden; }
    #stage { position: relative; overflow: hidden; background: #fff; box-shadow: ${exportMode ? "none" : "0 18px 55px rgba(0,0,0,.36)"}; transform-origin: top left; }
    .km-slide-layer { position: absolute; inset: 0; overflow: hidden; transform-origin: center center; }
    .km-object { position: absolute; box-sizing: border-box; overflow: hidden; transform-origin: center center; }
    .km-text, .km-shape-text { white-space: pre-wrap; display: flex; align-items: flex-start; }
    .km-text-char { display: inline-block; white-space: pre; }
    .km-image, .km-media { object-fit: contain; user-select: none; -webkit-user-drag: none; }
    .km-controls { display: ${controls ? "grid" : "none"}; grid-template-columns: auto auto auto 1fr auto auto; gap: 10px; align-items: center; padding: 12px 16px; background: #111827; border-top: 1px solid rgba(255,255,255,.12); }
    button { border: 1px solid rgba(255,255,255,.18); color: #f8fafc; background: #1f2937; border-radius: 6px; padding: 8px 11px; font: inherit; cursor: pointer; }
    button:hover { background: #374151; }
    input[type="range"] { width: 100%; }
    #status { font-variant-numeric: tabular-nums; color: #cbd5e1; min-width: 190px; text-align: right; }
  </style>
</head>
<body>
  <main id="viewport"><div id="stage" aria-label="KeyMorph slide stage"></div></main>
  <div class="km-controls">
    <button id="prev" type="button">Prev</button>
    <button id="play" type="button">Play</button>
    <button id="next" type="button">Next</button>
    <input id="scrub" type="range" min="0" max="1" step="1" value="0" aria-label="Timeline">
    <button id="step" type="button">Step</button>
    <span id="status"></span>
  </div>
  <script>
    window.__KEYMORPH_DECK__ = ${payload};
    window.__KEYMORPH_INITIAL_SLIDE__ = ${JSON.stringify(initialSlideIndex)};
    window.__KEYMORPH_RUNTIME_OPTIONS__ = ${runtimeOptions};
  </script>
  <script>${runtimeScript()}</script>
</body>
</html>`;
}

export function renderSlideMarkup(slide: Slide, deck?: DeckIR): string {
  return slide.objects.map((object) => renderObjectMarkup(object, deck)).join("");
}

export function getSlideDurationMs(slide: Slide | undefined): number {
  return Math.max(1, Number(slide?.timeline?.durationMs ?? DEFAULT_SLIDE_DURATION_MS));
}

export function getSlideTransitionDurationMs(slide: Slide | undefined, slideIndex = 0): number {
  if (slideIndex <= 0) return 0;
  return transitionDurationMs(slide?.transition);
}

export function createDeckTimeline(deck: DeckIR): DeckTimelinePlan {
  const slides: DeckSlideTimelineSpan[] = [];
  let cursor = 0;

  deck.deck.slides.forEach((slide, slideIndex) => {
    const transitionDurationMs = getSlideTransitionDurationMs(slide, slideIndex);
    const slideDurationMs = getSlideDurationMs(slide);
    const transitionStartMs = cursor;
    const contentStartMs = transitionStartMs + transitionDurationMs;
    const endMs = contentStartMs + slideDurationMs;

    slides.push({
      slideIndex,
      slideId: slide.id,
      startMs: transitionStartMs,
      transitionStartMs,
      contentStartMs,
      endMs,
      transitionDurationMs,
      slideDurationMs
    });
    cursor = endMs;
  });

  return { durationMs: Math.max(1, cursor), slides };
}

export function resolveDeckTime(deck: DeckIR, timeMs: number): DeckTimeResolution {
  const timeline = createDeckTimeline(deck);
  const globalTimeMs = clamp(Number(timeMs) || 0, 0, timeline.durationMs);
  const fallback = timeline.slides[timeline.slides.length - 1];
  const span = timeline.slides.find((candidate) => globalTimeMs < candidate.endMs) ?? fallback;

  if (!span) {
    return {
      globalTimeMs: 0,
      slideIndex: 0,
      slideId: "",
      slideTimeMs: 0,
      slideStartMs: 0,
      slideEndMs: 0,
      inTransition: false,
      transitionProgress: 1,
      transitionDurationMs: 0
    };
  }

  const inTransition = span.transitionDurationMs > 0 && globalTimeMs < span.contentStartMs;
  const transitionProgress = inTransition
    ? clamp((globalTimeMs - span.transitionStartMs) / span.transitionDurationMs, 0, 1)
    : 1;
  const slideTimeMs = inTransition ? 0 : clamp(globalTimeMs - span.contentStartMs, 0, span.slideDurationMs);

  return {
    globalTimeMs,
    slideIndex: span.slideIndex,
    slideId: span.slideId,
    slideTimeMs,
    slideStartMs: span.contentStartMs,
    slideEndMs: span.endMs,
    inTransition,
    transitionProgress,
    transitionDurationMs: span.transitionDurationMs,
    previousSlideIndex: inTransition ? span.slideIndex - 1 : undefined
  };
}

export function resolveSlideObjectStates(
  deck: DeckIR,
  slide: Slide,
  timeMs: number,
  options: SlideObjectStateResolutionOptions = {}
): Map<string, ObjectStateProperties> {
  const states = new Map<string, ObjectStateProperties>();
  const objects = flattenObjects(slide.objects);
  for (const object of objects) states.set(object.id, initialObjectState(object));

  const starts = computeEventStarts(slide);
  const events = [...(slide.timeline?.events ?? [])].sort((a, b) => (starts.get(a.id) ?? 0) - (starts.get(b.id) ?? 0));
  for (const event of events) {
    if (event.kind === "setState") {
      const progress = eventProgress(event, starts.get(event.id) ?? 0, timeMs);
      if (progress === undefined) continue;
      const stateRecord = slide.states?.find((item) => item.id === event.stateId);
      if (stateRecord) states.set(event.targetId, mergeObjectStateProperties(states.get(event.targetId) ?? {}, stateRecord.properties));
    } else if (event.kind === "visibility") {
      const progress = eventProgress(event, starts.get(event.id) ?? 0, timeMs);
      if (progress === undefined) continue;
      const current = mergeObjectStateProperties(states.get(event.targetId) ?? {}, {});
      current.visible = progress < 1 && event.fill === "backwards" ? !event.visible : event.visible;
      states.set(event.targetId, current);
    } else if (event.kind === "property") {
      const progress = eventProgress(event, starts.get(event.id) ?? 0, timeMs);
      if (progress === undefined) continue;
      const current = mergeObjectStateProperties(states.get(event.targetId) ?? {}, {});
      const from = event.from ?? getRuntimeProperty(current, event.property);
      setRuntimeProperty(current, event.property, interpolateValue(from, event.to, progress, interpolationModeForProperty(event.property, event.interpolation)));
      states.set(event.targetId, current);
    } else if (event.kind === "keyframes") {
      const progress = eventProgress(event, starts.get(event.id) ?? 0, timeMs);
      if (progress === undefined) continue;
      const current = mergeObjectStateProperties(states.get(event.targetId) ?? {}, {});
      for (const track of event.tracks) {
        setRuntimeProperty(current, track.property, keyframeTrackValue(track, progress));
      }
      states.set(event.targetId, current);
    } else if (event.kind === "morph") {
      const progress = eventProgress(event, starts.get(event.id) ?? 0, timeMs);
      if (progress === undefined) continue;
      const pairs = event.pairs?.length
        ? event.pairs
        : event.from?.objectId && event.to?.objectId
          ? [{ fromObjectId: event.from.objectId, toObjectId: event.to.objectId }]
          : [];
      for (const pair of pairs) {
        const targetId = pair.toObjectId || event.to?.objectId || pair.fromObjectId;
        const targetObject = objectById(slide, targetId);
        if (!targetObject) continue;
        const fromSlide = event.from?.slideId ? slideById(deck, event.from.slideId) : slide;
        const toSlide = event.to?.slideId ? slideById(deck, event.to.slideId) : slide;
        const fromState = stateForEndpoint(deck, { ...(event.from || {}), objectId: event.from?.objectId || pair.fromObjectId }, fromSlide, pair.fromObjectId);
        const toState =
          stateForEndpoint(deck, { ...(event.to || {}), objectId: event.to?.objectId || pair.toObjectId }, toSlide, pair.toObjectId) ||
          states.get(targetId) ||
          initialObjectState(targetObject);
        states.set(targetId, interpolateObjectState(fromState, toState, progress, event.properties));
      }
    } else if (event.kind === "media") {
      const progress = eventProgress(event, starts.get(event.id) ?? 0, timeMs);
      if (progress === undefined) continue;
      const current = mergeObjectStateProperties(states.get(event.targetId) ?? {}, {});
      current.media = { ...(current.media || {}) };
      if (event.action === "seek") current.media.startMs = event.seekMs || 0;
      if (event.action === "mute") current.media.muted = true;
      if (event.action === "unmute") current.media.muted = false;
      states.set(event.targetId, current);
    }
  }

  if (options.effectiveGroupStates) {
    applyEffectiveGroupStates(slide.objects, states);
  }

  return states;
}

export function resolveMorphTransitionObjectStates(
  deck: DeckIR,
  previousSlide: Slide | undefined,
  currentSlide: Slide | undefined,
  transition: SlideTransition | undefined | null,
  progress: number
): MorphTransitionStateResolution {
  const options = transition?.morph;
  const easedProgress = easeProgressValue(transition?.easing, progress);
  const pairs = inferMorphTransitionPairs(previousSlide, currentSlide, options);
  const properties = options?.properties ?? ["bounds", "transform", "opacity", "fill", "stroke"];
  const previousStates = previousSlide
    ? resolveSlideObjectStates(deck, previousSlide, getSlideDurationMs(previousSlide), { effectiveGroupStates: true })
    : new Map<string, ObjectStateProperties>();
  const currentStates = currentSlide
    ? resolveSlideObjectStates(deck, currentSlide, 0, { effectiveGroupStates: true })
    : new Map<string, ObjectStateProperties>();
  const states = new Map<string, ObjectStateProperties>();

  for (const pair of pairs) {
    const toObject = objectById(currentSlide, pair.toObjectId);
    if (!toObject) continue;
    const fromState = previousStates.get(pair.fromObjectId) || initialObjectState(objectById(previousSlide, pair.fromObjectId) || toObject);
    const toState = currentStates.get(pair.toObjectId) || initialObjectState(toObject);
    states.set(pair.toObjectId, interpolateObjectState(fromState, toState, easedProgress, properties));
  }

  return { pairs, states };
}

export function resolveTimelineEventDiagnostics(slide: Slide | undefined, timeMs: number): RuntimeTimelineEventDiagnostic[] {
  const timingPlan = createSlideTimingPlan(slide);
  const starts = timingPlan.starts;
  return (slide?.timeline?.events ?? []).map((event) => {
    const startMs = starts[event.id] ?? 0;
    const durationMs = eventDuration(event);
    const endMs = startMs + durationMs;
    const appliedProgress = eventProgress(event, startMs, timeMs);
    const rawProgress =
      durationMs === 0
        ? timeMs >= startMs
          ? 1
          : undefined
        : timeMs >= startMs && timeMs <= endMs
          ? clamp((timeMs - startMs) / durationMs, 0, 1)
          : undefined;
    const phase: RuntimeTimelineEventPhase =
      durationMs === 0 && timeMs >= startMs
        ? "instant"
        : timeMs < startMs
          ? "before"
          : timeMs > endMs
            ? "after"
            : "active";
    const targetId = "targetId" in event ? event.targetId : undefined;
    const property = event.kind === "property" ? event.property : undefined;
    const transitionType = event.kind === "transition" ? event.transition.type : undefined;

    return {
      eventId: event.id,
      kind: event.kind,
      label: event.label,
      targetId,
      property,
      transitionType,
      startMs,
      endMs,
      durationMs,
      phase,
      active: phase === "active" || phase === "instant",
      applied: appliedProgress !== undefined,
      fill: event.fill,
      rawProgress,
      easedProgress: rawProgress === undefined ? undefined : easeProgressValue(event.easing, rawProgress),
      appliedProgress
    };
  });
}

export function createRuntimeFrameSnapshot(
  deck: DeckIR,
  timeMs: number,
  options: RuntimeFrameSnapshotOptions = {}
): RuntimeFrameSnapshot {
  const resolution = resolveDeckTime(deck, timeMs);
  const includeInactiveSlides = options.includeInactiveSlides ?? false;
  const stateOptions = { effectiveGroupStates: options.effectiveGroupStates ?? true };
  const activeSlideIndexes = new Set<number>([resolution.slideIndex]);
  if (resolution.previousSlideIndex !== undefined) activeSlideIndexes.add(resolution.previousSlideIndex);
  const slideIndexes = includeInactiveSlides
    ? deck.deck.slides.map((_, slideIndex) => slideIndex)
    : Array.from(activeSlideIndexes).sort((a, b) => a - b);

  const slides = slideIndexes
    .map((slideIndex) => {
      const slide = deck.deck.slides[slideIndex];
      if (!slide) return undefined;
      const phase =
        slideIndex === resolution.slideIndex
          ? "current"
          : slideIndex === resolution.previousSlideIndex
            ? "previous"
            : "inactive";
      const slideTimeMs =
        phase === "current"
          ? resolution.slideTimeMs
          : phase === "previous"
            ? getSlideDurationMs(slide)
            : 0;
      return {
        slideIndex,
        slideId: slide.id,
        phase,
        timeMs: slideTimeMs,
        states: serializeObjectStates(resolveSlideObjectStates(deck, slide, slideTimeMs, stateOptions)),
        events: resolveTimelineEventDiagnostics(slide, slideTimeMs),
        timingPlan: createSlideTimingPlan(slide)
      } satisfies RuntimeSlideObjectStateSnapshot;
    })
    .filter((slide): slide is RuntimeSlideObjectStateSnapshot => Boolean(slide));

  const objects: Record<string, ObjectStateProperties> = {};
  for (const slide of slides) {
    if (slide.phase === "inactive") continue;
    for (const [objectId, state] of Object.entries(slide.states)) {
      objects[objects[objectId] ? `${slide.slideId}:${objectId}` : objectId] = state;
    }
  }

  const currentSlide = deck.deck.slides[resolution.slideIndex];
  const previousSlide =
    resolution.previousSlideIndex === undefined ? undefined : deck.deck.slides[resolution.previousSlideIndex];
  const transitionStates =
    resolution.inTransition && currentSlide
      ? resolveMorphTransitionObjectStates(deck, previousSlide, currentSlide, currentSlide.transition, resolution.transitionProgress)
      : undefined;

  return {
    deckId: deck.deck.id,
    globalTimeMs: resolution.globalTimeMs,
    resolution,
    objects,
    slides,
    transition:
      resolution.inTransition && currentSlide
        ? {
            previousSlideIndex: resolution.previousSlideIndex,
            previousSlideId: previousSlide?.id,
            currentSlideIndex: resolution.slideIndex,
            currentSlideId: currentSlide.id,
            progress: resolution.transitionProgress,
            pairs: transitionStates?.pairs ?? [],
            states: serializeObjectStates(transitionStates?.states ?? new Map<string, ObjectStateProperties>())
          }
        : undefined,
    warnings: slides.flatMap((slide) => slide.timingPlan.warnings)
  };
}

function runtimeScript(): string {
  return `
(() => {
  const deck = window.__KEYMORPH_DECK__;
  const runtimeOptions = window.__KEYMORPH_RUNTIME_OPTIONS__ || {};
  const stage = document.getElementById("stage");
  const viewport = document.getElementById("viewport");
  const playButton = document.getElementById("play");
  const prevButton = document.getElementById("prev");
  const nextButton = document.getElementById("next");
  const stepButton = document.getElementById("step");
  const scrub = document.getElementById("scrub");
  const status = document.getElementById("status");
  const DEFAULT_SLIDE_DURATION_MS = ${DEFAULT_SLIDE_DURATION_MS};
  const DEFAULT_TRANSITION_DURATION_MS = ${DEFAULT_TRANSITION_DURATION_MS};
  const state = {
    slideIndex: Math.max(0, Math.min((deck.deck.slides || []).length - 1, Number(window.__KEYMORPH_INITIAL_SLIDE__) || 0)),
    slideTimeMs: 0,
    globalTimeMs: 0,
    playing: false,
    playbackRate: 1,
    raf: 0,
    startedAt: 0,
    renderedKey: ""
  };
  const runtimeEvents = [];
  const assetsById = new Map((deck.deck.assets || []).map((asset) => [asset.id, asset]));

  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
  const clamp = (value, min, max) => Math.min(max, Math.max(min, value));
  const isNumber = (value) => typeof value === "number" && Number.isFinite(value);
  const clone = (value) => value === undefined ? undefined : JSON.parse(JSON.stringify(value));
  const sourceUrl = (source) => {
    if (!source) return "";
    if (source.dataUri || source.uri) return source.dataUri || source.uri;
    const asset = source.assetId ? assetsById.get(source.assetId) : undefined;
    return asset?.uri || asset?.dataUri || "";
  };
  const emitRuntimeEvent = (type, detail = {}) => {
    const entry = { type, detail: clone(detail), globalTimeMs: state.globalTimeMs, atMs: performance.now() };
    runtimeEvents.push(entry);
    if (runtimeEvents.length > 500) runtimeEvents.shift();
    window.dispatchEvent(new CustomEvent("keymorph:runtime", { detail: entry }));
    return entry;
  };
  const slideDuration = (slide) => Math.max(1, Number(slide?.timeline?.durationMs ?? DEFAULT_SLIDE_DURATION_MS));
  const transitionDuration = (transition) => {
    if (!transition || transition.type === "none" || transition.type === "cut") return 0;
    return Math.max(1, Number(transition.durationMs ?? DEFAULT_TRANSITION_DURATION_MS));
  };
  const easingCss = (value) => {
    if (!value) return "linear";
    if (typeof value !== "string") {
      if (value.type === "cubicBezier") return "cubic-bezier(" + value.x1 + "," + value.y1 + "," + value.x2 + "," + value.y2 + ")";
      if (value.type === "steps") return "steps(" + value.count + ", " + (value.position || "end") + ")";
      return "linear";
    }
    return ({ ease: "ease", easeOutCubic: "cubic-bezier(.22,1,.36,1)", easeInOutCubic: "cubic-bezier(.65,0,.35,1)", easeInCubic: "cubic-bezier(.32,0,.67,0)", easeOut: "ease-out", easeIn: "ease-in", easeInOut: "ease-in-out" }[value] || value);
  };
  const easeProgress = (easing, progress) => {
    const p = clamp(progress, 0, 1);
    if (!easing || easing === "linear") return p;
    if (typeof easing !== "string") {
      if (easing.type === "cubicBezier") return cubicBezierProgress(easing.x1, easing.y1, easing.x2, easing.y2, p);
      if (easing.type === "steps") return stepsProgress(easing.count, easing.position, p);
      if (easing.type === "spring") return springProgress(easing, p);
      return p;
    }
    if (easing === "ease") return cubicBezierProgress(0.25, 0.1, 0.25, 1, p);
    if (easing === "easeIn") return cubicBezierProgress(0.42, 0, 1, 1, p);
    if (easing === "easeOut") return cubicBezierProgress(0, 0, 0.58, 1, p);
    if (easing === "easeInOut") return cubicBezierProgress(0.42, 0, 0.58, 1, p);
    if (easing === "easeInCubic") return p * p * p;
    if (easing === "easeOutCubic") return 1 - Math.pow(1 - p, 3);
    if (easing === "easeInOutCubic") return p < 0.5 ? 4 * p * p * p : 1 - Math.pow(-2 * p + 2, 3) / 2;
    if (easing === "backIn") return backInProgress(p);
    if (easing === "backOut") return 1 - backInProgress(1 - p);
    if (easing === "backInOut") return p < 0.5 ? backInProgress(p * 2) / 2 : 1 - backInProgress((1 - p) * 2) / 2;
    if (easing === "bounceIn") return 1 - bounceOutProgress(1 - p);
    if (easing === "bounceOut") return bounceOutProgress(p);
    if (easing === "bounceInOut") return p < 0.5 ? (1 - bounceOutProgress(1 - 2 * p)) / 2 : (1 + bounceOutProgress(2 * p - 1)) / 2;
    if (easing === "elasticIn") return elasticInProgress(p);
    if (easing === "elasticOut") return 1 - elasticInProgress(1 - p);
    if (easing === "elasticInOut") return p < 0.5 ? elasticInProgress(2 * p) / 2 : 1 - elasticInProgress(2 - 2 * p) / 2;
    return p;
  };
  const cubicBezierProgress = (x1, y1, x2, y2, progress) => {
    const p = clamp(progress, 0, 1);
    const sample = (a1, a2, t) => {
      const inverse = 1 - t;
      return 3 * inverse * inverse * t * a1 + 3 * inverse * t * t * a2 + t * t * t;
    };
    let lower = 0;
    let upper = 1;
    let t = p;
    for (let index = 0; index < 20; index += 1) {
      t = (lower + upper) / 2;
      if (sample(x1, x2, t) < p) lower = t;
      else upper = t;
    }
    return clamp(sample(y1, y2, t), 0, 1);
  };
  const stepsProgress = (count, position, progress) => {
    const steps = Math.max(1, Math.floor(Number(count) || 1));
    const scaled = clamp(progress, 0, 1) * steps;
    return clamp((position === "start" ? Math.ceil(scaled) : Math.floor(scaled)) / steps, 0, 1);
  };
  const springProgress = (easing, progress) => {
    const p = clamp(progress, 0, 1);
    const mass = Math.max(0.001, Number(easing.mass ?? 1));
    const stiffness = Math.max(0.001, Number(easing.stiffness ?? 100));
    const damping = Math.max(0.001, Number(easing.damping ?? 10));
    const velocity = Number(easing.velocity ?? 0);
    const angularFrequency = Math.sqrt(stiffness / mass);
    const dampingRatio = damping / (2 * Math.sqrt(stiffness * mass));
    if (dampingRatio < 1) {
      const damped = angularFrequency * Math.sqrt(1 - dampingRatio * dampingRatio);
      const envelope = Math.exp(-dampingRatio * angularFrequency * p);
      return clamp(1 - envelope * (Math.cos(damped * p) + ((dampingRatio * angularFrequency - velocity) / damped) * Math.sin(damped * p)), 0, 1);
    }
    return clamp(1 - Math.exp(-angularFrequency * p), 0, 1);
  };
  const backInProgress = (progress) => {
    const c1 = 1.70158;
    const c3 = c1 + 1;
    return c3 * progress * progress * progress - c1 * progress * progress;
  };
  const bounceOutProgress = (progress) => {
    const n1 = 7.5625;
    const d1 = 2.75;
    if (progress < 1 / d1) return n1 * progress * progress;
    if (progress < 2 / d1) {
      const p = progress - 1.5 / d1;
      return n1 * p * p + 0.75;
    }
    if (progress < 2.5 / d1) {
      const p = progress - 2.25 / d1;
      return n1 * p * p + 0.9375;
    }
    const p = progress - 2.625 / d1;
    return n1 * p * p + 0.984375;
  };
  const elasticInProgress = (progress) => {
    if (progress === 0 || progress === 1) return progress;
    return -Math.pow(2, 10 * progress - 10) * Math.sin(((progress * 10 - 10.75) * (2 * Math.PI)) / 3);
  };
  const colorCss = (color) => {
    if (typeof color === "string") return color;
    if (color && typeof color === "object" && "value" in color) return String(color.value);
    return undefined;
  };
  const fillCss = (input) => {
    if (!input) return undefined;
    if (input.type === "solid") return colorCss(input.color);
    if (input.type === "none") return "transparent";
    if (input.type === "gradient") {
      const angle = Number(input.angleDeg ?? 180);
      const stops = (input.stops || []).map((stop) => (colorCss(stop.color) || "transparent") + " " + Math.round(Number(stop.offset || 0) * 100) + "%").join(", ");
      return "linear-gradient(" + angle + "deg, " + stops + ")";
    }
    if (input.type === "image") {
      const source = sourceUrl(input.source);
      return source ? "url('" + String(source).replace(/'/g, "%27") + "')" : undefined;
    }
    if (typeof input === "string") return input;
    return undefined;
  };
  const strokeColorCss = (stroke) => colorCss(stroke?.color) || "transparent";
  const textOf = (object, statePatch) => statePatch?.text?.plainText || (statePatch?.text?.runs || []).map((run) => run.text).join("") || object.text?.plainText || (object.text?.runs || []).map((run) => run.text).join("") || "";
  const graphemes = (text) => {
    const value = String(text || "");
    if (typeof Intl !== "undefined" && Intl.Segmenter) return Array.from(new Intl.Segmenter(undefined, { granularity: "grapheme" }).segment(value), (item) => item.segment);
    return Array.from(value);
  };
  const characterTextHtml = (text) => graphemes(text).map((char, index) => '<span class="km-text-char" data-char-index="' + index + '">' + esc(char) + '</span>').join("");
  const hasCharacterBuild = (object) => {
    for (const slide of deck.deck.slides || []) {
      for (const event of slide.timeline?.events || []) {
        if (isCharacterDissolveEvent(event, object.id) && characterOpacityTrack(event)) return true;
      }
    }
    return false;
  };
  const textStyleOf = (object, statePatch) => {
    const run = statePatch?.text?.runs?.[0] || object.text?.runs?.[0];
    return run?.style || statePatch?.style?.textStyle || object.initialState?.style?.textStyle || object.style?.textStyle || {};
  };
  const textStyleCss = (object, statePatch) => {
    const style = textStyleOf(object, statePatch);
    return [
      "font-family:" + (style.fontFamily || "Inter, Arial, sans-serif"),
      "font-size:" + Number(style.fontSize || 32) + "px",
      "font-weight:" + (style.fontWeight || 400),
      "color:" + (colorCss(style.color) || "#111827"),
      "line-height:" + (style.lineHeight || 1.15),
      "letter-spacing:" + Number(style.letterSpacing || 0) + "px"
    ].join(";");
  };
  const applyTextStyle = (el, object, statePatch) => {
    const style = textStyleOf(object, statePatch);
    el.style.fontFamily = style.fontFamily || "Inter, Arial, sans-serif";
    el.style.fontSize = Number(style.fontSize || 32) + "px";
    el.style.fontWeight = String(style.fontWeight || 400);
    el.style.color = colorCss(style.color) || "#111827";
    el.style.lineHeight = String(style.lineHeight || 1.15);
    el.style.letterSpacing = Number(style.letterSpacing || 0) + "px";
  };
  const normalizeTransform = (value) => {
    const input = value && typeof value === "object" ? value : {};
    const output = {};
    const setNumber = (key, raw) => {
      const numeric = Number(raw);
      if (Number.isFinite(numeric)) output[key] = numeric;
    };
    const applyScale = (raw) => {
      if (raw && typeof raw === "object") {
        setNumber("scaleX", raw.scaleX ?? raw.x);
        setNumber("scaleY", raw.scaleY ?? raw.y);
        const uniform = Number(raw.scale ?? raw.value);
        if (Number.isFinite(uniform)) {
          output.scaleX = uniform;
          output.scaleY = uniform;
        }
        return;
      }
      const numeric = Number(raw);
      if (Number.isFinite(numeric)) {
        output.scaleX = numeric;
        output.scaleY = numeric;
      }
    };
    setNumber("translateX", input.translateX ?? input.x ?? input.tx);
    setNumber("translateY", input.translateY ?? input.y ?? input.ty);
    setNumber("scaleX", input.scaleX);
    setNumber("scaleY", input.scaleY);
    applyScale(input.scale);
    setNumber("rotateDeg", input.rotateDeg ?? input.rotationDeg ?? input.rotate ?? input.rotation);
    setNumber("skewXDeg", input.skewXDeg ?? input.skewX);
    setNumber("skewYDeg", input.skewYDeg ?? input.skewY);
    if (input.origin && typeof input.origin === "object") {
      const x = Number(input.origin.x);
      const y = Number(input.origin.y);
      if (Number.isFinite(x) && Number.isFinite(y)) output.origin = { x, y };
    }
    return output;
  };
  const normalizeProperty = (property) => {
    const raw = String(property || "");
    const lower = raw.toLowerCase();
    const direct = {
      x: "bounds.x",
      y: "bounds.y",
      left: "bounds.x",
      top: "bounds.y",
      width: "bounds.width",
      height: "bounds.height",
      fill: "style.fill",
      stroke: "style.stroke",
      "stroke.width": "style.stroke.width",
      "stroke.color": "style.stroke.color",
      blur: "filter.blurPx",
      blurpx: "filter.blurPx",
      "filter.blur": "filter.blurPx",
      "filter.blurpx": "filter.blurPx",
      "filters.gaussianblur": "filter.blurPx",
      "custom.keynote.filters.gaussianblur": "filter.blurPx",
      tx: "transform.translateX",
      ty: "transform.translateY",
      translatex: "transform.translateX",
      translatey: "transform.translateY",
      scalex: "transform.scaleX",
      scaley: "transform.scaleY",
      scale: "transform.scale",
      rotate: "transform.rotateDeg",
      rotation: "transform.rotateDeg",
      rotatedeg: "transform.rotateDeg",
      rotationdeg: "transform.rotateDeg",
      skewx: "transform.skewXDeg",
      skewy: "transform.skewYDeg"
    };
    if (direct[lower]) return direct[lower];
    if (!lower.startsWith("transform.")) return raw;
    const suffix = lower.slice("transform.".length);
    return ({
      tx: "transform.translateX",
      ty: "transform.translateY",
      x: "transform.translateX",
      y: "transform.translateY",
      translatex: "transform.translateX",
      translatey: "transform.translateY",
      scalex: "transform.scaleX",
      scaley: "transform.scaleY",
      scale: "transform.scale",
      rotate: "transform.rotateDeg",
      rotation: "transform.rotateDeg",
      rotatedeg: "transform.rotateDeg",
      rotationdeg: "transform.rotateDeg",
      skewx: "transform.skewXDeg",
      skewy: "transform.skewYDeg"
    })[suffix] || raw;
  };
  const merge = (base, patch) => {
    if (!patch) return { ...(base || {}) };
    const next = { ...(base || {}), ...patch };
    if (base?.bounds || patch.bounds) next.bounds = { ...(base?.bounds || {}), ...(patch.bounds || {}) };
    if (base?.transform || patch.transform) next.transform = { ...(base?.transform || {}), ...normalizeTransform(patch.transform || {}) };
    if (base?.filter || patch.filter) next.filter = { ...(base?.filter || {}), ...(patch.filter || {}) };
    if (base?.style || patch.style) {
      next.style = { ...(base?.style || {}), ...(patch.style || {}) };
      if (base?.style?.stroke || patch.style?.stroke) next.style.stroke = { ...(base?.style?.stroke || {}), ...(patch.style?.stroke || {}) };
      if (base?.style?.textStyle || patch.style?.textStyle) next.style.textStyle = { ...(base?.style?.textStyle || {}), ...(patch.style?.textStyle || {}) };
    }
    return next;
  };
  const baseState = (object) => merge({
    visible: object.visible !== false,
    opacity: object.opacity ?? 1,
    bounds: object.bounds || { x: 0, y: 0, width: 100, height: 100 },
    transform: normalizeTransform(object.transform || {}),
    filter: object.filter || {},
    style: object.style || {},
    text: object.text,
    crop: object.crop,
    media: object.playback
  }, object.initialState);
  const transformCss = (transform) => {
    const t = normalizeTransform(transform || {});
    return "translate(" + Number(t.translateX || 0) + "px," + Number(t.translateY || 0) + "px) scale(" + Number(t.scaleX ?? 1) + "," + Number(t.scaleY ?? 1) + ") rotate(" + Number(t.rotateDeg || 0) + "deg) skew(" + Number(t.skewXDeg || 0) + "deg," + Number(t.skewYDeg || 0) + "deg)";
  };
  const filterCss = (filter) => {
    const blur = Number(filter?.blurPx || 0);
    return Number.isFinite(blur) && blur > 0 ? "blur(" + blur + "px)" : "none";
  };
  const boxStyle = (object, statePatch) => {
    const current = merge(baseState(object), statePatch);
    const b = current.bounds || { x: 0, y: 0, width: 100, height: 100 };
    const style = current.style || {};
    const stroke = style.stroke || {};
    const parts = [
      "left:" + Number(b.x || 0) + "px",
      "top:" + Number(b.y || 0) + "px",
      "width:" + Number(b.width || 0) + "px",
      "height:" + Number(b.height || 0) + "px",
      "opacity:" + Number(current.opacity ?? 1),
      "visibility:" + (current.visible === false ? "hidden" : "visible"),
      "transform:" + transformCss(current.transform),
      "filter:" + filterCss(current.filter)
    ];
    if (object.type === "shape" || object.type === "placeholder") {
      parts.push("background:" + (fillCss(style.fill) || "#e2e8f0"));
      parts.push("border:" + Number(stroke.width || 0) + "px solid " + strokeColorCss(stroke));
      if (object.shape === "ellipse") parts.push("border-radius:999px");
      if (object.shape === "roundRect") parts.push("border-radius:8px");
    }
    return parts.join(";");
  };
  const objectHtml = (object) => {
    const common = 'class="km-object km-' + esc(object.type) + '" data-object-id="' + esc(object.id) + '" style="' + boxStyle(object) + ';';
    if (object.type === "text") return '<div ' + common + textStyleCss(object) + '">' + (hasCharacterBuild(object) ? characterTextHtml(textOf(object)) : esc(textOf(object))) + '</div>';
    if (object.type === "image") return '<img ' + common + '" src="' + esc(sourceUrl(object.source)) + '" alt="' + esc(object.altText || object.name || "") + '">';
    if (object.type === "media") {
      const tag = object.mediaType === "audio" ? "audio" : "video";
      const poster = tag === "video" && object.posterSource ? ' poster="' + esc(sourceUrl(object.posterSource)) + '"' : "";
      return '<' + tag + ' ' + common + '" src="' + esc(sourceUrl(object.source)) + '"' + poster + ' muted playsinline></' + tag + '>';
    }
    if (object.type === "group") return '<div ' + common + 'overflow:visible">' + (object.children || []).map(objectHtml).join("") + '</div>';
    if (object.type === "shape" && object.text) return '<div ' + common + textStyleCss(object) + '">' + (hasCharacterBuild(object) ? characterTextHtml(textOf(object)) : esc(textOf(object))) + '</div>';
    return '<div ' + common + '"></div>';
  };
  const layerHtml = (slide, layerName) => {
    const background = fillCss(slide.background) || "#fff";
    return '<div class="km-slide-layer" data-layer="' + layerName + '" style="background:' + background + '">' + (slide.objects || []).map(objectHtml).join("") + '</div>';
  };
  const flattenObjects = (objects, out = []) => {
    for (const object of objects || []) {
      out.push(object);
      if (object.type === "group") flattenObjects(object.children || [], out);
    }
    return out;
  };
  const objectById = (slide, objectId) => flattenObjects(slide?.objects || []).find((object) => object.id === objectId);
  const slideById = (slideId) => (deck.deck.slides || []).find((slide) => slide.id === slideId);
  const findObjectAnywhere = (objectId) => {
    for (const slide of deck.deck.slides || []) {
      const object = objectById(slide, objectId);
      if (object) return { slide, object };
    }
    return undefined;
  };
  const getProperty = (stateValue, property) => {
    const normalized = normalizeProperty(property);
    if (normalized === "opacity") return stateValue.opacity;
    if (normalized === "visible") return stateValue.visible;
    if (normalized === "text") return stateValue.text;
    if (normalized === "crop") return stateValue.crop;
    if (normalized === "filter.blurPx") return stateValue.filter?.blurPx;
    if (normalized.startsWith("bounds.")) return stateValue.bounds?.[normalized.slice(7)];
    if (normalized === "transform.scale") return stateValue.transform?.scaleX ?? stateValue.transform?.scaleY ?? 1;
    if (normalized.startsWith("transform.")) return stateValue.transform?.[normalized.slice(10)];
    if (normalized === "style.fill") return stateValue.style?.fill;
    if (normalized === "style.stroke") return stateValue.style?.stroke;
    if (normalized.startsWith("style.stroke.")) return stateValue.style?.stroke?.[normalized.slice(13)];
    if (normalized.startsWith("style.textStyle.")) return stateValue.style?.textStyle?.[normalized.slice(16)];
    return undefined;
  };
  const setProperty = (stateValue, property, value) => {
    if (value === undefined) return;
    const normalized = normalizeProperty(property);
    const setNumber = (target, key, raw) => {
      const numeric = Number(raw);
      if (Number.isFinite(numeric)) target[key] = numeric;
    };
    const applyScale = (target, raw) => {
      if (raw && typeof raw === "object") {
        setNumber(target, "scaleX", raw.scaleX ?? raw.x);
        setNumber(target, "scaleY", raw.scaleY ?? raw.y);
        const uniform = Number(raw.scale ?? raw.value);
        if (Number.isFinite(uniform)) {
          target.scaleX = uniform;
          target.scaleY = uniform;
        }
        return;
      }
      const numeric = Number(raw);
      if (Number.isFinite(numeric)) {
        target.scaleX = numeric;
        target.scaleY = numeric;
      }
    };
    if (normalized === "opacity") {
      const numeric = Number(value);
      if (Number.isFinite(numeric)) stateValue.opacity = numeric;
    } else if (normalized === "visible") stateValue.visible = Boolean(value);
    else if (normalized === "text") stateValue.text = typeof value === "string" ? { plainText: value } : value;
    else if (normalized === "crop") stateValue.crop = value;
    else if (normalized === "filter.blurPx") {
      stateValue.filter = { ...(stateValue.filter || {}) };
      setNumber(stateValue.filter, "blurPx", value);
    }
    else if (normalized === "bounds" && value && typeof value === "object") {
      stateValue.bounds = { ...(stateValue.bounds || {}) };
      setNumber(stateValue.bounds, "x", value.x);
      setNumber(stateValue.bounds, "y", value.y);
      setNumber(stateValue.bounds, "width", value.width);
      setNumber(stateValue.bounds, "height", value.height);
    }
    else if (normalized.startsWith("bounds.")) {
      stateValue.bounds = { ...(stateValue.bounds || {}) };
      setNumber(stateValue.bounds, normalized.slice(7), value);
    } else if (normalized === "transform") {
      stateValue.transform = { ...(stateValue.transform || {}), ...normalizeTransform(value) };
    } else if (normalized === "transform.scale") {
      stateValue.transform = { ...(stateValue.transform || {}) };
      applyScale(stateValue.transform, value);
    } else if (normalized.startsWith("transform.")) {
      stateValue.transform = { ...(stateValue.transform || {}) };
      setNumber(stateValue.transform, normalized.slice(10), value);
    } else if (normalized === "style.fill") {
      stateValue.style = { ...(stateValue.style || {}), fill: value };
    } else if (normalized === "style.stroke") {
      stateValue.style = { ...(stateValue.style || {}), stroke: value };
    } else if (normalized.startsWith("style.stroke.")) {
      stateValue.style = { ...(stateValue.style || {}), stroke: { ...(stateValue.style?.stroke || {}) } };
      stateValue.style.stroke[normalized.slice(13)] = value;
    } else if (normalized.startsWith("style.textStyle.")) {
      stateValue.style = { ...(stateValue.style || {}), textStyle: { ...(stateValue.style?.textStyle || {}) } };
      stateValue.style.textStyle[normalized.slice(16)] = value;
    }
  };
  const interpolationModeForProperty = (property, mode) => {
    if (mode) return mode;
    const normalized = normalizeProperty(property);
    if (normalized === "style.fill" || normalized === "style.stroke" || normalized.endsWith(".color")) return "color";
    if (normalized === "bounds" || normalized === "transform" || normalized === "crop") return "matrix";
    return undefined;
  };
  const interpolateColor = (from, to, progress) => {
    const parse = (value) => {
      const css = colorCss(value);
      if (value && typeof value === "object" && "value" in value) {
        const parsed = parse(value.value);
        if (parsed) parsed[3] = Number(value.alpha ?? parsed[3] ?? 1);
        return parsed;
      }
      if (typeof css !== "string") return undefined;
      const hex = css.match(/^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i);
      if (hex) {
        const value = hex[1].length === 3 ? hex[1].split("").map((char) => char + char).join("") : hex[1];
        return [
          parseInt(value.slice(0, 2), 16),
          parseInt(value.slice(2, 4), 16),
          parseInt(value.slice(4, 6), 16),
          value.length === 8 ? parseInt(value.slice(6, 8), 16) / 255 : 1
        ];
      }
      const rgb = css.match(/^rgba?\\(([^)]+)\\)$/i);
      if (rgb) {
        const channels = rgb[1].split(",").map((part) => Number(part.trim().replace("%", "")));
        if (channels.length >= 3 && channels.slice(0, 3).every(Number.isFinite)) return [channels[0], channels[1], channels[2], Number.isFinite(channels[3]) ? channels[3] : 1];
      }
      return undefined;
    };
    const a = parse(from);
    const b = parse(to);
    if (!a || !b) return progress < 1 ? from : to;
    const channel = (index) => Math.round(a[index] + (b[index] - a[index]) * progress);
    const alpha = a[3] + (b[3] - a[3]) * progress;
    if (alpha < 0.999) return "rgba(" + channel(0) + ", " + channel(1) + ", " + channel(2) + ", " + Number(alpha.toFixed(3)) + ")";
    const hex = (index) => channel(index).toString(16).padStart(2, "0");
    return "#" + hex(0) + hex(1) + hex(2);
  };
  const interpolate = (from, to, progress, mode) => {
    const p = clamp(progress, 0, 1);
    if (to === undefined) return from;
    if (from === undefined) return to;
    if (mode === "matrix") return interpolatePlainObject(from, to, p);
    if (mode === "number") {
      const a = Number(from);
      const b = Number(to);
      return Number.isFinite(a) && Number.isFinite(b) ? a + (b - a) * p : p < 1 ? from : to;
    }
    if (mode === "color") {
      if (from && to && typeof from === "object" && typeof to === "object") {
        if (looksLikeFill(from) || looksLikeFill(to)) return interpolateFill(from, to, p);
        if (looksLikeStroke(from) || looksLikeStroke(to)) return interpolateStroke(from, to, p);
      }
      return interpolateColor(from, to, p);
    }
    if (mode === "discrete" || typeof from === "boolean" || typeof to === "boolean" || typeof from === "string" || typeof to === "string") return p < 1 ? from : to;
    if (isNumber(from) && isNumber(to)) return from + (to - from) * p;
    if (Array.isArray(from) && Array.isArray(to) && from.length === to.length) return from.map((value, index) => interpolate(value, to[index], p, mode));
    if (from && to && typeof from === "object" && typeof to === "object") {
      if (looksLikeFill(from) || looksLikeFill(to)) return interpolateFill(from, to, p);
      if (looksLikeStroke(from) || looksLikeStroke(to)) return interpolateStroke(from, to, p);
      return interpolatePlainObject(from, to, p);
    }
    return p < 1 ? from : to;
  };
  const interpolatePlainObject = (from, to, progress, numericKeys) => {
    if (!from || !to || typeof from !== "object" || typeof to !== "object" || Array.isArray(from) || Array.isArray(to)) return progress < 1 ? from : to;
    const result = { ...to };
    const keys = numericKeys || Array.from(new Set([...Object.keys(from), ...Object.keys(to)]));
    for (const key of keys) {
      const a = Number(from[key]);
      const b = Number(to[key]);
      result[key] = Number.isFinite(a) && Number.isFinite(b) ? a + (b - a) * progress : progress < 1 ? from[key] ?? to[key] : to[key] ?? from[key];
    }
    return result;
  };
  const interpolateFill = (from, to, progress) => {
    const p = clamp(progress, 0, 1);
    if (!from) return to;
    if (!to) return from;
    if (from.type === "solid" && to.type === "solid") return { ...to, color: interpolateColor(from.color, to.color, p) };
    if (from.type === "gradient" && to.type === "gradient" && (from.stops || []).length === (to.stops || []).length) {
      return {
        ...to,
        angleDeg: interpolate(from.angleDeg, to.angleDeg, p, "number"),
        stops: (to.stops || []).map((stop, index) => ({
          ...stop,
          offset: interpolate(from.stops?.[index]?.offset, stop.offset, p, "number"),
          color: interpolateColor(from.stops?.[index]?.color, stop.color, p)
        }))
      };
    }
    return p < 1 ? from : to;
  };
  const interpolateStroke = (from, to, progress) => {
    const p = clamp(progress, 0, 1);
    if (!from) return to;
    if (!to) return from;
    return { ...to, width: interpolate(from.width, to.width, p, "number"), color: interpolateColor(from.color, to.color, p) };
  };
  const looksLikeFill = (value) => typeof value?.type === "string" && ["none", "solid", "gradient", "image"].includes(value.type);
  const looksLikeStroke = (value) => !!value && typeof value === "object" && ("width" in value || "dash" in value || "lineCap" in value || "lineJoin" in value || "color" in value);
  const interpolateState = (from, to, progress, properties) => {
    const p = clamp(progress, 0, 1);
    const result = merge({}, to || {});
    const allowed = new Set(properties && properties.length ? properties : ["bounds", "transform", "opacity", "filter", "fill", "stroke"]);
    if (allowed.has("all") || allowed.has("bounds")) {
      result.bounds = { ...(to?.bounds || {}) };
      for (const key of ["x", "y", "width", "height"]) result.bounds[key] = interpolate(from?.bounds?.[key], to?.bounds?.[key], p, "number");
    }
    if (allowed.has("all") || allowed.has("transform")) {
      result.transform = { ...(to?.transform || {}) };
      for (const key of ["translateX", "translateY", "scaleX", "scaleY", "rotateDeg", "skewXDeg", "skewYDeg"]) {
        const fallback = key === "scaleX" || key === "scaleY" ? 1 : 0;
        result.transform[key] = interpolate(from?.transform?.[key] ?? fallback, to?.transform?.[key] ?? fallback, p, "number");
      }
    }
    if (allowed.has("all") || allowed.has("opacity")) result.opacity = interpolate(from?.opacity ?? 1, to?.opacity ?? 1, p, "number");
    if (allowed.has("all") || allowed.has("filter")) {
      result.filter = { ...(to?.filter || {}) };
      result.filter.blurPx = interpolate(from?.filter?.blurPx ?? 0, to?.filter?.blurPx ?? 0, p, "number");
      result.filter.brightness = interpolate(from?.filter?.brightness ?? 1, to?.filter?.brightness ?? 1, p, "number");
      result.filter.contrast = interpolate(from?.filter?.contrast ?? 1, to?.filter?.contrast ?? 1, p, "number");
      result.filter.saturate = interpolate(from?.filter?.saturate ?? 1, to?.filter?.saturate ?? 1, p, "number");
    }
    if (allowed.has("all") || allowed.has("fill")) {
      result.style = { ...(result.style || {}) };
      result.style.fill = interpolateFill(from?.style?.fill, to?.style?.fill, p);
    }
    if (allowed.has("all") || allowed.has("stroke")) {
      result.style = { ...(result.style || {}) };
      result.style.stroke = interpolateStroke(from?.style?.stroke, to?.style?.stroke, p);
    }
    if (allowed.has("all") || allowed.has("text")) result.text = p < 1 ? from?.text ?? to?.text : to?.text;
    if (allowed.has("all") || allowed.has("crop")) result.crop = interpolatePlainObject(from?.crop, to?.crop, p, ["x", "y", "width", "height"]);
    result.visible = p < 1 ? from?.visible ?? true : to?.visible ?? true;
    return result;
  };
  const eventDuration = (event) => Math.max(0, Number(event?.durationMs || 0));
  const createTimingPlan = (slide) => {
    const events = slide?.timeline?.events || [];
    const eventMap = new Map(events.map((event) => [event.id, event]));
    const warnings = [];
    const constraints = [];
    const nodeToEventId = new Map();
    for (const node of slide?.timeline?.dependencyGraph?.nodes || []) {
      if (node.eventId && eventMap.has(node.eventId)) nodeToEventId.set(node.id, node.eventId);
    }
    const resolveEventId = (id) => eventMap.has(id) ? id : nodeToEventId.get(id);
    const addUnresolvedWarning = (eventIds, message) => {
      const key = eventIds.join("|") + ":" + message;
      if (warnings.some((warning) => warning.code === "unresolved" && warning.eventIds.join("|") + ":" + warning.message === key)) return;
      warnings.push({ code: "unresolved", eventIds, message });
    };
    const addConstraint = (rawFromId, rawToId, relation = "after", offsetMs = 0) => {
      const fromId = resolveEventId(rawFromId);
      const toId = resolveEventId(rawToId);
      if (!fromId || !toId) {
        addUnresolvedWarning([rawFromId, rawToId].filter(Boolean), "Timing dependency could not resolve " + (rawFromId || "<missing>") + " -> " + (rawToId || "<missing>") + ".");
        return;
      }
      constraints.push({ fromId, toId, relation, offsetMs: Number(offsetMs || 0) });
    };
    const baseStarts = new Map();
    events.forEach((event, index) => {
      const previous = events[index - 1];
      let base = Number(event.delayMs || 0);
      const timing = event.start;
      if (timing?.type === "absolute") base += Number(timing.atMs || 0);
      else if (timing?.type === "with" || timing?.type === "after" || timing?.type === "before") addConstraint(timing.eventId, event.id, timing.type, Number(timing.offsetMs || 0));
      else if (timing?.type === "withPrevious" || timing?.type === "afterPrevious") addConstraint(previous?.id, event.id, timing.type === "withPrevious" ? "with" : "after", Number(timing.offsetMs || 0));
      else if (timing?.type === "trigger" || timing?.type === "onClick") base += Number(timing.type === "trigger" ? timing.offsetMs || 0 : 0);
      for (const dependency of event.dependencies || []) addConstraint(dependency.eventId, event.id, dependency.relation || "after", Number(dependency.offsetMs || 0));
      baseStarts.set(event.id, Math.max(0, base));
    });
    for (const edge of slide?.timeline?.dependencyGraph?.edges || []) addConstraint(edge.from, edge.to, edge.relation, Number(edge.offsetMs || 0));

    const indegree = new Map();
    const outgoing = new Map();
    for (const event of events) {
      indegree.set(event.id, 0);
      outgoing.set(event.id, new Set());
    }
    for (const constraint of constraints) {
      if (constraint.fromId === constraint.toId) continue;
      const targets = outgoing.get(constraint.fromId);
      if (!targets || targets.has(constraint.toId)) continue;
      targets.add(constraint.toId);
      indegree.set(constraint.toId, (indegree.get(constraint.toId) || 0) + 1);
    }
    const ready = events.filter((event) => (indegree.get(event.id) || 0) === 0).map((event) => event.id);
    const order = [];
    while (ready.length) {
      const id = ready.shift();
      order.push(id);
      for (const targetId of outgoing.get(id) || []) {
        const next = Math.max(0, (indegree.get(targetId) || 0) - 1);
        indegree.set(targetId, next);
        if (next === 0) ready.push(targetId);
      }
    }
    const cyclicIds = events.map((event) => event.id).filter((id) => !order.includes(id));
    if (cyclicIds.length) {
      warnings.push({ code: "cycle", eventIds: cyclicIds, message: "Timing dependency cycle detected among " + cyclicIds.join(", ") + "; source order fallback was used for those events." });
      order.push(...cyclicIds);
    }

    const starts = new Map();
    const constraintsByTarget = new Map();
    for (const constraint of constraints) {
      const list = constraintsByTarget.get(constraint.toId) || [];
      list.push(constraint);
      constraintsByTarget.set(constraint.toId, list);
    }
    for (const eventId of order) {
      const event = eventMap.get(eventId);
      if (!event) continue;
      let start = baseStarts.get(eventId) || 0;
      for (const constraint of constraintsByTarget.get(eventId) || []) {
        const ref = eventMap.get(constraint.fromId);
        if (!ref) continue;
        const refStart = starts.get(constraint.fromId) ?? baseStarts.get(constraint.fromId) ?? 0;
        if (constraint.relation === "with" || constraint.relation === "triggers") start = Math.max(start, refStart + constraint.offsetMs);
        else if (constraint.relation === "before") start = Math.max(start, refStart - eventDuration(event) - constraint.offsetMs);
        else start = Math.max(start, refStart + eventDuration(ref) + constraint.offsetMs);
      }
      starts.set(eventId, Math.max(0, start));
    }
    return { starts, order, warnings };
  };
  const computeEventStarts = (slide) => createTimingPlan(slide).starts;
  const eventProgress = (event, start, timeMs) => {
    const duration = eventDuration(event);
    const end = start + duration;
    if (duration === 0) return timeMs >= start ? 1 : undefined;
    if (timeMs < start) return event.fill === "backwards" || event.fill === "both" ? 0 : undefined;
    if (timeMs > end) return event.fill === "forwards" || event.fill === "both" ? 1 : undefined;
    return easeProgress(event.easing, (timeMs - start) / duration);
  };
  const keyframeValue = (track, progress) => {
    const frames = [...(track.keyframes || [])].sort((a, b) => Number(a.offset) - Number(b.offset));
    if (!frames.length) return undefined;
    if (progress <= Number(frames[0].offset)) return frames[0].value;
    for (let index = 1; index < frames.length; index += 1) {
      const prev = frames[index - 1];
      const next = frames[index];
      if (progress <= Number(next.offset)) {
        const local = (progress - Number(prev.offset)) / Math.max(0.0001, Number(next.offset) - Number(prev.offset));
        return interpolate(prev.value, next.value, easeProgress(next.easing, local), interpolationModeForProperty(track.property, track.interpolation));
      }
    }
    return frames[frames.length - 1].value;
  };
  const isCharacterDissolveEvent = (event, objectId) =>
    event?.kind === "keyframes" &&
    event.targetId === objectId &&
    event.metadata?.nativeBuildGranularity === "character" &&
    /^dissolve-(in|out)$/.test(String(event.metadata?.nativeBuildFallback || ""));
  const characterOpacityTrack = (event) => (event.tracks || []).find((track) => normalizeProperty(track.property) === "opacity");
  const stateById = (slide, stateId) => (slide?.states || []).find((item) => item.id === stateId);
  const stateForEndpoint = (endpoint, fallbackSlide, fallbackObjectId) => {
    if (!endpoint) return undefined;
    if (endpoint.snapshot) return endpoint.snapshot;
    const slide = endpoint.slideId ? slideById(endpoint.slideId) : fallbackSlide;
    const state = endpoint.stateId ? stateById(slide, endpoint.stateId) : undefined;
    if (state) return state.properties;
    const objectId = endpoint.objectId || fallbackObjectId;
    const object = objectId ? objectById(slide, objectId) || findObjectAnywhere(objectId)?.object : undefined;
    return object ? baseState(object) : undefined;
  };
  const inferMorphPairs = (fromSlide, toSlide, options) => {
    const explicit = options?.pairs || [];
    if (explicit.length) return explicit;
    const fromObjects = flattenObjects(fromSlide?.objects || []);
    const toObjects = flattenObjects(toSlide?.objects || []);
    const matchBy = options?.matching?.matchBy || ["morphKey", "objectId", "name"];
    const fallback = options?.matching?.fallback;
    const used = new Set();
    const pairs = [];
    for (const toObject of toObjects) {
      const fromObject = fromObjects.find((candidate) => {
        if (used.has(candidate.id)) return false;
        if (matchBy.includes("morphKey") && toObject.morphKey && candidate.morphKey === toObject.morphKey) return true;
        if (matchBy.includes("objectId") && candidate.id === toObject.id) return true;
        if (matchBy.includes("name") && toObject.name && candidate.name === toObject.name && candidate.type === toObject.type) return true;
        if (matchBy.includes("type") && candidate.type === toObject.type) return true;
        if (matchBy.includes("geometry") && similarGeometry(candidate, toObject, options?.matching?.tolerance)) return true;
        if (fallback === "name" && toObject.name && candidate.name === toObject.name) return true;
        if (fallback === "type" && candidate.type === toObject.type) return true;
        if (fallback === "geometry" && similarGeometry(candidate, toObject, options?.matching?.tolerance)) return true;
        return false;
      });
      if (fromObject) {
        used.add(fromObject.id);
        pairs.push({ fromObjectId: fromObject.id, toObjectId: toObject.id, morphKey: toObject.morphKey || fromObject.morphKey });
      }
    }
    return pairs;
  };
  const similarGeometry = (fromObject, toObject, tolerance = 0.2) => {
    const from = fromObject?.bounds;
    const to = toObject?.bounds;
    if (!from || !to) return false;
    const maxDimension = Math.max(from.width, from.height, to.width, to.height, 1);
    const normalizedDistance = (Math.abs(from.width - to.width) + Math.abs(from.height - to.height) + Math.abs(from.x - to.x) + Math.abs(from.y - to.y)) / maxDimension;
    return normalizedDistance <= tolerance;
  };
  const applyMorphEvent = (slide, states, event, progress) => {
    const pairs = event.pairs?.length
      ? event.pairs
      : event.from?.objectId && event.to?.objectId
        ? [{ fromObjectId: event.from.objectId, toObjectId: event.to.objectId }]
        : [];
    for (const pair of pairs) {
      const targetId = pair.toObjectId || event.to?.objectId || pair.fromObjectId;
      const targetObject = objectById(slide, targetId);
      if (!targetObject) continue;
      const fromSlide = event.from?.slideId ? slideById(event.from.slideId) : slide;
      const toSlide = event.to?.slideId ? slideById(event.to.slideId) : slide;
      const fromState = stateForEndpoint({ ...(event.from || {}), objectId: event.from?.objectId || pair.fromObjectId }, fromSlide, pair.fromObjectId);
      const toState = stateForEndpoint({ ...(event.to || {}), objectId: event.to?.objectId || pair.toObjectId }, toSlide, pair.toObjectId) || states.get(targetId) || baseState(targetObject);
      states.set(targetId, interpolateState(fromState, toState, progress, event.properties));
    }
  };
  const computeSlideObjectStates = (slide, timeMs) => {
    const states = new Map();
    const objects = flattenObjects(slide?.objects || []);
    for (const object of objects) states.set(object.id, baseState(object));
    const starts = computeEventStarts(slide);
    const events = [...(slide?.timeline?.events || [])].sort((a, b) => (starts.get(a.id) || 0) - (starts.get(b.id) || 0));
    for (const event of events) {
      if (event.kind === "setState") {
        const start = starts.get(event.id) || 0;
        const progress = eventProgress(event, start, timeMs);
        if (progress === undefined) continue;
        const stateRecord = stateById(slide, event.stateId);
        if (stateRecord) states.set(event.targetId, merge(states.get(event.targetId), stateRecord.properties));
      } else if (event.kind === "visibility") {
        const start = starts.get(event.id) || 0;
        const progress = eventProgress(event, start, timeMs);
        if (progress !== undefined) {
          const current = merge(states.get(event.targetId), {});
          current.visible = progress < 1 && event.fill === "backwards" ? !event.visible : event.visible;
          states.set(event.targetId, current);
        }
      } else if (event.kind === "property") {
        const start = starts.get(event.id) || 0;
        const progress = eventProgress(event, start, timeMs);
        if (progress === undefined) continue;
        const current = merge(states.get(event.targetId), {});
        const from = event.from ?? getProperty(current, event.property);
        setProperty(current, event.property, interpolate(from, event.to, progress, interpolationModeForProperty(event.property, event.interpolation)));
        states.set(event.targetId, current);
      } else if (event.kind === "keyframes") {
        const start = starts.get(event.id) || 0;
        const progress = eventProgress(event, start, timeMs);
        if (progress === undefined) continue;
        const current = merge(states.get(event.targetId), {});
        for (const track of event.tracks || []) {
          if (isCharacterDissolveEvent(event, event.targetId) && normalizeProperty(track.property) === "opacity") continue;
          setProperty(current, track.property, keyframeValue(track, progress));
        }
        states.set(event.targetId, current);
      } else if (event.kind === "morph") {
        const start = starts.get(event.id) || 0;
        const progress = eventProgress(event, start, timeMs);
        if (progress !== undefined) applyMorphEvent(slide, states, event, progress);
      } else if (event.kind === "media") {
        const start = starts.get(event.id) || 0;
        const progress = eventProgress(event, start, timeMs);
        if (progress !== undefined) {
          const current = merge(states.get(event.targetId), {});
          current.media = { ...(current.media || {}) };
          if (event.action === "seek") current.media.startMs = event.seekMs || 0;
          if (event.action === "mute") current.media.muted = true;
          if (event.action === "unmute") current.media.muted = false;
          states.set(event.targetId, current);
        }
      }
    }
    return states;
  };
  const composeTransforms = (parent, child) => {
    const p = normalizeTransform(parent || {});
    const c = normalizeTransform(child || {});
    return {
      ...c,
      translateX: (p.translateX || 0) + (c.translateX || 0),
      translateY: (p.translateY || 0) + (c.translateY || 0),
      scaleX: (p.scaleX ?? 1) * (c.scaleX ?? 1),
      scaleY: (p.scaleY ?? 1) * (c.scaleY ?? 1),
      rotateDeg: (p.rotateDeg || 0) + (c.rotateDeg || 0),
      skewXDeg: (p.skewXDeg || 0) + (c.skewXDeg || 0),
      skewYDeg: (p.skewYDeg || 0) + (c.skewYDeg || 0),
      origin: c.origin || p.origin
    };
  };
  const composeChildState = (parent, child) => {
    const result = merge(child, {});
    const parentBounds = parent.bounds || { x: 0, y: 0, width: 0, height: 0 };
    const childBounds = child.bounds || { x: 0, y: 0, width: 0, height: 0 };
    result.bounds = { ...childBounds, x: Number(parentBounds.x || 0) + Number(childBounds.x || 0), y: Number(parentBounds.y || 0) + Number(childBounds.y || 0) };
    result.opacity = Number(parent.opacity ?? 1) * Number(child.opacity ?? 1);
    result.visible = (parent.visible ?? true) && (child.visible ?? true);
    result.transform = composeTransforms(parent.transform, child.transform);
    return result;
  };
  const applyEffectiveGroupStates = (objects, states) => {
    const walk = (object, inherited) => {
      const own = merge(states.get(object.id) || baseState(object), {});
      const effective = inherited ? composeChildState(inherited, own) : own;
      states.set(object.id, effective);
      if (object.type === "group") for (const child of object.children || []) walk(child, effective);
    };
    for (const object of objects || []) walk(object, undefined);
    return states;
  };
  const computeEffectiveSlideObjectStates = (slide, timeMs) => applyEffectiveGroupStates(slide?.objects || [], computeSlideObjectStates(slide, timeMs));
  const characterOpacityAt = (event, index, count, timeMs, starts) => {
    const track = characterOpacityTrack(event);
    if (!track) return undefined;
    const duration = Math.max(1, Number(event.durationMs || 1));
    const start = starts?.get(event.id) ?? (event.start?.type === "absolute" ? Number(event.start.atMs || 0) : Number(event.delayMs || 0));
    const local = Number(timeMs || 0) - start;
    if (local <= 0) return event.fill === "backwards" || event.fill === "both" ? keyframeValue(track, 0) : undefined;
    if (local >= duration) return event.fill === "forwards" || event.fill === "both" ? keyframeValue(track, 1) : undefined;
    const span = duration / Math.max(1, count);
    const progress = clamp((local - index * span) / Math.max(1, span), 0, 1);
    return keyframeValue(track, easeProgress(event.easing, progress));
  };
  const applyCharacterTextState = (el, object, statePatch, slide, timeMs, starts) => {
    const text = textOf(object, statePatch);
    const chars = graphemes(text);
    const spans = Array.from(el.querySelectorAll(".km-text-char"));
    if (spans.length !== chars.length || spans.some((span, index) => span.textContent !== chars[index])) {
      el.innerHTML = characterTextHtml(text);
    }
    const activeEvents = (slide?.timeline?.events || []).filter((event) => isCharacterDissolveEvent(event, object.id) && characterOpacityTrack(event));
    for (const [index, span] of Array.from(el.querySelectorAll(".km-text-char")).entries()) {
      let opacity = 1;
      for (const event of activeEvents) {
        const value = characterOpacityAt(event, index, chars.length, timeMs, starts);
        if (value !== undefined) opacity = value;
      }
      span.style.opacity = String(opacity);
    }
  };
  const applyObjectState = (el, object, statePatch, slide, timeMs, starts) => {
    if (!el || !object || !statePatch) return;
    const b = statePatch.bounds || { x: 0, y: 0, width: 100, height: 100 };
    el.style.left = Number(b.x || 0) + "px";
    el.style.top = Number(b.y || 0) + "px";
    el.style.width = Number(b.width || 0) + "px";
    el.style.height = Number(b.height || 0) + "px";
    el.style.opacity = String(statePatch.opacity ?? 1);
    el.style.visibility = statePatch.visible === false ? "hidden" : "visible";
    el.style.transform = transformCss(statePatch.transform);
    el.style.filter = filterCss(statePatch.filter);
    if (object.type === "text" || (object.type === "shape" && object.text)) {
      if (hasCharacterBuild(object)) {
        el.style.opacity = String(object.opacity ?? 1);
        applyCharacterTextState(el, object, statePatch, slide, timeMs, starts);
      } else {
        el.textContent = textOf(object, statePatch);
      }
      applyTextStyle(el, object, statePatch);
    }
    if (object.type === "shape" || object.type === "placeholder") {
      const style = statePatch.style || {};
      const stroke = style.stroke || {};
      el.style.background = fillCss(style.fill) || "#e2e8f0";
      el.style.border = Number(stroke.width || 0) + "px solid " + strokeColorCss(stroke);
    }
  };
  const applySlideFrame = (layer, slide, timeMs) => {
    const states = computeSlideObjectStates(slide, timeMs);
    const starts = computeEventStarts(slide);
    for (const object of flattenObjects(slide?.objects || [])) {
      applyObjectState(layer.querySelector('[data-object-id="' + CSS.escape(object.id) + '"]'), object, states.get(object.id), slide, timeMs, starts);
    }
    return states;
  };
  const resetLayerTransitionStyle = (layer) => {
    if (!layer) return;
    layer.style.opacity = "1";
    layer.style.transform = "";
    layer.style.clipPath = "";
  };
  const applyTransitionLayerStyles = (previousLayer, currentLayer, transition, progress) => {
    const type = transition?.type || "cut";
    const direction = transition?.direction || "left";
    const p = easeProgress(transition?.easing, progress);
    previousLayer.style.opacity = "1";
    currentLayer.style.opacity = "1";
    previousLayer.style.transform = "";
    currentLayer.style.transform = "";
    currentLayer.style.clipPath = "";
    if (type === "fade" || type === "dissolve" || type === "morph" || type === "magicMove") {
      previousLayer.style.opacity = String(1 - p);
      currentLayer.style.opacity = String(type === "morph" || type === "magicMove" ? 1 : p);
    } else if (type === "push") {
      const axis = direction === "up" || direction === "down" ? "Y" : "X";
      const sign = direction === "right" || direction === "down" ? 1 : -1;
      previousLayer.style.transform = "translate" + axis + "(" + (sign * p * 100) + "%)";
      currentLayer.style.transform = "translate" + axis + "(" + (sign * (p - 1) * 100) + "%)";
    } else if (type === "wipe") {
      const inset = Math.round((1 - p) * 100);
      currentLayer.style.clipPath =
        direction === "right" ? "inset(0 " + inset + "% 0 0)" :
        direction === "up" ? "inset(" + inset + "% 0 0 0)" :
        direction === "down" ? "inset(0 0 " + inset + "% 0)" :
        "inset(0 0 0 " + inset + "%)";
    } else if (type === "zoom") {
      currentLayer.style.opacity = String(p);
      currentLayer.style.transform = "scale(" + (0.9 + 0.1 * p) + ")";
    } else {
      previousLayer.style.opacity = p >= 1 ? "0" : "1";
      currentLayer.style.opacity = p >= 1 ? "1" : "0";
    }
  };
  const activeTimelineTransition = (slide, timeMs) => {
    const starts = computeEventStarts(slide);
    let active;
    for (const event of slide?.timeline?.events || []) {
      if (event.kind !== "transition") continue;
      const progress = eventProgress(event, starts.get(event.id) || 0, timeMs);
      if (progress !== undefined) active = { event, progress };
    }
    return active;
  };
  const timelineEventDiagnostics = (slide, timeMs) => {
    const timingPlan = createTimingPlan(slide);
    return (slide?.timeline?.events || []).map((event) => {
      const startMs = timingPlan.starts.get(event.id) || 0;
      const durationMs = eventDuration(event);
      const endMs = startMs + durationMs;
      const appliedProgress = eventProgress(event, startMs, timeMs);
      const rawProgress = durationMs === 0
        ? timeMs >= startMs ? 1 : undefined
        : timeMs >= startMs && timeMs <= endMs ? clamp((timeMs - startMs) / durationMs, 0, 1) : undefined;
      const phase = durationMs === 0 && timeMs >= startMs ? "instant" : timeMs < startMs ? "before" : timeMs > endMs ? "after" : "active";
      return {
        eventId: event.id,
        kind: event.kind,
        label: event.label,
        targetId: event.targetId,
        property: event.kind === "property" ? event.property : undefined,
        transitionType: event.kind === "transition" ? event.transition?.type : undefined,
        startMs,
        endMs,
        durationMs,
        phase,
        active: phase === "active" || phase === "instant",
        applied: appliedProgress !== undefined,
        fill: event.fill,
        rawProgress,
        easedProgress: rawProgress === undefined ? undefined : easeProgress(event.easing, rawProgress),
        appliedProgress
      };
    });
  };
  const applyTimelineTransition = (layer, transition, progress) => {
    const type = transition?.type || "cut";
    const direction = transition?.direction || "left";
    const p = easeProgress(transition?.easing, progress);
    resetLayerTransitionStyle(layer);
    if (type === "fade" || type === "dissolve") {
      layer.style.opacity = String(p);
    } else if (type === "zoom") {
      layer.style.opacity = String(p);
      layer.style.transform = "scale(" + (0.9 + 0.1 * p) + ")";
    } else if (type === "push") {
      const axis = direction === "up" || direction === "down" ? "Y" : "X";
      const sign = direction === "right" || direction === "down" ? 1 : -1;
      layer.style.transform = "translate" + axis + "(" + (sign * (p - 1) * 100) + "%)";
    } else if (type === "wipe") {
      const inset = Math.round((1 - p) * 100);
      layer.style.clipPath =
        direction === "right" ? "inset(0 " + inset + "% 0 0)" :
        direction === "up" ? "inset(" + inset + "% 0 0 0)" :
        direction === "down" ? "inset(0 0 " + inset + "% 0)" :
        "inset(0 0 0 " + inset + "%)";
    }
  };
  const resolveMorphTransitionStates = (previousSlide, currentSlide, transition, progress) => {
    const options = transition?.morph || {};
    const easedProgress = easeProgress(transition?.easing, progress);
    const pairs = inferMorphPairs(previousSlide, currentSlide, options);
    const properties = options.properties || ["bounds", "transform", "opacity", "fill", "stroke"];
    const previousStates = computeEffectiveSlideObjectStates(previousSlide, slideDuration(previousSlide));
    const currentStates = computeEffectiveSlideObjectStates(currentSlide, 0);
    const states = new Map();
    for (const pair of pairs) {
      const toObject = objectById(currentSlide, pair.toObjectId);
      if (!toObject) continue;
      const fromState = previousStates.get(pair.fromObjectId) || baseState(objectById(previousSlide, pair.fromObjectId) || {});
      const toState = currentStates.get(pair.toObjectId) || baseState(toObject);
      states.set(pair.toObjectId, interpolateState(fromState, toState, easedProgress, properties));
    }
    return { pairs, states };
  };
  const applyMorphTransition = (previousLayer, currentLayer, previousSlide, currentSlide, transition, progress) => {
    const resolved = resolveMorphTransitionStates(previousSlide, currentSlide, transition, progress);
    for (const pair of resolved.pairs) {
      const toObject = objectById(currentSlide, pair.toObjectId);
      if (!toObject) continue;
      const blended = resolved.states.get(pair.toObjectId);
      if (!blended) continue;
      applyObjectState(currentLayer.querySelector('[data-object-id="' + CSS.escape(pair.toObjectId) + '"]'), toObject, blended, currentSlide, 0, undefined);
      const sourceEl = previousLayer.querySelector('[data-object-id="' + CSS.escape(pair.fromObjectId) + '"]');
      if (sourceEl) sourceEl.style.opacity = "0";
    }
  };
  const createTimeline = () => {
    const spans = [];
    let cursor = 0;
    (deck.deck.slides || []).forEach((slide, slideIndex) => {
      const transitionDurationMs = slideIndex > 0 ? transitionDuration(slide.transition) : 0;
      const slideDurationMs = slideDuration(slide);
      const transitionStartMs = cursor;
      const contentStartMs = cursor + transitionDurationMs;
      const endMs = contentStartMs + slideDurationMs;
      spans.push({ slideIndex, slideId: slide.id, startMs: transitionStartMs, transitionStartMs, contentStartMs, endMs, transitionDurationMs, slideDurationMs });
      cursor = endMs;
    });
    return { durationMs: Math.max(1, cursor), slides: spans };
  };
  let timeline = createTimeline();
  const resolveTime = (timeMs) => {
    const globalTimeMs = clamp(Number(timeMs) || 0, 0, timeline.durationMs);
    const fallback = timeline.slides[timeline.slides.length - 1];
    const span = timeline.slides.find((candidate) => globalTimeMs < candidate.endMs) || fallback;
    const inTransition = !!span && span.transitionDurationMs > 0 && globalTimeMs < span.contentStartMs;
    return {
      globalTimeMs,
      slideIndex: span?.slideIndex || 0,
      slideId: span?.slideId || "",
      slideTimeMs: inTransition ? 0 : clamp(globalTimeMs - (span?.contentStartMs || 0), 0, span?.slideDurationMs || 1),
      slideStartMs: span?.contentStartMs || 0,
      slideEndMs: span?.endMs || timeline.durationMs,
      inTransition,
      transitionProgress: inTransition ? clamp((globalTimeMs - span.transitionStartMs) / span.transitionDurationMs, 0, 1) : 1,
      transitionDurationMs: span?.transitionDurationMs || 0,
      previousSlideIndex: inTransition ? span.slideIndex - 1 : undefined
    };
  };
  const renderFrame = (resolution) => {
    const slide = deck.deck.slides[resolution.slideIndex];
    if (!slide) return;
    stage.style.width = deck.deck.size.width + "px";
    stage.style.height = deck.deck.size.height + "px";
    stage.style.background = fillCss(slide.background) || "#fff";
    const key = resolution.inTransition ? "transition:" + resolution.previousSlideIndex + ":" + resolution.slideIndex : "slide:" + resolution.slideIndex;
    if (state.renderedKey !== key) {
      if (resolution.inTransition) {
        const previousSlide = deck.deck.slides[resolution.previousSlideIndex];
        stage.innerHTML = layerHtml(previousSlide, "previous") + layerHtml(slide, "current");
      } else {
        stage.innerHTML = layerHtml(slide, "current");
      }
      state.renderedKey = key;
    }
    if (resolution.inTransition) {
      const previousSlide = deck.deck.slides[resolution.previousSlideIndex];
      const previousLayer = stage.querySelector('[data-layer="previous"]');
      const currentLayer = stage.querySelector('[data-layer="current"]');
      applySlideFrame(previousLayer, previousSlide, slideDuration(previousSlide));
      applySlideFrame(currentLayer, slide, 0);
      applyTransitionLayerStyles(previousLayer, currentLayer, slide.transition, resolution.transitionProgress);
      if (slide.transition?.type === "morph" || slide.transition?.type === "magicMove") {
        applyMorphTransition(previousLayer, currentLayer, previousSlide, slide, slide.transition, resolution.transitionProgress);
      }
    } else {
      const layer = stage.querySelector('[data-layer="current"]');
      applySlideFrame(layer, slide, resolution.slideTimeMs);
      const transition = activeTimelineTransition(slide, resolution.slideTimeMs);
      if (transition) applyTimelineTransition(layer, transition.event.transition, transition.progress);
      else resetLayerTransitionStyle(layer);
    }
    resize();
  };
  const updateControls = (resolution) => {
    if (scrub) {
      scrub.max = String(timeline.durationMs);
      scrub.value = String(Math.round(resolution.globalTimeMs));
    }
    if (status) {
      const marker = resolution.inTransition ? " transition" : "";
      status.textContent = "Slide " + (resolution.slideIndex + 1) + "/" + deck.deck.slides.length + marker + " · " + Math.round(resolution.globalTimeMs) + " ms";
    }
  };
  const seekGlobal = (timeMs) => {
    const resolution = resolveTime(timeMs);
    state.globalTimeMs = resolution.globalTimeMs;
    state.slideIndex = resolution.slideIndex;
    state.slideTimeMs = resolution.slideTimeMs;
    renderFrame(resolution);
    updateControls(resolution);
    emitRuntimeEvent("seek", { resolution });
    return resolution;
  };
  const seekSlide = (slideIndex, timeMs = 0) => {
    const span = timeline.slides[Math.max(0, Math.min(timeline.slides.length - 1, Number(slideIndex) || 0))];
    return seekGlobal((span?.contentStartMs || 0) + Number(timeMs || 0));
  };
  const seekFrame = (frame, fps = 30) => {
    const frameIndex = Math.max(0, Math.floor(Number(frame) || 0));
    const frameRate = Math.max(1, Number(fps) || 30);
    return seekGlobal((frameIndex / frameRate) * 1000);
  };
  const step = (frames = 1, fps = 30) => {
    pause();
    const frameCount = Number.isFinite(Number(frames)) ? Number(frames) : 1;
    const frameRate = Math.max(1, Number(fps) || 30);
    const resolution = seekGlobal(state.globalTimeMs + (1000 / frameRate) * frameCount);
    emitRuntimeEvent("step", { frames: frameCount, fps: frameRate, resolution });
    return resolution;
  };
  const setPlaybackRate = (rate) => {
    const next = Math.max(0.01, Number(rate) || 1);
    state.playbackRate = next;
    if (state.playing) state.startedAt = performance.now() - state.globalTimeMs / state.playbackRate;
    return state.playbackRate;
  };
  const resize = () => {
    const scaleOption = Number(runtimeOptions.stageScale);
    if (Number.isFinite(scaleOption) && scaleOption > 0) {
      stage.style.transform = "scale(" + scaleOption + ")";
      viewport.style.width = Math.round(deck.deck.size.width * scaleOption) + "px";
      viewport.style.height = Math.round(deck.deck.size.height * scaleOption) + "px";
      return;
    }
    const scale = Math.min((viewport.clientWidth - 8) / deck.deck.size.width, (viewport.clientHeight - 8) / deck.deck.size.height, 1);
    stage.style.transform = "scale(" + Math.max(0.1, scale) + ")";
  };
  const pause = () => {
    state.playing = false;
    cancelAnimationFrame(state.raf);
    if (playButton) playButton.textContent = "Play";
  };
  const tick = () => {
    if (!state.playing) return;
    seekGlobal((performance.now() - state.startedAt) * state.playbackRate);
    if (state.globalTimeMs >= timeline.durationMs) {
      pause();
      return;
    }
    state.raf = requestAnimationFrame(tick);
  };
  const play = () => {
    if (state.globalTimeMs >= timeline.durationMs) seekGlobal(0);
    state.playing = true;
    state.startedAt = performance.now() - state.globalTimeMs / state.playbackRate;
    if (playButton) playButton.textContent = "Pause";
    tick();
  };
  const previousSlide = () => seekSlide(Math.max(0, state.slideIndex - 1), 0);
  const nextSlide = () => seekSlide(Math.min(timeline.slides.length - 1, state.slideIndex + 1), 0);
  const serializeStates = (states) => Object.fromEntries(Array.from(states.entries()).map(([id, value]) => [id, JSON.parse(JSON.stringify(value || {}))]));
  const getTimelineEventDiagnostics = (slideIndex = state.slideIndex, timeMs) => {
    const index = Math.max(0, Math.min((deck.deck.slides || []).length - 1, Number(slideIndex) || 0));
    const slide = deck.deck.slides[index];
    const resolvedTime = timeMs === undefined ? (index === state.slideIndex ? state.slideTimeMs : 0) : Number(timeMs || 0);
    return timelineEventDiagnostics(slide, resolvedTime);
  };
  const getSlideStates = (slideIndex = state.slideIndex, timeMs, options = {}) => {
    const slide = deck.deck.slides[Math.max(0, Math.min((deck.deck.slides || []).length - 1, Number(slideIndex) || 0))];
    const resolvedTime = timeMs === undefined ? (slideIndex === state.slideIndex ? state.slideTimeMs : 0) : Number(timeMs || 0);
    const states = options.effectiveGroupStates ? computeEffectiveSlideObjectStates(slide, resolvedTime) : computeSlideObjectStates(slide, resolvedTime);
    return serializeStates(states);
  };
  const getCurrentFrameState = (options = {}) => {
    const resolution = resolveTime(state.globalTimeMs);
    const slide = deck.deck.slides[resolution.slideIndex];
    const frame = {
      resolution,
      states: getSlideStates(resolution.slideIndex, resolution.slideTimeMs, { effectiveGroupStates: true }),
      events: getTimelineEventDiagnostics(resolution.slideIndex, resolution.slideTimeMs),
      transition: undefined
    };
    if (resolution.inTransition) {
      const previousSlide = deck.deck.slides[resolution.previousSlideIndex];
      const morph = resolveMorphTransitionStates(previousSlide, slide, slide.transition, resolution.transitionProgress);
      frame.transition = {
        pairs: morph.pairs,
        states: serializeStates(morph.states)
      };
    }
    if (options.includeDom) frame.dom = getDomSnapshot();
    return frame;
  };
  const getGlobalSnapshot = (timeMs = state.globalTimeMs, options = {}) => {
    const resolution = resolveTime(timeMs);
    const includeInactiveSlides = options.includeInactiveSlides === true;
    const activeIndexes = new Set([resolution.slideIndex]);
    if (resolution.previousSlideIndex !== undefined) activeIndexes.add(resolution.previousSlideIndex);
    const slideIndexes = includeInactiveSlides ? (deck.deck.slides || []).map((_, index) => index) : Array.from(activeIndexes).sort((a, b) => a - b);
    const slides = slideIndexes.map((slideIndex) => {
      const slide = deck.deck.slides[slideIndex];
      const phase = slideIndex === resolution.slideIndex ? "current" : slideIndex === resolution.previousSlideIndex ? "previous" : "inactive";
      const slideTimeMs = phase === "current" ? resolution.slideTimeMs : phase === "previous" ? slideDuration(slide) : 0;
      const timingPlan = createTimingPlan(slide);
      return {
        slideIndex,
        slideId: slide?.id || "",
        phase,
        timeMs: slideTimeMs,
        states: getSlideStates(slideIndex, slideTimeMs, { effectiveGroupStates: options.effectiveGroupStates !== false }),
        events: timelineEventDiagnostics(slide, slideTimeMs),
        timingPlan: {
          starts: Object.fromEntries(timingPlan.starts),
          order: timingPlan.order,
          warnings: timingPlan.warnings
        }
      };
    });
    const objects = {};
    for (const slide of slides) {
      if (slide.phase === "inactive") continue;
      for (const [objectId, objectState] of Object.entries(slide.states)) {
        objects[objects[objectId] ? slide.slideId + ":" + objectId : objectId] = objectState;
      }
    }
    const currentSlide = deck.deck.slides[resolution.slideIndex];
    const previousSlide = resolution.previousSlideIndex === undefined ? undefined : deck.deck.slides[resolution.previousSlideIndex];
    const morph = resolution.inTransition && currentSlide ? resolveMorphTransitionStates(previousSlide, currentSlide, currentSlide.transition, resolution.transitionProgress) : undefined;
    return {
      deckId: deck.deck.id,
      globalTimeMs: resolution.globalTimeMs,
      resolution,
      objects,
      slides,
      transition: resolution.inTransition && currentSlide ? {
        previousSlideIndex: resolution.previousSlideIndex,
        previousSlideId: previousSlide?.id,
        currentSlideIndex: resolution.slideIndex,
        currentSlideId: currentSlide.id,
        progress: resolution.transitionProgress,
        pairs: morph?.pairs || [],
        states: serializeStates(morph?.states || new Map())
      } : undefined,
      warnings: slides.flatMap((slide) => slide.timingPlan.warnings),
      runtimeEvents: runtimeEvents.slice()
    };
  };
  const getDomSnapshot = () => ({
    stage: {
      width: stage.style.width,
      height: stage.style.height,
      transform: stage.style.transform,
      renderedKey: state.renderedKey
    },
    layers: Array.from(stage.querySelectorAll(".km-slide-layer")).map((layer) => ({
      layer: layer.getAttribute("data-layer") || "",
      opacity: layer.style.opacity,
      transform: layer.style.transform,
      clipPath: layer.style.clipPath,
      objects: Array.from(layer.querySelectorAll("[data-object-id]")).map((el) => ({
        objectId: el.getAttribute("data-object-id") || "",
        left: el.style.left,
        top: el.style.top,
        width: el.style.width,
        height: el.style.height,
        opacity: el.style.opacity,
        visibility: el.style.visibility,
        transform: el.style.transform
      }))
    }))
  });
  const captureFrame = async (timeMs = state.globalTimeMs, options = {}) => {
    const resolution = seekGlobal(timeMs);
    if (document.fonts?.ready) await document.fonts.ready;
    const settles = Math.max(1, Math.floor(Number(options.animationFrames ?? 2) || 2));
    for (let index = 0; index < settles; index += 1) {
      await new Promise((resolve) => requestAnimationFrame(resolve));
    }
    const snapshot = getGlobalSnapshot(resolution.globalTimeMs, options);
    emitRuntimeEvent("captureFrame", { resolution });
    return snapshot;
  };
  playButton?.addEventListener("click", () => state.playing ? pause() : play());
  prevButton?.addEventListener("click", previousSlide);
  nextButton?.addEventListener("click", nextSlide);
  stepButton?.addEventListener("click", () => step(1, 30));
  scrub?.addEventListener("input", () => seekGlobal(Number(scrub.value)));
  window.addEventListener("resize", resize);
  timeline = createTimeline();
  window.__keyMorphRuntime = {
    render: () => seekGlobal(state.globalTimeMs),
    seek: seekGlobal,
    seekGlobal,
    seekSlide,
    seekFrame,
    step,
    setPlaybackRate,
    play,
    pause,
    nextSlide,
    previousSlide,
    resolveTime,
    createTimeline,
    createTimingPlan,
    getTimelineEventDiagnostics,
    getSlideStates,
    getCurrentFrameState,
    getGlobalSnapshot,
    getDomSnapshot,
    captureFrame,
    runtimeEvents,
    state,
    deck,
    debug: {
      computeSlideObjectStates,
      computeEffectiveSlideObjectStates,
      createTimingPlan,
      resolveMorphTransitionStates,
      getTimelineEventDiagnostics,
      getSlideStates,
      getCurrentFrameState,
      getGlobalSnapshot,
      getDomSnapshot,
      captureFrame,
      easeProgress
    }
  };
  seekSlide(state.slideIndex, 0);
})();`;
}

function renderObjectMarkup(object: IRObject, deck?: DeckIR): string {
  const style = objectStyle(object);
  const common = `class="km-object km-${object.type}" data-object-id="${escapeHtml(object.id)}" style="${style}"`;

  if (object.type === "text") {
    return `<div ${common}>${renderTextObjectContent(object, deck)}</div>`;
  }
  if (object.type === "image") {
    return `<img ${common} src="${escapeHtml(resolveObjectSourceUrl(object.source, deck))}" alt="${escapeHtml(object.altText ?? object.name ?? "")}">`;
  }
  if (object.type === "media") {
    const tag = object.mediaType === "audio" ? "audio" : "video";
    return `<${tag} ${common} src="${escapeHtml(resolveObjectSourceUrl(object.source, deck))}" muted playsinline></${tag}>`;
  }
  if (object.type === "group") {
    return `<div ${common}>${object.children.map((child) => renderObjectMarkup(child, deck)).join("")}</div>`;
  }
  if (object.type === "shape" && object.text) {
    return `<div ${common}>${renderTextObjectContent(object as ShapeObject & TextObject, deck)}</div>`;
  }
  return `<div ${common}></div>`;
}

function renderTextObjectContent(object: TextObject | (ShapeObject & TextObject), deck?: DeckIR): string {
  const text = textContent(object);
  if (!hasCharacterBuildEvent(object.id, deck)) {
    return escapeHtml(text);
  }
  return splitGraphemes(text)
    .map((char, index) => `<span class="km-text-char" data-char-index="${index}">${escapeHtml(char)}</span>`)
    .join("");
}

function hasCharacterBuildEvent(objectId: string, deck?: DeckIR): boolean {
  return Boolean(
    deck?.deck.slides.some((slide) =>
      slide.timeline?.events.some((event) => isStaticCharacterDissolveEvent(event, objectId))
    )
  );
}

function isStaticCharacterDissolveEvent(event: { kind: string; targetId?: string; metadata?: JSONRecord; tracks?: Array<{ property: string }> }, objectId: string): boolean {
  return (
    event.kind === "keyframes" &&
    event.targetId === objectId &&
    event.metadata?.nativeBuildGranularity === "character" &&
    /^dissolve-(in|out)$/.test(String(event.metadata?.nativeBuildFallback ?? "")) &&
    Boolean(event.tracks?.some((track) => normalizeStaticRuntimeProperty(track.property) === "opacity"))
  );
}

function normalizeStaticRuntimeProperty(property: string): string {
  return property.toLowerCase() === "alpha" ? "opacity" : property;
}

function splitGraphemes(text: string): string[] {
  const Segmenter = (Intl as typeof Intl & { Segmenter?: new (locale?: string, options?: { granularity?: "grapheme" }) => { segment(value: string): Iterable<{ segment: string }> } }).Segmenter;
  if (Segmenter) {
    return Array.from(new Segmenter(undefined, { granularity: "grapheme" }).segment(text), (item) => item.segment);
  }
  return Array.from(text);
}

function resolveObjectSourceUrl(source: ObjectSource | undefined, deck: DeckIR | undefined): string {
  if (!source) return "";
  if (source.dataUri ?? source.uri) return source.dataUri ?? source.uri ?? "";
  const asset = source.assetId ? deck?.deck.assets?.find((candidate) => candidate.id === source.assetId) : undefined;
  return asset?.uri ?? asset?.dataUri ?? "";
}

function objectStyle(object: IRObject): string {
  const state = initialObjectState(object);
  const bounds = state.bounds ?? { x: 0, y: 0, width: 100, height: 100 };
  const transform = state.transform ?? {};
  const style = [
    `left:${bounds.x}px`,
    `top:${bounds.y}px`,
    `width:${bounds.width}px`,
    `height:${bounds.height}px`,
    `opacity:${state.opacity ?? 1}`,
    `visibility:${state.visible === false ? "hidden" : "visible"}`,
    `transform:${transformToCss(transform)}`,
    `filter:${filterToCss(state.filter)}`
  ];

  if (object.type === "text") {
    const textStyle = firstTextStyle(object, state);
    style.push(`font-family:${textStyle.fontFamily ?? "Inter, Arial, sans-serif"}`);
    style.push(`font-size:${textStyle.fontSize ?? 32}px`);
    style.push(`font-weight:${textStyle.fontWeight ?? 400}`);
    style.push(`color:${colorToCss(textStyle.color) ?? "#111827"}`);
    style.push(`letter-spacing:${textStyle.letterSpacing ?? 0}px`);
    style.push("white-space:pre-wrap");
  }

  if (object.type === "shape" || object.type === "placeholder") {
    style.push(`background:${fillToCss(state.style?.fill) ?? "#e2e8f0"}`);
    const stroke = state.style?.stroke;
    if (stroke) {
      style.push(`border:${stroke.width ?? 1}px solid ${colorToCss(stroke.color) ?? "transparent"}`);
    }
    if (object.type === "shape" && object.shape === "ellipse") style.push("border-radius:999px");
    if (object.type === "shape" && object.shape === "roundRect") style.push("border-radius:8px");
  }

  return style.join(";");
}

function initialObjectState(object: IRObject): ObjectStateProperties {
  return mergeObjectStateProperties(
    {
      visible: object.visible !== false,
      opacity: object.opacity ?? 1,
      bounds: object.bounds ?? { x: 0, y: 0, width: 100, height: 100 },
      transform: normalizeTransformPatch(object.transform),
      filter: "filter" in object ? object.filter : {},
      style: object.style ?? {},
      text: "text" in object ? object.text : undefined,
      crop: "crop" in object ? object.crop : undefined,
      media: "playback" in object ? object.playback : undefined
    },
    object.initialState
  );
}

function mergeObjectStateProperties(
  base: ObjectStateProperties,
  patch: ObjectStateProperties | undefined
): ObjectStateProperties {
  if (!patch) return { ...base };
  const transformPatch = patch.transform ? normalizeTransformPatch(patch.transform) : undefined;
  return {
    ...base,
    ...patch,
    bounds: patch.bounds ? { ...base.bounds, ...patch.bounds } : base.bounds,
    transform: transformPatch ? { ...base.transform, ...transformPatch } : base.transform,
    filter: patch.filter ? { ...base.filter, ...patch.filter } : base.filter,
    style: patch.style
      ? {
          ...base.style,
          ...patch.style,
          stroke: patch.style.stroke ? { ...base.style?.stroke, ...patch.style.stroke } : base.style?.stroke,
          textStyle: patch.style.textStyle
            ? { ...base.style?.textStyle, ...patch.style.textStyle }
            : base.style?.textStyle
        }
      : base.style
  };
}

function flattenObjects(objects: IRObject[] | undefined, out: IRObject[] = []): IRObject[] {
  for (const object of objects ?? []) {
    out.push(object);
    if (object.type === "group") flattenObjects(object.children, out);
  }
  return out;
}

function applyEffectiveGroupStates(objects: IRObject[] | undefined, states: Map<string, ObjectStateProperties>): void {
  const walk = (object: IRObject, inherited: ObjectStateProperties | undefined) => {
    const own = mergeObjectStateProperties(states.get(object.id) ?? initialObjectState(object), {});
    const effective = inherited ? composeChildState(inherited, own) : own;
    states.set(object.id, effective);
    if (object.type === "group") {
      for (const child of object.children) walk(child, effective);
    }
  };

  for (const object of objects ?? []) walk(object, undefined);
}

function composeChildState(parent: ObjectStateProperties, child: ObjectStateProperties): ObjectStateProperties {
  const result = mergeObjectStateProperties(child, {});
  const parentBounds = parent.bounds ?? { x: 0, y: 0, width: 0, height: 0 };
  const childBounds = child.bounds ?? { x: 0, y: 0, width: 0, height: 0 };
  result.bounds = {
    ...childBounds,
    x: (parentBounds.x ?? 0) + (childBounds.x ?? 0),
    y: (parentBounds.y ?? 0) + (childBounds.y ?? 0)
  };
  result.opacity = (parent.opacity ?? 1) * (child.opacity ?? 1);
  result.visible = (parent.visible ?? true) && (child.visible ?? true);
  result.transform = composeTransforms(parent.transform, child.transform);
  result.filter = composeFilters(parent.filter, child.filter);
  return result;
}

function composeFilters(
  parent: ObjectStateProperties["filter"] | undefined,
  child: ObjectStateProperties["filter"] | undefined
): ObjectStateProperties["filter"] {
  return {
    ...child,
    blurPx: (parent?.blurPx ?? 0) + (child?.blurPx ?? 0),
    brightness: (parent?.brightness ?? 1) * (child?.brightness ?? 1),
    contrast: (parent?.contrast ?? 1) * (child?.contrast ?? 1),
    saturate: (parent?.saturate ?? 1) * (child?.saturate ?? 1)
  };
}

function composeTransforms(parent: Transform2D | undefined, child: Transform2D | undefined): Transform2D {
  const p = normalizeTransformPatch(parent);
  const c = normalizeTransformPatch(child);
  return {
    ...c,
    translateX: (p.translateX ?? 0) + (c.translateX ?? 0),
    translateY: (p.translateY ?? 0) + (c.translateY ?? 0),
    scaleX: (p.scaleX ?? 1) * (c.scaleX ?? 1),
    scaleY: (p.scaleY ?? 1) * (c.scaleY ?? 1),
    rotateDeg: (p.rotateDeg ?? 0) + (c.rotateDeg ?? 0),
    skewXDeg: (p.skewXDeg ?? 0) + (c.skewXDeg ?? 0),
    skewYDeg: (p.skewYDeg ?? 0) + (c.skewYDeg ?? 0),
    origin: c.origin ?? p.origin
  };
}

function objectById(slide: Slide | undefined, objectId: string | undefined): IRObject | undefined {
  if (!objectId) return undefined;
  return flattenObjects(slide?.objects).find((object) => object.id === objectId);
}

function slideById(deck: DeckIR, slideId: string | undefined): Slide | undefined {
  if (!slideId) return undefined;
  return deck.deck.slides.find((slide) => slide.id === slideId);
}

function findObjectAnywhere(deck: DeckIR, objectId: string | undefined): IRObject | undefined {
  if (!objectId) return undefined;
  for (const slide of deck.deck.slides) {
    const object = objectById(slide, objectId);
    if (object) return object;
  }
  return undefined;
}

function stateForEndpoint(
  deck: DeckIR,
  endpoint: { slideId?: string; stateId?: string; objectId?: string; snapshot?: ObjectStateProperties } | undefined,
  fallbackSlide: Slide | undefined,
  fallbackObjectId: string | undefined
): ObjectStateProperties | undefined {
  if (!endpoint) return undefined;
  if (endpoint.snapshot) return endpoint.snapshot;
  const slide = endpoint.slideId ? slideById(deck, endpoint.slideId) : fallbackSlide;
  const state = endpoint.stateId ? slide?.states?.find((item) => item.id === endpoint.stateId) : undefined;
  if (state) return state.properties;
  const objectId = endpoint.objectId || fallbackObjectId;
  const object = objectById(slide, objectId) || findObjectAnywhere(deck, objectId);
  return object ? initialObjectState(object) : undefined;
}

function inferMorphTransitionPairs(
  fromSlide: Slide | undefined,
  toSlide: Slide | undefined,
  options: MorphTransitionOptions | undefined
): { fromObjectId: string; toObjectId: string; morphKey?: string }[] {
  const explicit = options?.pairs ?? [];
  if (explicit.length) return explicit;
  const fromObjects = flattenObjects(fromSlide?.objects);
  const toObjects = flattenObjects(toSlide?.objects);
  const matchBy = options?.matching?.matchBy ?? ["morphKey", "objectId", "name"];
  const fallback = options?.matching?.fallback;
  const used = new Set<string>();
  const pairs: { fromObjectId: string; toObjectId: string; morphKey?: string }[] = [];

  for (const toObject of toObjects) {
    const fromObject = fromObjects.find((candidate) => {
      if (used.has(candidate.id)) return false;
      if (matchBy.includes("morphKey") && toObject.morphKey && candidate.morphKey === toObject.morphKey) return true;
      if (matchBy.includes("objectId") && candidate.id === toObject.id) return true;
      if (matchBy.includes("name") && toObject.name && candidate.name === toObject.name && candidate.type === toObject.type) return true;
      if (matchBy.includes("type") && candidate.type === toObject.type) return true;
      if (matchBy.includes("geometry") && similarGeometry(candidate, toObject, options?.matching?.tolerance)) return true;
      if (fallback === "name" && toObject.name && candidate.name === toObject.name) return true;
      if (fallback === "type" && candidate.type === toObject.type) return true;
      if (fallback === "geometry" && similarGeometry(candidate, toObject, options?.matching?.tolerance)) return true;
      return false;
    });
    if (!fromObject) continue;
    used.add(fromObject.id);
    pairs.push({ fromObjectId: fromObject.id, toObjectId: toObject.id, morphKey: toObject.morphKey || fromObject.morphKey });
  }

  return pairs;
}

function similarGeometry(fromObject: IRObject, toObject: IRObject, tolerance = 0.2): boolean {
  const from = fromObject.bounds;
  const to = toObject.bounds;
  if (!from || !to) return false;
  const maxDimension = Math.max(from.width, from.height, to.width, to.height, 1);
  const normalizedDistance =
    (Math.abs(from.width - to.width) +
      Math.abs(from.height - to.height) +
      Math.abs(from.x - to.x) +
      Math.abs(from.y - to.y)) /
    maxDimension;
  return normalizedDistance <= tolerance;
}

export function createSlideTimingPlan(slide: Slide | undefined): SlideTimingPlan {
  const events = slide?.timeline?.events ?? [];
  const eventMap = new Map(events.map((event) => [event.id, event]));
  const warnings: RuntimeTimingWarning[] = [];
  const constraints: {
    fromId: string;
    toId: string;
    relation: "after" | "with" | "before" | "blocks" | "triggers";
    offsetMs: number;
  }[] = [];
  const nodeToEventId = new Map<string, string>();

  for (const node of slide?.timeline?.dependencyGraph?.nodes ?? []) {
    if (node.eventId && eventMap.has(node.eventId)) nodeToEventId.set(node.id, node.eventId);
  }

  const resolveEventId = (id: string | undefined): string | undefined => {
    if (!id) return undefined;
    if (eventMap.has(id)) return id;
    return nodeToEventId.get(id);
  };
  const addUnresolvedWarning = (eventIds: string[], message: string) => {
    const key = eventIds.join("|") + ":" + message;
    if (warnings.some((warning) => warning.code === "unresolved" && warning.eventIds.join("|") + ":" + warning.message === key)) return;
    warnings.push({ code: "unresolved", eventIds, message });
  };
  const addConstraint = (
    rawFromId: string | undefined,
    rawToId: string | undefined,
    relation: "after" | "with" | "before" | "blocks" | "triggers" = "after",
    offsetMs = 0
  ) => {
    const fromId = resolveEventId(rawFromId);
    const toId = resolveEventId(rawToId);
    if (!fromId || !toId) {
      addUnresolvedWarning(
        [rawFromId, rawToId].filter((id): id is string => Boolean(id)),
        `Timing dependency could not resolve ${rawFromId ?? "<missing>"} -> ${rawToId ?? "<missing>"}.`
      );
      return;
    }
    constraints.push({ fromId, toId, relation, offsetMs: Number(offsetMs || 0) });
  };
  const baseStarts = new Map<string, number>();

  events.forEach((event, index) => {
    const previous = events[index - 1];
    let base = Number(event.delayMs || 0);
    const timing = event.start;
    if (timing?.type === "absolute") {
      base += Number(timing.atMs || 0);
    } else if (timing?.type === "with" || timing?.type === "after" || timing?.type === "before") {
      addConstraint(timing.eventId, event.id, timing.type, Number(timing.offsetMs || 0));
    } else if (timing?.type === "withPrevious" || timing?.type === "afterPrevious") {
      addConstraint(previous?.id, event.id, timing.type === "withPrevious" ? "with" : "after", Number(timing.offsetMs || 0));
    } else if (timing?.type === "trigger" || timing?.type === "onClick") {
      base += Number(timing.type === "trigger" ? timing.offsetMs || 0 : 0);
    }

    for (const dependency of event.dependencies ?? []) {
      addConstraint(dependency.eventId, event.id, dependency.relation ?? "after", Number(dependency.offsetMs || 0));
    }

    baseStarts.set(event.id, Math.max(0, base));
  });

  for (const edge of slide?.timeline?.dependencyGraph?.edges ?? []) {
    addConstraint(edge.from, edge.to, edge.relation, Number(edge.offsetMs || 0));
  }

  const indegree = new Map<string, number>();
  const outgoing = new Map<string, Set<string>>();
  for (const event of events) {
    indegree.set(event.id, 0);
    outgoing.set(event.id, new Set());
  }
  for (const constraint of constraints) {
    if (constraint.fromId === constraint.toId) continue;
    const targets = outgoing.get(constraint.fromId);
    if (!targets || targets.has(constraint.toId)) continue;
    targets.add(constraint.toId);
    indegree.set(constraint.toId, (indegree.get(constraint.toId) ?? 0) + 1);
  }

  const ready = events.filter((event) => (indegree.get(event.id) ?? 0) === 0).map((event) => event.id);
  const order: string[] = [];
  while (ready.length > 0) {
    const id = ready.shift();
    if (!id) continue;
    order.push(id);
    for (const targetId of outgoing.get(id) ?? []) {
      const next = Math.max(0, (indegree.get(targetId) ?? 0) - 1);
      indegree.set(targetId, next);
      if (next === 0) ready.push(targetId);
    }
  }

  const cyclicIds = events.map((event) => event.id).filter((id) => !order.includes(id));
  if (cyclicIds.length > 0) {
    warnings.push({
      code: "cycle",
      eventIds: cyclicIds,
      message: `Timing dependency cycle detected among ${cyclicIds.join(", ")}; source order fallback was used for those events.`
    });
    order.push(...cyclicIds);
  }

  const starts = new Map<string, number>();
  const constraintsByTarget = new Map<string, typeof constraints>();
  for (const constraint of constraints) {
    const list = constraintsByTarget.get(constraint.toId) ?? [];
    list.push(constraint);
    constraintsByTarget.set(constraint.toId, list);
  }

  for (const eventId of order) {
    const event = eventMap.get(eventId);
    if (!event) continue;
    let start = baseStarts.get(eventId) ?? 0;
    for (const constraint of constraintsByTarget.get(eventId) ?? []) {
      const ref = eventMap.get(constraint.fromId);
      if (!ref) continue;
      const refStart = starts.get(constraint.fromId) ?? baseStarts.get(constraint.fromId) ?? 0;
      if (constraint.relation === "with" || constraint.relation === "triggers") {
        start = Math.max(start, refStart + constraint.offsetMs);
      } else if (constraint.relation === "before") {
        start = Math.max(start, refStart - eventDuration(event) - constraint.offsetMs);
      } else {
        start = Math.max(start, refStart + eventDuration(ref) + constraint.offsetMs);
      }
    }
    starts.set(eventId, Math.max(0, start));
  }

  return {
    starts: Object.fromEntries(starts),
    order,
    warnings
  };
}

function computeEventStarts(slide: Slide | undefined): Map<string, number> {
  return new Map(Object.entries(createSlideTimingPlan(slide).starts));
}

function serializeObjectStates(states: Map<string, ObjectStateProperties>): Record<string, ObjectStateProperties> {
  return Object.fromEntries(
    Array.from(states.entries()).map(([objectId, state]) => [objectId, JSON.parse(JSON.stringify(state || {})) as ObjectStateProperties])
  );
}

function eventDuration(event: AnimationEvent | undefined): number {
  return Math.max(0, Number(event?.durationMs || 0));
}

function eventProgress(event: AnimationEvent, start: number, timeMs: number): number | undefined {
  const duration = eventDuration(event);
  const end = start + duration;
  if (duration === 0) return timeMs >= start ? 1 : undefined;
  if (timeMs < start) return event.fill === "backwards" || event.fill === "both" ? 0 : undefined;
  if (timeMs > end) return event.fill === "forwards" || event.fill === "both" ? 1 : undefined;
  return easeProgressValue(event.easing, (timeMs - start) / duration);
}

function keyframeTrackValue(track: KeyframeTrack, progress: number): unknown {
  const frames = [...track.keyframes].sort((a, b) => Number(a.offset) - Number(b.offset));
  if (!frames.length) return undefined;
  const first = frames[0];
  if (first && progress <= Number(first.offset)) return first.value;
  for (let index = 1; index < frames.length; index += 1) {
    const previous = frames[index - 1];
    const next = frames[index];
    if (!previous || !next) continue;
    if (progress <= Number(next.offset)) {
      const local = (progress - Number(previous.offset)) / Math.max(0.0001, Number(next.offset) - Number(previous.offset));
        return interpolateValue(
          previous.value,
          next.value,
          easeProgressValue(next.easing, local),
          interpolationModeForProperty(track.property, track.interpolation)
        );
    }
  }
  return frames[frames.length - 1]?.value;
}

function getRuntimeProperty(stateValue: ObjectStateProperties, property: string): unknown {
  const normalized = normalizeRuntimeProperty(property);
  if (normalized === "bounds") return stateValue.bounds;
  if (normalized === "transform") return stateValue.transform;
  if (normalized === "opacity") return stateValue.opacity;
  if (normalized === "visible") return stateValue.visible;
  if (normalized === "text") return stateValue.text;
  if (normalized === "crop") return stateValue.crop;
  if (normalized === "filter.blurPx") return stateValue.filter?.blurPx;
  if (normalized.startsWith("bounds.")) return stateValue.bounds?.[normalized.slice(7) as keyof NonNullable<ObjectStateProperties["bounds"]>];
  if (normalized === "transform.scale") return stateValue.transform?.scaleX ?? stateValue.transform?.scaleY ?? 1;
  if (normalized.startsWith("transform.")) return stateValue.transform?.[normalized.slice(10) as keyof Transform2D];
  if (normalized === "style.fill") return stateValue.style?.fill;
  if (normalized === "style.stroke") return stateValue.style?.stroke;
  if (normalized.startsWith("style.stroke.")) return stateValue.style?.stroke?.[normalized.slice(13) as keyof NonNullable<NonNullable<ObjectStateProperties["style"]>["stroke"]>];
  if (normalized.startsWith("style.textStyle.")) {
    return stateValue.style?.textStyle?.[
      normalized.slice(16) as keyof NonNullable<NonNullable<ObjectStateProperties["style"]>["textStyle"]>
    ];
  }
  return undefined;
}

function setRuntimeProperty(stateValue: ObjectStateProperties, property: string, value: unknown): void {
  if (value === undefined) return;
  const normalized = normalizeRuntimeProperty(property);
  if (normalized === "opacity") {
    const numeric = finiteNumber(value);
    if (numeric !== undefined) stateValue.opacity = numeric;
  } else if (normalized === "visible") {
    stateValue.visible = Boolean(value);
  } else if (normalized === "text") {
    stateValue.text = normalizeTextValue(value);
  } else if (normalized === "crop") {
    stateValue.crop = value as ObjectStateProperties["crop"];
  } else if (normalized === "filter.blurPx") {
    stateValue.filter = { ...(stateValue.filter || {}) };
    setNumericField(stateValue.filter, "blurPx", value);
  } else if (normalized === "bounds" && isRecord(value)) {
    stateValue.bounds = { ...(stateValue.bounds || { x: 0, y: 0, width: 0, height: 0 }) };
    setNumericField(stateValue.bounds, "x", value.x);
    setNumericField(stateValue.bounds, "y", value.y);
    setNumericField(stateValue.bounds, "width", value.width);
    setNumericField(stateValue.bounds, "height", value.height);
  } else if (normalized.startsWith("bounds.")) {
    stateValue.bounds = { ...(stateValue.bounds || { x: 0, y: 0, width: 0, height: 0 }) };
    setNumericField(stateValue.bounds, normalized.slice(7), value);
  } else if (normalized === "transform") {
    stateValue.transform = { ...(stateValue.transform || {}), ...normalizeTransformPatch(value) };
  } else if (normalized === "transform.scale") {
    stateValue.transform = { ...(stateValue.transform || {}) };
    applyScaleValue(stateValue.transform, value);
  } else if (normalized.startsWith("transform.")) {
    stateValue.transform = { ...(stateValue.transform || {}) };
    setNumericField(stateValue.transform, normalized.slice(10), value);
  } else if (normalized === "style.fill") {
    stateValue.style = { ...(stateValue.style || {}), fill: value as ObjectStateProperties["style"] extends { fill?: infer T } ? T : never };
  } else if (normalized === "style.stroke") {
    stateValue.style = { ...(stateValue.style || {}), stroke: value as ObjectStateProperties["style"] extends { stroke?: infer T } ? T : never };
  } else if (normalized.startsWith("style.stroke.")) {
    stateValue.style = { ...(stateValue.style || {}), stroke: { ...(stateValue.style?.stroke || {}) } };
    setStyleField(stateValue.style.stroke, normalized.slice(13), value);
  } else if (normalized.startsWith("style.textStyle.")) {
    stateValue.style = { ...(stateValue.style || {}), textStyle: { ...(stateValue.style?.textStyle || {}) } };
    setStyleField(stateValue.style.textStyle, normalized.slice(16), value);
  }
}

function interpolationModeForProperty(property: string, mode?: string): string | undefined {
  if (mode) return mode;
  const normalized = normalizeRuntimeProperty(property);
  if (normalized === "style.fill" || normalized === "style.stroke" || normalized.endsWith(".color")) return "color";
  if (normalized === "bounds" || normalized === "transform" || normalized === "crop") return "matrix";
  return undefined;
}

function normalizeRuntimeProperty(property: string): NormalizedRuntimeProperty {
  const raw = String(property || "");
  const lower = raw.toLowerCase();
  const direct: Record<string, NormalizedRuntimeProperty> = {
    x: "bounds.x",
    y: "bounds.y",
    left: "bounds.x",
    top: "bounds.y",
    width: "bounds.width",
    height: "bounds.height",
    fill: "style.fill",
    stroke: "style.stroke",
    "stroke.width": "style.stroke.width",
    "stroke.color": "style.stroke.color",
    blur: "filter.blurPx",
    blurpx: "filter.blurPx",
    "filter.blur": "filter.blurPx",
    "filter.blurpx": "filter.blurPx",
    "filters.gaussianblur": "filter.blurPx",
    "custom.keynote.filters.gaussianblur": "filter.blurPx",
    tx: "transform.translateX",
    ty: "transform.translateY",
    translatex: "transform.translateX",
    translatey: "transform.translateY",
    scalex: "transform.scaleX",
    scaley: "transform.scaleY",
    scale: "transform.scale",
    rotate: "transform.rotateDeg",
    rotation: "transform.rotateDeg",
    rotatedeg: "transform.rotateDeg",
    rotationdeg: "transform.rotateDeg",
    skewx: "transform.skewXDeg",
    skewy: "transform.skewYDeg"
  };
  if (direct[lower]) return direct[lower];
  if (!lower.startsWith("transform.")) return raw;

  const suffix = lower.slice("transform.".length);
  const transformAliases: Record<string, NormalizedRuntimeProperty> = {
    tx: "transform.translateX",
    ty: "transform.translateY",
    x: "transform.translateX",
    y: "transform.translateY",
    translatex: "transform.translateX",
    translatey: "transform.translateY",
    scalex: "transform.scaleX",
    scaley: "transform.scaleY",
    scale: "transform.scale",
    rotate: "transform.rotateDeg",
    rotation: "transform.rotateDeg",
    rotatedeg: "transform.rotateDeg",
    rotationdeg: "transform.rotateDeg",
    skewx: "transform.skewXDeg",
    skewy: "transform.skewYDeg"
  };
  return transformAliases[suffix] ?? raw;
}

function normalizeTransformPatch(value: unknown): Partial<Transform2D> {
  if (!isRecord(value)) return {};
  const transform: Partial<Transform2D> = {};
  setNumericField(transform, "translateX", value.translateX ?? value.x ?? value.tx);
  setNumericField(transform, "translateY", value.translateY ?? value.y ?? value.ty);
  setNumericField(transform, "scaleX", value.scaleX);
  setNumericField(transform, "scaleY", value.scaleY);
  applyScaleValue(transform, value.scale);
  setNumericField(transform, "rotateDeg", value.rotateDeg ?? value.rotationDeg ?? value.rotate ?? value.rotation);
  setNumericField(transform, "skewXDeg", value.skewXDeg ?? value.skewX);
  setNumericField(transform, "skewYDeg", value.skewYDeg ?? value.skewY);
  if (isRecord(value.origin)) {
    const x = finiteNumber(value.origin.x);
    const y = finiteNumber(value.origin.y);
    if (x !== undefined && y !== undefined) transform.origin = { x, y };
  }
  return transform;
}

function applyScaleValue(transform: Partial<Transform2D>, value: unknown): void {
  if (isRecord(value)) {
    setNumericField(transform, "scaleX", value.scaleX ?? value.x);
    setNumericField(transform, "scaleY", value.scaleY ?? value.y);
    const uniform = value.scale ?? value.value;
    const numeric = finiteNumber(uniform);
    if (numeric !== undefined) {
      transform.scaleX = numeric;
      transform.scaleY = numeric;
    }
    return;
  }
  const numeric = finiteNumber(value);
  if (numeric !== undefined) {
    transform.scaleX = numeric;
    transform.scaleY = numeric;
  }
}

function normalizeTextValue(value: unknown): NonNullable<ObjectStateProperties["text"]> {
  return typeof value === "string" ? { plainText: value } : (value as NonNullable<ObjectStateProperties["text"]>);
}

function interpolateObjectState(
  from: ObjectStateProperties | undefined,
  to: ObjectStateProperties | undefined,
  progress: number,
  properties: MorphProperty[] | undefined
): ObjectStateProperties {
  const p = clamp(progress, 0, 1);
  const result = mergeObjectStateProperties({}, to || {});
  const allowed = new Set<string>(properties && properties.length ? properties : ["bounds", "transform", "opacity", "filter", "fill", "stroke"]);
  if (allowed.has("all") || allowed.has("bounds")) {
    result.bounds = { ...(to?.bounds || { x: 0, y: 0, width: 0, height: 0 }) };
    for (const key of ["x", "y", "width", "height"] as const) {
      result.bounds[key] = interpolateValue(from?.bounds?.[key], to?.bounds?.[key], p, "number") as number;
    }
  }
  if (allowed.has("all") || allowed.has("transform")) {
    result.transform = { ...(to?.transform || {}) };
    for (const key of ["translateX", "translateY", "scaleX", "scaleY", "rotateDeg", "skewXDeg", "skewYDeg"] as const) {
      const fallback = key === "scaleX" || key === "scaleY" ? 1 : 0;
      result.transform[key] = interpolateValue(from?.transform?.[key] ?? fallback, to?.transform?.[key] ?? fallback, p, "number") as number;
    }
  }
  if (allowed.has("all") || allowed.has("opacity")) result.opacity = interpolateValue(from?.opacity ?? 1, to?.opacity ?? 1, p, "number") as number;
  if (allowed.has("all") || allowed.has("filter")) {
    result.filter = { ...(to?.filter || {}) };
    result.filter.blurPx = interpolateValue(from?.filter?.blurPx ?? 0, to?.filter?.blurPx ?? 0, p, "number") as number;
    result.filter.brightness = interpolateValue(from?.filter?.brightness ?? 1, to?.filter?.brightness ?? 1, p, "number") as number;
    result.filter.contrast = interpolateValue(from?.filter?.contrast ?? 1, to?.filter?.contrast ?? 1, p, "number") as number;
    result.filter.saturate = interpolateValue(from?.filter?.saturate ?? 1, to?.filter?.saturate ?? 1, p, "number") as number;
  }
  if (allowed.has("all") || allowed.has("fill")) {
    result.style = { ...(result.style || {}) };
    result.style.fill = interpolateFillValue(from?.style?.fill, to?.style?.fill, p);
  }
  if (allowed.has("all") || allowed.has("stroke")) {
    result.style = { ...(result.style || {}) };
    result.style.stroke = interpolateStrokeValue(from?.style?.stroke, to?.style?.stroke, p);
  }
  if (allowed.has("all") || allowed.has("text")) result.text = p < 1 ? from?.text ?? to?.text : to?.text;
  if (allowed.has("all") || allowed.has("crop")) {
    result.crop = interpolatePlainObject(from?.crop, to?.crop, p, ["x", "y", "width", "height"]);
  }
  result.visible = p < 1 ? from?.visible ?? true : to?.visible ?? true;
  return result;
}

function interpolateValue(from: unknown, to: unknown, progress: number, mode?: string): unknown {
  const p = clamp(progress, 0, 1);
  if (to === undefined) return from;
  if (from === undefined) return to;
  if (mode === "matrix") return interpolatePlainObject(from, to, p);
  if (mode === "number") {
    const a = finiteNumber(from);
    const b = finiteNumber(to);
    return a !== undefined && b !== undefined ? a + (b - a) * p : p < 1 ? from : to;
  }
  if (mode === "color") {
    if (isRecord(from) && isRecord(to)) {
      if (looksLikeFill(from) || looksLikeFill(to)) return interpolateFillValue(from as Fill, to as Fill, p);
      if (looksLikeStroke(from) || looksLikeStroke(to)) return interpolateStrokeValue(from, to, p);
    }
    return interpolateColorValue(from, to, p);
  }
  if (mode === "discrete" || typeof from === "boolean" || typeof to === "boolean" || typeof from === "string" || typeof to === "string") {
    return p < 1 ? from : to;
  }
  if (isFiniteNumber(from) && isFiniteNumber(to)) return from + (to - from) * p;
  if (Array.isArray(from) && Array.isArray(to) && from.length === to.length) {
    return from.map((value, index) => interpolateValue(value, to[index], p, mode));
  }
  if (isRecord(from) && isRecord(to)) {
    if (looksLikeFill(from) || looksLikeFill(to)) return interpolateFillValue(from as Fill, to as Fill, p);
    if (looksLikeStroke(from) || looksLikeStroke(to)) return interpolateStrokeValue(from, to, p);
    return interpolatePlainObject(from, to, p);
  }
  return p < 1 ? from : to;
}

function interpolatePlainObject(
  from: unknown,
  to: unknown,
  progress: number,
  numericKeys?: string[]
): Record<string, unknown> | undefined {
  if (!isRecord(from) || !isRecord(to)) return progress < 1 ? (from as Record<string, unknown> | undefined) : (to as Record<string, unknown> | undefined);
  const result: Record<string, unknown> = { ...to };
  const keys = numericKeys ?? Array.from(new Set([...Object.keys(from), ...Object.keys(to)]));
  for (const key of keys) {
    const a = from[key];
    const b = to[key];
    const numberA = finiteNumber(a);
    const numberB = finiteNumber(b);
    if (numberA !== undefined && numberB !== undefined) result[key] = numberA + (numberB - numberA) * clamp(progress, 0, 1);
    else result[key] = progress < 1 ? a ?? b : b ?? a;
  }
  return result;
}

function interpolateFillValue(from: Fill | undefined, to: Fill | undefined, progress: number): Fill | undefined {
  const p = clamp(progress, 0, 1);
  if (!from) return to;
  if (!to) return from;
  if (from.type === "solid" && to.type === "solid") {
    return { ...to, color: interpolateColorValue(from.color, to.color, p) as Color };
  }
  if (from.type === "gradient" && to.type === "gradient" && from.stops.length === to.stops.length) {
    return {
      ...to,
      angleDeg:
        finiteNumber(from.angleDeg) !== undefined && finiteNumber(to.angleDeg) !== undefined
          ? (interpolateValue(from.angleDeg, to.angleDeg, p, "number") as number)
          : to.angleDeg,
      stops: to.stops.map((stop, index) => ({
        ...stop,
        offset: interpolateValue(from.stops[index]?.offset, stop.offset, p, "number") as number,
        color: interpolateColorValue(from.stops[index]?.color, stop.color, p) as typeof stop.color
      }))
    };
  }
  return p < 1 ? from : to;
}

function interpolateStrokeValue(
  from: Stroke | undefined,
  to: Stroke | undefined,
  progress: number
): Stroke | undefined {
  const p = clamp(progress, 0, 1);
  if (!from) return to;
  if (!to) return from;
  return {
    ...to,
    width: interpolateValue(from.width, to.width, p, "number") as number,
    color: interpolateColorValue(from.color, to.color, p) as typeof to.color
  };
}

function looksLikeFill(value: Record<string, unknown>): boolean {
  return typeof value.type === "string" && ["none", "solid", "gradient", "image"].includes(value.type);
}

function looksLikeStroke(value: Record<string, unknown>): boolean {
  return "width" in value || "dash" in value || "lineCap" in value || "lineJoin" in value || "color" in value;
}

function interpolateColorValue(from: unknown, to: unknown, progress: number): unknown {
  const parse = (value: unknown): [number, number, number, number] | undefined => {
    if (isRecord(value) && "value" in value) {
      const parsed = parse(value.value);
      if (parsed) parsed[3] = finiteNumber(value.alpha) ?? parsed[3];
      return parsed;
    }
    const css = colorToCss(value);
    if (typeof css !== "string") return undefined;
    const hex = css.match(/^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i);
    if (hex?.[1]) {
      const value = hex[1].length === 3 ? hex[1].split("").map((char) => char + char).join("") : hex[1];
      return [
        parseInt(value.slice(0, 2), 16),
        parseInt(value.slice(2, 4), 16),
        parseInt(value.slice(4, 6), 16),
        value.length === 8 ? parseInt(value.slice(6, 8), 16) / 255 : 1
      ];
    }
    const rgb = css.match(/^rgba?\(([^)]+)\)$/i);
    if (rgb?.[1]) {
      const channels = rgb[1].split(",").map((part) => Number(part.trim().replace("%", "")));
      if (channels.length >= 3 && channels.slice(0, 3).every(Number.isFinite)) {
        return [channels[0] ?? 0, channels[1] ?? 0, channels[2] ?? 0, Number.isFinite(channels[3]) ? channels[3] ?? 1 : 1];
      }
    }
    return undefined;
  };
  const a = parse(from);
  const b = parse(to);
  if (!a || !b) return progress < 1 ? from : to;
  const channelNumber = (index: 0 | 1 | 2) => Math.round(a[index] + (b[index] - a[index]) * progress);
  const alpha = a[3] + (b[3] - a[3]) * progress;
  if (alpha < 0.999) return `rgba(${channelNumber(0)}, ${channelNumber(1)}, ${channelNumber(2)}, ${Number(alpha.toFixed(3))})`;
  const channel = (index: 0 | 1 | 2) => channelNumber(index).toString(16).padStart(2, "0");
  return `#${channel(0)}${channel(1)}${channel(2)}`;
}

export function easeProgressValue(easing: Easing | undefined, progress: number): number {
  const p = clamp(progress, 0, 1);
  if (!easing || easing === "linear") return p;
  if (typeof easing !== "string") {
    if (easing.type === "cubicBezier") return cubicBezierProgress(easing.x1, easing.y1, easing.x2, easing.y2, p);
    if (easing.type === "steps") return stepsProgress(easing.count, easing.position, p);
    if (easing.type === "spring") return springProgress(easing, p);
    return p;
  }
  if (easing === "ease") return cubicBezierProgress(0.25, 0.1, 0.25, 1, p);
  if (easing === "easeIn") return cubicBezierProgress(0.42, 0, 1, 1, p);
  if (easing === "easeOut") return cubicBezierProgress(0, 0, 0.58, 1, p);
  if (easing === "easeInOut") return cubicBezierProgress(0.42, 0, 0.58, 1, p);
  if (easing === "easeInCubic") return p * p * p;
  if (easing === "easeOutCubic") return 1 - Math.pow(1 - p, 3);
  if (easing === "easeInOutCubic") return p < 0.5 ? 4 * p * p * p : 1 - Math.pow(-2 * p + 2, 3) / 2;
  if (easing === "backIn") return backInProgress(p);
  if (easing === "backOut") return 1 - backInProgress(1 - p);
  if (easing === "backInOut") return p < 0.5 ? backInProgress(p * 2) / 2 : 1 - backInProgress((1 - p) * 2) / 2;
  if (easing === "bounceIn") return 1 - bounceOutProgress(1 - p);
  if (easing === "bounceOut") return bounceOutProgress(p);
  if (easing === "bounceInOut") return p < 0.5 ? (1 - bounceOutProgress(1 - 2 * p)) / 2 : (1 + bounceOutProgress(2 * p - 1)) / 2;
  if (easing === "elasticIn") return elasticInProgress(p);
  if (easing === "elasticOut") return 1 - elasticInProgress(1 - p);
  if (easing === "elasticInOut") return p < 0.5 ? elasticInProgress(2 * p) / 2 : 1 - elasticInProgress(2 - 2 * p) / 2;
  return p;
}

function cubicBezierProgress(x1: number, y1: number, x2: number, y2: number, progress: number): number {
  const p = clamp(progress, 0, 1);
  const sample = (a1: number, a2: number, t: number) => {
    const inverse = 1 - t;
    return 3 * inverse * inverse * t * a1 + 3 * inverse * t * t * a2 + t * t * t;
  };
  let lower = 0;
  let upper = 1;
  let t = p;
  for (let index = 0; index < 20; index += 1) {
    t = (lower + upper) / 2;
    if (sample(x1, x2, t) < p) lower = t;
    else upper = t;
  }
  return clamp(sample(y1, y2, t), 0, 1);
}

function stepsProgress(count: number, position: "start" | "end" | undefined, progress: number): number {
  const steps = Math.max(1, Math.floor(Number(count) || 1));
  const scaled = clamp(progress, 0, 1) * steps;
  return clamp((position === "start" ? Math.ceil(scaled) : Math.floor(scaled)) / steps, 0, 1);
}

function springProgress(easing: Extract<Easing, { type: "spring" }>, progress: number): number {
  const p = clamp(progress, 0, 1);
  const mass = Math.max(0.001, Number(easing.mass ?? 1));
  const stiffness = Math.max(0.001, Number(easing.stiffness ?? 100));
  const damping = Math.max(0.001, Number(easing.damping ?? 10));
  const velocity = Number(easing.velocity ?? 0);
  const angularFrequency = Math.sqrt(stiffness / mass);
  const dampingRatio = damping / (2 * Math.sqrt(stiffness * mass));
  if (dampingRatio < 1) {
    const damped = angularFrequency * Math.sqrt(1 - dampingRatio * dampingRatio);
    const envelope = Math.exp(-dampingRatio * angularFrequency * p);
    return clamp(1 - envelope * (Math.cos(damped * p) + ((dampingRatio * angularFrequency - velocity) / damped) * Math.sin(damped * p)), 0, 1);
  }
  return clamp(1 - Math.exp(-angularFrequency * p), 0, 1);
}

function backInProgress(progress: number): number {
  const c1 = 1.70158;
  const c3 = c1 + 1;
  return c3 * progress * progress * progress - c1 * progress * progress;
}

function bounceOutProgress(progress: number): number {
  const n1 = 7.5625;
  const d1 = 2.75;
  if (progress < 1 / d1) return n1 * progress * progress;
  if (progress < 2 / d1) {
    const p = progress - 1.5 / d1;
    return n1 * p * p + 0.75;
  }
  if (progress < 2.5 / d1) {
    const p = progress - 2.25 / d1;
    return n1 * p * p + 0.9375;
  }
  const p = progress - 2.625 / d1;
  return n1 * p * p + 0.984375;
}

function elasticInProgress(progress: number): number {
  if (progress === 0 || progress === 1) return progress;
  return -Math.pow(2, 10 * progress - 10) * Math.sin(((progress * 10 - 10.75) * (2 * Math.PI)) / 3);
}

function setNumericField(target: object, key: string, value: unknown): void {
  const numeric = finiteNumber(value);
  if (numeric !== undefined) (target as Record<string, unknown>)[key] = numeric;
}

function setStyleField(target: object, key: string, value: unknown): void {
  const numericKeys = new Set(["width", "fontSize", "letterSpacing", "lineHeight"]);
  const numeric = numericKeys.has(key) ? finiteNumber(value) : undefined;
  (target as Record<string, unknown>)[key] = numeric ?? value;
}

function finiteNumber(value: unknown): number | undefined {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : undefined;
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function transformToCss(transform: Transform2D | undefined): string {
  const normalized = normalizeTransformPatch(transform);
  return `translate(${normalized.translateX ?? 0}px,${normalized.translateY ?? 0}px) scale(${normalized.scaleX ?? 1},${normalized.scaleY ?? 1}) rotate(${normalized.rotateDeg ?? 0}deg) skew(${normalized.skewXDeg ?? 0}deg,${normalized.skewYDeg ?? 0}deg)`;
}

function filterToCss(filter: ObjectStateProperties["filter"] | undefined): string {
  const parts: string[] = [];
  const blurPx = finiteNumber(filter?.blurPx) ?? 0;
  const brightness = finiteNumber(filter?.brightness);
  const contrast = finiteNumber(filter?.contrast);
  const saturate = finiteNumber(filter?.saturate);
  if (blurPx > 0) parts.push(`blur(${blurPx}px)`);
  if (brightness !== undefined && brightness !== 1) parts.push(`brightness(${brightness})`);
  if (contrast !== undefined && contrast !== 1) parts.push(`contrast(${contrast})`);
  if (saturate !== undefined && saturate !== 1) parts.push(`saturate(${saturate})`);
  return parts.length > 0 ? parts.join(" ") : "none";
}

function firstTextStyle(object: TextObject | ShapeObject, state?: ObjectStateProperties) {
  const text = "text" in object ? object.text : undefined;
  return text?.runs?.[0]?.style ?? state?.style?.textStyle ?? object.style?.textStyle ?? {};
}

function textContent(object: TextObject | ShapeObject): string {
  const text = "text" in object ? object.text : undefined;
  return text?.plainText ?? text?.runs?.map((run) => run.text).join("") ?? "";
}

function fillToCss(fill: Fill | undefined): string | undefined {
  if (!fill) return undefined;
  if (fill.type === "solid") return colorToCss(fill.color);
  if (fill.type === "none") return "transparent";
  if (fill.type === "gradient") {
    const stops = fill.stops
      .map((stop) => `${colorToCss(stop.color) ?? "transparent"} ${Math.round(stop.offset * 100)}%`)
      .join(", ");
    return `linear-gradient(${fill.angleDeg ?? 180}deg, ${stops})`;
  }
  if (fill.type === "image") {
    const source = fill.source.dataUri ?? fill.source.uri;
    return source ? `url("${source.replaceAll('"', "%22")}")` : undefined;
  }
  return undefined;
}

function colorToCss(color: unknown): string | undefined {
  if (typeof color === "string") return color;
  if (color && typeof color === "object" && "value" in color) return String(color.value);
  return undefined;
}

export function easingToCss(easing: Easing | undefined): string {
  if (!easing) return "linear";
  if (typeof easing !== "string") {
    if (easing.type === "cubicBezier") return `cubic-bezier(${easing.x1},${easing.y1},${easing.x2},${easing.y2})`;
    if (easing.type === "steps") return `steps(${easing.count}, ${easing.position ?? "end"})`;
    return "linear";
  }
  return (
    {
      ease: "ease",
      easeIn: "ease-in",
      easeOut: "ease-out",
      easeInOut: "ease-in-out",
      easeInCubic: "cubic-bezier(.32,0,.67,0)",
      easeOutCubic: "cubic-bezier(.22,1,.36,1)",
      easeInOutCubic: "cubic-bezier(.65,0,.35,1)",
      backIn: "cubic-bezier(.36,0,.66,-.56)",
      backOut: "cubic-bezier(.34,1.56,.64,1)",
      backInOut: "cubic-bezier(.68,-.6,.32,1.6)",
      bounceIn: "cubic-bezier(.11,0,.5,0)",
      bounceOut: "cubic-bezier(.5,1,.89,1)",
      bounceInOut: "cubic-bezier(.76,0,.24,1)",
      elasticIn: "cubic-bezier(.7,0,.84,0)",
      elasticOut: "cubic-bezier(.16,1,.3,1)",
      elasticInOut: "cubic-bezier(.87,0,.13,1)"
    } satisfies Record<string, string>
  )[easing] ?? easing;
}

export function keyframeEventToCssFrames(event: KeyframeAnimationEvent): Record<string, string | number>[] {
  const grouped = new Map<number, Record<string, string | number>>();
  const transforms = new Map<number, Partial<Record<keyof Transform2D, number>>>();

  for (const track of event.tracks) {
    for (const keyframe of track.keyframes) {
      const frame = grouped.get(keyframe.offset) ?? { offset: keyframe.offset };
      const transformValue = transformTrackValue(track, keyframe.value);
      if (transformValue) {
        const current = transforms.get(keyframe.offset) ?? {};
        Object.assign(current, transformValue);
        transforms.set(keyframe.offset, current);
      } else {
        Object.assign(frame, propertyValueToCss(track.property, keyframe.value));
      }
      grouped.set(keyframe.offset, frame);
    }
  }

  for (const [offset, transform] of transforms) {
    const frame = grouped.get(offset) ?? { offset };
    frame.transform = transformToCss(transform);
    grouped.set(offset, frame);
  }

  return Array.from(grouped.values()).sort((a, b) => Number(a.offset) - Number(b.offset));
}

function transformTrackValue(track: KeyframeTrack, value: unknown): Partial<Record<keyof Transform2D, number>> | undefined {
  const property = normalizeRuntimeProperty(track.property);
  if (property === "transform.scale") {
    const numeric = Number(value);
    return Number.isFinite(numeric) ? { scaleX: numeric, scaleY: numeric } : undefined;
  }
  if (!property.startsWith("transform.")) return undefined;
  const key = property.slice("transform.".length) as keyof Transform2D;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? { [key]: numeric } : undefined;
}

function propertyValueToCss(property: string, value: unknown): Record<string, string> {
  const normalized = normalizeRuntimeProperty(property);
  if (normalized === "opacity") return { opacity: String(value) };
  if (normalized === "visible") return { visibility: value ? "visible" : "hidden" };
  if (normalized === "bounds.x") return { left: `${value}px` };
  if (normalized === "bounds.y") return { top: `${value}px` };
  if (normalized === "bounds.width") return { width: `${value}px` };
  if (normalized === "bounds.height") return { height: `${value}px` };
  if (normalized === "filter.blurPx") return { filter: filterToCss({ blurPx: finiteNumber(value) ?? 0 }) };
  if (normalized === "style.fill") return { background: fillToCss(value as Fill) ?? String(value) };
  return {};
}

function transitionDurationMs(transition: SlideTransition | null | undefined): number {
  if (!transition || transition.type === "none" || transition.type === "cut") return 0;
  return Math.max(1, Number(transition.durationMs ?? DEFAULT_TRANSITION_DURATION_MS));
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function escapeHtml(value: unknown): string {
  return String(value ?? "").replace(/[&<>"']/g, (char) => {
    switch (char) {
      case "&":
        return "&amp;";
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case '"':
        return "&quot;";
      default:
        return "&#39;";
    }
  });
}
