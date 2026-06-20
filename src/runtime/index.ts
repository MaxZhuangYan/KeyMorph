import type {
  DeckIR,
  Easing,
  Fill,
  IRObject,
  KeyframeAnimationEvent,
  KeyframeTrack,
  ObjectStateProperties,
  ShapeObject,
  Slide,
  SlideTransition,
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

export function renderSlideMarkup(slide: Slide): string {
  return slide.objects.map((object) => renderObjectMarkup(object)).join("");
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
    raf: 0,
    startedAt: 0,
    renderedKey: ""
  };

  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
  const clamp = (value, min, max) => Math.min(max, Math.max(min, value));
  const isNumber = (value) => typeof value === "number" && Number.isFinite(value);
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
    if (easing === "easeIn" || easing === "easeInCubic") return p * p * p;
    if (easing === "easeOut" || easing === "easeOutCubic") return 1 - Math.pow(1 - p, 3);
    if (easing === "easeInOut" || easing === "easeInOutCubic") return p < 0.5 ? 4 * p * p * p : 1 - Math.pow(-2 * p + 2, 3) / 2;
    return p;
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
      const source = input.source?.dataUri || input.source?.uri || "";
      return source ? "url('" + String(source).replace(/'/g, "%27") + "')" : undefined;
    }
    if (typeof input === "string") return input;
    return undefined;
  };
  const strokeColorCss = (stroke) => colorCss(stroke?.color) || "transparent";
  const textOf = (object, statePatch) => statePatch?.text?.plainText || (statePatch?.text?.runs || []).map((run) => run.text).join("") || object.text?.plainText || (object.text?.runs || []).map((run) => run.text).join("") || "";
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
  const merge = (base, patch) => {
    if (!patch) return { ...(base || {}) };
    const next = { ...(base || {}), ...patch };
    if (base?.bounds || patch.bounds) next.bounds = { ...(base?.bounds || {}), ...(patch.bounds || {}) };
    if (base?.transform || patch.transform) next.transform = { ...(base?.transform || {}), ...(patch.transform || {}) };
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
    transform: object.transform || {},
    style: object.style || {},
    text: object.text,
    crop: object.crop,
    media: object.playback
  }, object.initialState);
  const transformCss = (transform) => {
    const t = transform || {};
    return "translate(" + Number(t.translateX || 0) + "px," + Number(t.translateY || 0) + "px) scale(" + Number(t.scaleX ?? 1) + "," + Number(t.scaleY ?? 1) + ") rotate(" + Number(t.rotateDeg || 0) + "deg) skew(" + Number(t.skewXDeg || 0) + "deg," + Number(t.skewYDeg || 0) + "deg)";
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
      "transform:" + transformCss(current.transform)
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
    if (object.type === "text") return '<div ' + common + textStyleCss(object) + '">' + esc(textOf(object)) + '</div>';
    if (object.type === "image") return '<img ' + common + '" src="' + esc(object.source?.dataUri || object.source?.uri || "") + '" alt="' + esc(object.altText || object.name || "") + '">';
    if (object.type === "media") {
      const tag = object.mediaType === "audio" ? "audio" : "video";
      const poster = tag === "video" && object.posterSource ? ' poster="' + esc(object.posterSource.dataUri || object.posterSource.uri || "") + '"' : "";
      return '<' + tag + ' ' + common + '" src="' + esc(object.source?.dataUri || object.source?.uri || "") + '"' + poster + ' muted playsinline></' + tag + '>';
    }
    if (object.type === "group") return '<div ' + common + 'overflow:visible">' + (object.children || []).map(objectHtml).join("") + '</div>';
    if (object.type === "shape" && object.text) return '<div ' + common + textStyleCss(object) + '">' + esc(textOf(object)) + '</div>';
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
    if (property === "opacity") return stateValue.opacity;
    if (property === "visible") return stateValue.visible;
    if (property === "text") return stateValue.text;
    if (property === "crop") return stateValue.crop;
    if (property.startsWith("bounds.")) return stateValue.bounds?.[property.slice(7)];
    if (property.startsWith("transform.")) return stateValue.transform?.[property.slice(10)];
    if (property === "style.fill") return stateValue.style?.fill;
    if (property === "style.stroke") return stateValue.style?.stroke;
    return undefined;
  };
  const setProperty = (stateValue, property, value) => {
    if (property === "opacity") stateValue.opacity = Number(value);
    else if (property === "visible") stateValue.visible = Boolean(value);
    else if (property === "text") stateValue.text = value;
    else if (property === "crop") stateValue.crop = value;
    else if (property.startsWith("bounds.")) {
      stateValue.bounds = { ...(stateValue.bounds || {}) };
      stateValue.bounds[property.slice(7)] = Number(value);
    } else if (property.startsWith("transform.")) {
      stateValue.transform = { ...(stateValue.transform || {}) };
      stateValue.transform[property.slice(10)] = Number(value);
    } else if (property === "style.fill") {
      stateValue.style = { ...(stateValue.style || {}), fill: value };
    } else if (property === "style.stroke") {
      stateValue.style = { ...(stateValue.style || {}), stroke: value };
    }
  };
  const interpolateColor = (from, to, progress) => {
    const parse = (value) => {
      const css = colorCss(value);
      const match = typeof css === "string" ? css.match(/^#([0-9a-f]{6})$/i) : null;
      return match ? [parseInt(match[1].slice(0, 2), 16), parseInt(match[1].slice(2, 4), 16), parseInt(match[1].slice(4, 6), 16)] : undefined;
    };
    const a = parse(from);
    const b = parse(to);
    if (!a || !b) return progress < 1 ? from : to;
    const channel = (index) => Math.round(a[index] + (b[index] - a[index]) * progress).toString(16).padStart(2, "0");
    return "#" + channel(0) + channel(1) + channel(2);
  };
  const interpolate = (from, to, progress, mode) => {
    const p = clamp(progress, 0, 1);
    if (mode === "discrete" || typeof from === "boolean" || typeof to === "boolean" || typeof from === "string" || typeof to === "string") return p < 1 ? from : to;
    if (isNumber(from) && isNumber(to)) return from + (to - from) * p;
    if (mode === "color") return interpolateColor(from, to, p);
    return p < 1 ? from : to;
  };
  const interpolateState = (from, to, progress, properties) => {
    const result = merge({}, to || {});
    const allowed = new Set(properties && properties.length ? properties : ["bounds", "transform", "opacity", "fill", "stroke"]);
    if (allowed.has("all") || allowed.has("bounds")) {
      result.bounds = { ...(to?.bounds || {}) };
      for (const key of ["x", "y", "width", "height"]) result.bounds[key] = interpolate(from?.bounds?.[key], to?.bounds?.[key], progress);
    }
    if (allowed.has("all") || allowed.has("transform")) {
      result.transform = { ...(to?.transform || {}) };
      for (const key of ["translateX", "translateY", "scaleX", "scaleY", "rotateDeg", "skewXDeg", "skewYDeg"]) {
        const fallback = key === "scaleX" || key === "scaleY" ? 1 : 0;
        result.transform[key] = interpolate(from?.transform?.[key] ?? fallback, to?.transform?.[key] ?? fallback, progress);
      }
    }
    if (allowed.has("all") || allowed.has("opacity")) result.opacity = interpolate(from?.opacity ?? 1, to?.opacity ?? 1, progress);
    if (allowed.has("all") || allowed.has("fill")) {
      result.style = { ...(result.style || {}) };
      result.style.fill = progress < 1 ? from?.style?.fill || to?.style?.fill : to?.style?.fill;
    }
    if (allowed.has("all") || allowed.has("stroke")) {
      result.style = { ...(result.style || {}) };
      result.style.stroke = progress < 1 ? from?.style?.stroke || to?.style?.stroke : to?.style?.stroke;
    }
    result.visible = progress < 1 ? from?.visible ?? true : to?.visible ?? true;
    return result;
  };
  const eventDuration = (event) => Math.max(0, Number(event.durationMs || 0));
  const computeEventStarts = (slide) => {
    const events = slide?.timeline?.events || [];
    const starts = new Map();
    const eventMap = new Map(events.map((event) => [event.id, event]));
    const graphEdges = slide?.timeline?.dependencyGraph?.edges || [];
    for (let pass = 0; pass < Math.max(2, events.length + 1); pass += 1) {
      events.forEach((event, index) => {
        const previous = events[index - 1];
        let start = Number(event.delayMs || 0);
        const timing = event.start;
        if (timing?.type === "absolute") start += Number(timing.atMs || 0);
        else if (timing?.type === "with" || timing?.type === "withPrevious") {
          const ref = timing.type === "with" ? eventMap.get(timing.eventId) : previous;
          start += (starts.get(ref?.id) || 0) + Number(timing.offsetMs || 0);
        } else if (timing?.type === "after" || timing?.type === "afterPrevious") {
          const ref = timing.type === "after" ? eventMap.get(timing.eventId) : previous;
          start += (starts.get(ref?.id) || 0) + eventDuration(ref || {}) + Number(timing.offsetMs || 0);
        } else if (timing?.type === "before") {
          const ref = eventMap.get(timing.eventId);
          start += (starts.get(ref?.id) || 0) - eventDuration(event) - Number(timing.offsetMs || 0);
        } else if (timing?.type === "trigger" || timing?.type === "onClick") {
          start += Number(timing.offsetMs || 0);
        }
        for (const dep of event.dependencies || []) {
          const ref = eventMap.get(dep.eventId);
          if (!ref) continue;
          const refStart = starts.get(ref.id) || 0;
          if (dep.relation === "with") start = Math.max(start, refStart + Number(dep.offsetMs || 0));
          else if (dep.relation === "before") start = Math.max(start, refStart - eventDuration(event) - Number(dep.offsetMs || 0));
          else start = Math.max(start, refStart + eventDuration(ref) + Number(dep.offsetMs || 0));
        }
        for (const edge of graphEdges.filter((edge) => edge.to === event.id)) {
          const ref = eventMap.get(edge.from);
          if (!ref) continue;
          const refStart = starts.get(ref.id) || 0;
          if (edge.relation === "with") start = Math.max(start, refStart + Number(edge.offsetMs || 0));
          else if (edge.relation === "before") start = Math.max(start, refStart - eventDuration(event) - Number(edge.offsetMs || 0));
          else if (edge.relation === "after") start = Math.max(start, refStart + eventDuration(ref) + Number(edge.offsetMs || 0));
        }
        starts.set(event.id, Math.max(0, start));
      });
    }
    return starts;
  };
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
        return interpolate(prev.value, next.value, easeProgress(next.easing, local), track.interpolation);
      }
    }
    return frames[frames.length - 1].value;
  };
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
    const pairs = [];
    for (const toObject of toObjects) {
      const fromObject = fromObjects.find((candidate) =>
        (toObject.morphKey && candidate.morphKey === toObject.morphKey) ||
        candidate.id === toObject.id ||
        (toObject.name && candidate.name === toObject.name && candidate.type === toObject.type)
      );
      if (fromObject) pairs.push({ fromObjectId: fromObject.id, toObjectId: toObject.id, morphKey: toObject.morphKey || fromObject.morphKey });
    }
    return pairs;
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
        setProperty(current, event.property, interpolate(from, event.to, progress, event.interpolation));
        states.set(event.targetId, current);
      } else if (event.kind === "keyframes") {
        const start = starts.get(event.id) || 0;
        const progress = eventProgress(event, start, timeMs);
        if (progress === undefined) continue;
        const current = merge(states.get(event.targetId), {});
        for (const track of event.tracks || []) setProperty(current, track.property, keyframeValue(track, progress));
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
  const applyObjectState = (el, object, statePatch) => {
    if (!el || !object || !statePatch) return;
    const b = statePatch.bounds || { x: 0, y: 0, width: 100, height: 100 };
    el.style.left = Number(b.x || 0) + "px";
    el.style.top = Number(b.y || 0) + "px";
    el.style.width = Number(b.width || 0) + "px";
    el.style.height = Number(b.height || 0) + "px";
    el.style.opacity = String(statePatch.opacity ?? 1);
    el.style.visibility = statePatch.visible === false ? "hidden" : "visible";
    el.style.transform = transformCss(statePatch.transform);
    if (object.type === "text" || (object.type === "shape" && object.text)) {
      el.textContent = textOf(object, statePatch);
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
    for (const object of flattenObjects(slide?.objects || [])) {
      applyObjectState(layer.querySelector('[data-object-id="' + CSS.escape(object.id) + '"]'), object, states.get(object.id));
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
  const applyMorphTransition = (previousLayer, currentLayer, previousSlide, currentSlide, transition, progress) => {
    const options = transition?.morph || {};
    const pairs = inferMorphPairs(previousSlide, currentSlide, options);
    const properties = options.properties || ["bounds", "transform", "opacity", "fill", "stroke"];
    const previousStates = computeSlideObjectStates(previousSlide, slideDuration(previousSlide));
    const currentStates = computeSlideObjectStates(currentSlide, 0);
    for (const pair of pairs) {
      const toObject = objectById(currentSlide, pair.toObjectId);
      if (!toObject) continue;
      const fromState = previousStates.get(pair.fromObjectId) || baseState(objectById(previousSlide, pair.fromObjectId) || {});
      const toState = currentStates.get(pair.toObjectId) || baseState(toObject);
      const blended = interpolateState(fromState, toState, progress, properties);
      applyObjectState(currentLayer.querySelector('[data-object-id="' + CSS.escape(pair.toObjectId) + '"]'), toObject, blended);
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
    return resolution;
  };
  const seekSlide = (slideIndex, timeMs = 0) => {
    const span = timeline.slides[Math.max(0, Math.min(timeline.slides.length - 1, Number(slideIndex) || 0))];
    return seekGlobal((span?.contentStartMs || 0) + Number(timeMs || 0));
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
    seekGlobal(performance.now() - state.startedAt);
    if (state.globalTimeMs >= timeline.durationMs) {
      pause();
      return;
    }
    state.raf = requestAnimationFrame(tick);
  };
  const play = () => {
    if (state.globalTimeMs >= timeline.durationMs) seekGlobal(0);
    state.playing = true;
    state.startedAt = performance.now() - state.globalTimeMs;
    if (playButton) playButton.textContent = "Pause";
    tick();
  };
  const previousSlide = () => seekSlide(Math.max(0, state.slideIndex - 1), 0);
  const nextSlide = () => seekSlide(Math.min(timeline.slides.length - 1, state.slideIndex + 1), 0);
  playButton?.addEventListener("click", () => state.playing ? pause() : play());
  prevButton?.addEventListener("click", previousSlide);
  nextButton?.addEventListener("click", nextSlide);
  stepButton?.addEventListener("click", () => seekGlobal(state.globalTimeMs + 1000 / 30));
  scrub?.addEventListener("input", () => seekGlobal(Number(scrub.value)));
  window.addEventListener("resize", resize);
  timeline = createTimeline();
  window.__keyMorphRuntime = {
    render: () => seekGlobal(state.globalTimeMs),
    seek: seekGlobal,
    seekGlobal,
    seekSlide,
    play,
    pause,
    nextSlide,
    previousSlide,
    resolveTime,
    createTimeline,
    state,
    deck
  };
  seekSlide(state.slideIndex, 0);
})();`;
}

function renderObjectMarkup(object: IRObject): string {
  const style = objectStyle(object);
  const common = `class="km-object km-${object.type}" data-object-id="${escapeHtml(object.id)}" style="${style}"`;

  if (object.type === "text") {
    return `<div ${common}>${escapeHtml(textContent(object))}</div>`;
  }
  if (object.type === "image") {
    return `<img ${common} src="${escapeHtml(object.source.dataUri ?? object.source.uri ?? "")}" alt="${escapeHtml(object.altText ?? object.name ?? "")}">`;
  }
  if (object.type === "media") {
    const tag = object.mediaType === "audio" ? "audio" : "video";
    return `<${tag} ${common} src="${escapeHtml(object.source.dataUri ?? object.source.uri ?? "")}" muted playsinline></${tag}>`;
  }
  if (object.type === "group") {
    return `<div ${common}>${object.children.map(renderObjectMarkup).join("")}</div>`;
  }
  if (object.type === "shape" && object.text) {
    return `<div ${common}>${escapeHtml(textContent(object as ShapeObject & TextObject))}</div>`;
  }
  return `<div ${common}></div>`;
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
    `transform:${transformToCss(transform)}`
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
      transform: object.transform ?? {},
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
  return {
    ...base,
    ...patch,
    bounds: patch.bounds ? { ...base.bounds, ...patch.bounds } : base.bounds,
    transform: patch.transform ? { ...base.transform, ...patch.transform } : base.transform,
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

function transformToCss(transform: Transform2D | undefined): string {
  return `translate(${transform?.translateX ?? 0}px,${transform?.translateY ?? 0}px) scale(${transform?.scaleX ?? 1},${transform?.scaleY ?? 1}) rotate(${transform?.rotateDeg ?? 0}deg) skew(${transform?.skewXDeg ?? 0}deg,${transform?.skewYDeg ?? 0}deg)`;
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
      easeInOutCubic: "cubic-bezier(.65,0,.35,1)"
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
  if (!track.property.startsWith("transform.")) return undefined;
  const key = track.property.slice("transform.".length) as keyof Transform2D;
  return { [key]: Number(value) };
}

function propertyValueToCss(property: string, value: unknown): Record<string, string> {
  if (property === "opacity") return { opacity: String(value) };
  if (property === "visible") return { visibility: value ? "visible" : "hidden" };
  if (property === "bounds.x") return { left: `${value}px` };
  if (property === "bounds.y") return { top: `${value}px` };
  if (property === "bounds.width") return { width: `${value}px` };
  if (property === "bounds.height") return { height: `${value}px` };
  if (property === "style.fill") return { background: fillToCss(value as Fill) ?? String(value) };
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
