import type {
  DeckIR,
  Easing,
  Fill,
  IRObject,
  KeyframeAnimationEvent,
  ShapeObject,
  Slide,
  TextObject
} from "../ir/index.ts";

export interface HtmlRuntimeOptions {
  controls?: boolean;
  initialSlideIndex?: number;
}

export function renderHtmlDocument(deck: DeckIR, options: HtmlRuntimeOptions = {}): string {
  const payload = JSON.stringify(deck).replaceAll("<", "\\u003c");
  const controls = options.controls ?? true;
  const initialSlideIndex = options.initialSlideIndex ?? 0;

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(deck.metadata?.title ?? deck.deck.title ?? "KeyMorph Runtime")}</title>
  <style>
    :root { color-scheme: light; font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
    body { margin: 0; min-height: 100vh; background: #202124; color: #f8fafc; display: grid; grid-template-rows: 1fr auto; }
    #viewport { min-height: 0; display: grid; place-items: center; padding: 24px; }
    #stage { position: relative; overflow: hidden; background: #fff; box-shadow: 0 18px 55px rgba(0,0,0,.36); transform-origin: center center; }
    .km-object { position: absolute; box-sizing: border-box; overflow: hidden; transform-origin: center center; }
    .km-text { white-space: pre-wrap; display: flex; align-items: flex-start; }
    .km-image { object-fit: contain; user-select: none; -webkit-user-drag: none; }
    .km-controls { display: ${controls ? "grid" : "none"}; grid-template-columns: auto auto 1fr auto auto; gap: 10px; align-items: center; padding: 12px 16px; background: #111827; border-top: 1px solid rgba(255,255,255,.12); }
    button { border: 1px solid rgba(255,255,255,.18); color: #f8fafc; background: #1f2937; border-radius: 6px; padding: 8px 11px; font: inherit; cursor: pointer; }
    button:hover { background: #374151; }
    input[type="range"] { width: 100%; }
    #status { font-variant-numeric: tabular-nums; color: #cbd5e1; min-width: 170px; text-align: right; }
  </style>
</head>
<body>
  <main id="viewport"><div id="stage" aria-label="KeyMorph slide stage"></div></main>
  <div class="km-controls">
    <button id="prev" type="button">Prev</button>
    <button id="play" type="button">Play</button>
    <input id="scrub" type="range" min="0" max="1" step="1" value="0" aria-label="Timeline">
    <button id="step" type="button">Step</button>
    <span id="status"></span>
  </div>
  <script>
    window.__KEYMORPH_DECK__ = ${payload};
    window.__KEYMORPH_INITIAL_SLIDE__ = ${JSON.stringify(initialSlideIndex)};
  </script>
  <script>${runtimeScript()}</script>
</body>
</html>`;
}

export function renderSlideMarkup(slide: Slide): string {
  return slide.objects.map((object) => renderObjectMarkup(object)).join("");
}

function runtimeScript(): string {
  return `
(() => {
  const deck = window.__KEYMORPH_DECK__;
  const stage = document.getElementById("stage");
  const viewport = document.getElementById("viewport");
  const playButton = document.getElementById("play");
  const prevButton = document.getElementById("prev");
  const stepButton = document.getElementById("step");
  const scrub = document.getElementById("scrub");
  const status = document.getElementById("status");
  const state = { slideIndex: window.__KEYMORPH_INITIAL_SLIDE__ || 0, timeMs: 0, playing: false, raf: 0, startedAt: 0 };

  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
  const slideDuration = (slide) => Math.max(1, Number(slide?.timeline?.durationMs || 2500));
  const easing = (value) => ({ easeOutCubic: "cubic-bezier(.22,1,.36,1)", easeInOutCubic: "cubic-bezier(.65,0,.35,1)", easeInCubic: "cubic-bezier(.32,0,.67,0)", easeOut: "ease-out", easeIn: "ease-in", easeInOut: "ease-in-out" }[value] || "linear");
  const fill = (input) => input?.type === "solid" ? input.color : typeof input === "string" ? input : "#fff";
  const text = (object) => object.text?.plainText || (object.text?.runs || []).map((run) => run.text).join("") || "";
  const textStyle = (object) => {
    const run = object.text?.runs?.[0];
    const style = run?.style || object.style?.textStyle || {};
    return [
      "font-family:" + (style.fontFamily || "Inter, Arial, sans-serif"),
      "font-size:" + Number(style.fontSize || 32) + "px",
      "font-weight:" + (style.fontWeight || 400),
      "color:" + (style.color || "#111827"),
      "line-height:" + (style.lineHeight || 1.15)
    ].join(";");
  };
  const boxStyle = (object) => {
    const b = object.bounds || { x: 0, y: 0, width: 100, height: 100 };
    const t = object.transform || {};
    return [
      "left:" + b.x + "px",
      "top:" + b.y + "px",
      "width:" + b.width + "px",
      "height:" + b.height + "px",
      "opacity:" + (object.opacity ?? 1),
      "transform:translate(" + (t.translateX || 0) + "px," + (t.translateY || 0) + "px) scale(" + (t.scaleX ?? 1) + "," + (t.scaleY ?? 1) + ") rotate(" + (t.rotateDeg || 0) + "deg)"
    ].join(";");
  };
  const objectHtml = (object) => {
    const common = 'class="km-object km-' + esc(object.type) + '" data-object-id="' + esc(object.id) + '" style="' + boxStyle(object) + ';';
    if (object.type === "text") return '<div ' + common + textStyle(object) + '">' + esc(text(object)) + '</div>';
    if (object.type === "image") return '<img ' + common + '" src="' + esc(object.source?.dataUri || object.source?.uri || "") + '" alt="' + esc(object.altText || object.name || "") + '">';
    if (object.type === "group") return '<div ' + common + 'overflow:visible">' + (object.children || []).map(objectHtml).join("") + '</div>';
    const style = object.style || {};
    const bg = fill(style.fill);
    const stroke = style.stroke || {};
    return '<div ' + common + 'background:' + bg + ';border:' + Number(stroke.width || 0) + 'px solid ' + (stroke.color || "transparent") + ';border-radius:' + (object.shape === "ellipse" ? "999px" : "0") + '"></div>';
  };
  const cssFrame = (property, value) => {
    if (property === "opacity") return { opacity: String(value) };
    if (property === "bounds.x") return { left: value + "px" };
    if (property === "bounds.y") return { top: value + "px" };
    if (property === "bounds.width") return { width: value + "px" };
    if (property === "bounds.height") return { height: value + "px" };
    if (property === "transform.scaleX" || property === "transform.scaleY") return {};
    return {};
  };
  const applyTimeline = (slide) => {
    for (const event of slide.timeline?.events || []) {
      if (event.kind !== "keyframes") continue;
      const target = stage.querySelector('[data-object-id="' + CSS.escape(event.targetId) + '"]');
      if (!target) continue;
      const framesByOffset = new Map();
      for (const track of event.tracks || []) {
        for (const frame of track.keyframes || []) {
          const current = framesByOffset.get(frame.offset) || { offset: frame.offset };
          Object.assign(current, cssFrame(track.property, frame.value));
          framesByOffset.set(frame.offset, current);
        }
      }
      const frames = Array.from(framesByOffset.values()).sort((a, b) => a.offset - b.offset);
      const animation = target.animate(frames, { duration: event.durationMs || 0, delay: event.start?.atMs || 0, easing: easing(event.easing), fill: event.fill || "both" });
      animation.pause();
      animation.currentTime = 0;
      target.__kmAnimation = target.__kmAnimation || [];
      target.__kmAnimation.push(animation);
    }
  };
  const render = () => {
    const slide = deck.deck.slides[state.slideIndex];
    stage.style.width = deck.deck.size.width + "px";
    stage.style.height = deck.deck.size.height + "px";
    stage.style.background = fill(slide.background);
    stage.innerHTML = (slide.objects || []).map(objectHtml).join("");
    resize();
    scrub.max = String(slideDuration(slide));
    scrub.value = String(state.timeMs);
    applyTimeline(slide);
    seek(state.timeMs);
  };
  const resize = () => {
    const scale = Math.min((viewport.clientWidth - 8) / deck.deck.size.width, (viewport.clientHeight - 8) / deck.deck.size.height, 1);
    stage.style.transform = "scale(" + Math.max(0.1, scale) + ")";
  };
  const seek = (timeMs) => {
    const slide = deck.deck.slides[state.slideIndex];
    state.timeMs = Math.max(0, Math.min(slideDuration(slide), Number(timeMs) || 0));
    for (const el of stage.querySelectorAll("[data-object-id]")) {
      for (const animation of el.__kmAnimation || []) animation.currentTime = state.timeMs;
    }
    scrub.value = String(Math.round(state.timeMs));
    status.textContent = "Slide " + (state.slideIndex + 1) + "/" + deck.deck.slides.length + " · " + Math.round(state.timeMs) + " ms";
  };
  const tick = () => {
    if (!state.playing) return;
    seek(performance.now() - state.startedAt);
    if (state.timeMs >= slideDuration(deck.deck.slides[state.slideIndex])) {
      state.playing = false;
      playButton.textContent = "Play";
      return;
    }
    state.raf = requestAnimationFrame(tick);
  };
  playButton?.addEventListener("click", () => {
    state.playing = !state.playing;
    playButton.textContent = state.playing ? "Pause" : "Play";
    if (state.playing) {
      state.startedAt = performance.now() - state.timeMs;
      tick();
    } else {
      cancelAnimationFrame(state.raf);
    }
  });
  prevButton?.addEventListener("click", () => {
    state.slideIndex = Math.max(0, state.slideIndex - 1);
    state.timeMs = 0;
    render();
  });
  stepButton?.addEventListener("click", () => seek(state.timeMs + 1000 / 30));
  scrub?.addEventListener("input", () => seek(Number(scrub.value)));
  window.addEventListener("resize", resize);
  window.__keyMorphRuntime = { render, seek, state, deck };
  render();
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
  if (object.type === "group") {
    return `<div ${common}>${object.children.map(renderObjectMarkup).join("")}</div>`;
  }
  return `<div ${common}></div>`;
}

function objectStyle(object: IRObject): string {
  const bounds = object.bounds ?? { x: 0, y: 0, width: 100, height: 100 };
  const transform = object.transform ?? {};
  const style = [
    `left:${bounds.x}px`,
    `top:${bounds.y}px`,
    `width:${bounds.width}px`,
    `height:${bounds.height}px`,
    `opacity:${object.opacity ?? 1}`,
    `transform:translate(${transform.translateX ?? 0}px,${transform.translateY ?? 0}px) scale(${transform.scaleX ?? 1},${transform.scaleY ?? 1}) rotate(${transform.rotateDeg ?? 0}deg)`
  ];

  if (object.type === "text") {
    const textStyle = firstTextStyle(object);
    style.push(`font-family:${textStyle.fontFamily ?? "Inter, Arial, sans-serif"}`);
    style.push(`font-size:${textStyle.fontSize ?? 32}px`);
    style.push(`font-weight:${textStyle.fontWeight ?? 400}`);
    style.push(`color:${colorToCss(textStyle.color) ?? "#111827"}`);
    style.push("white-space:pre-wrap");
  }

  if (object.type === "shape") {
    style.push(`background:${fillToCss(object.style?.fill) ?? "#e2e8f0"}`);
    const stroke = object.style?.stroke;
    if (stroke) {
      style.push(`border:${stroke.width ?? 1}px solid ${colorToCss(stroke.color) ?? "transparent"}`);
    }
    if ((object as ShapeObject).shape === "ellipse") style.push("border-radius:999px");
    if ((object as ShapeObject).shape === "roundRect") style.push("border-radius:8px");
  }

  return style.join(";");
}

function firstTextStyle(object: TextObject) {
  return object.text.runs?.[0]?.style ?? object.style?.textStyle ?? {};
}

function textContent(object: TextObject): string {
  return object.text.plainText ?? object.text.runs?.map((run) => run.text).join("") ?? "";
}

function fillToCss(fill: Fill | undefined): string | undefined {
  if (!fill) return undefined;
  if (fill.type === "solid") return colorToCss(fill.color);
  if (fill.type === "none") return "transparent";
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
  const frames = new Map<number, Record<string, string | number>>();
  for (const track of event.tracks) {
    for (const keyframe of track.keyframes) {
      const frame = frames.get(keyframe.offset) ?? { offset: keyframe.offset };
      Object.assign(frame, propertyValueToCss(track.property, keyframe.value));
      frames.set(keyframe.offset, frame);
    }
  }
  return Array.from(frames.values()).sort((a, b) => Number(a.offset) - Number(b.offset));
}

function propertyValueToCss(property: string, value: unknown): Record<string, string> {
  if (property === "opacity") return { opacity: String(value) };
  if (property === "bounds.x") return { left: `${value}px` };
  if (property === "bounds.y") return { top: `${value}px` };
  if (property === "bounds.width") return { width: `${value}px` };
  if (property === "bounds.height") return { height: `${value}px` };
  return {};
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
