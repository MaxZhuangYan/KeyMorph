export const IR_VERSION = "keymorph.ir.v1" as const;

export type IRVersion = typeof IR_VERSION;
export type IRID = string;

export type JSONPrimitive = string | number | boolean | null;
export type JSONValue = JSONPrimitive | JSONValue[] | { [key: string]: JSONValue };
export type JSONRecord = { [key: string]: JSONValue };

export interface KeyMorphIR {
  irVersion: IRVersion;
  deck: Deck;
  metadata?: IRMetadata;
  conversion?: ConversionReport;
  extensions?: JSONRecord;
}

export type DeckIR = KeyMorphIR;

export interface IRMetadata {
  title?: string;
  author?: string;
  createdAt?: string;
  updatedAt?: string;
  locale?: string;
  sourceApplication?: string;
  tags?: string[];
  custom?: JSONRecord;
}

export interface Deck {
  id: IRID;
  title?: string;
  size: DeckSize;
  slides: Slide[];
  assets?: Asset[];
  theme?: Theme;
  metadata?: JSONRecord;
}

export interface DeckSize {
  width: number;
  height: number;
  unit: "px" | "pt" | "in" | "cm";
}

export interface Asset {
  id: IRID;
  kind: "image" | "video" | "audio" | "font" | "data" | "other";
  uri?: string;
  dataUri?: string;
  mimeType?: string;
  name?: string;
  width?: number;
  height?: number;
  durationMs?: number;
  checksum?: string;
  metadata?: JSONRecord;
}

export interface Theme {
  id?: IRID;
  name?: string;
  colors?: Record<string, Color>;
  fonts?: Record<string, FontRef>;
  styles?: Record<string, ObjectStyle>;
  metadata?: JSONRecord;
}

export interface FontRef {
  family: string;
  assetId?: IRID;
  weight?: number | string;
  style?: "normal" | "italic" | "oblique";
}

export interface Slide {
  id: IRID;
  index?: number;
  name?: string;
  background?: Fill;
  objects: IRObject[];
  states?: ObjectState[];
  timeline?: Timeline;
  transition?: SlideTransition | null;
  notes?: string;
  metadata?: JSONRecord;
}

export type IRObject =
  | ShapeObject
  | TextObject
  | ImageObject
  | MediaObject
  | PathObject
  | GroupObject
  | EmbedObject
  | PlaceholderObject;

export interface IRObjectBase {
  id: IRID;
  type: IRObjectType;
  name?: string;
  morphKey?: string;
  visible?: boolean;
  locked?: boolean;
  bounds?: Rect;
  transform?: Transform2D;
  opacity?: number;
  filter?: FilterState;
  style?: ObjectStyle;
  effects?: VisualEffect[];
  altText?: string;
  tags?: string[];
  initialState?: ObjectStateProperties;
  metadata?: JSONRecord;
}

export type IRObjectType =
  | "shape"
  | "text"
  | "image"
  | "media"
  | "path"
  | "group"
  | "embed"
  | "placeholder";

export interface ShapeObject extends IRObjectBase {
  type: "shape";
  shape: ShapeKind;
  text?: TextContent;
}

export type ShapeKind =
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "line"
  | "arrow"
  | "polygon"
  | "star"
  | "freeform"
  | "custom";

export interface TextObject extends IRObjectBase {
  type: "text";
  text: TextContent;
}

export interface ImageObject extends IRObjectBase {
  type: "image";
  source: ObjectSource;
  crop?: Crop;
}

export interface MediaObject extends IRObjectBase {
  type: "media";
  mediaType: "video" | "audio";
  source: ObjectSource;
  posterSource?: ObjectSource;
  playback?: MediaPlayback;
}

export interface PathObject extends IRObjectBase {
  type: "path";
  path: PathCommand[];
}

export interface GroupObject extends IRObjectBase {
  type: "group";
  children: IRObject[];
  layout?: GroupLayout;
  isolation?: boolean;
}

export interface EmbedObject extends IRObjectBase {
  type: "embed";
  embedType: "html" | "iframe" | "chart" | "table" | "unknown";
  source?: ObjectSource;
  data?: JSONValue;
}

export interface PlaceholderObject extends IRObjectBase {
  type: "placeholder";
  placeholderType: "title" | "body" | "image" | "media" | "footer" | "slideNumber" | "custom";
}

export interface ObjectSource {
  assetId?: IRID;
  uri?: string;
  dataUri?: string;
  metadata?: JSONRecord;
}

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface Point {
  x: number;
  y: number;
}

export interface Transform2D {
  translateX?: number;
  translateY?: number;
  scaleX?: number;
  scaleY?: number;
  rotateDeg?: number;
  skewXDeg?: number;
  skewYDeg?: number;
  origin?: Point;
}

export type Color =
  | string
  | {
      space: "rgb" | "hsl" | "hex" | "theme";
      value: string;
      alpha?: number;
    };

export type Fill =
  | { type: "none" }
  | { type: "solid"; color: Color }
  | { type: "gradient"; stops: GradientStop[]; angleDeg?: number }
  | { type: "image"; source: ObjectSource; fit?: "cover" | "contain" | "stretch" | "tile" };

export interface GradientStop {
  offset: number;
  color: Color;
}

export interface Stroke {
  color?: Color;
  width?: number;
  dash?: number[];
  lineCap?: "butt" | "round" | "square";
  lineJoin?: "miter" | "round" | "bevel";
}

export interface ObjectStyle {
  fill?: Fill;
  stroke?: Stroke;
  shadow?: Shadow;
  blendMode?: string;
  textStyle?: TextStyle;
  custom?: JSONRecord;
}

export interface Shadow {
  color?: Color;
  offsetX?: number;
  offsetY?: number;
  blur?: number;
  spread?: number;
}

export interface VisualEffect {
  type: "blur" | "brightness" | "contrast" | "saturate" | "opacity" | "custom";
  value?: number | string;
  metadata?: JSONRecord;
}

export interface TextContent {
  plainText?: string;
  runs?: TextRun[];
  paragraphs?: TextParagraph[];
}

export interface TextRun {
  text: string;
  style?: TextStyle;
}

export interface TextParagraph {
  runs: TextRun[];
  alignment?: "left" | "center" | "right" | "justify";
  bullet?: BulletStyle;
}

export interface TextStyle {
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: number | string;
  italic?: boolean;
  underline?: boolean;
  color?: Color;
  lineHeight?: number;
  letterSpacing?: number;
}

export interface BulletStyle {
  type: "bullet" | "number";
  level?: number;
  marker?: string;
}

export interface Crop {
  x: number;
  y: number;
  width: number;
  height: number;
  unit?: "ratio" | "px";
}

export interface MediaPlayback {
  autoplay?: boolean;
  loop?: boolean;
  muted?: boolean;
  startMs?: number;
  endMs?: number;
}

export type PathCommand =
  | { command: "moveTo"; x: number; y: number }
  | { command: "lineTo"; x: number; y: number }
  | { command: "curveTo"; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { command: "quadTo"; x1: number; y1: number; x: number; y: number }
  | { command: "close" };

export interface GroupLayout {
  mode: "free" | "stack" | "grid" | "locked";
  gap?: number;
  columns?: number;
  rows?: number;
}

export interface ObjectState {
  id: IRID;
  targetId: IRID;
  name?: string;
  inheritFrom?: IRID;
  properties: ObjectStateProperties;
  metadata?: JSONRecord;
}

export interface ObjectStateProperties {
  visible?: boolean;
  opacity?: number;
  bounds?: Rect;
  transform?: Transform2D;
  filter?: FilterState;
  style?: ObjectStyle;
  text?: TextContent;
  crop?: Crop;
  media?: MediaPlayback;
  custom?: JSONRecord;
}

export interface FilterState {
  blurPx?: number;
  brightness?: number;
  contrast?: number;
  saturate?: number;
}

export interface Timeline {
  id?: IRID;
  durationMs?: number;
  defaultEasing?: Easing;
  triggers?: TimelineTrigger[];
  events: AnimationEvent[];
  dependencyGraph?: TimingDependencyGraph;
  metadata?: JSONRecord;
}

export type TimelineTrigger =
  | { id: IRID; type: "onEnter"; offsetMs?: number }
  | { id: IRID; type: "onClick"; targetId?: IRID; clickIndex?: number }
  | { id: IRID; type: "afterPrevious"; offsetMs?: number }
  | { id: IRID; type: "custom"; name: string; metadata?: JSONRecord };

export type AnimationEvent =
  | SetStateEvent
  | PropertyAnimationEvent
  | KeyframeAnimationEvent
  | MorphAnimationEvent
  | TransitionAnimationEvent
  | MediaAnimationEvent
  | VisibilityAnimationEvent;

export interface AnimationEventBase {
  id: IRID;
  kind: AnimationEventKind;
  label?: string;
  start?: TimingStart;
  durationMs?: number;
  delayMs?: number;
  easing?: Easing;
  fill?: "none" | "forwards" | "backwards" | "both";
  dependencies?: TimingDependencyRef[];
  metadata?: JSONRecord;
}

export type AnimationEventKind =
  | "setState"
  | "property"
  | "keyframes"
  | "morph"
  | "transition"
  | "media"
  | "visibility";

export interface SetStateEvent extends AnimationEventBase {
  kind: "setState";
  targetId: IRID;
  stateId: IRID;
}

export interface PropertyAnimationEvent extends AnimationEventBase {
  kind: "property";
  targetId: IRID;
  property: AnimatableProperty | string;
  from?: JSONValue;
  to: JSONValue;
  interpolation?: InterpolationMode;
}

export interface KeyframeAnimationEvent extends AnimationEventBase {
  kind: "keyframes";
  targetId: IRID;
  tracks: KeyframeTrack[];
}

export interface KeyframeTrack {
  property: AnimatableProperty | string;
  interpolation?: InterpolationMode;
  keyframes: Keyframe[];
}

export interface Keyframe {
  offset: number;
  value: JSONValue;
  easing?: Easing;
}

export type AnimatableProperty =
  | "bounds.x"
  | "bounds.y"
  | "bounds.width"
  | "bounds.height"
  | "transform.translateX"
  | "transform.translateY"
  | "transform.scaleX"
  | "transform.scaleY"
  | "transform.rotateDeg"
  | "opacity"
  | "visible"
  | "filter.blurPx"
  | "style.fill"
  | "style.stroke"
  | "text"
  | "crop";

export type InterpolationMode = "linear" | "discrete" | "color" | "path" | "number" | "matrix";

export interface MorphAnimationEvent extends AnimationEventBase {
  kind: "morph";
  strategy: MorphStrategy;
  from: MorphEndpoint;
  to: MorphEndpoint;
  matching?: MorphMatching;
  pairs?: MorphObjectPair[];
  properties?: MorphProperty[];
}

export type MorphStrategy = "morph" | "magicMove";

export interface MorphEndpoint {
  slideId?: IRID;
  objectId?: IRID;
  stateId?: IRID;
  snapshot?: ObjectStateProperties;
}

export interface MorphMatching {
  matchBy: MorphMatchKey[];
  fallback?: "none" | "geometry" | "name" | "type";
  tolerance?: number;
}

export type MorphMatchKey = "morphKey" | "objectId" | "name" | "type" | "geometry" | "imageHash";

export interface MorphObjectPair {
  fromObjectId: IRID;
  toObjectId: IRID;
  morphKey?: string;
}

export type MorphProperty =
  | "bounds"
  | "transform"
  | "opacity"
  | "filter"
  | "fill"
  | "stroke"
  | "text"
  | "path"
  | "crop"
  | "children"
  | "all";

export interface TransitionAnimationEvent extends AnimationEventBase {
  kind: "transition";
  transition: SlideTransition;
  fromSlideId?: IRID;
  toSlideId?: IRID;
}

export interface MediaAnimationEvent extends AnimationEventBase {
  kind: "media";
  targetId: IRID;
  action: "play" | "pause" | "seek" | "stop" | "mute" | "unmute";
  seekMs?: number;
}

export interface VisibilityAnimationEvent extends AnimationEventBase {
  kind: "visibility";
  targetId: IRID;
  visible: boolean;
}

export type TimingStart =
  | { type: "absolute"; atMs: number }
  | { type: "with"; eventId: IRID; offsetMs?: number }
  | { type: "after"; eventId: IRID; offsetMs?: number }
  | { type: "before"; eventId: IRID; offsetMs?: number }
  | { type: "onClick"; targetId?: IRID; clickIndex?: number }
  | { type: "withPrevious"; offsetMs?: number }
  | { type: "afterPrevious"; offsetMs?: number }
  | { type: "trigger"; triggerId: IRID; offsetMs?: number };

export interface TimingDependencyRef {
  eventId: IRID;
  relation?: "after" | "with" | "before";
  offsetMs?: number;
}

export interface TimingDependencyGraph {
  nodes?: TimingNode[];
  edges: TimingDependencyEdge[];
}

export interface TimingNode {
  id: IRID;
  eventId?: IRID;
  label?: string;
  kind?: "event" | "trigger" | "state" | "custom";
}

export interface TimingDependencyEdge {
  id?: IRID;
  from: IRID;
  to: IRID;
  relation: "after" | "with" | "before" | "blocks" | "triggers";
  offsetMs?: number;
}

export type Easing = EasingName | CubicBezierEasing | SpringEasing | StepsEasing | CustomEasing;

export type EasingName =
  | "linear"
  | "ease"
  | "easeIn"
  | "easeOut"
  | "easeInOut"
  | "easeInCubic"
  | "easeOutCubic"
  | "easeInOutCubic"
  | "backIn"
  | "backOut"
  | "backInOut"
  | "bounceIn"
  | "bounceOut"
  | "bounceInOut"
  | "elasticIn"
  | "elasticOut"
  | "elasticInOut";

export interface CubicBezierEasing {
  type: "cubicBezier";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export interface SpringEasing {
  type: "spring";
  mass?: number;
  stiffness?: number;
  damping?: number;
  velocity?: number;
}

export interface StepsEasing {
  type: "steps";
  count: number;
  position?: "start" | "end";
}

export interface CustomEasing {
  type: "custom";
  name: string;
  parameters?: JSONRecord;
}

export interface SlideTransition {
  id?: IRID;
  type: TransitionType;
  durationMs?: number;
  easing?: Easing;
  direction?: "left" | "right" | "up" | "down" | "in" | "out";
  trigger?: "auto" | "click" | "timeline";
  morph?: MorphTransitionOptions;
  fromSlideId?: IRID;
  toSlideId?: IRID;
  metadata?: JSONRecord;
}

export type TransitionType =
  | "none"
  | "cut"
  | "fade"
  | "push"
  | "wipe"
  | "zoom"
  | "dissolve"
  | "morph"
  | "magicMove";

export interface MorphTransitionOptions {
  strategy: MorphStrategy;
  matching?: MorphMatching;
  pairs?: MorphObjectPair[];
  properties?: MorphProperty[];
}

export interface ConversionReport {
  id?: IRID;
  source?: ConversionSource;
  status: "success" | "partial" | "failed";
  generatedAt?: string;
  tool?: string;
  messages: ConversionMessage[];
  mappings?: ConversionMapping[];
  unsupportedFeatures?: UnsupportedFeature[];
  degradedFeatures?: DegradedFeature[];
  uncertainMappings?: UncertainMapping[];
  statistics?: ConversionStatistics;
  metadata?: JSONRecord;
}

export interface ConversionSource {
  kind: "pptx" | "keynote" | "googleSlides" | "html" | "json" | "unknown";
  uri?: string;
  name?: string;
  checksum?: string;
  application?: string;
  applicationVersion?: string;
}

export interface ConversionMessage {
  severity: "info" | "warning" | "error";
  code: string;
  message: string;
  path?: string;
  sourceLocation?: SourceLocation;
  metadata?: JSONRecord;
}

export interface SourceLocation {
  slideIndex?: number;
  objectId?: string;
  page?: number;
  line?: number;
  column?: number;
}

export interface ConversionMapping {
  sourceId?: string;
  sourcePath?: string;
  irPath: string;
  kind: "deck" | "slide" | "object" | "state" | "event" | "asset" | "style";
  confidence?: number;
  notes?: string;
}

export interface UnsupportedFeature {
  code: string;
  severity: "info" | "warning" | "error";
  area?: ConversionFeatureArea;
  description: string;
  sourcePath?: string;
  fallback?: string;
}

export interface DegradedFeature {
  code: string;
  severity: "info" | "warning" | "error";
  area?: ConversionFeatureArea;
  description: string;
  sourcePath?: string;
  fallback: string;
}

export interface UncertainMapping {
  code: string;
  severity: "info" | "warning" | "error";
  description: string;
  sourcePath?: string;
  confidence?: number;
}

export type ConversionFeatureArea =
  | "layout"
  | "text"
  | "image"
  | "shape"
  | "animation"
  | "transition"
  | "media"
  | "asset"
  | "unknown";

export interface ConversionStatistics {
  slideCount?: number;
  objectCount?: number;
  animationCount?: number;
  assetCount?: number;
  unsupportedFeatureCount?: number;
  degradedFeatureCount?: number;
  uncertainMappingCount?: number;
  durationMs?: number;
}
