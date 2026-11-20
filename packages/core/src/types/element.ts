import type { Box, Insets, Point, Size } from "./geometry.js";
import type { DynamicDateUpdateFormat } from "./field.js";
import type {
  NormalizedPaint,
  NormalizedShapeStyle,
  NormalizedStrokeStyle,
  NormalizedTableCellStyle,
  NormalizedTextFrameStyle,
  NormalizedTextParagraphStyle,
  NormalizedTextRunStyle,
  NormalizedTransform,
  ColorValue,
  FontFamilyValue,
  PaintInput,
  ShapeStyleInput,
  StrokeStyleInput,
  TableCellStyleInput,
  TableStyleInput,
  TextFrameStyleInput,
  TextParagraphStyleInput,
  TextRunStyleInput,
  TextStylePresetResetProperty,
  TransformInput,
} from "./style.js";

export type ElementAction =
  | { type: "url"; url: string; tooltip?: string }
  | { type: "slide"; slideId: string; tooltip?: string };

export interface ElementAccessibilityInput {
  description?: string;
  decorative?: boolean;
}

export interface NormalizedElementAccessibility {
  description?: string;
  decorative: boolean;
}

export interface ElementBaseInput {
  id?: string;
  name?: string;
  box?: Box;
  transform?: TransformInput;
  opacity?: number;
  hidden?: boolean;
  accessibility?: ElementAccessibilityInput;
  action?: ElementAction;
  placeholderKey?: string;
}

export interface ElementBase extends Omit<ElementBaseInput, "id"> {
  id: string;
}

export interface NormalizedElementBase {
  id: string;
  name: string;
  box: Box;
  transform: NormalizedTransform;
  opacity: number;
  hidden: boolean;
  accessibility: NormalizedElementAccessibility;
  action?: ElementAction;
  placeholderKey?: string;
}

export interface TextRunInput {
  text: string;
  /** Optional dynamic field identity; text is its caller-supplied display cache. */
  field?: { id: string; kind: "slide-number" } | { id: string; kind: "dynamic-date"; updateFormat: DynamicDateUpdateFormat };
  style?: TextRunStyleInput;
  action?: ElementAction;
}

export interface TextParagraphInput {
  runs: TextRunInput[];
  style?: TextParagraphStyleInput;
}

export type TextContentInput = string | TextParagraphInput[];

export interface NormalizedTextRun {
  text: string;
  field?: { id: string; kind: "slide-number" } | { id: string; kind: "dynamic-date"; updateFormat: DynamicDateUpdateFormat };
  style: NormalizedTextRunStyle;
  action?: ElementAction;
}

export interface NormalizedTextParagraph {
  runs: NormalizedTextRun[];
  style: NormalizedTextParagraphStyle;
}

export type TextBoxInput = Omit<Box, "height"> & { height?: number };

export interface TextElementInput extends Omit<ElementBaseInput, "box"> {
  type: "text";
  content: TextContentInput;
  textStylePreset?: string;
  /** Preset properties that inherit from the placeholder instead. */
  textStylePresetReset?: readonly TextStylePresetResetProperty[];
  box?: TextBoxInput;
  frame?: TextFrameStyleInput;
}

export interface TextElement extends ElementBase, Omit<TextElementInput, keyof ElementBaseInput> {}

export interface NormalizedTextElement extends NormalizedElementBase {
  type: "text";
  content: NormalizedTextParagraph[];
  plainText: string;
  frame: NormalizedTextFrameStyle;
}

export interface ShapeTextInput {
  content: TextContentInput;
  textStylePreset?: string;
  /** Preset properties that inherit from the placeholder instead. */
  textStylePresetReset?: readonly TextStylePresetResetProperty[];
  frame?: TextFrameStyleInput;
}

export interface NormalizedShapeText {
  content: NormalizedTextParagraph[];
  plainText: string;
  frame: NormalizedTextFrameStyle;
}

export type ImageFit = "stretch" | "contain" | "cover" | "crop";

export interface ImageCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface ImageElementInput extends ElementBaseInput {
  type: "image";
  assetId: string;
  fit?: ImageFit;
  /** Used with fit crop; other fit modes ignore or derive crop fractions. */
  crop?: Partial<ImageCrop>;
}

export interface ImageElement extends ElementBase, Omit<ImageElementInput, keyof ElementBaseInput> {}

export interface NormalizedImageElement extends NormalizedElementBase {
  type: "image";
  assetId: string;
  fit: ImageFit;
  crop: ImageCrop;
}

/** Audio and its optional static poster are independently owned assets. */
export interface AudioElementInput extends ElementBaseInput {
  type: "audio";
  audioAssetId: string;
  posterAssetId?: string;
  /** Defaults to click-sequence; object-click targets this audio object. */
  playbackStart?: "click-sequence" | "object-click";
  /** PowerPoint-style playback volume; defaults to high. */
  playbackVolume?: "low" | "medium" | "high" | "mute";
}
export interface AudioElement extends ElementBase, Omit<AudioElementInput, keyof ElementBaseInput> {}
export interface NormalizedAudioElement extends NormalizedElementBase {
  type: "audio";
  audioAssetId: string;
  posterAssetId?: string;
  playbackStart?: "click-sequence" | "object-click";
  /** PowerPoint-style playback volume; defaults to high. */
  playbackVolume?: "low" | "medium" | "high" | "mute";
}

/** Video bytes and the optional static poster are independently owned assets. */
export interface VideoElementInput extends ElementBaseInput {
  type: "video";
  videoAssetId: string;
  posterAssetId?: string;
}
export interface VideoElement extends ElementBase, Omit<VideoElementInput, keyof ElementBaseInput> {}
export interface NormalizedVideoElement extends NormalizedElementBase {
  type: "video";
  videoAssetId: string;
  posterAssetId?: string;
}

export type ShapeKind = "rect" | "roundRect" | "ellipse" | "triangle" | "diamond" | "arrow" | "chevron" | "parallelogram";

export interface ShapeElementInput extends ElementBaseInput {
  type: "shape";
  shape: ShapeKind;
  /** Parallelogram horizontal slant as a fraction of the shorter frame side, from 0 to 1. */
  slant?: number;
  style?: ShapeStyleInput;
  text?: ShapeTextInput;
}

export interface ShapeElement extends ElementBase, Omit<ShapeElementInput, keyof ElementBaseInput> {}

export interface NormalizedShapeElement extends NormalizedElementBase {
  type: "shape";
  shape: ShapeKind;
  /** Absent means the default parallelogram slant of 0.25. */
  slant?: number;
  style: NormalizedShapeStyle;
  text?: NormalizedShapeText;
}

export type ConnectorAnchor = "top" | "right" | "bottom" | "left" | "center";
export type ConnectorEndpointInput = Point | { elementId: string; anchor?: ConnectorAnchor };

export interface ConnectorElementInput extends ElementBaseInput {
  type: "connector";
  start: ConnectorEndpointInput;
  end: ConnectorEndpointInput;
  route?: Point[];
  /** Custom path with round caps and joins; closed paths end at their start. */
  pathMode?: "open-round" | "closed-round";
  style?: StrokeStyleInput;
}

export interface ConnectorElement extends ElementBase, Omit<ConnectorElementInput, keyof ElementBaseInput> {}

export interface NormalizedConnectorElement extends NormalizedElementBase {
  type: "connector";
  start: ConnectorEndpointInput;
  end: ConnectorEndpointInput;
  route: Point[];
  pathMode?: "open-round" | "closed-round";
  style: NormalizedStrokeStyle;
}

export interface GroupElementInput extends ElementBaseInput {
  type: "group";
  coordinateSize: Size;
  children: PresentationElementInput[];
}

export interface GroupElement extends ElementBase {
  type: "group";
  coordinateSize: Size;
  children: readonly PresentationElement[];
}

export interface NormalizedGroupElement extends NormalizedElementBase {
  type: "group";
  coordinateSize: Size;
  children: NormalizedElement[];
}

export interface TableCellInput {
  content: TextContentInput;
  /** Explicit accessibility role for a row or column header. */
  headerScope?: "row" | "column";
  textStylePreset?: string;
  rowSpan?: number;
  colSpan?: number;
  style?: TableCellStyleInput;
}

export interface TableRowInput {
  cells: TableCellInput[];
  height?: number;
}

/** Named native table presets verified by the PPTX adapter. */
export type NativeTableStylePreset = "medium-2-accent-1" | "medium-2-accent-2";

/** Native table style region switches, separate from Core's effective cell formatting. */
export interface NativeTableStyleFlags {
  firstColumn?: boolean;
  lastColumn?: boolean;
  lastRow?: boolean;
  bandedRows?: boolean;
  bandedColumns?: boolean;
}

export interface TableElementInput extends ElementBaseInput {
  type: "table";
  /** Native first-row table style toggle; independent of per-cell header scope. */
  nativeFirstRowStyle?: boolean;
  nativeTableStyleFlags?: NativeTableStyleFlags;
  nativeTableStylePreset?: NativeTableStylePreset;
  style?: TableStyleInput;
  columns: number[];
  rows: TableRowInput[];
}

export interface TableElement extends ElementBase, Omit<TableElementInput, keyof ElementBaseInput> {}

export interface NormalizedTableCell {
  content: NormalizedTextParagraph[];
  headerScope?: "row" | "column";
  rowSpan: number;
  colSpan: number;
  style: NormalizedTableCellStyle;
}

export interface NormalizedTableRow {
  cells: NormalizedTableCell[];
  height: number;
}

export interface NormalizedTableElement extends NormalizedElementBase {
  type: "table";
  nativeFirstRowStyle?: boolean;
  nativeTableStyleFlags?: NativeTableStyleFlags;
  nativeTableStylePreset?: NativeTableStylePreset;
  /** Cells contain effective values; the authoring region rules are not native table styles. */
  styleSource?: "materialized-regions";
  columns: number[];
  rows: NormalizedTableRow[];
}

export type ChartType = "bar" | "line" | "area" | "pie" | "doughnut" | "scatter" | "bubble" | "combo";
export type ChartOrientation = "vertical" | "horizontal";
export type ChartGrouping = "clustered" | "stacked" | "percentStacked";
export type AreaChartGrouping = "standard" | "stacked" | "percentStacked";
/** Blank display strategy; span is supported for standard line and scatter charts. */
export type ChartBlankDisplay = "gap" | "zero" | "span";
export type ChartLegendPosition = "right" | "bottom";
export type ChartTickMark = "none" | "inside" | "outside" | "cross";
export type ChartMarkerShape = "diamond" | "square" | "circle" | "triangle" | "star" | "x" | "plus" | "dash";

export interface ChartTextStyleInput {
  fontFamily?: FontFamilyValue;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: ColorValue;
}

export interface NormalizedChartTextStyle {
  fontFamily: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  color: string;
}

export interface ChartLegendInput {
  visible?: boolean;
  position?: ChartLegendPosition;
  textStyle?: ChartTextStyleInput;
}

export interface NormalizedChartLegend {
  visible: boolean;
  position: ChartLegendPosition;
  textStyle: NormalizedChartTextStyle;
}

export interface CategoryAxisInput {
  visible?: boolean;
  labels?: boolean;
  line?: StrokeStyleInput;
  majorTick?: ChartTickMark;
  majorGridlines?: false | StrokeStyleInput;
  labelStyle?: ChartTextStyleInput;
}

export type ValueAxisScaleInput = "auto" | { min: number; max: number; majorUnit: number };

export interface ValueAxisInput extends CategoryAxisInput {
  scale?: ValueAxisScaleInput;
}

export interface ChartAxesInput {
  category?: CategoryAxisInput;
  value?: ValueAxisInput;
}

export interface NormalizedCategoryAxis {
  visible: boolean;
  labels: boolean;
  line: NormalizedStrokeStyle;
  majorTick: ChartTickMark;
  majorGridlines: false | NormalizedStrokeStyle;
  labelStyle: NormalizedChartTextStyle;
}

export interface NormalizedValueAxis extends NormalizedCategoryAxis {
  scale: ValueAxisScaleInput;
}

export interface NormalizedChartAxes {
  category: NormalizedCategoryAxis;
  value: NormalizedValueAxis;
}

export interface ChartStyleInput {
  chartArea?: PaintInput;
  plotArea?: PaintInput;
  textStyle?: ChartTextStyleInput;
}

export interface NormalizedChartStyle {
  chartArea: NormalizedPaint;
  plotArea: NormalizedPaint;
  textStyle: NormalizedChartTextStyle;
}

export type ChartDataValue = number | null;

export interface ChartSeriesBaseInput {
  name: string;
  /** Explicit null is a blank point; sparse array holes remain invalid. */
  values: ChartDataValue[];
  color?: ColorValue;
}

export interface ChartMarkerStyleInput {
  shape: ChartMarkerShape;
  size?: number;
  fill?: PaintInput;
  stroke?: StrokeStyleInput;
}

export type ChartMarkerInput = "auto" | false | ChartMarkerStyleInput;

export interface BarChartSeriesInput extends ChartSeriesBaseInput {}
export interface LineChartSeriesInput extends ChartSeriesBaseInput { line?: StrokeStyleInput; marker?: ChartMarkerInput }
export interface ScatterChartSeriesInput {
  name: string;
  xValues: ChartDataValue[];
  yValues: ChartDataValue[];
  color?: ColorValue;
  line?: false | StrokeStyleInput;
  marker?: ChartMarkerInput;
}
export interface BubbleChartSeriesInput {
  name: string;
  xValues: number[];
  yValues: number[];
  sizes: number[];
  color?: ColorValue;
}
export interface NormalizedBubbleChartSeries extends Omit<BubbleChartSeriesInput, "color"> { color: string }

export interface PieChartSeriesInput extends ChartSeriesBaseInput { pointColors?: ColorValue[] }

export interface NormalizedChartMarkerStyle {
  shape: ChartMarkerShape;
  size: number;
  fill: NormalizedPaint;
  stroke: NormalizedStrokeStyle;
}

export interface NormalizedChartSeriesBase {
  name: string;
  values: ChartDataValue[];
  color: string;
}

export interface NormalizedBarChartSeries extends NormalizedChartSeriesBase {}

export interface NormalizedLineChartSeries extends NormalizedChartSeriesBase {
  line: false | NormalizedStrokeStyle;
  marker: false | NormalizedChartMarkerStyle;
}

export interface NormalizedScatterChartSeries {
  name: string;
  xValues: ChartDataValue[];
  yValues: ChartDataValue[];
  color: string;
  line: false | NormalizedStrokeStyle;
  marker: false | NormalizedChartMarkerStyle;
}

export interface ScatterChartAxesInput { x?: ValueAxisInput; y?: ValueAxisInput }
export interface NormalizedScatterChartAxes { x: NormalizedValueAxis; y: NormalizedValueAxis }

export interface NormalizedPieChartSeries extends NormalizedChartSeriesBase {
  pointColors: string[];
}

interface ChartElementInputBase extends ElementBaseInput {
  type: "chart";
  /** Defaults to gap when values contain null, otherwise zero. */
  blankDisplay?: ChartBlankDisplay;
  categories: string[];
  title?: string;
  titleStyle?: ChartTextStyleInput;
  legend?: ChartLegendInput;
  axes?: ChartAxesInput;
  style?: ChartStyleInput;
}

export interface BarChartElementInput extends ChartElementInputBase {
  chartType: "bar";
  series: BarChartSeriesInput[];
  orientation?: ChartOrientation;
  grouping?: ChartGrouping;
  seriesGap?: number;
  categoryGap?: number;
}

export interface LineChartElementInput extends ChartElementInputBase {
  chartType: "line";
  series: LineChartSeriesInput[];
}

export interface AreaChartElementInput extends ChartElementInputBase {
  chartType: "area";
  series: ChartSeriesBaseInput[];
  grouping?: AreaChartGrouping;
}

export interface ScatterChartElementInput extends Omit<ChartElementInputBase, "categories" | "axes"> {
  chartType: "scatter";
  series: ScatterChartSeriesInput[];
  axes?: ScatterChartAxesInput;
}

export interface BubbleChartElementInput extends Omit<ChartElementInputBase, "categories" | "axes" | "blankDisplay"> {
  chartType: "bubble";
  series: BubbleChartSeriesInput[];
  axes?: ScatterChartAxesInput;
}

/** Clustered columns on the primary axis, followed by lines on the secondary axis. */
export type ComboChartSeriesInput = ({ kind: "bar" } & BarChartSeriesInput) | ({ kind: "line" } & LineChartSeriesInput);
export interface ComboChartAxesInput extends ChartAxesInput { secondaryValue?: ValueAxisInput }
export interface ComboChartElementInput extends Omit<ChartElementInputBase, "axes"> {
  chartType: "combo";
  series: ComboChartSeriesInput[];
  axes?: ComboChartAxesInput;
}

export interface PieChartElementInput extends Omit<ChartElementInputBase, "axes"> {
  chartType: "pie";
  series: PieChartSeriesInput[];
  axes?: never;
}

export interface DoughnutChartElementInput extends Omit<ChartElementInputBase, "axes"> {
  chartType: "doughnut";
  series: PieChartSeriesInput[];
  axes?: never;
  /** Integer percentage of the radius reserved for the central hole, from 1 to 90. */
  holeSize?: number;
  /** Clockwise angle in integer degrees from 12 o'clock, from 0 to 360. */
  firstSliceAngle?: number;
}

export type ChartElementInput = BarChartElementInput | LineChartElementInput | AreaChartElementInput | PieChartElementInput | DoughnutChartElementInput | ScatterChartElementInput | BubbleChartElementInput | ComboChartElementInput;

interface ChartElementBase extends ElementBase {
  type: "chart";
  blankDisplay?: ChartBlankDisplay;
  categories: string[];
  title?: string;
  titleStyle?: ChartTextStyleInput;
  legend?: ChartLegendInput;
  axes?: ChartAxesInput;
  style?: ChartStyleInput;
}

export interface BarChartElement extends ChartElementBase, Omit<BarChartElementInput, keyof ElementBaseInput | "type" | "categories" | "title" | "titleStyle" | "legend" | "axes" | "style"> {
  chartType: "bar";
}
export interface LineChartElement extends ChartElementBase, Omit<LineChartElementInput, keyof ElementBaseInput | "type" | "categories" | "title" | "titleStyle" | "legend" | "axes" | "style"> {
  chartType: "line";
}
export interface AreaChartElement extends ChartElementBase, Omit<AreaChartElementInput, keyof ElementBaseInput | "type" | "categories" | "title" | "titleStyle" | "legend" | "axes" | "style"> {
  chartType: "area";
}
export interface ScatterChartElement extends Omit<ChartElementBase, "categories" | "axes">, Omit<ScatterChartElementInput, keyof ElementBaseInput | "type" | "title" | "titleStyle" | "legend" | "axes" | "style"> {
  chartType: "scatter";
  axes?: ScatterChartAxesInput;
}
export interface BubbleChartElement extends Omit<ChartElementBase, "categories" | "axes" | "blankDisplay"> {
  chartType: "bubble";
  series: BubbleChartSeriesInput[];
  axes?: ScatterChartAxesInput;
}
export interface ComboChartElement extends Omit<ChartElementBase, "axes"> {
  chartType: "combo";
  series: ComboChartSeriesInput[];
  axes?: ComboChartAxesInput;
}
export interface PieChartElement extends Omit<ChartElementBase, "axes">, Omit<PieChartElementInput, keyof ElementBaseInput | "type" | "categories" | "title" | "titleStyle" | "legend" | "style"> {
  chartType: "pie";
}
export interface DoughnutChartElement extends Omit<ChartElementBase, "axes">, Omit<DoughnutChartElementInput, keyof ElementBaseInput | "type" | "categories" | "title" | "titleStyle" | "legend" | "style"> {
  chartType: "doughnut";
}
export type ChartElement = BarChartElement | LineChartElement | AreaChartElement | PieChartElement | DoughnutChartElement | ScatterChartElement | BubbleChartElement | ComboChartElement;

interface NormalizedChartElementBase<TSeries extends NormalizedChartSeriesBase> extends NormalizedElementBase {
  type: "chart";
  /** Normalization supplies this; absent legacy IR retains adaptive defaults. */
  blankDisplay?: ChartBlankDisplay;
  categories: string[];
  series: TSeries[];
  title?: string;
  titleStyle: NormalizedChartTextStyle;
  legend: NormalizedChartLegend;
  style: NormalizedChartStyle;
}

export interface NormalizedBarChartElement extends NormalizedChartElementBase<NormalizedBarChartSeries> {
  chartType: "bar";
  orientation: ChartOrientation;
  grouping: ChartGrouping;
  seriesGap: number;
  categoryGap: number;
  axes: NormalizedChartAxes;
}

export interface NormalizedLineChartElement extends NormalizedChartElementBase<NormalizedLineChartSeries> {
  chartType: "line";
  axes: NormalizedChartAxes;
}

export interface NormalizedAreaChartElement extends NormalizedChartElementBase<NormalizedChartSeriesBase> {
  chartType: "area";
  grouping: AreaChartGrouping;
  axes: NormalizedChartAxes;
}
export interface NormalizedScatterChartElement extends Omit<NormalizedChartElementBase<NormalizedChartSeriesBase>, "categories" | "series"> {
  chartType: "scatter";
  series: NormalizedScatterChartSeries[];
  axes: NormalizedScatterChartAxes;
}

export interface NormalizedBubbleChartElement extends Omit<NormalizedChartElementBase<NormalizedChartSeriesBase>, "categories" | "series" | "blankDisplay"> {
  chartType: "bubble";
  series: NormalizedBubbleChartSeries[];
  axes: NormalizedScatterChartAxes;
}

export type NormalizedComboChartSeries = ({ kind: "bar" } & NormalizedBarChartSeries) | ({ kind: "line"; line: NormalizedStrokeStyle } & Omit<NormalizedLineChartSeries, "line">);
export interface NormalizedComboChartElement extends NormalizedChartElementBase<NormalizedComboChartSeries> {
  chartType: "combo";
  axes: NormalizedChartAxes & { secondaryValue: NormalizedValueAxis };
}

export interface NormalizedPieChartElement extends NormalizedChartElementBase<NormalizedPieChartSeries> {
  chartType: "pie";
}

export interface NormalizedDoughnutChartElement extends NormalizedChartElementBase<NormalizedPieChartSeries> {
  chartType: "doughnut";
  holeSize: number;
  firstSliceAngle: number;
}

export type NormalizedChartElement = NormalizedBarChartElement | NormalizedLineChartElement | NormalizedAreaChartElement | NormalizedPieChartElement | NormalizedDoughnutChartElement | NormalizedScatterChartElement | NormalizedBubbleChartElement | NormalizedComboChartElement;

export type PlaceholderKind = "title" | "subtitle" | "body" | "image" | "table" | "footer" | "slideNumber";

export interface PlaceholderTextStyleInput {
  frame?: TextFrameStyleInput;
  paragraph?: TextParagraphStyleInput;
  run?: TextRunStyleInput;
}

export interface PlaceholderDefinitionInput {
  key: string;
  kind: PlaceholderKind;
  /** Explicit layout-local numeric identity (0…4294967295), never an array position. */
  index?: number;
  box: Box;
  textStyle?: PlaceholderTextStyleInput;
  textStylePreset?: string;
}

export interface NormalizedPlaceholderDefinition {
  key: string;
  kind: PlaceholderKind;
  /** Only present when authored; serialization-generated ordinals are not identities. */
  index?: number;
  box: Box;
  textStyle: {
    frame: NormalizedTextFrameStyle;
    paragraph: NormalizedTextParagraphStyle;
    run: NormalizedTextRunStyle;
  };
}

export type PresentationElementInput =
  | TextElementInput
  | ImageElementInput
  | AudioElementInput
  | VideoElementInput
  | ShapeElementInput
  | ConnectorElementInput
  | GroupElementInput
  | TableElementInput
  | ChartElementInput;

export type PresentationElement =
  | TextElement
  | ImageElement
  | AudioElement
  | VideoElement
  | ShapeElement
  | ConnectorElement
  | GroupElement
  | TableElement
  | ChartElement;

export type NormalizedElement =
  | NormalizedTextElement
  | NormalizedImageElement
  | NormalizedAudioElement
  | NormalizedVideoElement
  | NormalizedShapeElement
  | NormalizedConnectorElement
  | NormalizedGroupElement
  | NormalizedTableElement
  | NormalizedChartElement;

export interface PlaceholderBindingStyle {
  frame?: TextFrameStyleInput;
  paragraph?: TextParagraphStyleInput;
  run?: TextRunStyleInput;
  fill?: PaintInput;
  margin?: number | Partial<Insets>;
}

export type ElementPaint = NormalizedPaint;
