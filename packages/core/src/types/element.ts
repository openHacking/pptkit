import type { Box, Insets, Point, Size } from "./geometry.js";
import type {
  NormalizedPaint,
  NormalizedShapeStyle,
  NormalizedStrokeStyle,
  NormalizedTableCellStyle,
  NormalizedTextFrameStyle,
  NormalizedTextParagraphStyle,
  NormalizedTextRunStyle,
  NormalizedTransform,
  PaintInput,
  ShapeStyleInput,
  StrokeStyleInput,
  TableCellStyleInput,
  TextFrameStyleInput,
  TextParagraphStyleInput,
  TextRunStyleInput,
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
  style: NormalizedTextRunStyle;
  action?: ElementAction;
}

export interface NormalizedTextParagraph {
  runs: NormalizedTextRun[];
  style: NormalizedTextParagraphStyle;
}

export interface TextElementInput extends ElementBaseInput {
  type: "text";
  content: TextContentInput;
  frame?: TextFrameStyleInput;
}

export interface TextElement extends ElementBase, Omit<TextElementInput, keyof ElementBaseInput> {}

export interface NormalizedTextElement extends NormalizedElementBase {
  type: "text";
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
  crop?: Partial<ImageCrop>;
}

export interface ImageElement extends ElementBase, Omit<ImageElementInput, keyof ElementBaseInput> {}

export interface NormalizedImageElement extends NormalizedElementBase {
  type: "image";
  assetId: string;
  fit: ImageFit;
  crop: ImageCrop;
}

export type ShapeKind = "rect" | "roundRect" | "ellipse" | "triangle" | "diamond" | "arrow" | "chevron";

export interface ShapeElementInput extends ElementBaseInput {
  type: "shape";
  shape: ShapeKind;
  style?: ShapeStyleInput;
}

export interface ShapeElement extends ElementBase, Omit<ShapeElementInput, keyof ElementBaseInput> {}

export interface NormalizedShapeElement extends NormalizedElementBase {
  type: "shape";
  shape: ShapeKind;
  style: NormalizedShapeStyle;
}

export type ConnectorAnchor = "top" | "right" | "bottom" | "left" | "center";
export type ConnectorEndpointInput = Point | { elementId: string; anchor?: ConnectorAnchor };

export interface ConnectorElementInput extends ElementBaseInput {
  type: "connector";
  start: ConnectorEndpointInput;
  end: ConnectorEndpointInput;
  route?: Point[];
  style?: StrokeStyleInput;
}

export interface ConnectorElement extends ElementBase, Omit<ConnectorElementInput, keyof ElementBaseInput> {}

export interface NormalizedConnectorElement extends NormalizedElementBase {
  type: "connector";
  start: ConnectorEndpointInput;
  end: ConnectorEndpointInput;
  route: Point[];
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
  rowSpan?: number;
  colSpan?: number;
  style?: TableCellStyleInput;
}

export interface TableRowInput {
  cells: TableCellInput[];
  height?: number;
}

export interface TableElementInput extends ElementBaseInput {
  type: "table";
  columns: number[];
  rows: TableRowInput[];
}

export interface TableElement extends ElementBase, Omit<TableElementInput, keyof ElementBaseInput> {}

export interface NormalizedTableCell {
  content: NormalizedTextParagraph[];
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
  columns: number[];
  rows: NormalizedTableRow[];
}

export type PlaceholderKind = "title" | "subtitle" | "body" | "image" | "table" | "footer" | "slideNumber";

export interface PlaceholderTextStyleInput {
  frame?: TextFrameStyleInput;
  paragraph?: TextParagraphStyleInput;
  run?: TextRunStyleInput;
}

export interface PlaceholderDefinitionInput {
  key: string;
  kind: PlaceholderKind;
  box: Box;
  textStyle?: PlaceholderTextStyleInput;
}

export interface NormalizedPlaceholderDefinition {
  key: string;
  kind: PlaceholderKind;
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
  | ShapeElementInput
  | ConnectorElementInput
  | GroupElementInput
  | TableElementInput;

export type PresentationElement =
  | TextElement
  | ImageElement
  | ShapeElement
  | ConnectorElement
  | GroupElement
  | TableElement;

export type NormalizedElement =
  | NormalizedTextElement
  | NormalizedImageElement
  | NormalizedShapeElement
  | NormalizedConnectorElement
  | NormalizedGroupElement
  | NormalizedTableElement;

export interface PlaceholderBindingStyle {
  frame?: TextFrameStyleInput;
  paragraph?: TextParagraphStyleInput;
  run?: TextRunStyleInput;
  fill?: PaintInput;
  margin?: number | Partial<Insets>;
}

export type ElementPaint = NormalizedPaint;
