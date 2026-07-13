import type { PresentationSize } from "../types/geometry.js";
import type {
  NormalizedPaint,
  NormalizedShapeStyle,
  NormalizedStrokeStyle,
  NormalizedTableCellStyle,
  NormalizedTextFrameStyle,
  NormalizedTextParagraphStyle,
  NormalizedTextRunStyle,
} from "../types/style.js";
import type { NormalizedPresentationTheme } from "../types/theme.js";

export const DEFAULT_PRESENTATION_SIZE: PresentationSize = {
  width: 960,
  height: 540,
  unit: "pt",
};

export const DEFAULT_LAYOUT_ID = "layout-blank";

export const DEFAULT_THEME: NormalizedPresentationTheme = {
  name: "PPTKit",
  colors: {
    background1: "FFFFFF",
    background2: "F2F2F2",
    text1: "000000",
    text2: "404040",
    accent1: "4472C4",
    accent2: "ED7D31",
    accent3: "A5A5A5",
    accent4: "FFC000",
    accent5: "5B9BD5",
    accent6: "70AD47",
    hyperlink: "0563C1",
    followedHyperlink: "954F72",
  },
  fonts: {
    heading: "Aptos Display",
    body: "Aptos",
  },
};

export const DEFAULT_BACKGROUND: NormalizedPaint = {
  type: "solid",
  color: { theme: "background1" },
  opacity: 1,
};

export const DEFAULT_STROKE: NormalizedStrokeStyle = {
  paint: { type: "none" },
  width: 1,
  dash: "solid",
  beginArrow: "none",
  endArrow: "none",
};

export const DEFAULT_SHAPE_STYLE: NormalizedShapeStyle = {
  fill: { type: "none" },
  stroke: DEFAULT_STROKE,
};

export const DEFAULT_TEXT_FRAME_STYLE: NormalizedTextFrameStyle = {
  margin: { top: 3.5, right: 7, bottom: 3.5, left: 7 },
  verticalAlign: "top",
  wrap: true,
  autoFit: { mode: "none" },
};

export const DEFAULT_TEXT_PARAGRAPH_STYLE: NormalizedTextParagraphStyle = {
  align: "left",
  indent: 0,
  hanging: 0,
  lineSpacing: 1,
  spaceBefore: 0,
  spaceAfter: 0,
  bullet: { type: "none" },
};

export const DEFAULT_TEXT_RUN_STYLE: NormalizedTextRunStyle = {
  fontFamily: { theme: "body" },
  fontSize: 18,
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  color: { theme: "text1" },
  language: "en-US",
};

export const DEFAULT_TABLE_CELL_STYLE: NormalizedTableCellStyle = {
  fill: { type: "none" },
  stroke: DEFAULT_STROKE,
  margin: { top: 3.5, right: 7, bottom: 3.5, left: 7 },
  verticalAlign: "top",
};
