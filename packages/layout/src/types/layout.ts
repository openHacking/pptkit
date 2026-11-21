import type {
  NormalizedSlideEventTimeline,
  ConnectorAnchor,
  Box,
  NormalizedAsset,
  NormalizedConnectorElement,
  NormalizedChartElement,
  NormalizedChartMarkerStyle,
  NormalizedElement,
  NormalizedGroupElement,
  NormalizedImageElement,
  NormalizedTableElement,
  NormalizedPaint,
  NormalizedBackgroundPaint,
  NormalizedPlaceholderDefinition,
  NormalizedPresentationMetadata,
  NormalizedPresentationTheme,
  NormalizedSlideMaster,
  NormalizedTextParagraph,
  SlideTransitionInput,
  SlideReadingOrder,
  Point,
  PresentationSize,
  PresentationSection,
  JsonValue,
} from "@pptkit/core";
import type { LayoutAnalysis } from "./analysis.js";
import type { LayoutPagination } from "./pagination.js";
import type { LayoutTableBorderGrid } from "./table-borders.js";

export interface LayoutConnectorElement extends Omit<NormalizedConnectorElement, "start" | "end"> {
  start: Point;
  end: Point;
  /** Authored endpoint identities retained alongside resolved coordinates. */
  connections?: {
    start?: { elementId: string; anchor: ConnectorAnchor };
    end?: { elementId: string; anchor: ConnectorAnchor };
  };
}

export interface LayoutGroupElement extends Omit<NormalizedGroupElement, "children"> {
  children: LayoutElement[];
}

export interface ChartValueScale {
  min: number;
  max: number;
  majorUnit: number;
  ticks: number[];
}

export interface ChartLegendItemLayout {
  label: string;
  color: string;
  marker: false | NormalizedChartMarkerStyle;
  box: Box;
  seriesIndex?: number;
  pointIndex?: number;
}

export interface ChartBarLayout {
  seriesIndex: number;
  categoryIndex: number;
  value: number;
  /** True when a blank point is displayed as zero. */
  blank?: true;
  box: Box;
}

export interface ChartLineLayout {
  seriesIndex: number;
  /** Displayed points in category order, including zero-display blanks. */
  points: Point[];
  /** Paths for gap or span display when the series contains blanks. */
  segments?: Point[][];
}

export interface ChartAreaLayout {
  seriesIndex: number;
  /** Closed polygons; each is top boundary followed by reversed baseline. */
  polygons: Point[][];
}

export interface ChartPieSliceLayout {
  pointIndex: number;
  startAngle: number;
  endAngle: number;
  color: string;
}

export interface ResolvedChartLayout {
  plotBox: Box;
  titleBox?: Box;
  legendBox?: Box;
  valueScale: ChartValueScale;
  secondaryValueScale?: ChartValueScale;
  /** Horizontal numeric scale and paired point geometry for XY charts. */
  xScale?: ChartValueScale;
  bubbles?: { seriesIndex: number; points: (Point & { sourceIndex: number; radius: number })[] }[];
  scatter?: import("../chart/scatter-geometry.js").ScatterGeometrySeriesLayout[];
  categoryPositions: Point[];
  categoryTickPositions: Point[];
  bars: ChartBarLayout[];
  lines: ChartLineLayout[];
  /** Present for area charts; absent on older layout results. */
  areas?: ChartAreaLayout[];
  pie: { center: Point; radius: number; slices: ChartPieSliceLayout[] } | undefined;
  /** Concentric rings in series order, from the center outward. */
  doughnut?: { center: Point; rings: { seriesIndex: number; innerRadius: number; outerRadius: number; slices: ChartPieSliceLayout[] }[] };
  legendItems: ChartLegendItemLayout[];
  categoryAxisPlacement?: "onCategory" | "betweenCategories";
}

export type LayoutChartElement = NormalizedChartElement & { chartLayout: ResolvedChartLayout };

export interface LayoutTableElement extends NormalizedTableElement {
  /** Deterministic physical edge grid for tables with explicit per-edge overrides. */
  borderGrid?: LayoutTableBorderGrid;
  /** Exact table-local cell geometry when explicit sizing or pagination succeeded. */
  cellLayout?: { rowIndex: number; cellIndex: number; columnIndex: number; box: Box }[];
}

export type LayoutElement =
  | Exclude<NormalizedElement, NormalizedConnectorElement | NormalizedGroupElement | NormalizedChartElement | NormalizedTableElement>
  | LayoutConnectorElement
  | LayoutChartElement
  | LayoutTableElement
  | LayoutGroupElement;

export interface LayoutSlideMaster extends Omit<NormalizedSlideMaster, "elements"> {
  elements?: LayoutElement[];
}

export interface LayoutSlideLayout {
  showMasterShapes?: boolean;
  backgroundSource?: "layout" | "master" | "theme";
  masterId?: string;
  theme?: NormalizedPresentationTheme;
  id: string;
  name: string;
  background: NormalizedBackgroundPaint;
  elements: LayoutElement[];
  placeholders: NormalizedPlaceholderDefinition[];
}

export interface LayoutSlide {
  showInheritedShapes?: boolean;
  id: string;
  layoutId: string;
  background: NormalizedBackgroundPaint;
  backgroundSource: "slide" | "layout" | "master" | "theme";
  elements: LayoutElement[];
  notes: NormalizedTextParagraph[];
  transition?: SlideTransitionInput;
  eventTimeline?: NormalizedSlideEventTimeline;
  readingOrder?: readonly SlideReadingOrder[];
  hidden: boolean;
  section?: string;
  tags: string[];
  customData: Record<string, JsonValue>;
}

export interface LayoutResult {
  sections?: PresentationSection[];
  firstSlideNumber?: number;
  masters?: LayoutSlideMaster[];
  pagination?: LayoutPagination;
  analysis?: LayoutAnalysis;
  size: PresentationSize;
  metadata: NormalizedPresentationMetadata;
  theme: NormalizedPresentationTheme;
  assets: NormalizedAsset[];
  layouts: LayoutSlideLayout[];
  slides: LayoutSlide[];
  slideCount: number;
  status: "resolved";
}
