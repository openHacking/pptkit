import type {
  Box,
  NormalizedAsset,
  NormalizedConnectorElement,
  NormalizedChartElement,
  NormalizedChartMarkerStyle,
  NormalizedElement,
  NormalizedGroupElement,
  NormalizedImageElement,
  NormalizedPaint,
  NormalizedPlaceholderDefinition,
  NormalizedPresentationMetadata,
  NormalizedPresentationTheme,
  NormalizedTextParagraph,
  Point,
  PresentationSize,
  JsonValue,
} from "@pptkit/core";

export interface LayoutConnectorElement extends Omit<NormalizedConnectorElement, "start" | "end"> {
  start: Point;
  end: Point;
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
  box: Box;
}

export interface ChartLineLayout {
  seriesIndex: number;
  points: Point[];
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
  categoryPositions: Point[];
  categoryTickPositions: Point[];
  bars: ChartBarLayout[];
  lines: ChartLineLayout[];
  pie: { center: Point; radius: number; slices: ChartPieSliceLayout[] } | undefined;
  legendItems: ChartLegendItemLayout[];
  categoryAxisPlacement?: "onCategory" | "betweenCategories";
}

export type LayoutChartElement = NormalizedChartElement & { chartLayout: ResolvedChartLayout };

export type LayoutElement =
  | Exclude<NormalizedElement, NormalizedConnectorElement | NormalizedGroupElement | NormalizedChartElement>
  | LayoutConnectorElement
  | LayoutChartElement
  | LayoutGroupElement;

export interface LayoutSlideLayout {
  id: string;
  name: string;
  background: NormalizedPaint;
  elements: LayoutElement[];
  placeholders: NormalizedPlaceholderDefinition[];
}

export interface LayoutSlide {
  id: string;
  layoutId: string;
  background: NormalizedPaint;
  backgroundSource: "slide" | "layout" | "theme";
  elements: LayoutElement[];
  notes: NormalizedTextParagraph[];
  hidden: boolean;
  section?: string;
  tags: string[];
  customData: Record<string, JsonValue>;
}

export interface LayoutResult {
  size: PresentationSize;
  metadata: NormalizedPresentationMetadata;
  theme: NormalizedPresentationTheme;
  assets: NormalizedAsset[];
  layouts: LayoutSlideLayout[];
  slides: LayoutSlide[];
  slideCount: number;
  status: "resolved";
}
