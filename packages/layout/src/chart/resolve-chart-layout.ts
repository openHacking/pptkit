import type { Box, ChartBlankDisplay, NormalizedChartElement, NormalizedBubbleChartElement, NormalizedScatterChartElement, NormalizedValueAxis, Point } from "@pptkit/core";
import type { ChartAreaLayout, ChartBarLayout, ChartLineLayout, ChartValueScale, ResolvedChartLayout } from "../types/layout.js";

import { scatterGeometry } from "./scatter-geometry.js";

type XYChart = NormalizedScatterChartElement | NormalizedBubbleChartElement;
type CategoryChart = Exclude<NormalizedChartElement, XYChart>;
const EPSILON = 1e-9;

/** Older transported normalized IR may predate the explicit display property. */
function blankDisplay(chart: CategoryChart): ChartBlankDisplay {
  return chart.blankDisplay ?? (chart.chartType === "area" && chart.grouping !== "standard" ? "zero" : chart.series.some(series => series.values.includes(null)) ? "gap" : "zero");
}

function niceNumber(value: number, round: boolean): number {
  if (!Number.isFinite(value) || value <= 0) return 1;
  // Powers of ten below the subnormal range underflow to zero.
  if (value < 1e-300) return value;
  const exponent = Math.floor(Math.log10(value));
  const fraction = value / 10 ** exponent;
  const niceFraction = round
    ? fraction < 1.5 ? 1 : fraction < 3 ? 2 : fraction < 7 ? 5 : 10
    : fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
  return niceFraction * 10 ** exponent;
}

type AreaChart = Extract<NormalizedChartElement, { chartType: "area" }>;
/** Series-order signed accumulation, as used by native stacked area charts. */
function areaBands(chart: AreaChart): { start: number; end: number }[][] {
  const result = chart.series.map(() => [] as { start: number; end: number }[]);
  for (let i = 0; i < chart.categories.length; i++) {
    const max = chart.grouping === "percentStacked" ? chart.series.reduce((m, series) => Math.max(m, Math.abs(series.values[i] ?? 0)), 0) : 0;
    const total = max === 0 ? 0 : chart.series.reduce((sum, series) => sum + Math.abs((series.values[i] ?? 0) / max), 0);
    let sum = 0;
    chart.series.forEach((series, j) => {
      const raw = series.values[i] ?? 0;
      const value = chart.grouping === "percentStacked" ? (total === 0 ? 0 : (raw / max) / total) : raw;
      const start = chart.grouping === "standard" ? 0 : sum;
      const end = start + value;
      result[j]!.push({ start, end });
      if (chart.grouping !== "standard") sum = end;
    });
  }
  return result;
}

function chartDataExtent(chart: CategoryChart): { min: number; max: number } {
  const mode = blankDisplay(chart);
  if (chart.chartType === "area") {
    let min = 0, max = 0;
    for (const band of areaBands(chart)) for (const point of band) { min = Math.min(min, point.start, point.end); max = Math.max(max, point.start, point.end); }
    return { min, max };
  }
  if (chart.chartType === "bar" && chart.grouping !== "clustered") {
    let min = 0;
    let max = 0;
    for (let categoryIndex = 0; categoryIndex < chart.categories.length; categoryIndex += 1) {
      let positive = 0;
      let negative = 0;
      for (const series of chart.series) {
        const value = series.values[categoryIndex] ?? 0;
        if (value >= 0) positive += value;
        else negative += value;
      }
      if (chart.grouping === "percentStacked") {
        if (positive > 0) max = 100;
        if (negative < 0) min = -100;
      } else {
        min = Math.min(min, negative);
        max = Math.max(max, positive);
      }
    }
    return { min, max };
  }
  let min = Infinity, max = -Infinity, hasValue = false;
  for (const series of chart.series) for (let index = 0; index < series.values.length; index++) {
    if (!(index in series.values)) continue;
    const raw = series.values[index] as number | null;
    const value = raw === null && mode === "zero" ? 0 : raw;
    if (value !== null) { hasValue = true; min = Math.min(min, value); max = Math.max(max, value); }
  }
  return hasValue ? { min, max } : { min: 0, max: 0 };
}

function automaticScale(chart: CategoryChart, plotLength: number): ChartValueScale {
  const extent = chartDataExtent(chart);
  if (chart.chartType === "area" && chart.grouping === "percentStacked") return makeScale(extent.min < 0 ? -1 : 0, extent.max > 0 || extent.min === 0 ? 1 : 0, 0.2);
  if (chart.chartType === "bar" && chart.grouping === "percentStacked") {
    const intervalCount = Math.max(5, Math.min(7, Math.round(plotLength / 40)));
    const majorUnit = niceNumber((extent.max - extent.min) / intervalCount, true);
    return makeScale(extent.min, extent.max > extent.min ? extent.max : extent.min + majorUnit, majorUnit);
  }
  return ordinaryScale(extent, plotLength, chart.chartType === "bar" || chart.chartType === "area");
}

function ordinaryScale(extent: { min: number; max: number }, plotLength: number, includeZero: boolean): ChartValueScale {
  let min = extent.min;
  let max = extent.max;
  const span = max - min;
  if (includeZero) {
    min = Math.min(min, 0);
    max = Math.max(max, 0);
  } else if (span > 0) {
    if (min >= 0 && min <= span * 0.2) min = 0;
    if (max <= 0 && Math.abs(max) <= span * 0.2) max = 0;
  }
  if (max === min) {
    const magnitude = min !== 0 && Math.abs(min) < EPSILON ? Math.abs(min) : Math.max(Math.abs(min), 1);
    if (min === 0) { min = 0; max = 1; }
    else {
      const lower = min - magnitude * 0.1, upper = max + magnitude * 0.1;
      min = Number.isFinite(lower) ? lower : -Number.MAX_VALUE;
      max = Number.isFinite(upper) ? upper : Number.MAX_VALUE;
      // Subnormal magnitudes can make the ten-percent expansion round to zero.
      if (min === max) { if (min > 0) min = 0; else max = 0; }
    }
  }
  const paddedSpan = max - min;
  if (!Number.isFinite(paddedSpan)) return makeScale(min, max, Math.max(Math.abs(min), Math.abs(max)) / 2);
  const paddedMin = min < 0 ? min - paddedSpan * 0.05 : min;
  const paddedMax = max > 0 ? max + paddedSpan * 0.05 : max;
  if (!Number.isFinite(paddedMin) || !Number.isFinite(paddedMax) || !Number.isFinite(paddedMax - paddedMin)) return makeScale(min, max, Math.max(Math.abs(min), Math.abs(max)) / 2);
  const intervalCount = Math.max(5, Math.min(7, Math.round(plotLength / 40)));
  const majorUnit = niceNumber(Math.max(Number.MIN_VALUE, (paddedMax - paddedMin) / intervalCount), true);
  const niceMin = Math.floor(paddedMin / majorUnit) * majorUnit;
  const niceMax = Math.ceil(paddedMax / majorUnit) * majorUnit;
  if (!Number.isFinite(niceMin) || !Number.isFinite(niceMax)) return makeScale(min, max, Math.max(Math.abs(min), Math.abs(max)) / 2);
  return makeScale(niceMin, niceMax > niceMin ? niceMax : niceMin + majorUnit, majorUnit);
}

function roundedTick(value: number, majorUnit: number): number {
  if (Math.abs(value) < Math.abs(majorUnit) * 1e-9) return 0;
  if (Math.abs(majorUnit) < 1e-9) {
    const rounded = Number(value.toPrecision(12));
    return Math.abs(rounded - value) <= Math.abs(majorUnit) * 1e-6 ? rounded : value;
  }
  return Number.isFinite(value * 1e12) ? Math.round(value * 1e12) / 1e12 : value;
}

function makeScale(min: number, max: number, majorUnit: number): ChartValueScale {
  const ticks: number[] = [];
  const limit = 1000;
  for (let index = 0; index < limit; index += 1) {
    const offset = index * majorUnit;
    const value = Number.isFinite(offset) ? min + offset : (min / 2 + index * (majorUnit / 2)) * 2;
    if (!Number.isFinite(value) || value > max + majorUnit / 2) break;
    ticks.push(roundedTick(value, majorUnit));
  }
  return { min, max, majorUnit, ticks };
}

function valueScale(chart: CategoryChart, plotLength: number): ChartValueScale {
  if (chart.chartType !== "pie" && chart.chartType !== "doughnut" && chart.axes.value.scale !== "auto") {
    const { min, max, majorUnit } = chart.axes.value.scale;
    return makeScale(min, max, majorUnit);
  }
  return automaticScale(chart, plotLength);
}

function valueRatio(value: number, scale: ChartValueScale): number {
  const span = scale.max - scale.min;
  return Number.isFinite(span) ? (span > 0 ? (value - scale.min) / span : 0) : (value / 2 - scale.min / 2) / (scale.max / 2 - scale.min / 2);
}

function categoryPositions(chart: CategoryChart, plot: Box): Point[] {
  if (chart.chartType === "bar" && chart.orientation === "horizontal") {
    const step = plot.height / Math.max(chart.categories.length, 1);
    return chart.categories.map((_, index) => ({ x: plot.x, y: plot.y + (index + 0.5) * step }));
  }
  if (chart.chartType === "line" || chart.chartType === "area") {
    const step = plot.width / Math.max(chart.categories.length - 1, 1);
    return chart.categories.map((_, index) => ({ x: plot.x + (chart.categories.length === 1 ? plot.width / 2 : index * step), y: plot.y + plot.height }));
  }
  const step = plot.width / Math.max(chart.categories.length, 1);
  return chart.categories.map((_, index) => ({ x: plot.x + (index + 0.5) * step, y: plot.y + plot.height }));
}

function categoryTickPositions(chart: CategoryChart, plot: Box, categories: Point[]): Point[] {
  if (chart.chartType === "line" || chart.chartType === "area") return categories;
  if (chart.chartType === "bar" && chart.orientation === "horizontal") {
    const step = plot.height / Math.max(chart.categories.length, 1);
    return Array.from({ length: chart.categories.length + 1 }, (_, index) => ({ x: plot.x, y: plot.y + index * step }));
  }
  const step = plot.width / Math.max(chart.categories.length, 1);
  return Array.from({ length: chart.categories.length + 1 }, (_, index) => ({ x: plot.x + index * step, y: plot.y + plot.height }));
}

function normalizedStackValue(chart: Extract<NormalizedChartElement, { chartType: "bar" }>, categoryIndex: number, value: number): number {
  if (chart.grouping !== "percentStacked") return value;
  if (value === 0) return 0;
  const sameSign = (candidate: number): boolean => value > 0 ? candidate > 0 : candidate < 0;
  const maximum = chart.series.reduce((max, series) => {
    const candidate = series.values[categoryIndex] ?? 0;
    return sameSign(candidate) ? Math.max(max, Math.abs(candidate)) : max;
  }, 0);
  const total = chart.series.reduce((sum, series) => {
    const candidate = series.values[categoryIndex] ?? 0;
    return sameSign(candidate) ? sum + Math.abs(candidate / maximum) : sum;
  }, 0);
  return maximum === 0 ? 0 : (value / maximum) / total * 100;
}

function barGeometry(chart: Extract<NormalizedChartElement, { chartType: "bar" }>, plot: Box, scale: ChartValueScale): ChartBarLayout[] {
  const mode = blankDisplay(chart);
  const horizontal = chart.orientation === "horizontal";
  const slot = (horizontal ? plot.height : plot.width) / Math.max(chart.categories.length, 1);
  const seriesCount = chart.grouping === "clustered" ? Math.max(chart.series.length, 1) : 1;
  const gapRatio = chart.seriesGap / 100;
  const categoryGapRatio = chart.categoryGap / 100;
  const barThickness = slot / Math.max(seriesCount + Math.max(0, seriesCount - 1) * gapRatio + categoryGapRatio, EPSILON);
  const groupThickness = seriesCount * barThickness + Math.max(0, seriesCount - 1) * gapRatio * barThickness;
  const bars: ChartBarLayout[] = [];
  const positive = new Array(chart.categories.length).fill(0) as number[];
  const negative = new Array(chart.categories.length).fill(0) as number[];
  for (let categoryIndex = 0; categoryIndex < chart.categories.length; categoryIndex += 1) {
    for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex += 1) {
      const sourceValue = chart.series[seriesIndex]?.values[categoryIndex];
      if (sourceValue === undefined || (sourceValue === null && mode !== "zero")) continue;
      const rawValue = sourceValue ?? 0;
      const blank = sourceValue === null ? { blank: true as const } : {};
      const value = normalizedStackValue(chart, categoryIndex, rawValue);
      let start = 0;
      let end = value;
      if (chart.grouping !== "clustered") {
        if (value >= 0) {
          start = positive[categoryIndex] ?? 0;
          end = start + value;
          positive[categoryIndex] = end;
        } else {
          start = negative[categoryIndex] ?? 0;
          end = start + value;
          negative[categoryIndex] = end;
        }
      }
      const categoryStart = (horizontal ? plot.y : plot.x) + categoryIndex * slot + (slot - groupThickness) / 2;
      const thicknessOffset = chart.grouping === "clustered" ? seriesIndex * barThickness * (1 + gapRatio) : 0;
      if (horizontal) {
        const x1 = plot.x + valueRatio(start, scale) * plot.width;
        const x2 = plot.x + valueRatio(end, scale) * plot.width;
        bars.push({ seriesIndex, categoryIndex, value: rawValue, ...blank, box: { x: Math.min(x1, x2), y: categoryStart + thicknessOffset, width: Math.abs(x2 - x1), height: barThickness } });
      } else {
        const y1 = plot.y + plot.height - valueRatio(start, scale) * plot.height;
        const y2 = plot.y + plot.height - valueRatio(end, scale) * plot.height;
        bars.push({ seriesIndex, categoryIndex, value: rawValue, ...blank, box: { x: categoryStart + thicknessOffset, y: Math.min(y1, y2), width: barThickness, height: Math.abs(y2 - y1) } });
      }
    }
  }
  return bars;
}

function lineGeometry(chart: Extract<NormalizedChartElement, { chartType: "line" }>, plot: Box, scale: ChartValueScale, categories: Point[]): ChartLineLayout[] {
  const mode = blankDisplay(chart);
  return chart.series.map((series, seriesIndex) => {
    const points: Point[] = [], segments: Point[][] = [];
    let current: Point[] = [];
    for (const [categoryIndex, value] of series.values.entries()) {
      if (value === null && mode !== "zero") {
        if (mode === "gap") { if (current.length) segments.push(current); current = []; }
        continue;
      }
      const point = { x: categories[categoryIndex]?.x ?? plot.x, y: plot.y + plot.height - valueRatio(value ?? 0, scale) * plot.height };
      points.push(point); current.push({ ...point });
    }
    if (current.length) segments.push(current);
    return { seriesIndex, points, ...(series.values.includes(null) && mode !== "zero" ? { segments } : {}) };
  });
}

function areaGeometry(chart: AreaChart, plot: Box, scale: ChartValueScale, categories: Point[]): ChartAreaLayout[] {
  const bands = areaBands(chart), mode = blankDisplay(chart);
  return chart.series.map((series, seriesIndex) => {
    const polygons: Point[][] = [], upper: Point[] = [], lower: Point[] = [];
    const flush = () => { if (upper.length) { polygons.push([...upper, ...lower.slice().reverse()]); upper.length = 0; lower.length = 0; } };
    series.values.forEach((value, index) => {
      if (value === null && mode === "gap") { flush(); return; }
      const band = bands[seriesIndex]![index]!, x = categories[index]!.x;
      upper.push({ x, y: plot.y + plot.height - valueRatio(band.end, scale) * plot.height });
      lower.push({ x, y: plot.y + plot.height - valueRatio(band.start, scale) * plot.height });
    });
    flush(); return { seriesIndex, polygons };
  });
}


function xyAxisScale(chart: XYChart, axis: NormalizedValueAxis, coordinate: "xValues" | "yValues", length: number): ChartValueScale {
  if (axis.scale !== "auto") return makeScale(axis.scale.min, axis.scale.max, axis.scale.majorUnit);
  const mode = chart.chartType === "scatter" ? chart.blankDisplay ?? "gap" : "gap";
  let min = Infinity, max = -Infinity;
  for (const series of chart.series) for (let i = 0; i < series.xValues.length; i++) {
    const x = series.xValues[i]!, y = series.yValues[i]!;
    if ((x === null || y === null) && mode !== "zero") continue;
    const value = (coordinate === "xValues" ? x : y) ?? 0; min = Math.min(min, value); max = Math.max(max, value);
  }
  return ordinaryScale(min === Infinity ? { min: 0, max: 0 } : { min, max }, length, false);
}

function resolveXYLayout(chart: XYChart): ResolvedChartLayout {
  const padding = 8, titleHeight = chart.title === undefined ? 0 : 22;
  const legendWidth = chart.legend.visible && chart.legend.position === "right" ? Math.min(100, chart.box.width * .3) : 0;
  const legendHeight = chart.legend.visible && chart.legend.position === "bottom" ? 20 : 0;
  const left = chart.axes.y.visible && chart.axes.y.labels ? 38 : padding, bottom = chart.axes.x.visible && chart.axes.x.labels ? 22 : padding;
  const plotBox = { x: padding + left, y: padding + titleHeight, width: Math.max(1, chart.box.width - padding * 2 - left - legendWidth), height: Math.max(1, chart.box.height - padding * 2 - titleHeight - bottom - legendHeight) };
  const xScale = xyAxisScale(chart, chart.axes.x, "xValues", plotBox.width), valueScale = xyAxisScale(chart, chart.axes.y, "yValues", plotBox.height);
  const legendBox = !chart.legend.visible ? undefined : chart.legend.position === "right" ? { x: chart.box.width - legendWidth, y: plotBox.y, width: legendWidth, height: Math.min(plotBox.height, chart.series.length * Math.max(14, chart.legend.textStyle.fontSize + 5)) } : { x: plotBox.x, y: chart.box.height - legendHeight, width: plotBox.width, height: legendHeight };
  const legendItems = legendBox === undefined ? [] : chart.series.map((series, index) => ({ label: series.name, color: series.color, marker: "marker" in series ? series.marker : false, seriesIndex: index, box: chart.legend.position === "right" ? { x: legendBox.x, y: legendBox.y + index * Math.max(14, chart.legend.textStyle.fontSize + 5), width: legendBox.width, height: Math.max(14, chart.legend.textStyle.fontSize + 5) } : { x: legendBox.x + index * legendBox.width / chart.series.length, y: legendBox.y, width: legendBox.width / chart.series.length, height: legendBox.height } }));
  const titleBox = chart.title === undefined ? undefined : { x: padding, y: padding, width: chart.box.width - padding * 2, height: titleHeight };
  const geometry = scatterGeometry(chart.series, plotBox, xScale, valueScale, chart.chartType === "scatter" ? chart.blankDisplay ?? "gap" : "gap");
  let bubbles: ResolvedChartLayout["bubbles"];
  if (chart.chartType === "bubble") {
    let largest = 0;
    for (const series of chart.series) for (const size of series.sizes) largest = Math.max(largest, size);
    const scale = Math.sqrt(largest), radius = Math.min(plotBox.width, plotBox.height) * .12;
    bubbles = geometry.map(series => ({ seriesIndex: series.seriesIndex, points: series.points.map(point => ({
      ...point, radius: radius * (Math.sqrt(chart.series[series.seriesIndex]!.sizes[point.sourceIndex]!) / scale),
    })) }));
  }
  return { plotBox, ...(titleBox ? { titleBox } : {}), ...(legendBox ? { legendBox } : {}), valueScale, xScale,
    categoryPositions: [], categoryTickPositions: [], bars: [], lines: [], ...(bubbles ? { bubbles } : { scatter: geometry }),
    pie: undefined, legendItems };
}

export function resolveChartLayout(chart: NormalizedChartElement): ResolvedChartLayout {
  if (chart.chartType === "scatter" || chart.chartType === "bubble") return resolveXYLayout(chart);
  const padding = 8;
  const titleHeight = chart.title === undefined ? 0 : 22;
  const legendItems = (chart.chartType === "pie" || chart.chartType === "doughnut")
    ? chart.categories.map((label, pointIndex) => ({ label, color: chart.series[0]?.pointColors[pointIndex] ?? chart.series[0]?.color ?? "#000000", marker: false as const, pointIndex }))
    : chart.chartType === "line"
      ? chart.series.map((series, seriesIndex) => ({ label: series.name, color: series.color, marker: series.marker, seriesIndex }))
      : chart.series.map((series, seriesIndex) => ({ label: series.name, color: series.color, marker: "marker" in series ? series.marker : false as const, seriesIndex }));
  const legendWidth = chart.legend.visible && chart.legend.position === "right" ? Math.min(100, chart.box.width * 0.3) : 0;
  const legendHeight = chart.legend.visible && chart.legend.position === "bottom" ? 20 : 0;
  const horizontal = chart.chartType === "bar" && chart.orientation === "horizontal";
  const categoryMargin = chart.chartType !== "pie" && chart.chartType !== "doughnut" && chart.axes.category.visible && chart.axes.category.labels ? (horizontal ? 52 : 22) : padding;
  const valueMargin = chart.chartType !== "pie" && chart.chartType !== "doughnut" && chart.axes.value.visible && chart.axes.value.labels ? (horizontal ? 22 : 38) : padding;
  const secondaryMargin = chart.chartType === "combo" && chart.axes.secondaryValue.visible ? (chart.axes.secondaryValue.labels ? 45 : 12) : 0;
  const plotBox: Box = (chart.chartType === "pie" || chart.chartType === "doughnut")
    ? { x: padding, y: padding + titleHeight, width: Math.max(1, chart.box.width - padding * 2 - legendWidth), height: Math.max(1, chart.box.height - padding * 2 - titleHeight - legendHeight) }
    : horizontal
      ? { x: padding + categoryMargin, y: padding + titleHeight, width: Math.max(1, chart.box.width - padding * 2 - categoryMargin - legendWidth), height: Math.max(1, chart.box.height - padding * 2 - titleHeight - valueMargin - legendHeight) }
      : { x: padding + valueMargin, y: padding + titleHeight, width: Math.max(1, chart.box.width - padding * 2 - valueMargin - legendWidth - secondaryMargin), height: Math.max(1, chart.box.height - padding * 2 - titleHeight - categoryMargin - legendHeight) };
  // Doughnuts have no axes; avoid deriving an unused scale from potentially huge values.
  const primary = chart.chartType === "combo" ? { ...chart, chartType: "bar" as const, orientation: "vertical" as const,
    grouping: "clustered" as const, seriesGap: 0, categoryGap: 150, series: chart.series.filter(item => item.kind === "bar") } : undefined;
  const secondary = chart.chartType === "combo" ? { ...chart, chartType: "line" as const,
    axes: { category: chart.axes.category, value: chart.axes.secondaryValue }, series: chart.series.filter(item => item.kind === "line") } : undefined;
  const scale = chart.chartType === "doughnut" ? makeScale(0, 1, 1) : valueScale(primary ?? chart, horizontal ? plotBox.width : plotBox.height);
  const secondaryValueScale = secondary ? valueScale(secondary, plotBox.height) : undefined;
  const categories = categoryPositions(chart, plotBox);
  const categoryTicks = categoryTickPositions(chart, plotBox, categories);
  const legendBox = chart.legend.visible
    ? chart.legend.position === "right"
      ? (() => {
          const rowHeight = Math.max(14, chart.legend.textStyle.fontSize + 5);
          const height = Math.min(plotBox.height, Math.max(rowHeight, legendItems.length * rowHeight));
          return { x: chart.box.width - legendWidth, y: plotBox.y + (plotBox.height - height) / 2, width: legendWidth, height };
        })()
      : { x: plotBox.x, y: chart.box.height - legendHeight, width: plotBox.width, height: legendHeight }
    : undefined;
  const resolvedLegendItems = legendBox === undefined ? [] : legendItems.map((item, index) => {
    if (chart.legend.position === "right") {
      const rowHeight = Math.max(14, chart.legend.textStyle.fontSize + 5);
      return { ...item, box: { x: legendBox.x, y: legendBox.y + index * rowHeight, width: legendBox.width, height: rowHeight } };
    }
    const itemWidth = legendItems.length === 0 ? legendBox.width : legendBox.width / legendItems.length;
    return { ...item, box: { x: legendBox.x + index * itemWidth, y: legendBox.y, width: itemWidth, height: legendBox.height } };
  });
  const titleBox = chart.title === undefined ? undefined : { x: padding, y: padding, width: chart.box.width - padding * 2, height: titleHeight };
  const bars = primary ? barGeometry(primary, plotBox, scale) : chart.chartType === "bar" ? barGeometry(chart, plotBox, scale) : [];
  const lines = primary && secondary && secondaryValueScale
    ? lineGeometry(secondary, plotBox, secondaryValueScale, categories).map(line => ({ ...line, seriesIndex: line.seriesIndex + primary.series.length }))
    : chart.chartType === "line" ? lineGeometry(chart, plotBox, scale, categories) : [];
  const areas = chart.chartType === "area" ? areaGeometry(chart, plotBox, scale, categories) : undefined;
  let pie: ResolvedChartLayout["pie"];
  if (chart.chartType === "pie") {
    const series = chart.series[0];
    const maximum = series?.values.reduce<number>((max, value) => Math.max(max, value ?? 0), 0) ?? 0;
    const total = maximum === 0 ? 0 : series!.values.reduce<number>((sum, value) => sum + (value ?? 0) / maximum, 0);
    const center = { x: plotBox.x + plotBox.width / 2, y: plotBox.y + plotBox.height / 2 };
    const radius = Math.min(plotBox.width, plotBox.height) * 0.425;
    let angle = -Math.PI / 2;
    const slices = chart.categories.map((_, pointIndex) => {
      const value = Math.max(0, series?.values[pointIndex] ?? 0);
      const startAngle = angle;
      angle += total === 0 ? 0 : (value / maximum) / total * Math.PI * 2;
      return { pointIndex, startAngle, endAngle: angle, color: series?.pointColors[pointIndex] ?? series?.color ?? "#000000" };
    });
    pie = { center, radius, slices };
  }
  let doughnut: ResolvedChartLayout["doughnut"];
  if (chart.chartType === "doughnut") {
    const center = { x: plotBox.x + plotBox.width / 2, y: plotBox.y + plotBox.height / 2 };
    const radius = Math.min(plotBox.width, plotBox.height) * 0.5;
    const inner = radius * chart.holeSize / 100;
    const ringWidth = (radius - inner) / chart.series.length;
    const rings = chart.series.map((series, seriesIndex) => {
      // Scaling avoids overflowing the sum of individually finite values.
      const maximum = series.values.reduce<number>((max, value) => Math.max(max, value ?? 0), 0);
      const total = maximum === 0 ? 0 : series.values.reduce<number>((sum, value) => sum + (value ?? 0) / maximum, 0);
      let angle = -Math.PI / 2 + chart.firstSliceAngle * Math.PI / 180;
      const slices = series.values.map((value, pointIndex) => {
        const startAngle = angle;
        if (maximum !== 0 && total !== 0) angle += ((value ?? 0) / maximum) / total * Math.PI * 2;
        return { pointIndex, startAngle, endAngle: angle, color: series.pointColors[pointIndex] ?? series.color };
      });
      return { seriesIndex, innerRadius: inner + ringWidth * seriesIndex, outerRadius: inner + ringWidth * (seriesIndex + 1), slices };
    });
    doughnut = { center, rings };
  }
  return {
    plotBox,
    ...(titleBox === undefined ? {} : { titleBox }),
    ...(legendBox === undefined ? {} : { legendBox }),
    valueScale: scale,
    ...(secondaryValueScale ? { secondaryValueScale } : {}),
    categoryPositions: categories,
    categoryTickPositions: categoryTicks,
    bars,
    lines,
    ...(areas === undefined ? {} : { areas }),
    pie,
    ...(doughnut === undefined ? {} : { doughnut }),
    legendItems: resolvedLegendItems,
    ...((chart.chartType === "pie" || chart.chartType === "doughnut") ? {} : { categoryAxisPlacement: (chart.chartType === "line" || chart.chartType === "area") ? "onCategory" as const : "betweenCategories" as const }),
  };
}
