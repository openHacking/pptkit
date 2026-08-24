import type { Box, NormalizedChartElement, Point } from "@pptkit/core";
import type { ChartBarLayout, ChartLineLayout, ChartValueScale, ResolvedChartLayout } from "../types/layout.js";

const EPSILON = 1e-9;

function niceNumber(value: number, round: boolean): number {
  if (!Number.isFinite(value) || value <= 0) return 1;
  const exponent = Math.floor(Math.log10(value));
  const fraction = value / 10 ** exponent;
  const niceFraction = round
    ? fraction < 1.5 ? 1 : fraction < 3 ? 2 : fraction < 7 ? 5 : 10
    : fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
  return niceFraction * 10 ** exponent;
}

function chartDataExtent(chart: NormalizedChartElement): { min: number; max: number } {
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
  const values = chart.series.flatMap((series) => series.values);
  return values.length === 0 ? { min: 0, max: 0 } : { min: Math.min(...values), max: Math.max(...values) };
}

function automaticScale(chart: NormalizedChartElement, plotLength: number): ChartValueScale {
  const extent = chartDataExtent(chart);
  if (chart.chartType === "bar" && chart.grouping === "percentStacked") {
    const intervalCount = Math.max(5, Math.min(7, Math.round(plotLength / 40)));
    const majorUnit = niceNumber((extent.max - extent.min) / intervalCount, true);
    return makeScale(extent.min, extent.max > extent.min ? extent.max : extent.min + majorUnit, majorUnit);
  }
  let min = extent.min;
  let max = extent.max;
  const span = max - min;
  if (chart.chartType === "bar") {
    min = Math.min(min, 0);
    max = Math.max(max, 0);
  } else if (span > EPSILON) {
    if (min >= 0 && min <= span * 0.2) min = 0;
    if (max <= 0 && Math.abs(max) <= span * 0.2) max = 0;
  }
  if (Math.abs(max - min) <= EPSILON) {
    const magnitude = Math.max(Math.abs(min), 1);
    min = min === 0 ? 0 : min - magnitude * 0.1;
    max = max === 0 ? 1 : max + magnitude * 0.1;
  }
  const paddedSpan = max - min;
  const paddedMin = min < 0 ? min - paddedSpan * 0.05 : min;
  const paddedMax = max > 0 ? max + paddedSpan * 0.05 : max;
  const intervalCount = Math.max(5, Math.min(7, Math.round(plotLength / 40)));
  const majorUnit = niceNumber((paddedMax - paddedMin) / intervalCount, true);
  const niceMin = Math.floor(paddedMin / majorUnit) * majorUnit;
  const niceMax = Math.ceil(paddedMax / majorUnit) * majorUnit;
  return makeScale(niceMin, niceMax > niceMin ? niceMax : niceMin + majorUnit, majorUnit);
}

function makeScale(min: number, max: number, majorUnit: number): ChartValueScale {
  const ticks: number[] = [];
  const limit = 1000;
  for (let index = 0; index < limit; index += 1) {
    const value = min + index * majorUnit;
    if (value > max + majorUnit / 2) break;
    const rounded = Math.abs(value) < EPSILON ? 0 : Math.round(value * 1e12) / 1e12;
    ticks.push(rounded);
  }
  return { min, max, majorUnit, ticks };
}

function valueScale(chart: NormalizedChartElement, plotLength: number): ChartValueScale {
  if (chart.chartType !== "pie" && chart.axes.value.scale !== "auto") {
    const { min, max, majorUnit } = chart.axes.value.scale;
    return makeScale(min, max, majorUnit);
  }
  return automaticScale(chart, plotLength);
}

function valueRatio(value: number, scale: ChartValueScale): number {
  return (value - scale.min) / Math.max(scale.max - scale.min, EPSILON);
}

function categoryPositions(chart: NormalizedChartElement, plot: Box): Point[] {
  if (chart.chartType === "bar" && chart.orientation === "horizontal") {
    const step = plot.height / Math.max(chart.categories.length, 1);
    return chart.categories.map((_, index) => ({ x: plot.x, y: plot.y + (index + 0.5) * step }));
  }
  if (chart.chartType === "line") {
    const step = plot.width / Math.max(chart.categories.length - 1, 1);
    return chart.categories.map((_, index) => ({ x: plot.x + (chart.categories.length === 1 ? plot.width / 2 : index * step), y: plot.y + plot.height }));
  }
  const step = plot.width / Math.max(chart.categories.length, 1);
  return chart.categories.map((_, index) => ({ x: plot.x + (index + 0.5) * step, y: plot.y + plot.height }));
}

function categoryTickPositions(chart: NormalizedChartElement, plot: Box, categories: Point[]): Point[] {
  if (chart.chartType === "line") return categories;
  if (chart.chartType === "bar" && chart.orientation === "horizontal") {
    const step = plot.height / Math.max(chart.categories.length, 1);
    return Array.from({ length: chart.categories.length + 1 }, (_, index) => ({ x: plot.x, y: plot.y + index * step }));
  }
  const step = plot.width / Math.max(chart.categories.length, 1);
  return Array.from({ length: chart.categories.length + 1 }, (_, index) => ({ x: plot.x + index * step, y: plot.y + plot.height }));
}

function normalizedStackValue(chart: Extract<NormalizedChartElement, { chartType: "bar" }>, categoryIndex: number, value: number): number {
  if (chart.grouping !== "percentStacked") return value;
  const sameSignTotal = chart.series.reduce((sum, series) => {
    const candidate = series.values[categoryIndex] ?? 0;
    return candidate * value >= 0 ? sum + Math.abs(candidate) : sum;
  }, 0);
  return sameSignTotal <= EPSILON ? 0 : (value / sameSignTotal) * 100;
}

function barGeometry(chart: Extract<NormalizedChartElement, { chartType: "bar" }>, plot: Box, scale: ChartValueScale): ChartBarLayout[] {
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
      const rawValue = chart.series[seriesIndex]?.values[categoryIndex] ?? 0;
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
        bars.push({ seriesIndex, categoryIndex, value: rawValue, box: { x: Math.min(x1, x2), y: categoryStart + thicknessOffset, width: Math.abs(x2 - x1), height: barThickness } });
      } else {
        const y1 = plot.y + plot.height - valueRatio(start, scale) * plot.height;
        const y2 = plot.y + plot.height - valueRatio(end, scale) * plot.height;
        bars.push({ seriesIndex, categoryIndex, value: rawValue, box: { x: categoryStart + thicknessOffset, y: Math.min(y1, y2), width: barThickness, height: Math.abs(y2 - y1) } });
      }
    }
  }
  return bars;
}

function lineGeometry(chart: Extract<NormalizedChartElement, { chartType: "line" }>, plot: Box, scale: ChartValueScale, categories: Point[]): ChartLineLayout[] {
  return chart.series.map((series, seriesIndex) => ({
    seriesIndex,
    points: series.values.map((value, categoryIndex) => ({
      x: categories[categoryIndex]?.x ?? plot.x,
      y: plot.y + plot.height - valueRatio(value, scale) * plot.height,
    })),
  }));
}

export function resolveChartLayout(chart: NormalizedChartElement): ResolvedChartLayout {
  const padding = 8;
  const titleHeight = chart.title === undefined ? 0 : 22;
  const legendItems = chart.chartType === "pie"
    ? chart.categories.map((label, pointIndex) => ({ label, color: chart.series[0]?.pointColors[pointIndex] ?? chart.series[0]?.color ?? "#000000", marker: false as const, pointIndex }))
    : chart.chartType === "line"
      ? chart.series.map((series, seriesIndex) => ({ label: series.name, color: series.color, marker: series.marker, seriesIndex }))
      : chart.series.map((series, seriesIndex) => ({ label: series.name, color: series.color, marker: false as const, seriesIndex }));
  const legendWidth = chart.legend.visible && chart.legend.position === "right" ? Math.min(100, chart.box.width * 0.3) : 0;
  const legendHeight = chart.legend.visible && chart.legend.position === "bottom" ? 20 : 0;
  const horizontal = chart.chartType === "bar" && chart.orientation === "horizontal";
  const categoryMargin = chart.chartType !== "pie" && chart.axes.category.visible && chart.axes.category.labels ? (horizontal ? 52 : 22) : padding;
  const valueMargin = chart.chartType !== "pie" && chart.axes.value.visible && chart.axes.value.labels ? (horizontal ? 22 : 38) : padding;
  const plotBox: Box = chart.chartType === "pie"
    ? { x: padding, y: padding + titleHeight, width: Math.max(1, chart.box.width - padding * 2 - legendWidth), height: Math.max(1, chart.box.height - padding * 2 - titleHeight - legendHeight) }
    : horizontal
      ? { x: padding + categoryMargin, y: padding + titleHeight, width: Math.max(1, chart.box.width - padding * 2 - categoryMargin - legendWidth), height: Math.max(1, chart.box.height - padding * 2 - titleHeight - valueMargin - legendHeight) }
      : { x: padding + valueMargin, y: padding + titleHeight, width: Math.max(1, chart.box.width - padding * 2 - valueMargin - legendWidth), height: Math.max(1, chart.box.height - padding * 2 - titleHeight - categoryMargin - legendHeight) };
  const scale = valueScale(chart, horizontal ? plotBox.width : plotBox.height);
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
  const bars = chart.chartType === "bar" ? barGeometry(chart, plotBox, scale) : [];
  const lines = chart.chartType === "line" ? lineGeometry(chart, plotBox, scale, categories) : [];
  let pie: ResolvedChartLayout["pie"];
  if (chart.chartType === "pie") {
    const series = chart.series[0];
    const total = series?.values.reduce((sum, value) => sum + Math.max(0, value), 0) ?? 0;
    const center = { x: plotBox.x + plotBox.width / 2, y: plotBox.y + plotBox.height / 2 };
    const radius = Math.min(plotBox.width, plotBox.height) * 0.425;
    let angle = -Math.PI / 2;
    const slices = chart.categories.map((_, pointIndex) => {
      const value = Math.max(0, series?.values[pointIndex] ?? 0);
      const startAngle = angle;
      angle += total <= EPSILON ? 0 : value / total * Math.PI * 2;
      return { pointIndex, startAngle, endAngle: angle, color: series?.pointColors[pointIndex] ?? series?.color ?? "#000000" };
    });
    pie = { center, radius, slices };
  }
  return {
    plotBox,
    ...(titleBox === undefined ? {} : { titleBox }),
    ...(legendBox === undefined ? {} : { legendBox }),
    valueScale: scale,
    categoryPositions: categories,
    categoryTickPositions: categoryTicks,
    bars,
    lines,
    pie,
    legendItems: resolvedLegendItems,
    ...(chart.chartType === "pie" ? {} : { categoryAxisPlacement: chart.chartType === "line" ? "onCategory" as const : "betweenCategories" as const }),
  };
}
