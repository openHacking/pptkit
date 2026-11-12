import type { Box, ChartBlankDisplay, Point } from "@pptkit/core";
import type { ChartValueScale } from "../types/layout.js";

/** Internal geometry input for XY scatter charts. */
export interface ScatterGeometrySeries {
  xValues: readonly (number | null)[];
  yValues: readonly (number | null)[];
}

export interface ScatterGeometryPoint extends Point {
  /** Index in the original series, including gaps. */
  sourceIndex: number;
  /** True if either coordinate is an explicit blank displayed at zero. */
  blank?: true;
}

export interface ScatterGeometrySeriesLayout {
  seriesIndex: number;
  points: ScatterGeometryPoint[];
  segments: ScatterGeometryPoint[][];
}

function numericRatio(value: number, scale: ChartValueScale): number {
  if (!Number.isFinite(value) || !Number.isFinite(scale.min) || !Number.isFinite(scale.max) || scale.max <= scale.min) {
    throw new RangeError("Scatter values and ordered axis bounds must be finite.");
  }
  const span = scale.max - scale.min;
  const result = Number.isFinite(span)
    ? (value - scale.min) / span
    : (value / 2 - scale.min / 2) / (scale.max / 2 - scale.min / 2);
  if (!Number.isFinite(result)) throw new RangeError("Scatter point is outside the representable geometry range.");
  return result;
}

/** Map independent X/Y series through two numeric scales without changing order. */
export function scatterGeometry(
  series: readonly ScatterGeometrySeries[], plot: Box, xScale: ChartValueScale, yScale: ChartValueScale, mode: ChartBlankDisplay,
): ScatterGeometrySeriesLayout[] {
  if (![plot.x, plot.y, plot.width, plot.height].every(Number.isFinite) || plot.width <= 0 || plot.height <= 0) {
    throw new RangeError("Scatter plot box must be finite with positive dimensions.");
  }
  if (!["gap", "zero", "span"].includes(mode)) throw new RangeError("Unsupported scatter blank display.");
  // Validate both scales even when every point is absent or displayed as a gap.
  numericRatio(xScale.min, xScale);
  numericRatio(yScale.min, yScale);
  return Array.from(series, (item, seriesIndex) => {
    if (item === undefined) throw new RangeError("Scatter series entries must be present.");
    if (item.xValues.length !== item.yValues.length) throw new RangeError("Scatter X and Y arrays must have equal lengths.");
    const points: ScatterGeometryPoint[] = [], segments: ScatterGeometryPoint[][] = [];
    let segment: ScatterGeometryPoint[] = [];
    for (let index = 0; index < item.xValues.length; index++) {
      const x = item.xValues[index], y = item.yValues[index];
      if ((x !== null && (x === undefined || !Number.isFinite(x))) || (y !== null && (y === undefined || !Number.isFinite(y)))) {
        throw new RangeError("Scatter coordinates must be finite numbers or explicit null.");
      }
      if ((x === null || y === null) && mode !== "zero") {
        if (mode === "gap") { if (segment.length) segments.push(segment); segment = []; }
        continue;
      }
      const point: ScatterGeometryPoint = {
        x: plot.x + numericRatio(x ?? 0, xScale) * plot.width,
        y: plot.y + plot.height - numericRatio(y ?? 0, yScale) * plot.height,
        sourceIndex: index,
        ...(x === null || y === null ? { blank: true as const } : {}),
      };
      if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new RangeError("Scatter point is outside the representable geometry range.");
      points.push(point); segment.push({ ...point });
    }
    if (segment.length) segments.push(segment);
    return { seriesIndex, points, segments };
  });
}
