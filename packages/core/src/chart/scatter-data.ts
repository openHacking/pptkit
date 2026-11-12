import type { ChartBlankDisplay, ChartDataValue } from "../types/element.js";

/** Internal data contract for XY charts; each series owns both arrays. */
export interface ScatterSeriesData {
  xValues: readonly ChartDataValue[];
  yValues: readonly ChartDataValue[];
}

export interface ScatterDataIssue {
  code: "scatter-empty-series" | "scatter-missing-series" | "scatter-invalid-values" | "scatter-values-length" | "scatter-coordinate";
  seriesIndex?: number;
  coordinate?: "xValues" | "yValues";
  pointIndex?: number;
}

/** Validate paired data without treating missing array entries as explicit blanks. */
export function inspectScatterData(series: readonly ScatterSeriesData[]): ScatterDataIssue[] {
  const issues: ScatterDataIssue[] = [];
  if (series.length === 0) issues.push({ code: "scatter-empty-series" });
  for (let seriesIndex = 0; seriesIndex < series.length; seriesIndex++) {
    const item = series[seriesIndex];
    if (item === undefined || item === null || typeof item !== "object") { issues.push({ code: "scatter-missing-series", seriesIndex }); continue; }
    const { xValues, yValues } = item;
    if (!Array.isArray(xValues) || !Array.isArray(yValues)) {
      if (!Array.isArray(xValues)) issues.push({ code: "scatter-invalid-values", seriesIndex, coordinate: "xValues" });
      if (!Array.isArray(yValues)) issues.push({ code: "scatter-invalid-values", seriesIndex, coordinate: "yValues" });
      continue;
    }
    if (xValues.length !== yValues.length) issues.push({ code: "scatter-values-length", seriesIndex });
    for (const [coordinate, values] of [["xValues", xValues], ["yValues", yValues]] as const) {
      for (let pointIndex = 0; pointIndex < values.length; pointIndex++) {
        const value = values[pointIndex];
        if (value !== null && (typeof value !== "number" || !Number.isFinite(value))) issues.push({ code: "scatter-coordinate", seriesIndex, coordinate, pointIndex });
      }
    }
  }
  return issues;
}

/** Snapshot validated inputs and resolve the adaptive blank-display default. */
export function snapshotScatterData(series: readonly ScatterSeriesData[], display?: ChartBlankDisplay): { series: ScatterSeriesData[]; blankDisplay: ChartBlankDisplay } {
  if (inspectScatterData(series).length) throw new RangeError("Scatter data must contain series with paired arrays of finite values or explicit null blanks.");
  if (display !== undefined && !["gap", "zero", "span"].includes(display)) throw new RangeError("Unsupported scatter blank display.");
  const snapshot = series.map(({ xValues, yValues }) => ({ xValues: [...xValues], yValues: [...yValues] }));
  return { series: snapshot, blankDisplay: display ?? (snapshot.some(item => item.xValues.includes(null) || item.yValues.includes(null)) ? "gap" : "zero") };
}
