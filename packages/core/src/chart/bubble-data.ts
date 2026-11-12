import type { BubbleChartSeriesInput } from "../types/element.js";

export function bubbleDataIssue(series: readonly BubbleChartSeriesInput[]): string | undefined {
  if (!Array.isArray(series) || !series.length || series.length > 5461) return "requires 1–5461 series";
  for (const item of series) {
    if (!item || typeof item.name !== "string" || !item.name || !Array.isArray(item.xValues) || !Array.isArray(item.yValues) || !Array.isArray(item.sizes)) {
      return "each series requires a name and X, Y, and size arrays";
    }
    const count = item.xValues.length;
    if (!count || count > 1_048_575 || item.yValues.length !== count || item.sizes.length !== count) {
      return "series arrays must be nonempty and equal-length, with at most 1,048,575 points";
    }
    for (let point = 0; point < count; point++) {
      if (!Number.isFinite(item.xValues[point]) || !Number.isFinite(item.yValues[point]) ||
        !Number.isFinite(item.sizes[point]) || item.sizes[point]! <= 0) return "coordinates must be finite and sizes must be finite and positive";
    }
  }
  return undefined;
}
