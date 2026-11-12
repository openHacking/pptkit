import type { NormalizedChartMarkerStyle, NormalizedStrokeStyle } from "@pptkit/core";
import { chartNumberReference, chartStringReference } from "./chart-reference.js";
import { scatterWorkbookRanges, type ScatterWorkbookSeries } from "./scatter-workbook.js";
import { paintXml, strokeXml } from "./xml.js";

/** Internal scatter-series XML input; integration into public export is pending. */
export interface ScatterExportSeries extends ScatterWorkbookSeries {
  line: false | NormalizedStrokeStyle;
  marker: false | NormalizedChartMarkerStyle;
}

function markerXml(marker: false | NormalizedChartMarkerStyle): string {
  if (marker === false) return `<c:marker><c:symbol val="none"/></c:marker>`;
  return `<c:marker><c:symbol val="${marker.shape}"/><c:size val="${Math.max(2, Math.min(72, Math.round(marker.size)))}"/><c:spPr>${paintXml(marker.fill)}${strokeXml(marker.stroke)}</c:spPr></c:marker>`;
}

function valuesXml(formula: string | undefined, values: readonly (number | null)[]): string {
  return formula === undefined ? `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="0"/></c:numLit>` : chartNumberReference(formula, values);
}

/** Emit series in source order; styles determine marker and line visibility. */
export function scatterSeriesXml(series: readonly ScatterExportSeries[]): string {
  const ranges = scatterWorkbookRanges(series);
  return series.map((item, index) => {
    const range = ranges[index]!, line = item.line === false ? `<a:ln><a:noFill/></a:ln>` : strokeXml(item.line);
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx>${chartStringReference(range.nameFormula, [item.name])}</c:tx><c:spPr>${line}</c:spPr>${markerXml(item.marker)}<c:xVal>${valuesXml(range.xFormula, item.xValues)}</c:xVal><c:yVal>${valuesXml(range.yFormula, item.yValues)}</c:yVal><c:smooth val="0"/></c:ser>`;
  }).join("");
}
