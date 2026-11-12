import { chartNumberText, chartWorkbookFromWorksheet, spreadsheetColumn } from "./chart-workbook.js";
import { chartStringXml } from "./chart-string.js";

/** Internal XY workbook shape; each series owns distinct X and Y columns. */
export interface ScatterWorkbookSeries {
  name: string;
  xValues: readonly (number | null)[];
  yValues: readonly (number | null)[];
}
export interface ScatterWorkbookRange {
  xColumn: string;
  yColumn: string;
  nameFormula: string;
  /** Empty series have no cell range and must use an empty numeric literal. */
  xFormula?: string;
  yFormula?: string;
}

const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;

/** Validate worksheet bounds before reading elements or allocating XML. */
export function scatterWorkbookRanges(series: readonly ScatterWorkbookSeries[]): ScatterWorkbookRange[] {
  if (!Array.isArray(series) || series.length === 0 || series.length > MAX_COLUMNS / 2) throw new RangeError("Scatter workbook needs between 1 and 8192 series.");
  return Array.from(series, (item, index) => {
    if (!item || typeof item.name !== "string" || item.name.length === 0 || !Array.isArray(item.xValues) || !Array.isArray(item.yValues)) throw new RangeError("Scatter workbook series requires a name and paired arrays.");
    if (item.xValues.length !== item.yValues.length || item.xValues.length > MAX_ROWS - 1) throw new RangeError("Scatter workbook series has incompatible or oversized point arrays.");
    for (let i = 0; i < item.xValues.length; i++) {
      for (const value of [item.xValues[i], item.yValues[i]]) if (value !== null && (typeof value !== "number" || !Number.isFinite(value))) throw new RangeError("Scatter workbook coordinates must be finite numbers or explicit null.");
    }
    const xColumn = spreadsheetColumn(index * 2 + 1), yColumn = spreadsheetColumn(index * 2 + 2);
    const nameFormula = `Sheet1!$${yColumn}$1`, lastRow = item.xValues.length + 1;
    return { xColumn, yColumn, nameFormula, ...(lastRow === 1 ? {} : { xFormula: `Sheet1!$${xColumn}$2:$${xColumn}$${lastRow}`, yFormula: `Sheet1!$${yColumn}$2:$${yColumn}$${lastRow}` }) };
  });
}

function cell(reference: string, value: number | null): string {
  return value === null ? `<c r="${reference}"/>` : `<c r="${reference}"><v>${chartNumberText(value)}</v></c>`;
}

export function scatterWorkbook(series: readonly ScatterWorkbookSeries[]): Uint8Array {
  const ranges = scatterWorkbookRanges(series), lastColumn = ranges.at(-1)!.yColumn;
  const count = series.reduce((max, item) => Math.max(max, item.xValues.length), 0);
  const headers = series.map((item, i) => `<c r="${ranges[i]!.yColumn}1" t="inlineStr"><is><t xml:space="preserve">${chartStringXml(item.name)}</t></is></c>`).join("");
  const rows: string[] = [];
  // Drop exhausted series so uneven lengths do not scan all series for every row.
  let active = series.flatMap((item, index) => item.xValues.length ? [index] : []);
  for (let pointIndex = 0; pointIndex < count; pointIndex++) {
    const row = pointIndex + 2, cells: string[] = [], next: number[] = [];
    for (const i of active) {
      const item = series[i]!, range = ranges[i]!;
      cells.push(cell(`${range.xColumn}${row}`, item.xValues[pointIndex]!), cell(`${range.yColumn}${row}`, item.yValues[pointIndex]!));
      if (pointIndex + 1 < item.xValues.length) next.push(i);
    }
    active = next;
    rows.push(`<row r="${row}">${cells.join("")}</row>`);
  }
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${lastColumn}${count + 1}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1">${headers}</row>${rows.join("")}</sheetData></worksheet>`;
  return chartWorkbookFromWorksheet(sheet);
}
