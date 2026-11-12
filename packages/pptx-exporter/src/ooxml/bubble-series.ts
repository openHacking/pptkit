import { chartNumberText, chartWorkbookFromWorksheet, spreadsheetColumn } from "./chart-workbook.js";
import { chartNumberReference, chartStringReference } from "./chart-reference.js";
import { chartStringXml } from "./chart-string.js";
import { paintXml } from "./xml.js";

export interface BubbleExportSeries {
  name: string;
  xValues: readonly number[];
  yValues: readonly number[];
  sizes: readonly number[];
  color: string;
}

const MAX_POINTS = 1_048_575;
const MAX_SERIES = Math.floor(16_384 / 3);

function columns(series: readonly BubbleExportSeries[]) {
  if (!Array.isArray(series) || !series.length || series.length > MAX_SERIES) {
    throw new RangeError(`Bubble workbook requires between 1 and ${MAX_SERIES} series.`);
  }
  return Array.from(series, (item, index) => {
    if (!item || typeof item.name !== "string" || !item.name || !Array.isArray(item.xValues) || !Array.isArray(item.yValues) || !Array.isArray(item.sizes)) {
      throw new RangeError("Bubble series requires a name and X, Y, and size arrays.");
    }
    const count = item.xValues.length;
    if (!count || count > MAX_POINTS || item.yValues.length !== count || item.sizes.length !== count) {
      throw new RangeError("Bubble series arrays must be nonempty, equal-length, and within the worksheet row limit.");
    }
    for (let point = 0; point < count; point++) {
      if (!Number.isFinite(item.xValues[point]) || !Number.isFinite(item.yValues[point]) ||
        !Number.isFinite(item.sizes[point]) || item.sizes[point]! <= 0) {
        throw new RangeError("Bubble coordinates must be finite and sizes must be finite and positive.");
      }
    }
    const [x, y, size] = [1, 2, 3].map(offset => spreadsheetColumn(index * 3 + offset)) as [string, string, string];
    const values = (column: string) => `Sheet1!$${column}$2:$${column}$${count + 1}`;
    return { x, y, size, name: `Sheet1!$${y}$1`, xFormula: values(x), yFormula: values(y), sizeFormula: values(size) };
  });
}

/** Build a matching worksheet and series fragment from the same dense data. */
export function bubbleSeriesParts(series: readonly BubbleExportSeries[]): { xml: string; workbook: Uint8Array } {
  const ranges = columns(series);
  const xml = series.map((item, index) => {
    const range = ranges[index]!;
    const paint = paintXml({ type: "solid", color: item.color, opacity: 1 });
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx>${chartStringReference(range.name, [item.name])}</c:tx><c:spPr>${paint}</c:spPr><c:xVal>${chartNumberReference(range.xFormula, item.xValues)}</c:xVal><c:yVal>${chartNumberReference(range.yFormula, item.yValues)}</c:yVal><c:bubbleSize>${chartNumberReference(range.sizeFormula, item.sizes)}</c:bubbleSize><c:bubble3D val="0"/></c:ser>`;
  }).join("");
  const count = series.reduce((max, item) => Math.max(max, item.xValues.length), 0);
  const cell = (column: string, row: number, value: number) => `<c r="${column}${row}"><v>${chartNumberText(value)}</v></c>`;
  const text = (column: string, value: string) => `<c r="${column}1" t="inlineStr"><is><t xml:space="preserve">${chartStringXml(value)}</t></is></c>`;
  const headers = series.map((item, index) => {
    const { x, y, size } = ranges[index]!;
    return text(x, "X") + text(y, item.name) + text(size, "Size");
  }).join("");
  const rows: string[] = [];
  let active = series.map((_, index) => index);
  for (let point = 0; point < count; point++) {
    const row = point + 2, cells: string[] = [], next: number[] = [];
    for (const index of active) {
      const item = series[index]!, range = ranges[index]!;
      cells.push(cell(range.x, row, item.xValues[point]!), cell(range.y, row, item.yValues[point]!),
        cell(range.size, row, item.sizes[point]!));
      if (point + 1 < item.xValues.length) next.push(index);
    }
    active = next;
    rows.push(`<row r="${row}">${cells.join("")}</row>`);
  }
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${ranges.at(-1)!.size}${count + 1}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1">${headers}</row>${rows.join("")}</sheetData></worksheet>`;
  return { xml, workbook: chartWorkbookFromWorksheet(sheet) };
}
