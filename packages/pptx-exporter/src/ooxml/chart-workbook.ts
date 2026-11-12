import type { LayoutChartElement } from "@pptkit/layout";
import { createZip } from "../archive/create-zip.js";
import { encodeUtf8 } from "../binary/encode.js";
import { chartStringXml } from "./chart-string.js";

export function spreadsheetColumn(index: number): string {
  let value = index;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + value % 26) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

/** Keep the same finite numeric spelling in chart caches and worksheet cells. */
export function chartNumberText(value: number): string {
  if (!Number.isFinite(value)) throw new Error("Chart data must contain only finite numeric values.");
  return Object.is(value, -0) ? "-0" : String(value);
}

function inlineStringCell(reference: string, value: string): string {
  return `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${chartStringXml(value)}</t></is></c>`;
}

function worksheetXml(chart: Exclude<LayoutChartElement, { chartType: "scatter" | "bubble" }>): string {
  const lastColumn = spreadsheetColumn(chart.series.length + 1);
  const lastRow = chart.categories.length + 1;
  const header = chart.series.map((series, index) => inlineStringCell(`${spreadsheetColumn(index + 2)}1`, series.name)).join("");
  const rows = chart.categories.map((category, categoryIndex) => {
    const row = categoryIndex + 2;
    const values = chart.series.map((series, seriesIndex) => {
      const reference = `${spreadsheetColumn(seriesIndex + 2)}${row}`, value = series.values[categoryIndex];
      return value === null ? `<c r="${reference}"/>` : `<c r="${reference}"><v>${chartNumberText(value!)}</v></c>`;
    }).join("");
    return `<row r="${row}">${inlineStringCell(`A${row}`, category)}${values}</row>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${lastColumn}${lastRow}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1">${header}</row>${rows}</sheetData></worksheet>`;
}

/** Package a chart worksheet without changing the existing workbook relationships. */
export function chartWorkbookFromWorksheet(sheetXml: string): Uint8Array {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
  return createZip([
    { name: "[Content_Types].xml", data: encodeUtf8(contentTypes) },
    { name: "_rels/.rels", data: encodeUtf8(rootRels) },
    { name: "xl/workbook.xml", data: encodeUtf8(workbook) },
    { name: "xl/_rels/workbook.xml.rels", data: encodeUtf8(workbookRels) },
    { name: "xl/worksheets/sheet1.xml", data: encodeUtf8(sheetXml) },
  ]);
}

export function chartWorkbook(chart: Exclude<LayoutChartElement, { chartType: "scatter" | "bubble" }>): Uint8Array {
  return chartWorkbookFromWorksheet(worksheetXml(chart));
}
