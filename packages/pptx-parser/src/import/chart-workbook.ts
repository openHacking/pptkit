import { readOpcSource, snapshotOptions } from "../opc/read-package.js";
import { attribute, children, integer, invalidStructure, singleChild } from "../presentation/structure-xml.js";
import { PackageInspectionError } from "../types/archive.js";
import { PptxImportError, type PptxSessionLimits } from "../types/import.js";
import type { OpcInspectionOptions, OpcRelationshipInfo } from "../types/opc.js";
import type { XmlElement } from "../xml/read-xml.js";
import { MCE } from "../compatibility/configuration.js";
import { chartString } from "./chart-string.js";

export interface ChartWorkbookCell { worksheetPart?: string; value?: string; numeric: boolean; formula: boolean; blank?: boolean; source?: { part: string; elementIndex: number; namespaceUri: string; lexical: string }; stringSource?: { part: string; elementIndex: number; namespaceUri: string; lexical: string; preserveSpace: boolean; space?: "default" | "preserve" } }
type Cell = ChartWorkbookCell;
export interface ChartWorkbook {
  readonly sheets: ReadonlyMap<string, ReadonlyMap<string, Cell>>;
  readonly reason?: string;
  readonly editBlockReason?: string;
  readonly captureBindings?: boolean;
}
const SHEET_NS = ["http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main"];
const OFFICE_REL = ["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"];
const sheetType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const worksheetType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

/** A blank is a plain numeric cell, not an error, formula or opaque value. */
function isPlainNumericBlank(cell: XmlElement, type: string): boolean {
  if (type !== "n") return false;
  const attributes = cell.attributes.every(attribute => attribute.namespaceUri === "" && ["r", "s", "t"].includes(attribute.local)
    && (attribute.local !== "s" || (/^\d+$/u.test(attribute.value) && Number(attribute.value) <= 4294967295)));
  const content = cell.children.every(node => node.kind === "comment" || (node.kind === "text" && /^[\t\r\n ]*$/u.test(node.value)));
  return attributes && content;
}

/** Nested workbooks share invocation budgets and are inspected once per part. */
export function chartWorkbooks(options: OpcInspectionOptions, limits: Required<PptxSessionLimits>, countPoint: () => void, checkpoint: () => Promise<void>, captureBindings = false) {
  const captured = snapshotOptions(options), workbooks = new Map<string, ChartWorkbook>();
  let totalExpanded = 0;
  return async (part: string, bytes: Uint8Array): Promise<ChartWorkbook> => {
    const previous = workbooks.get(part);
    if (previous) return previous;
    const remaining = limits.maxTotalExpandedChartWorkbookBytes - totalExpanded;
    if (remaining <= 0) throw new PptxImportError("import-limit-exceeded", "Embedded chart workbooks exceed maxTotalExpandedChartWorkbookBytes.", { part });
    try {
      const source = await readOpcSource(bytes, {
        limits: captured.limits, xmlLimits: captured.xmlLimits,
        zipLimits: { ...captured.zipLimits, maxExpandedBytes: Math.min(captured.zipLimits.maxExpandedBytes, limits.maxExpandedChartWorkbookBytes, remaining) },
        ...(captured.signal ? { signal: captured.signal } : {}),
      }, true);
      totalExpanded += source.inspection.archive.expandedBytes;
      let location = "/";
      const fail = (message: string): never => invalidStructure(message, location);
      const entries = new Map(source.inspection.parts.map(p => [p.name, p]));
      const internal = (relation: OpcRelationshipInfo | undefined, type: string) => {
        if (!relation || !OFFICE_REL.some(r => relation.type === `${r}/${type}`) || relation.targetMode !== "Internal" || !relation.targetPart || relation.fragment !== undefined) fail(`Expected an internal workbook ${type} relationship.`);
        return relation!.targetPart!;
      };
      const rootRels = source.inspection.relationships.filter(r => r.source === "/" && OFFICE_REL.some(ns => r.type === `${ns}/officeDocument`));
      if (rootRels.length !== 1) fail("Chart workbook requires exactly one officeDocument entry.");
      const mainPart = internal(rootRels[0], "officeDocument"); location = mainPart;
      const main = source.xml.get(mainPart);
      if (!main || main.local !== "workbook" || !SHEET_NS.includes(main.namespaceUri) || entries.get(mainPart)?.contentType !== sheetType) fail("Embedded chart package is not a supported XLSX workbook.");
      const s = main!.namespaceUri;
      const at = (node: XmlElement, name: string, required = false) => singleChild(node, s, name, location, required);
      const text = (node: XmlElement): string => {
        if (children(node).length) fail("Workbook scalar text contains child elements.");
        return node.children.map(n => n.kind === "text" || n.kind === "cdata" ? n.value : "").join("");
      };
      const inline = (node: XmlElement): string => children(node).map(n => n.namespaceUri !== s ? "" : n.local === "t" ? chartString(text(n)) : n.local === "r" ? children(n).filter(t => t.namespaceUri === s && t.local === "t").map(t => chartString(text(t))).join("") : "").join("");
      // Workbook MCE is a separate spreadsheet consumer profile. Do not
      // interpret raw unselected branches as if they were effective cells.
      if ([...source.xml.values()].some(root => root.attributes.some(a => a.namespaceUri === MCE) || root.children.some(n => n.kind === "element" && n.namespaceUri === MCE)) ||
          source.inspection.parts.some(p => p.xml?.namespaceUris.includes(MCE))) {
        const result = { sheets: new Map(), reason: "Spreadsheet MCE requires a separate workbook interpretation profile; caches remain the available chart snapshot." };
        workbooks.set(part, result); return result;
      }
      const strings: string[] = [];
      const sharedRels = source.inspection.relationships.filter(r => r.source === mainPart && OFFICE_REL.some(ns => r.type === `${ns}/sharedStrings`));
      if (sharedRels.length > 1) fail("Duplicate workbook shared-string relationships.");
      if (sharedRels.length) {
        const shared = source.xml.get(internal(sharedRels[0], "sharedStrings"));
        if (!shared || shared.local !== "sst" || shared.namespaceUri !== s || entries.get(internal(sharedRels[0], "sharedStrings"))?.contentType !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml") fail("Invalid shared-string root.");
        for (const item of children(shared!).filter(n => n.namespaceUri === s && n.local === "si")) { await checkpoint(); countPoint(); strings.push(inline(item)); }
      }
      let editBlockReason: string | undefined;
      if (captureBindings) {
      if (source.inspection.relationships.some(r => r.targetMode === "External" || /vbaProject|digital-signature|calcChain|externalLink/i.test(r.type)) || source.inspection.parts.some(p => /vbaProject|digital-signature|drawingml.chart|pivot|query|connections/i.test(p.contentType))) editBlockReason = "Workbook signatures/macros/external dependencies are outside numeric preservation edits.";
      for (const tree of source.xml.values()) {
        const pending = [tree];
        while (pending.length) { await checkpoint(); const n = pending.pop()!; if (SHEET_NS.includes(n.namespaceUri) && (n.local === "f" || (/formula|calcChain/i.test(n.local) || n.local === "definedName"))) editBlockReason = "Workbook formulas/named calculations require a dependency-aware editing profile."; pending.push(...children(n)); }
      }
      }
      const sheets = new Map<string, ReadonlyMap<string, Cell>>();
      const sheetList = at(main!, "sheets", true)!;
      for (const sheet of children(sheetList).filter(n => n.namespaceUri === s && n.local === "sheet")) {
        await checkpoint();
        const name = attribute(sheet, "name");
        if (!name || sheets.has(name)) fail("Missing or duplicate workbook sheet name.");
        const ids = sheet.attributes.filter(a => a.local === "id" && OFFICE_REL.includes(a.namespaceUri));
        if (ids.length !== 1) fail("Workbook sheet requires a relationship identity.");
        const rel = source.inspection.relationships.find(r => r.source === mainPart && r.id === ids[0]!.value);
        const sheetPart = internal(rel, "worksheet"), root = source.xml.get(sheetPart);
        location = sheetPart;
        if (!root || root.local !== "worksheet" || root.namespaceUri !== s || entries.get(sheetPart)?.contentType !== worksheetType) fail("Invalid workbook worksheet target.");
        const rawIndices = new Map<XmlElement, number>(), pending = [root!]; let ordinal = 0;
        while (captureBindings && pending.length) { await checkpoint(); const n = pending.pop()!; rawIndices.set(n, ++ordinal); const nested = children(n); for (let i = nested.length - 1; i >= 0; i--) pending.push(nested[i]!); }
        const data = at(root!, "sheetData", true)!, cells = new Map<string, Cell>();
        for (const row of children(data).filter(n => n.namespaceUri === s && n.local === "row")) {
          for (const cell of children(row).filter(n => n.namespaceUri === s && n.local === "c")) {
            await checkpoint(); countPoint();
            const ref = attribute(cell, "r");
            if (!ref || !/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(ref) || cells.has(ref) || Number(ref.replace(/^[A-Z]+/, "")) > 1048576 || [...ref.replace(/\d+$/, "")].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) > 16384) fail("Workbook cell reference is missing, duplicated or unsupported.");
            const type = attribute(cell, "t") ?? "n", value = at(cell, "v"), formula = at(cell, "f") !== undefined;
            let raw: string | undefined;
            const inlineContainer = type === "inlineStr" ? at(cell, "is") : undefined;
            if (type === "inlineStr") { raw = inlineContainer ? inline(inlineContainer) : ""; }
            else if (type === "s" && value) {
              const index = integer(text(value), 0, strings.length - 1, "shared-string index", sheetPart);
              raw = strings[index];
            } else if (["n", "str"].includes(type) && value) raw = type === "str" ? chartString(text(value)) : text(value);
            const scalar = captureBindings && value && value.children.every(n => n.kind === "text" || n.kind === "cdata");
            const stringNodes = inlineContainer ? children(inlineContainer) : [];
            const stringNode = stringNodes.length === 1 && stringNodes[0]!.namespaceUri === s && stringNodes[0]!.local === "t" ? stringNodes[0] : undefined;
            const simpleString = captureBindings && stringNode && !formula && cell.attributes.every(a => a.namespaceUri === "" && ["r", "s", "t"].includes(a.local) && (a.local !== "s" || (/^\d+$/u.test(a.value) && Number(a.value) <= 4294967295))) &&
              cell.children.every(n => n === inlineContainer || (n.kind === "text" && /^[\t\r\n ]*$/u.test(n.value))) &&
              inlineContainer!.attributes.length === 0 && inlineContainer!.children.every(n => n === stringNode || (n.kind === "text" && /^[\t\r\n ]*$/u.test(n.value))) &&
              stringNode.attributes.every(a => a.namespaceUri === "http://www.w3.org/XML/1998/namespace" && a.local === "space" && (a.value === "preserve" || a.value === "default")) &&
              stringNode.children.every(n => n.kind === "text" || n.kind === "cdata");
            cells.set(ref!, { ...(captureBindings ? { worksheetPart: sheetPart } : {}), ...(raw === undefined ? {} : { value: raw }), numeric: type === "n", formula, ...(type === "n" && cell.attributes.every(a => a.namespaceUri === "" && ["r", "s", "t"].includes(a.local) && (a.local !== "s" || (/^\d+$/u.test(a.value) && Number(a.value) <= 4294967295))) && cell.children.every(n => n.kind === "comment" || (n.kind === "text" && /^[\t\r\n ]*$/u.test(n.value))) ? { blank: true } : {}), ...(scalar ? { source: { part: sheetPart, elementIndex: rawIndices.get(value!)!, namespaceUri: s, lexical: text(value!) } } : {}), ...(simpleString ? { stringSource: { part: sheetPart, elementIndex: rawIndices.get(stringNode)!, namespaceUri: s, lexical: text(stringNode), preserveSpace: stringNode.attributes.some(a => a.namespaceUri === "http://www.w3.org/XML/1998/namespace" && a.local === "space" && a.value === "preserve"), ...(stringNode.attributes.find(a => a.namespaceUri === "http://www.w3.org/XML/1998/namespace" && a.local === "space") ? { space: stringNode.attributes.find(a => a.namespaceUri === "http://www.w3.org/XML/1998/namespace" && a.local === "space")!.value as "default" | "preserve" } : {}) } } : {}) });
          }
        }
        sheets.set(name!, cells);
      }
      const result = { sheets, ...(captureBindings ? { captureBindings: true } : {}), ...(editBlockReason ? { editBlockReason } : {}) }; workbooks.set(part, result); return result;
    } catch (error) {
      if (error instanceof PackageInspectionError) throw new PackageInspectionError(error.code, error.message, `${part}!/${error.entry?.replace(/^\//, "") ?? ""}`, error.location);
      throw error;
    }
  };
}

/** Only A1 scalar and one-dimensional local ranges; no expression evaluation. */
export async function workbookRange(workbook: ChartWorkbook, formula: string, numeric: boolean, countPoint: () => void, checkpoint: () => Promise<void>, expected?: number, allowBlanks = false, captureSparseSources = false): Promise<{ values?: string[]; blankValues?: (string | null)[]; reason?: string; cachedFormula?: boolean; cells?: readonly { reference: string; cell: Cell }[]; sparseCells?: ReadonlyMap<number, { reference: string; cell: Cell }> }> {
  const match = /^(?:'((?:[^']|'')+)'|([^'!\[\]]+))!\$?([A-Z]{1,3})\$?([1-9]\d{0,6})(?::\$?([A-Z]{1,3})\$?([1-9]\d{0,6}))?$/.exec(formula.trim());
  if (!match) return { reason: "Chart formula is not a supported local A1 reference; names, expressions and external links are not evaluated." };
  const name = match[1]?.replace(/''/g, "'") ?? match[2]!, cells = workbook.sheets.get(name);
  if (!cells) return { reason: workbook.reason ?? "Chart formula sheet is absent from the inspected workbook." };
  const column = (value: string) => [...value].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const x1 = column(match[3]!), y1 = Number(match[4]), x2 = column(match[5] ?? match[3]!), y2 = Number(match[6] ?? match[4]);
  if (x1 > 16384 || x2 > 16384 || y1 > 1048576 || y2 > 1048576 || x2 < x1 || y2 < y1 || (x1 !== x2 && y1 !== y2)) return { reason: "Chart formula is outside a forward one-dimensional worksheet range." };
  const length = Math.max(x2 - x1, y2 - y1) + 1;
  // A blank-aware range is bounded by an already inspected category array;
  // otherwise work is bounded by actual workbook cells, never the formula alone.
  const boundedBlanks = numeric && allowBlanks && expected !== undefined && expected === length;
  if ((!boundedBlanks && length > cells.size) || (expected !== undefined && expected !== length)) return { reason: "Workbook range length disagrees with complete cached points or available cells." };
  const columnName = (index: number) => { let text = ""; for (let n = index; n; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + (n - 1) % 26) + text; return text; };
  const values: (string | null)[] = [], selected: { reference: string; cell: Cell; index: number }[] = []; let cachedFormula = false;
  for (let i = 0; i < length; i++) {
    await checkpoint(); countPoint();
    const reference = `${columnName(x1 === x2 ? x1 : x1 + i)}${y1 === y2 ? y1 : y1 + i}`;
    const cell = cells.get(reference);
    if (boundedBlanks && (!cell || cell.blank === true)) { values.push(null); continue; }
    if (!cell || cell.value === undefined || (numeric && !cell.numeric)) return { reason: "Workbook range contains missing, error or incompatible cell values." };
    cachedFormula ||= cell.formula; values.push(cell.value); if (workbook.captureBindings) selected.push({ reference, cell, index: i });
  }
  if (values.includes(null)) {
    const sparseCells = workbook.captureBindings && captureSparseSources ? new Map<number, { reference: string; cell: Cell }>() : undefined;
    if (sparseCells) for (const item of selected) { await checkpoint(); sparseCells.set(item.index, item); }
    return { blankValues: values, cachedFormula, ...(sparseCells ? { sparseCells } : {}) };
  }
  return { values: values as string[], cachedFormula, ...(workbook.captureBindings ? { cells: selected } : {}) };
}
