import type { readOpcSource } from "../opc/read-package.js";
import { readOpcSource as readSource } from "../opc/read-package.js";
import { captureXmlTextSource, type XmlTextEdit } from "../preservation/xml-text-source.js";
import { replaceArchiveEntries } from "../preservation/replace-archive-entries.js";
import { comparePackages } from "../comparison/compare-packages.js";
import type { PptxInspectionOptions } from "../types/structure.js";
import { PptxImportError, type PptxSessionLimits } from "../types/import.js";
import type { ChartValueProposal } from "./chart-value-bindings.js";
import type { ChartLabelClosure } from "./chart-label-closure.js";
import { chartString } from "./chart-string.js";
import { editXmlTextAttributes } from "../preservation/replace-xml-relationship-id.js";
import { planChartSeriesOrder, type ChartSeriesOrderEdit, type ChartSeriesOrderSource } from "./chart-series-order.js";
import { moveChartSeries } from "./chart-series-move.js";
import { captureColumnLegendSources } from "./chart-legend-source.js";
import { planChartLegendOrder } from "./chart-legend-order.js";

/** Private staging input; a closure alone never authorizes a session save. */
export interface ChartLabelProposal { readonly closure: ChartLabelClosure; readonly value: string }
/** Internal candidate only; callers must separately prove package-level authority. */
export interface ChartSeriesOrderProposal { readonly part: string; readonly sources: readonly ChartSeriesOrderSource[]; readonly requested: readonly number[]; readonly complete: boolean; readonly researchColumnLegend?: boolean }

function encodeChartString(value: string): string {
  return value.replace(/_(?=x[0-9a-fA-F]{4}_)|[\u0000-\u0008\u000b-\u000d\u000e-\u001f\uD800-\uDFFF\ufffe\uffff]/g,
    character => character === "_" ? "_x005F_" : `_x${character.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`);
}

/** Stages both layers. Nothing publishes or advances a session baseline here. */
export async function prepareChartValues(original: Awaited<ReturnType<typeof readOpcSource>>, proposals: readonly ChartValueProposal[], options: PptxInspectionOptions,
  limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>, priorBytes: number, labels: readonly ChartLabelProposal[] = [], seriesOrders: readonly ChartSeriesOrderProposal[] = []) {
  const charts = new Map<string, XmlTextEdit[]>(), books = new Map<string, Map<string, XmlTextEdit[]>>();
  const orderEdits = new Map<string, readonly ChartSeriesOrderEdit[]>();
  const ordersByPart = new Map<string, ChartSeriesOrderProposal>(), movedElementIndexes = new Map<string, ReadonlyMap<number, number>>(), legendSourceProofs = new Map<string, Awaited<ReturnType<typeof captureColumnLegendSources>>>();
  const spaceEdits = new Map<string, Map<string, { elementIndex: number; elementNamespace: string; elementLocal: string; attributeNamespace: string; attributeLocal: string; expected?: string; replacement: string }[]>>();
  const append = (parts: Map<string, XmlTextEdit[]>, part: string, edit: XmlTextEdit) => { const list = parts.get(part) ?? []; list.push(edit); parts.set(part, list); };
  const targets = new Set<string>();
  for (const { binding: b, value } of proposals) {
    await checkpoint();
    if (b.status !== "bound" || !b.workbook || !b.sourceElementIndex || b.cacheLexical === undefined || b.workbookLexical === undefined || !b.cacheNamespace || !b.workbookNamespace || !Number.isFinite(value)) throw new PptxImportError("save-candidate-invalid", "Chart numeric source binding is incomplete.", b.source);
    const key = JSON.stringify([b.workbook.part, b.workbook.worksheetPart, b.workbook.sourceElementIndex]);
    if (targets.has(key)) throw new PptxImportError("save-candidate-invalid", "Chart numeric cell has competing write proposals.", b.source); targets.add(key);
    const text = Object.is(value, -0) ? "-0" : String(value);
    append(charts, b.source.part, { elementIndex: b.sourceElementIndex, namespaceUri: b.cacheNamespace, local: "v", expectedText: b.cacheLexical, text });
    const sheets = books.get(b.workbook.part) ?? new Map<string, XmlTextEdit[]>(); books.set(b.workbook.part, sheets);
    append(sheets, b.workbook.worksheetPart, { elementIndex: b.workbook.sourceElementIndex, namespaceUri: b.workbookNamespace, local: "v", expectedText: b.workbookLexical, text });
  }
  const cacheTargets = new Set<string>();
  for (const { closure, value } of labels) {
    await checkpoint();
    if (closure.status !== "bound" || !closure.uses.length || typeof value !== "string" || !value.length) throw new PptxImportError("save-candidate-invalid", "Chart label source binding is incomplete.");
    const text = encodeChartString(value), first = closure.uses[0]!;
    const cellKey = JSON.stringify([first.cell.workbookPart, first.cell.worksheetPart, first.cell.elementIndex]);
    if (targets.has(cellKey)) throw new PptxImportError("save-candidate-invalid", "Chart workbook cell has competing write proposals.");
    targets.add(cellKey);
    for (const use of closure.uses) {
      await checkpoint();
      const cell = use.cell, cache = use.cache, key = JSON.stringify([cache.part, cache.elementIndex]);
      if (use.logicalKey !== closure.logicalKey || JSON.stringify([cell.workbookPart, cell.worksheetPart, cell.elementIndex]) !== cellKey ||
          cell.reference !== first.cell.reference || cell.namespaceUri !== first.cell.namespaceUri || cell.lexical !== first.cell.lexical || cell.preserveSpace !== first.cell.preserveSpace || cell.space !== first.cell.space ||
          !cache.namespaceUri || !cell.namespaceUri || chartString(cache.lexical) !== use.projected || chartString(cell.lexical) !== use.projected || use.projected !== first.projected || cacheTargets.has(key)) {
        throw new PptxImportError("save-candidate-invalid", "Chart label sources conflict or no longer agree.");
      }
      cacheTargets.add(key);
      append(charts, cache.part, { elementIndex: cache.elementIndex, namespaceUri: cache.namespaceUri, local: "v", expectedText: cache.lexical, text });
    }
    if (!first.cell.preserveSpace && /^[\t\n\r ]|[\t\n\r ]$/u.test(value)) {
      const book = spaceEdits.get(first.cell.workbookPart) ?? new Map(); spaceEdits.set(first.cell.workbookPart, book);
      const edits = book.get(first.cell.worksheetPart) ?? []; book.set(first.cell.worksheetPart, edits);
      edits.push({ elementIndex: first.cell.elementIndex, elementNamespace: first.cell.namespaceUri, elementLocal: "t", attributeNamespace: "http://www.w3.org/XML/1998/namespace", attributeLocal: "space", ...(first.cell.space ? { expected: first.cell.space } : {}), replacement: "preserve" });
    }
    const sheets = books.get(first.cell.workbookPart) ?? new Map<string, XmlTextEdit[]>(); books.set(first.cell.workbookPart, sheets);
    append(sheets, first.cell.worksheetPart, { elementIndex: first.cell.elementIndex, namespaceUri: first.cell.namespaceUri, local: "t", expectedText: first.cell.lexical, text });
  }
  let preparedBytes = priorBytes, validationBytes = 0, validationWork = 0, nestedExpanded = 0;
  for (const order of seriesOrders) {
    if (!order.part || orderEdits.has(order.part)) throw new PptxImportError("save-candidate-invalid", "Chart series order has duplicate or missing source parts.");
    const cost = order.sources.length * (2 + Math.ceil(Math.log2(order.sources.length + 1)));
    if (!Number.isSafeInteger(cost) || cost > limits.maxSaveValidationWork - validationWork) throw new PptxImportError("save-limit-exceeded", "Chart order planning exhausted the shared work budget.");
    let edits: readonly ChartSeriesOrderEdit[];
    try { edits = await planChartSeriesOrder(order.sources, order.requested, order.complete, Math.max(1, limits.maxSaveValidationWork - validationWork), checkpoint); }
    catch (error) { if (error instanceof TypeError) throw new PptxImportError("save-candidate-invalid", error.message, { part: order.part }, { cause: error }); if (error instanceof RangeError) throw new PptxImportError("save-limit-exceeded", error.message, { part: order.part }, { cause: error }); throw error; }
    validationWork += cost; orderEdits.set(order.part, edits); ordersByPart.set(order.part, order);
    if (edits.length && !charts.has(order.part)) charts.set(order.part, []);
  }
  const charge = (bytes: number) => { preparedBytes += bytes; if (preparedBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Staged chart/workbook payloads exceed maxPreparedOutputBytes."); };
  const payload = (part: string) => {
    const entry = original.inspection.parts.find(p => p.name === part), bytes = entry && original.payloads?.get(entry.entryName);
    if (!entry || !bytes) throw new PptxImportError("save-candidate-invalid", "Chart/workbook source payload is unavailable.", { part });
    return { entry, bytes };
  };
  const replacements: { name: string; bytes: Uint8Array }[] = [];
  const patch = async (source: Uint8Array, part: string, edits: readonly XmlTextEdit[]) => {
    const xml = await captureXmlTextSource(source, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part });
    try { const changed = await xml.replaceText(edits, { ...(options.signal ? { signal: options.signal } : {}), limits: { maxEdits: limits.maxSaveChanges, maxOutputBytes: limits.maxPreparedOutputBytes } }); await checkpoint(); charge(changed.bytes.length); return changed.bytes; }
    finally { xml.dispose(); }
  };
  for (const [part, edits] of charts) {
    const source = payload(part), order = ordersByPart.get(part); let attributes: readonly import("./chart-legend-order.js").ChartLegendIndexEdit[] | readonly ChartSeriesOrderEdit[] | (import("./chart-legend-order.js").ChartLegendIndexEdit | ChartSeriesOrderEdit)[] = orderEdits.get(part) ?? []; let bytes = source.bytes;
    if (order?.researchColumnLegend) {
      try {
        const legend = await captureColumnLegendSources(bytes, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part }, limits.maxSaveValidationWork - validationWork, checkpoint);
        legendSourceProofs.set(part, legend);
        validationWork += legend.workUnits;
        validationWork += legend.physicalIds.length * order.sources.length;
        if (!Number.isSafeInteger(validationWork) || validationWork >= limits.maxSaveValidationWork) throw new RangeError("Legend source identity checks exhaust validation work.");
        if (legend.physicalIds.length !== order.sources.length || legend.physicalIds.some(id => !order.sources.some(item => item.identity === id))) throw new TypeError("Legend and chart series identities disagree.");
        const cost = legend.entries.length * 2 + legend.physicalIds.length * 2;
        const changes = await planChartLegendOrder(legend.entries, legend.physicalIds, order.requested, limits.maxSaveValidationWork - validationWork, checkpoint);
        validationWork += cost; attributes = [...attributes, ...changes];
      } catch (cause) { if (cause instanceof TypeError) throw new PptxImportError("save-candidate-invalid", cause.message, { part }, { cause }); if (cause instanceof RangeError) throw new PptxImportError("save-limit-exceeded", cause.message, { part }, { cause }); throw cause; }
    }
    if (attributes?.length) {
      if (validationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Chart order editing exhausted the shared work budget.");
      const updated = await editXmlTextAttributes(bytes, attributes, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part, maxOutputBytes: limits.maxPreparedOutputBytes, maxEdits: limits.maxSaveChanges, maxWorkUnits: limits.maxSaveValidationWork - validationWork });
      bytes = updated.bytes; validationWork += updated.workUnits; charge(bytes.length);
    }
    if (edits.length) bytes = await patch(bytes, part, edits);
    if (order) {
      const moved = await moveChartSeries(bytes, part, order.sources, order.requested, options, limits.maxPreparedOutputBytes, limits.maxSaveValidationWork - validationWork, limits.maxChartDataPoints);
      bytes = moved.bytes; validationWork += moved.workUnits; movedElementIndexes.set(part, moved.movedElementIndexes); charge(bytes.length);
    }
    replacements.push({ name: source.entry.entryName, bytes });
  }
  for (const [part, sheets] of books) {
    await checkpoint(); const source = payload(part);
    const remaining = limits.maxTotalExpandedChartWorkbookBytes - nestedExpanded;
    if (remaining <= 0) throw new PptxImportError("save-limit-exceeded", "Chart workbooks exceed the aggregate expansion budget.");
    const nested = await readSource(source.bytes, { ...options, zipLimits: { ...options.zipLimits, maxExpandedBytes: Math.min(remaining, limits.maxExpandedChartWorkbookBytes, options.zipLimits?.maxExpandedBytes ?? Infinity) } }, false, true);
    nestedExpanded += nested.inspection.archive.expandedBytes;
    const changed: { name: string; bytes: Uint8Array }[] = [];
    for (const [sheetPart, edits] of sheets) {
      const entry = nested.inspection.parts.find(p => p.name === sheetPart), bytes = entry && nested.payloads?.get(entry.entryName);
      if (!entry || !bytes) throw new PptxImportError("save-candidate-invalid", "Bound worksheet payload is unavailable.", { part });
      const attributes = spaceEdits.get(part)?.get(sheetPart);
      let sheetBytes = bytes;
      if (attributes?.length) {
        if (validationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Chart attribute editing exhausted the shared work budget.");
        const updated = await editXmlTextAttributes(sheetBytes, attributes, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: `${part}!${sheetPart}`, maxOutputBytes: limits.maxPreparedOutputBytes, maxEdits: limits.maxSaveChanges, maxWorkUnits: limits.maxSaveValidationWork - validationWork });
        sheetBytes = updated.bytes; validationWork += updated.workUnits; charge(sheetBytes.length);
      }
      changed.push({ name: entry.entryName, bytes: await patch(sheetBytes, `${part}!${sheetPart}`, edits) });
    }
    const output = await replaceArchiveEntries(source.bytes, changed, { ...(options.zipLimits ? { limits: options.zipLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), maxOutputBytes: limits.maxPreparedOutputBytes }); charge(output.length);
    if (validationBytes >= limits.maxSaveValidationBytes || validationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Chart validation exhausted the shared comparison budget.");
    const compared = await comparePackages(source.bytes, output, { ...options, zipLimits: { ...options.zipLimits, maxInputBytes: Math.max(source.bytes.length, output.length) }, comparisonLimits: { maxTotalExpandedBytes: limits.maxSaveValidationBytes - validationBytes, maxWorkUnits: limits.maxSaveValidationWork - validationWork, maxDifferences: Math.max(limits.maxSaveChanges, sheets.size + 1) } });
    validationBytes += compared.before.expandedBytes + compared.after.expandedBytes; validationWork += compared.workUnits;
    if (compared.truncated || compared.differenceCount !== sheets.size || compared.differences.some(d => d.code !== "payload-changed" || !sheets.has(d.part))) throw new PptxImportError("save-candidate-invalid", "Nested workbook changed outside approved numeric worksheet values.", { part });
    replacements.push({ name: source.entry.entryName, bytes: output });
  }
  return { replacements, changedParts: new Set([...charts.keys(), ...books.keys()]), validationBytes, validationWork, movedElementIndexes, legendSourceProofs };
}
