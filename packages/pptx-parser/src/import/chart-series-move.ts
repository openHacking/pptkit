import { captureXmlTextSource } from "../preservation/xml-text-source.js";
import { reorderXmlSiblings } from "../preservation/reorder-xml-siblings.js";
import type { PptxInspectionOptions } from "../types/structure.js";
import { PptxImportError } from "../types/import.js";
import type { ChartSeriesOrderSource } from "./chart-series-order.js";

/** Move each family's physically adjacent series after scalar source edits. */
export async function moveChartSeries(bytes: Uint8Array, part: string, sources: readonly ChartSeriesOrderSource[], requested: readonly number[],
  options: PptxInspectionOptions, maxOutputBytes: number, maxWorkUnits: number, maxElements: number): Promise<{ bytes: Uint8Array; workUnits: number; movedElementIndexes: ReadonlyMap<number, number> }> {
  let workUnits = bytes.length * 2 + sources.length * (2 + sources.length + Math.ceil(Math.log2(sources.length + 1)));
  if (!Number.isSafeInteger(workUnits) || workUnits >= maxWorkUnits) throw new PptxImportError("save-limit-exceeded", "Chart series source inspection exhausted validation work.", { part });
  const captured = await captureXmlTextSource(bytes, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part });
  const groups = new Map<number, { namespaceUri: string; sources: { identity: number; index: number }[] }>();
  try {
    for (const source of sources) {
      const order = captured.elements[source.elementIndex - 1], series = order?.parentElementIndex ? captured.elements[order.parentElementIndex - 1] : undefined;
      const parent = series?.parentElementIndex ? captured.elements[series.parentElementIndex - 1] : undefined;
      if (!order || order.namespaceUri !== source.namespaceUri || order.local !== "order" || !series || series.namespaceUri !== source.namespaceUri || series.local !== "ser" || !parent || parent.namespaceUri !== source.namespaceUri || parent.local !== source.family) throw new PptxImportError("save-candidate-invalid", "Chart series movement source ancestry differs from its binding.", { part });
      const group = groups.get(parent.elementIndex) ?? { namespaceUri: source.namespaceUri, sources: [] };
      group.sources.push({ identity: source.identity, index: series.elementIndex }); groups.set(parent.elementIndex, group);
    }
  } finally { captured.dispose(); }
  let output = bytes; const movedElementIndexes = new Map<number, number>();
  for (const [parentElementIndex, group] of groups) {
    group.sources.sort((a, b) => a.index - b.index);
    const byId = new Map(group.sources.map(source => [source.identity, source.index])), desired = requested.filter(identity => byId.has(identity)).map(identity => byId.get(identity)!);
    const elementIndexes = group.sources.map(source => source.index);
    if (desired.length !== elementIndexes.length || new Set(desired).size !== desired.length) throw new PptxImportError("save-candidate-invalid", "Chart series movement identities are incomplete.", { part });
    if (elementIndexes.every((index, i) => index === desired[i])) continue;
    const remaining = maxWorkUnits - workUnits;
    if (remaining <= 0) throw new PptxImportError("save-limit-exceeded", "Chart series movement exhausted validation work.", { part });
    const result = await reorderXmlSiblings(output, { parentElementIndex, namespaceUri: group.namespaceUri, local: "ser", elementIndexes, requestedIndexes: desired }, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part, maxOutputBytes, maxWorkUnits: remaining, maxElements });
    output = result.bytes; workUnits += result.workUnits;
    for (const [before, after] of result.movedElementIndexes) { if (movedElementIndexes.has(before)) throw new PptxImportError("save-candidate-invalid", "Chart series movement regions overlap.", { part }); movedElementIndexes.set(before, after); }
  }
  return { bytes: output, workUnits, movedElementIndexes };
}
