import type { ChartSeriesOrderEdit } from "./chart-series-order.js";

/** Private source description; caller must prove raw ancestry and chart semantics. */
export interface ChartLegendIndexSource {
  readonly elementIndex: number;
  readonly namespaceUri: string;
  readonly position: number;
  readonly lexical: string;
}
export interface ChartLegendIndexEdit extends Omit<ChartSeriesOrderEdit, "elementLocal"> { readonly elementLocal: "idx" }

/** Map position-based legend entries through explicit physical-series identities. */
export async function planChartLegendOrder(entries: readonly ChartLegendIndexSource[], physicalIds: readonly number[], requestedIds: readonly number[], maxWork: number, checkpoint: () => Promise<void>): Promise<readonly ChartLegendIndexEdit[]> {
  if (!Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError("Legend order requires a positive work limit.");
  if (!Array.isArray(entries) || !Array.isArray(physicalIds) || !Array.isArray(requestedIds) || !physicalIds.length || requestedIds.length !== physicalIds.length) throw new TypeError("Legend order requires complete series identities.");
  const cost = entries.length * 2 + physicalIds.length * 2;
  if (!Number.isSafeInteger(cost) || cost > maxWork) throw new RangeError("Legend order exceeds its work limit.");
  entries = entries.map(entry => ({ ...entry })); physicalIds = [...physicalIds]; requestedIds = [...requestedIds];
  const validId = (id: number) => Number.isSafeInteger(id) && id >= 0 && id <= 4294967295;
  const ids = new Set(physicalIds), requested = new Map<number, number>();
  if (ids.size !== physicalIds.length || physicalIds.some(id => !validId(id)) || requestedIds.some(id => !validId(id) || !ids.has(id)) || new Set(requestedIds).size !== ids.size) throw new TypeError("Legend order requires an exact series identity permutation.");
  for (const [position, id] of requestedIds.entries()) { await checkpoint(); requested.set(id, position); }
  const positions = new Set<number>(), elements = new Set<number>(), edits: ChartLegendIndexEdit[] = [];
  for (const entry of entries) {
    await checkpoint();
    if (!Number.isSafeInteger(entry.elementIndex) || entry.elementIndex <= 0 || !entry.namespaceUri || !Number.isSafeInteger(entry.position) || entry.position < 0 || entry.position >= physicalIds.length || !/^\d+$/u.test(entry.lexical) || Number(entry.lexical) !== entry.position || positions.has(entry.position) || elements.has(entry.elementIndex)) throw new TypeError("Legend index source is invalid, duplicated or outside the series collection.");
    positions.add(entry.position); elements.add(entry.elementIndex);
    const position = requested.get(physicalIds[entry.position]!)!;
    if (position !== entry.position) edits.push({ elementIndex: entry.elementIndex, elementNamespace: entry.namespaceUri, elementLocal: "idx", attributeNamespace: "", attributeLocal: "val", expected: entry.lexical, replacement: String(position) });
  }
  return edits;
}
