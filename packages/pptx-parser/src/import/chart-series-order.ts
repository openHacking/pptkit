/** A raw series and its independently located c:order attribute. */
export interface ChartSeriesOrderSource {
  readonly identity: number;
  readonly order: number;
  readonly orderLexical: string;
  readonly elementIndex: number;
  readonly namespaceUri: string;
  readonly family: string;
}
export interface ChartSeriesOrderEdit {
  readonly elementIndex: number;
  readonly elementNamespace: string;
  readonly elementLocal: "order";
  readonly attributeNamespace: "";
  readonly attributeLocal: "val";
  readonly expected: string;
  readonly replacement: string;
}
/** Private link between Core chart identity and raw series-order sources. */
export interface ChartSeriesOrderBinding {
  readonly slideId: string;
  readonly elementId: string;
  readonly part: string;
  readonly slidePart: string;
  readonly sources: readonly ChartSeriesOrderSource[];
  readonly complete: boolean;
  readonly researchColumnLegend?: true;
  readonly localDependencyReason?: string;
}
export interface ChartSeriesOrderProposal { readonly binding: ChartSeriesOrderBinding; readonly requested: readonly number[] }

/**
 * Private plan only: the caller must prove dependency closure and save authority.
 * Source IDs, not names or array positions, identify series. Existing order
 * lexemes are reassigned by rank, preserving their distinct integer values.
 */
export async function planChartSeriesOrder(sources: readonly ChartSeriesOrderSource[], requested: readonly number[], complete: boolean,
  maxWork: number, checkpoint: () => Promise<void>): Promise<readonly ChartSeriesOrderEdit[]> {
  if (!Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError("Chart series-order work limit must be a positive safe integer.");
  if (!complete || !sources.length || sources.length !== requested.length) throw new TypeError("Chart series-order sources must be complete and match the requested permutation.");
  if (sources.length > maxWork) throw new RangeError("Chart series-order planning exceeds its work limit.");
  // Snapshot the caller's records before the first checkpoint yields.
  sources = sources.map(source => ({ ...source })); requested = [...requested];
  let work = 0; const charge = async (units: number) => { work += units; if (!Number.isSafeInteger(work) || work > maxWork) throw new RangeError("Chart series-order planning exceeds its work limit."); await checkpoint(); };
  const step = () => charge(1);
  const ids = new Map<number, ChartSeriesOrderSource>(), orders = new Set<number>(), elements = new Set<number>();
  let namespace: string | undefined;
  for (const source of sources) {
    await step();
    if (!Number.isSafeInteger(source.identity) || source.identity < 0 || source.identity > 4294967295 || !Number.isSafeInteger(source.order) || source.order < 0 || source.order > 4294967295 ||
        !/^\d+$/u.test(source.orderLexical) || Number(source.orderLexical) !== source.order || !Number.isSafeInteger(source.elementIndex) || source.elementIndex <= 0 || !source.namespaceUri || !source.family ||
        ids.has(source.identity) || orders.has(source.order) || elements.has(source.elementIndex) || namespace !== undefined && namespace !== source.namespaceUri) throw new TypeError("Chart series-order source identity is missing, duplicated or inconsistent.");
    ids.set(source.identity, source); orders.add(source.order); elements.add(source.elementIndex); namespace = source.namespaceUri;
  }
  await charge(sources.length * Math.ceil(Math.log2(sources.length + 1)));
  const ranked = [...sources].sort((a, b) => a.order - b.order), seen = new Set<number>(), edits: ChartSeriesOrderEdit[] = [];
  for (const [index, identity] of requested.entries()) {
    await step();
    const source = ids.get(identity), slot = ranked[index]!;
    if (!source || seen.has(identity)) throw new TypeError("Chart series-order request must be an exact identity permutation.");
    seen.add(identity);
    if (source.family !== slot.family) throw new TypeError("Chart series-order request cannot cross chart-family boundaries.");
    if (source.order !== slot.order) edits.push({ elementIndex: source.elementIndex, elementNamespace: source.namespaceUri, elementLocal: "order", attributeNamespace: "", attributeLocal: "val", expected: source.orderLexical, replacement: slot.orderLexical });
  }
  return edits;
}

/** Reconcile reimported raw identities; Core snapshot equality is a separate check. */
export async function verifyChartSeriesOrderRebound(original: readonly ChartSeriesOrderSource[], requested: readonly number[], rebound: readonly ChartSeriesOrderSource[],
  originalComplete: boolean, reboundComplete: boolean, maxWork: number, checkpoint: () => Promise<void>, movedElementIndexes: ReadonlyMap<number, number> = new Map()): Promise<void> {
  if (!Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError("Chart series-order verification work limit must be a positive safe integer.");
  if (!originalComplete || !reboundComplete || !original.length || original.length !== rebound.length || original.length !== requested.length) throw new TypeError("Chart series-order rebinding requires complete, equally sized sources.");
  const count = original.length, sortWork = count * Math.ceil(Math.log2(count + 1)), perPlan = count * 2 + sortWork, required = perPlan * 2 + sortWork + count;
  if (!Number.isSafeInteger(required) || required > maxWork) throw new RangeError("Chart series-order verification exceeds its work limit.");
  // Capture both generations and requested order before the first yield.
  original = original.map(source => ({ ...source })); rebound = rebound.map(source => ({ ...source })); requested = [...requested];
  await planChartSeriesOrder(original, requested, true, perPlan, checkpoint);
  if ((await planChartSeriesOrder(rebound, requested, true, perPlan, checkpoint)).length) throw new TypeError("Reimported series order differs from the requested order.");
  const ranked = [...original].sort((a, b) => a.order - b.order), originalById = new Map(original.map(source => [source.identity, source])), reboundById = new Map(rebound.map(source => [source.identity, source]));
  for (const [index, identity] of requested.entries()) {
    await checkpoint();
    const before = originalById.get(identity), after = reboundById.get(identity), slot = ranked[index];
    if (!before || !after || !slot || (movedElementIndexes.get(before.elementIndex) ?? before.elementIndex) !== after.elementIndex || before.namespaceUri !== after.namespaceUri || before.family !== after.family || after.order !== slot.order || after.orderLexical !== slot.orderLexical) throw new TypeError("Reimported series identity, source location or order differs from its plan.");
  }
}
