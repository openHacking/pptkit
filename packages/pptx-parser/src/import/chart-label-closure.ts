import { chartString } from './chart-string.js';

/** A physical cache point and its resolved, direct workbook scalar candidate. */
export interface ChartLabelUse {
  readonly logicalKey: string;
  readonly projected: string;
  readonly cache: { readonly part: string; readonly elementIndex: number; readonly lexical: string; readonly namespaceUri: string };
  readonly cell: { readonly workbookPart: string; readonly worksheetPart: string; readonly reference: string; readonly elementIndex: number; readonly lexical: string; readonly namespaceUri: string; readonly preserveSpace: boolean; readonly space?: "default" | "preserve" };
}
export interface ChartLabelClosure {
  readonly logicalKey: string;
  readonly status: 'bound' | 'unmapped';
  readonly reason?: string;
  readonly uses: readonly ChartLabelUse[];
}

/**
 * Reconcile already-resolved observations; the caller must count every formula
 * use, including non-label references, and prove the source scan is complete.
 * This alone does not grant save permission or establish raw XML ancestry.
 */
export async function chartLabelClosure(expectedKeys: readonly string[], uses: readonly ChartLabelUse[], cellUseCounts: ReadonlyMap<string, number>,
  scanComplete: boolean, maxWork: number, checkpoint: () => Promise<void>): Promise<readonly ChartLabelClosure[]> {
  if (!Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError('Chart label closure work limit must be a positive safe integer.');
  let work = 0;
  const step = async () => { if (++work > maxWork) throw new RangeError('Chart label closure exceeds its work limit.'); await checkpoint(); };
  const cellKey = (use: ChartLabelUse) => JSON.stringify([use.cell.workbookPart, use.cell.worksheetPart, use.cell.reference]);
  const cacheKey = (use: ChartLabelUse) => JSON.stringify([use.cache.part, use.cache.elementIndex]);
  const logical = new Map<string, { uses: ChartLabelUse[]; values: Set<string>; cells: Set<string> }>();
  const cells = new Map<string, { count: number; owners: Set<string>; indices: Set<number> }>(), caches = new Map<string, number>();
  const expected = new Set<string>();
  for (const key of expectedKeys) {
    await step();
    if (!key || expected.has(key)) throw new TypeError('Expected chart label identities must be distinct and nonempty.');
    expected.add(key);
  }
  for (const use of uses) {
    await step();
    if (!expected.has(use.logicalKey)) throw new TypeError('Chart label observation has no expected logical identity.');
    const key = cellKey(use), group = logical.get(use.logicalKey) ?? { uses: [], values: new Set<string>(), cells: new Set<string>() };
    group.uses.push(use); group.values.add(use.projected); group.cells.add(key); logical.set(use.logicalKey, group);
    const owners = cells.get(key) ?? { count: 0, owners: new Set<string>(), indices: new Set<number>() };
    owners.count++; owners.owners.add(use.logicalKey); owners.indices.add(use.cell.elementIndex); cells.set(key, owners);
    const cache = cacheKey(use); caches.set(cache, (caches.get(cache) ?? 0) + 1);
  }
  const result: ChartLabelClosure[] = [];
  for (const key of expectedKeys) {
    await step();
    const group = logical.get(key);
    if (!group) { result.push({ logicalKey: key, status: 'unmapped', reason: 'No source observations for this label.', uses: [] }); continue; }
    let reason: string | undefined;
    if (!scanComplete) reason = 'Chart reference scan is incomplete.';
    for (const use of group.uses) {
      await step();
      const cell = cellKey(use), owners = cells.get(cell)!;
      if (!use.logicalKey || !use.cache.part || !Number.isSafeInteger(use.cache.elementIndex) || use.cache.elementIndex <= 0 || !use.cell.workbookPart || !use.cell.worksheetPart || !use.cell.reference || !Number.isSafeInteger(use.cell.elementIndex) || use.cell.elementIndex <= 0) reason ??= 'A direct cache or workbook source is missing.';
      else if (caches.get(cacheKey(use)) !== 1) reason ??= 'A cache point is assigned more than once.';
      else if (owners.owners.size !== 1 || owners.indices.size !== 1) reason ??= 'Workbook cell has conflicting logical owners or physical sources.';
      else if (cellUseCounts.get(cell) !== owners.count) reason ??= 'Workbook cell has unaccounted or ambiguous references.';
      else if (group.values.size !== 1 || group.cells.size !== 1 || chartString(use.cache.lexical) !== use.projected || chartString(use.cell.lexical) !== use.projected) reason ??= 'Cache, workbook, or logical label values disagree.';
    }
    result.push({ logicalKey: key, status: reason ? 'unmapped' : 'bound', ...(reason ? { reason } : {}), uses: group.uses });
  }
  return result;
}
