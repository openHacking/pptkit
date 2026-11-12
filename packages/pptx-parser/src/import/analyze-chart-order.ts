import type { NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import { PptxImportError, type PptxSaveChange } from "../types/import.js";
import { planChartSeriesOrder, type ChartSeriesOrderBinding, type ChartSeriesOrderProposal } from "./chart-series-order.js";

/** Private explicit request; identities are raw chart series IDs, never names. */
export interface ChartSeriesOrderRequest { readonly slideId: string; readonly elementId: string; readonly identities: readonly number[] }

/** Capture untrusted save options synchronously, without invoking accessors. */
export function captureChartSeriesOrderRequests(input: unknown, maxChanges: number, maxWork: number): readonly ChartSeriesOrderRequest[] {
  if (!Number.isSafeInteger(maxChanges) || maxChanges <= 0 || !Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError("Chart order capture requires positive safe limits.");
  if (input === undefined) return [];
  const dense = (value: unknown, label: string): readonly unknown[] => {
    if (!Array.isArray(value) || value.length > maxWork || Reflect.ownKeys(value).length !== value.length + 1) throw new TypeError(`${label} requires a bounded dense array.`);
    const result: unknown[] = [];
    for (let i = 0; i < value.length; i++) {
      const descriptor = Object.getOwnPropertyDescriptor(value, String(i));
      if (!descriptor || !("value" in descriptor)) throw new TypeError(`${label} requires data properties.`);
      result.push(descriptor.value);
    }
    return result;
  };
  const items = dense(input, "Chart order requests");
  if (items.length > maxChanges) throw new PptxImportError("save-limit-exceeded", "Chart order requests exceed maxSaveChanges.");
  let work = items.length; const owners = new Set<string>(), result: ChartSeriesOrderRequest[] = [];
  for (const item of items) {
    if (!item || typeof item !== "object" || Array.isArray(item) || ![Object.prototype, null].includes(Object.getPrototypeOf(item)) || Reflect.ownKeys(item).length !== 3) throw new TypeError("Chart order requests require plain records with three fields.");
    const descriptors = Object.getOwnPropertyDescriptors(item);
    if (!["slideId", "elementId", "identities"].every(key => descriptors[key] && "value" in descriptors[key]!)) throw new TypeError("Chart order request fields must be data properties.");
    const slideId = descriptors.slideId!.value, elementId = descriptors.elementId!.value;
    if (typeof slideId !== "string" || !slideId || slideId.length > 256 || typeof elementId !== "string" || !elementId || elementId.length > 256) throw new TypeError("Chart order requests require bounded owner identities.");
    const key = JSON.stringify([slideId, elementId]);
    if (owners.has(key)) throw new TypeError("Chart order targets must be unique.");
    owners.add(key);
    const identities = dense(descriptors.identities!.value, "Chart series identities");
    work += slideId.length + elementId.length + identities.length;
    if (!Number.isSafeInteger(work) || work > maxWork) throw new PptxImportError("save-limit-exceeded", "Chart order capture exceeds its work budget.");
    if (!identities.length || identities.some(value => typeof value !== "number" || !Number.isSafeInteger(value) || Object.is(value, -0) || value < 0 || value > 4294967295) || new Set(identities).size !== identities.length) throw new TypeError("Chart series identities must be unique unsigned 32-bit integers.");
    result.push({ slideId, elementId, identities: identities as number[] });
  }
  return result;
}

/**
 * Restore candidate series to baseline indexing for ordinary scalar analysis.
 * The caller must separately verify source closure and the final candidate.
 */
export async function analyzeChartOrders(baseline: NormalizedPresentation, candidate: NormalizedPresentation, bindings: readonly ChartSeriesOrderBinding[],
  requests: readonly ChartSeriesOrderRequest[], maxWork: number, checkpoint: () => Promise<void>): Promise<{ candidateForScalars: NormalizedPresentation; proposals: readonly ChartSeriesOrderProposal[]; changes: readonly PptxSaveChange[]; work: number }> {
  if (!Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError("Chart order analysis requires a positive safe work limit.");
  let work = 0; const charge = async (units = 1) => { work += units; if (!Number.isSafeInteger(work) || work > maxWork) throw new PptxImportError("save-limit-exceeded", "Chart order analysis exceeds its work budget."); await checkpoint(); };
  if (requests.length > maxWork || bindings.length > maxWork) throw new PptxImportError("save-limit-exceeded", "Chart order analysis exceeds its work budget.");
  let identities = 0; for (const request of requests) { identities += request.identities.length; if (identities > maxWork) throw new PptxImportError("save-limit-exceeded", "Chart order identities exceed the work budget."); }
  const captured = requests.map(request => ({ slideId: request.slideId, elementId: request.elementId, identities: [...request.identities] }));
  const output = structuredClone(candidate), byOwner = new Map<string, ChartSeriesOrderBinding>(), ambiguous = new Set<string>();
  for (const binding of bindings) { await charge(); const key = JSON.stringify([binding.slideId, binding.elementId]); if (byOwner.has(key)) ambiguous.add(key); else byOwner.set(key, binding); }
  const find = async (document: NormalizedPresentation, slideId: string, elementId: string) => {
    let found: { element: Extract<NormalizedElement, { type: "chart" }>; path: string } | undefined;
    for (const [i, slide] of document.slides.entries()) {
      await charge(); if (slide.id !== slideId) continue;
      const pending = slide.elements.map((element, index) => ({ element, path: `/slides/${i}/elements/${index}` }));
      while (pending.length) {
        await charge(); const { element, path } = pending.pop()!;
        if (element.id === elementId) {
          if (found || element.type !== "chart") throw new PptxImportError("save-candidate-invalid", "Chart order target is ambiguous or not a chart.");
          found = { element, path };
        }
        if (element.type === "group") for (const [index, child] of element.children.entries()) pending.push({ element: child, path: `${path}/children/${index}` });
      }
    }
    return found;
  };
  const proposals: ChartSeriesOrderProposal[] = [], changes: PptxSaveChange[] = [], seen = new Set<string>();
  for (const request of captured) {
    await charge();
    const key = JSON.stringify([request.slideId, request.elementId]), binding = byOwner.get(key);
    if (seen.has(key) || ambiguous.has(key) || !binding || !binding.complete) throw new PptxImportError("save-candidate-invalid", "Chart order target is duplicate, unavailable or not source-bound.");
    seen.add(key);
    const old = await find(baseline, request.slideId, request.elementId), next = await find(output, request.slideId, request.elementId);
    if (!old || !next || !binding.sources.length || old.element.series.length !== binding.sources.length || next.element.series.length !== binding.sources.length) throw new PptxImportError("save-candidate-invalid", "Chart order target or series count differs from its baseline.");
    const count = binding.sources.length, sort = count * Math.ceil(Math.log2(count + 1)), planWork = count * 2 + sort;
    await charge(planWork + sort + count * 2);
    let edits;
    try { edits = await planChartSeriesOrder(binding.sources, request.identities, true, planWork, checkpoint); }
    catch (cause) { if (cause instanceof TypeError) throw new PptxImportError("save-candidate-invalid", cause.message, { part: binding.part }, { cause }); throw cause; }
    if (!edits.length) throw new PptxImportError("save-candidate-invalid", "Chart order request does not change the source order.");
    const baselineIds = [...binding.sources].sort((a, b) => a.order - b.order).map(source => source.identity), positions = new Map(request.identities.map((identity, index) => [identity, index]));
    const originalOrder = baselineIds.map(identity => next.element.series[positions.get(identity)!]!);
    (next.element.series as unknown[]).splice(0, next.element.series.length, ...originalOrder);
    proposals.push({ binding, requested: request.identities });
    changes.push({ kind: "structure", documentPath: `${old.path}/series`, slideId: request.slideId, elementId: request.elementId, source: { part: binding.part } });
  }
  return { candidateForScalars: output, proposals, changes, work };
}
