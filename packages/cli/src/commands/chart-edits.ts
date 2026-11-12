import type { ImportedPresentation } from "@pptkit/pptx-parser";
import type { PresentationSlide, PresentationElement, PresentationElementInput } from "@pptkit/core";
import { CliError } from "./arguments.js";
import { checkAbort } from "./read-file.js";

const record = (v: unknown): v is Record<string, unknown> => typeof v === "object" && v !== null && !Array.isArray(v);
export const numberText = (value: number): string => Object.is(value, -0) ? "-0" : String(value);
function numberValue(value: unknown): number | undefined {
  if (typeof value !== "string" || !value) return undefined;
  const number = Number(value); return Number.isFinite(number) && numberText(number) === value ? number : undefined;
}
export type ChartManifestVersion = "v1" | "v2" | "v3" | "v4" | "v5" | "v6";
export const chartManifestNames = { v1: "pptkit-chart-edits-v1", v2: "pptkit-chart-edits-v2", v3: "pptkit-chart-edits-v3", v4: "pptkit-chart-edits-v4", v5: "pptkit-chart-edits-v5", v6: "pptkit-chart-edits-v6" } as const;

export function chartPointArray(
  chart: Extract<PresentationElement | PresentationElementInput, { type: "chart" }>, series: number, coordinate: unknown,
): (number | null)[] | undefined {
  if (chart.chartType === "bubble" || chart.chartType === "scatter") {
    if (coordinate === "x") return chart.series[series]?.xValues;
    if (coordinate === "y") return chart.series[series]?.yValues;
    return chart.chartType === "bubble" && coordinate === "size" ? chart.series[series]?.sizes : undefined;
  }
  return coordinate === undefined ? chart.series[series]?.values : undefined;
}

export function parseChartEdits(bytes: Uint8Array, sourceSha256: string) {
  let value: unknown;
  try { value = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)); }
  catch { throw new CliError("chart-edits-invalid", "Chart edits must be valid UTF-8 JSON.", 3); }
  if (!record(value) || !Object.values(chartManifestNames).some(name => name === value.version) || Object.keys(value).some(k => !["version", "sourceSha256", "edits"].includes(k)) || typeof value.sourceSha256 !== "string" || !/^[0-9a-f]{64}$/.test(value.sourceSha256) || !Array.isArray(value.edits) || !value.edits.length || value.edits.length > 10000) throw new CliError("chart-edits-invalid", "Expected a bounded pptkit-chart-edits-v1 through v6 manifest.", 3);
  if (value.sourceSha256 !== sourceSha256) throw new CliError("chart-edits-source-mismatch", "Chart edits do not match the captured source SHA-256.", 3);
  return { version: value.version === chartManifestNames.v6 ? "v6" as const : value.version === chartManifestNames.v5 ? "v5" as const : value.version === chartManifestNames.v4 ? "v4" as const : value.version === chartManifestNames.v3 ? "v3" as const : value.version === chartManifestNames.v2 ? "v2" as const : "v1" as const, edits: value.edits as unknown[] };
}
export function applyChartEdits(session: ImportedPresentation, edits: readonly unknown[], signal?: AbortSignal, version: ChartManifestVersion = "v1"): void {
  const owners = new Map<string, { slide: PresentationSlide; root: PresentationElement; path: number[] } | null>();
  for (const slide of session.document.slides) for (const root of slide.elements) {
    const pending: { element: PresentationElement; path: number[] }[] = [{ element: root, path: [] }];
    while (pending.length) {
      checkAbort(signal); const { element, path } = pending.pop()!;
      if (element.type === "chart") { const key = JSON.stringify([slide.id, element.id]); owners.set(key, owners.has(key) ? null : { slide, root, path }); }
      else if (element.type === "group") for (let i = element.children.length - 1; i >= 0; i--) pending.push({ element: element.children[i]!, path: [...path, i] });
    }
  }
  const bindings = new Map<string, { count: number; bound: boolean }>();
  for (const binding of session.report.chartValues ?? []) {
    checkAbort(signal);
    if (binding.coordinate && (version === "v1" || (binding.coordinate === "size" && version !== "v4" && version !== "v5" && version !== "v6"))) continue;
    const key = JSON.stringify([binding.slideId, binding.elementId, binding.seriesIndex, binding.pointIndex, binding.coordinate ?? null]), prior = bindings.get(key);
    bindings.set(key, { count: (prior?.count ?? 0) + 1, bound: binding.status === "bound" });
  }
  const seen = new Set<string>(), staged = new Map<string, { slide: PresentationSlide; rootId: string; root: PresentationElementInput }>();
  const labels = new Map<string, { count: number; bound: boolean }>();
  for (const binding of session.report.chartLabels ?? []) { checkAbort(signal); const key = JSON.stringify([binding.slideId, binding.elementId, binding.categoryIndex]), prior = labels.get(key); labels.set(key, { count: (prior?.count ?? 0) + 1, bound: binding.status === "bound" }); }
  const names = new Map<string, { count: number; bound: boolean }>();
  for (const binding of session.report.chartSeriesNames ?? []) { checkAbort(signal); const key = JSON.stringify([binding.slideId, binding.elementId, binding.seriesIndex]), prior = names.get(key); names.set(key, { count: (prior?.count ?? 0) + 1, bound: binding.status === "bound" }); }
  for (const edit of edits) {
    checkAbort(signal);
    if (record(edit) && version === "v6" && edit.kind === "series-name") {
      if (Object.keys(edit).some(k => !["kind", "slideId", "elementId", "seriesIndex", "expected", "replacement"].includes(k)) || typeof edit.slideId !== "string" || !edit.slideId || typeof edit.elementId !== "string" || !edit.elementId || !Number.isSafeInteger(edit.seriesIndex) || (edit.seriesIndex as number) < 0 || typeof edit.expected !== "string" || typeof edit.replacement !== "string" || !edit.replacement.length) throw new CliError("chart-edits-invalid", "Series name edits require a target index and nonempty replacement text.", 3);
      const key = JSON.stringify([edit.slideId, edit.elementId, edit.seriesIndex]), seenKey = `series-name:${key}`;
      if (seen.has(seenKey)) throw new CliError("chart-edits-invalid", "Duplicate chart series-name target in manifest.", 3); seen.add(seenKey);
      const binding = names.get(key); if (!binding || binding.count !== 1 || !binding.bound) throw new CliError("chart-edits-target-unbound", "Chart series name is missing, ambiguous or not source-bound.", 3);
      const owner = owners.get(JSON.stringify([edit.slideId, edit.elementId])); if (!owner) throw new CliError("chart-edits-target-unbound", "Chart owner is missing or ambiguous.", 3);
      const rootKey = JSON.stringify([edit.slideId, owner.root.id]); let item = staged.get(rootKey);
      if (!item) { item = { slide: owner.slide, rootId: owner.root.id, root: structuredClone(owner.root) as PresentationElementInput }; staged.set(rootKey, item); }
      let element = item.root; for (const index of owner.path) { if (element.type !== "group" || !element.children[index]) throw new CliError("chart-edits-target-unbound", "Chart ancestry is inconsistent.", 3); element = element.children[index]!; }
      if (element.type !== "chart" || element.id !== edit.elementId || element.series[edit.seriesIndex as number]?.name !== edit.expected) throw new CliError("chart-edits-baseline-mismatch", "Expected series name does not match the imported baseline.", 3);
      element.series[edit.seriesIndex as number]!.name = edit.replacement; continue;
    }
    if (record(edit) && (version === "v5" || version === "v6") && edit.kind === "category") {
      if (Object.keys(edit).some(k => !["kind", "slideId", "elementId", "categoryIndex", "expected", "replacement"].includes(k)) || typeof edit.slideId !== "string" || !edit.slideId || typeof edit.elementId !== "string" || !edit.elementId || !Number.isSafeInteger(edit.categoryIndex) || (edit.categoryIndex as number) < 0 || typeof edit.expected !== "string" || typeof edit.replacement !== "string" || !edit.replacement.length) throw new CliError("chart-edits-invalid", "Category edits require a target index and nonempty replacement text.", 3);
      const key = JSON.stringify([edit.slideId, edit.elementId, edit.categoryIndex]), seenKey = `category:${key}`;
      if (seen.has(seenKey)) throw new CliError("chart-edits-invalid", "Duplicate chart category target in manifest.", 3); seen.add(seenKey);
      const binding = labels.get(key); if (!binding || binding.count !== 1 || !binding.bound) throw new CliError("chart-edits-target-unbound", "Chart category is missing, ambiguous or not source-bound.", 3);
      const owner = owners.get(JSON.stringify([edit.slideId, edit.elementId])); if (!owner) throw new CliError("chart-edits-target-unbound", "Chart owner is missing or ambiguous.", 3);
      const rootKey = JSON.stringify([edit.slideId, owner.root.id]); let item = staged.get(rootKey);
      if (!item) { item = { slide: owner.slide, rootId: owner.root.id, root: structuredClone(owner.root) as PresentationElementInput }; staged.set(rootKey, item); }
      let element = item.root; for (const index of owner.path) { if (element.type !== "group" || !element.children[index]) throw new CliError("chart-edits-target-unbound", "Chart ancestry is inconsistent.", 3); element = element.children[index]!; }
      if (element.type !== "chart" || element.id !== edit.elementId || !("categories" in element) || !element.categories || element.categories[edit.categoryIndex as number] !== edit.expected) throw new CliError("chart-edits-baseline-mismatch", "Expected category does not match the imported baseline.", 3);
      element.categories[edit.categoryIndex as number] = edit.replacement; continue;
    }
    if (!record(edit) || Object.keys(edit).some(k => !["slideId", "elementId", "seriesIndex", "pointIndex", "expected", "replacement", ...(version !== "v1" ? ["coordinate"] : [])].includes(k)) || ("coordinate" in edit && edit.coordinate !== "x" && edit.coordinate !== "y" && (edit.coordinate !== "size" || version !== "v4" && version !== "v5" && version !== "v6")) || typeof edit.slideId !== "string" || !edit.slideId || typeof edit.elementId !== "string" || !edit.elementId || !Number.isSafeInteger(edit.seriesIndex) || (edit.seriesIndex as number) < 0 || !Number.isSafeInteger(edit.pointIndex) || (edit.pointIndex as number) < 0 || numberValue(edit.expected) === undefined || numberValue(edit.replacement) === undefined) throw new CliError("chart-edits-invalid", "Chart edits require target indices and canonical finite numeric strings.", 3);
    const key = JSON.stringify([edit.slideId, edit.elementId, edit.seriesIndex, edit.pointIndex, edit.coordinate ?? null]);
    if (seen.has(key)) throw new CliError("chart-edits-invalid", "Duplicate chart target in manifest.", 3); seen.add(key);
    const binding = bindings.get(key); if (!binding || binding.count !== 1 || !binding.bound) throw new CliError("chart-edits-target-unbound", "Chart value is missing, ambiguous or not source-bound.", 3);
    const owner = owners.get(JSON.stringify([edit.slideId, edit.elementId])); if (!owner) throw new CliError("chart-edits-target-unbound", "Chart owner is missing or ambiguous.", 3);
    const rootKey = JSON.stringify([edit.slideId, owner.root.id]); let item = staged.get(rootKey);
    if (!item) { item = { slide: owner.slide, rootId: owner.root.id, root: structuredClone(owner.root) as PresentationElementInput }; staged.set(rootKey, item); }
    let element = item.root; for (const index of owner.path) { if (element.type !== "group" || !element.children[index]) throw new CliError("chart-edits-target-unbound", "Chart ancestry is inconsistent.", 3); element = element.children[index]!; }
    if (element.type !== "chart" || element.id !== edit.elementId) throw new CliError("chart-edits-target-unbound", "Chart ancestry is inconsistent.", 3);
    if ((element.chartType === "bubble" || element.chartType === "combo") && version !== "v4" && version !== "v5" && version !== "v6") throw new CliError("chart-edits-target-unbound", "Bubble and combination edits require manifest v4 or v5.", 3);
    const replacement = numberValue(edit.replacement)!;
    if (edit.coordinate === "size" && replacement <= 0) throw new CliError("chart-edits-invalid", "Bubble size replacements must be strictly positive.", 3);
    const values = chartPointArray(element, edit.seriesIndex as number, edit.coordinate);
    const prior = values?.[edit.pointIndex as number];
    if (typeof prior !== "number" || !Object.is(prior, numberValue(edit.expected))) throw new CliError("chart-edits-baseline-mismatch", "Expected value does not match the imported baseline.", 3);
    values![edit.pointIndex as number] = replacement;
  }
  for (const item of staged.values()) { checkAbort(signal); item.slide.replaceElement(item.rootId, item.root); }
}
