import type { ImportedPresentation } from "@pptkit/pptx-parser";
import type { PresentationElement } from "@pptkit/core";
import { CliError } from "./arguments.js";
import { chartManifestNames, chartPointArray, numberText, type ChartManifestVersion } from "./chart-edits.js";
import { checkAbort } from "./read-file.js";

export function chartCandidates(session: ImportedPresentation, sourceSha256: string, signal?: AbortSignal, version: ChartManifestVersion = "v1") {
  const owners = new Map<string, PresentationElement | null>();
  for (const slide of session.document.slides) {
    const pending = [...slide.elements];
    while (pending.length) {
      checkAbort(signal); const element = pending.pop()!;
      if (element.type === "group") for (const child of element.children) pending.push(child);
      else if (element.type === "chart") { const key = JSON.stringify([slide.id, element.id]); owners.set(key, owners.has(key) ? null : element); }
    }
  }
  const targets: Record<string, unknown>[] = [], seen = new Set<string>(); let unboundCount = 0;
  for (const binding of session.report.chartValues ?? []) {
    checkAbort(signal);
    // The v1 manifest has no coordinate field and cannot identify scatter targets.
    if (binding.coordinate && version === "v1") { unboundCount++; continue; }
    const identity = { slideId: binding.slideId, elementId: binding.elementId, seriesIndex: binding.seriesIndex, pointIndex: binding.pointIndex, ...(binding.coordinate ? { coordinate: binding.coordinate } : {}) }, key = JSON.stringify(identity);
    if (seen.has(key)) throw new CliError("chart-candidates-inconsistent", "Chart target identity is ambiguous.", 3); seen.add(key);
    if (binding.status !== "bound") { unboundCount++; continue; }
    const element = owners.get(JSON.stringify([binding.slideId, binding.elementId]));
    if (element?.type === "chart" && (element.chartType === "bubble" || element.chartType === "combo") && version !== "v4" && version !== "v5" && version !== "v6") { unboundCount++; continue; }
    const value = element?.type === "chart" ? chartPointArray(element, binding.seriesIndex, binding.coordinate)?.[binding.pointIndex] : undefined;
    if (typeof value !== "number" || !Number.isFinite(value)) throw new CliError("chart-candidates-inconsistent", "Bound chart target has unavailable or non-finite Core value.", 3);
    targets.push({ ...identity, expected: numberText(value) });
  }
  if (version === "v5" || version === "v6") for (const binding of session.report.chartLabels ?? []) {
    checkAbort(signal);
    const identity = { kind: "category", slideId: binding.slideId, elementId: binding.elementId, categoryIndex: binding.categoryIndex }, key = JSON.stringify(identity);
    if (seen.has(key)) throw new CliError("chart-candidates-inconsistent", "Chart category target identity is ambiguous.", 3); seen.add(key);
    if (binding.status !== "bound") { unboundCount++; continue; }
    const element = owners.get(JSON.stringify([binding.slideId, binding.elementId]));
    const value = element?.type === "chart" && "categories" in element ? element.categories[binding.categoryIndex] : undefined;
    if (typeof value !== "string" || !value) throw new CliError("chart-candidates-inconsistent", "Bound chart category has no Core value.", 3);
    targets.push({ ...identity, expected: value });
  }
  if (version === "v6") for (const binding of session.report.chartSeriesNames ?? []) {
    checkAbort(signal);
    const identity = { kind: "series-name", slideId: binding.slideId, elementId: binding.elementId, seriesIndex: binding.seriesIndex }, key = JSON.stringify(identity);
    if (seen.has(key)) throw new CliError("chart-candidates-inconsistent", "Chart series-name target identity is ambiguous.", 3); seen.add(key);
    if (binding.status !== "bound") { unboundCount++; continue; }
    const element = owners.get(JSON.stringify([binding.slideId, binding.elementId]));
    const value = element?.type === "chart" ? element.series[binding.seriesIndex]?.name : undefined;
    if (typeof value !== "string" || !value) throw new CliError("chart-candidates-inconsistent", "Bound chart series name has no Core value.", 3);
    targets.push({ ...identity, expected: value });
  }
  return { version: chartManifestNames[version], sourceSha256, targets, unboundCount };
}
