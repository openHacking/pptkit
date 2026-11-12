import type { ImportedPresentation } from "@pptkit/pptx-parser";
import type { PresentationElement } from "@pptkit/core";
import { CliError } from "./arguments.js";
import { checkAbort } from "./read-file.js";

/** Extract only source-bound scalar values; preparation still supplies edit authority. */
export function textCandidates(session: ImportedPresentation, profile: "text-runs-v1" | "text-runs-v2" | "text-runs-v3", sourceSha256: string, signal?: AbortSignal) {
  const owners = new Map<string, PresentationElement | null>(), slides = new Map(session.document.slides.map(s => [s.id, s]));
  for (const slide of session.document.slides) {
    const pending = [...slide.elements];
    while (pending.length) {
      checkAbort(signal); const element = pending.pop()!, key = JSON.stringify([slide.id, element.id]);
      if (element.type === "group") for (const child of element.children) pending.push(child);
      else if (element.type === "text" || element.type === "shape" || element.type === "table") owners.set(key, owners.has(key) ? null : element);
    }
  }
  const targets: Record<string, unknown>[] = [], seen = new Set<string>(); let unboundCount = 0;
  const add = (binding: { status: string; paragraphIndex: number; runIndex: number }, identity: Record<string, unknown>, content: unknown) => {
    checkAbort(signal);
    const key = JSON.stringify(identity);
    if (seen.has(key)) throw new CliError("text-candidates-inconsistent", "Text target identity is ambiguous.", 3);
    seen.add(key);
    if (binding.status !== "bound") { unboundCount++; return; }
    if (!Array.isArray(content) || typeof content[binding.paragraphIndex]?.runs?.[binding.runIndex]?.text !== "string") throw new CliError("text-candidates-inconsistent", "Bound text target has unavailable Core content.", 3);
    targets.push({ ...identity, expected: content[binding.paragraphIndex].runs[binding.runIndex].text });
  };
  for (const binding of session.report.textRuns ?? []) {
    const owner = binding.owner ?? "text-box", element = owners.get(JSON.stringify([binding.slideId, binding.elementId]));
    const content = owner === "text-box" && element?.type === "text" ? element.content : owner === "shape-text" && element?.type === "shape" ? element.text?.content : undefined;
    add(binding, { ...(profile === "text-runs-v1" ? {} : { owner }), slideId: binding.slideId, elementId: binding.elementId, paragraphIndex: binding.paragraphIndex, runIndex: binding.runIndex }, content);
  }
  for (const binding of session.report.notesTextRuns ?? []) add(binding, { owner: "notes", slideId: binding.slideId, paragraphIndex: binding.paragraphIndex, runIndex: binding.runIndex }, slides.get(binding.slideId)?.notes);
  for (const binding of session.report.tableTextRuns ?? []) {
    const element = owners.get(JSON.stringify([binding.slideId, binding.elementId]));
    add(binding, { owner: "table-cell", slideId: binding.slideId, elementId: binding.elementId, rowIndex: binding.rowIndex, cellIndex: binding.cellIndex, columnIndex: binding.columnIndex, paragraphIndex: binding.paragraphIndex, runIndex: binding.runIndex }, element?.type === "table" ? element.rows[binding.rowIndex]?.cells[binding.cellIndex]?.content : undefined);
  }
  return { version: profile.replace("text-runs", "pptkit-text-edits"), sourceSha256, targets, unboundCount };
}
