import type { ImportedPresentation } from "@pptkit/pptx-parser";
import type { PresentationSlide } from "@pptkit/core";
import { CliError } from "./arguments.js";
import { checkAbort } from "./read-file.js";

const record = (value: unknown): value is Record<string, unknown> => typeof value === "object" && value !== null && !Array.isArray(value);
function sources(session: ImportedPresentation, signal?: AbortSignal) {
  const byId = new Map<string, PresentationSlide | null>();
  for (const slide of session.document.slides) { checkAbort(signal); byId.set(slide.id, byId.has(slide.id) ? null : slide); }
  const bindings = new Map<string, boolean>();
  for (const binding of session.report.slideVisibilitySources ?? []) {
    checkAbort(signal); if (bindings.has(binding.slideId)) throw new CliError("visibility-candidates-inconsistent", "Slide visibility identity is ambiguous.", 3);
    bindings.set(binding.slideId, binding.status === "bound");
  }
  return { byId, bindings };
}
export function visibilityCandidates(session: ImportedPresentation, sourceSha256: string, signal?: AbortSignal) {
  const { byId, bindings } = sources(session, signal), targets: { slideId: string; expected: boolean }[] = []; let unboundCount = 0;
  for (const [slideId, bound] of bindings) {
    checkAbort(signal); if (!bound) { unboundCount++; continue; }
    const slide = byId.get(slideId); if (!slide || typeof slide.hidden !== "boolean" || typeof slide.setHidden !== "function") throw new CliError("visibility-candidates-inconsistent", "Bound slide visibility is unavailable or ambiguous.", 3);
    targets.push({ slideId, expected: slide.hidden });
  }
  return { version: "pptkit-visibility-edits-v1", sourceSha256, targets, unboundCount };
}
export function parseVisibilityEdits(bytes: Uint8Array, sourceSha256: string): readonly unknown[] {
  let value: unknown; try { value = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)); }
  catch { throw new CliError("visibility-edits-invalid", "Visibility edits must be valid UTF-8 JSON.", 3); }
  if (!record(value) || value.version !== "pptkit-visibility-edits-v1" || Object.keys(value).some(k => !["version", "sourceSha256", "edits"].includes(k)) || typeof value.sourceSha256 !== "string" || !/^[0-9a-f]{64}$/.test(value.sourceSha256) || !Array.isArray(value.edits) || !value.edits.length || value.edits.length > 10000) throw new CliError("visibility-edits-invalid", "Expected a bounded pptkit-visibility-edits-v1 manifest.", 3);
  if (value.sourceSha256 !== sourceSha256) throw new CliError("visibility-edits-source-mismatch", "Visibility edits do not match the captured source SHA-256.", 3);
  return value.edits as unknown[];
}
export function applyVisibilityEdits(session: ImportedPresentation, edits: readonly unknown[], signal?: AbortSignal): void {
  const { byId, bindings } = sources(session, signal), seen = new Set<string>(), staged: { slide: PresentationSlide; replacement: boolean }[] = [];
  for (const edit of edits) {
    checkAbort(signal);
    if (!record(edit) || Object.keys(edit).some(k => !["slideId", "expected", "replacement"].includes(k)) || typeof edit.slideId !== "string" || !edit.slideId || typeof edit.expected !== "boolean" || typeof edit.replacement !== "boolean") throw new CliError("visibility-edits-invalid", "Visibility edits require a slide ID and Boolean values.", 3);
    if (seen.has(edit.slideId)) throw new CliError("visibility-edits-invalid", "Duplicate visibility target in manifest.", 3); seen.add(edit.slideId);
    const slide = byId.get(edit.slideId); if (!slide || !bindings.get(edit.slideId) || typeof slide.setHidden !== "function") throw new CliError("visibility-edits-target-unbound", "Slide visibility is missing, ambiguous or not source-bound.", 3);
    if (slide.hidden !== edit.expected) throw new CliError("visibility-edits-baseline-mismatch", "Expected visibility does not match the imported baseline.", 3);
    staged.push({ slide, replacement: edit.replacement });
  }
  for (const { slide, replacement } of staged) { checkAbort(signal); slide.setHidden!(replacement); }
}
