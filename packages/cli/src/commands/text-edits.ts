import type { ImportedPresentation } from "@pptkit/pptx-parser";
import type { PresentationSlide, PresentationElement, PresentationElementInput, TextParagraphInput } from "@pptkit/core";
import { CliError } from "./arguments.js";
import { checkAbort } from "./read-file.js";

const record = (v: unknown): v is Record<string, unknown> => typeof v === "object" && v !== null && !Array.isArray(v);
export function parseTextEdits(bytes: Uint8Array, sourceSha256: string) {
  let value: unknown;
  try { value = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)); }
  catch { throw new CliError("text-edits-invalid", "Text edits must be valid UTF-8 JSON.", 3); }
  if (!record(value) || value.version !== "pptkit-text-edits-v1" && value.version !== "pptkit-text-edits-v2" && value.version !== "pptkit-text-edits-v3" || Object.keys(value).some(k => !["version", "sourceSha256", "edits"].includes(k)) || typeof value.sourceSha256 !== "string" || !/^[0-9a-f]{64}$/.test(value.sourceSha256) || !Array.isArray(value.edits) || !value.edits.length || value.edits.length > 10000) throw new CliError("text-edits-invalid", "Expected a bounded pptkit-text-edits-v1, v2 or v3 manifest.", 3);
  if (value.sourceSha256 !== sourceSha256) throw new CliError("text-edits-source-mismatch", "Text edits do not match the captured source SHA-256.", 3);
  return { version: value.version as "pptkit-text-edits-v1" | "pptkit-text-edits-v2" | "pptkit-text-edits-v3", edits: value.edits as unknown[] };
}

/** Apply explicit, baseline-bound edits only after all targets have been checked. */
export function applyTextEdits(session: ImportedPresentation, plan: ReturnType<typeof parseTextEdits>, signal?: AbortSignal): number {
  const staged = new Map<string, { slide: PresentationSlide; rootId: string; root: PresentationElementInput }>(), targets = new Set<string>();
  const notes = new Map<string, { slide: PresentationSlide; content: TextParagraphInput[] }>();
  const slides = new Map(session.document.slides.map(slide => [slide.id, slide]));
  const owners = new Map<string, { slide: PresentationSlide; root: PresentationElement; path: number[] } | null>();
  for (const slide of session.document.slides) for (const root of slide.elements) {
    const pending: { element: PresentationElement; path: number[] }[] = [{ element: root, path: [] }];
    while (pending.length) {
      checkAbort(signal); const { element, path } = pending.pop()!;
      if (element.type === "text" || element.type === "shape" && element.text || element.type === "table") {
        const key = JSON.stringify([slide.id, element.id]);
        owners.set(key, owners.has(key) ? null : { slide, root, path });
      } else if (element.type === "group") {
        for (let i = element.children.length - 1; i >= 0; i--) pending.push({ element: element.children[i]!, path: [...path, i] });
      }
    }
  }
  const bindings = new Map<string, { count: number; bound: boolean }>();
  for (const binding of session.report.textRuns ?? []) {
    checkAbort(signal);
    const key = JSON.stringify([binding.owner ?? "text-box", binding.slideId, binding.elementId, binding.paragraphIndex, binding.runIndex]), previous = bindings.get(key);
    bindings.set(key, { count: (previous?.count ?? 0) + 1, bound: binding.status === "bound" });
  }
  for (const binding of session.report.notesTextRuns ?? []) {
    checkAbort(signal); const key = JSON.stringify(["notes", binding.slideId, null, binding.paragraphIndex, binding.runIndex]), previous = bindings.get(key);
    bindings.set(key, { count: (previous?.count ?? 0) + 1, bound: binding.status === "bound" });
  }
  for (const binding of session.report.tableTextRuns ?? []) {
    checkAbort(signal); const key = JSON.stringify(["table-cell", binding.slideId, binding.elementId, binding.rowIndex, binding.cellIndex, binding.columnIndex, binding.paragraphIndex, binding.runIndex]), previous = bindings.get(key);
    bindings.set(key, { count: (previous?.count ?? 0) + 1, bound: binding.status === "bound" });
  }
  for (const edit of plan.edits) {
    checkAbort(signal);
    if (!record(edit)) throw new CliError("text-edits-invalid", "Each text edit must be an object.", 3);
    const owner = plan.version === "pptkit-text-edits-v1" ? "text-box" : edit.owner;
    const index = (v: unknown): v is number => Number.isSafeInteger(v) && (v as number) >= 0;
    if (Object.keys(edit).some(k => !["slideId", "elementId", "paragraphIndex", "runIndex", "expected", "replacement", ...(plan.version !== "pptkit-text-edits-v1" ? ["owner"] : []), ...(owner === "table-cell" && plan.version === "pptkit-text-edits-v3" ? ["rowIndex", "cellIndex", "columnIndex"] : [])].includes(k)) || typeof owner !== "string" || !["text-box", "shape-text", "notes", ...(plan.version === "pptkit-text-edits-v3" ? ["table-cell"] : [])].includes(owner) || typeof edit.slideId !== "string" || !edit.slideId || (owner === "notes" ? edit.elementId !== undefined : typeof edit.elementId !== "string" || !edit.elementId) || !index(edit.paragraphIndex) || !index(edit.runIndex) || owner === "table-cell" && (!index(edit.rowIndex) || !index(edit.cellIndex) || !index(edit.columnIndex)) || typeof edit.expected !== "string" || typeof edit.replacement !== "string") throw new CliError("text-edits-invalid", "Each text edit needs a valid owner, target, indices, expected text and replacement.", 3);
    const paragraph = edit.paragraphIndex as number, run = edit.runIndex as number, key = owner === "table-cell" ? JSON.stringify([owner, edit.slideId, edit.elementId, edit.rowIndex, edit.cellIndex, edit.columnIndex, paragraph, run]) : JSON.stringify([owner, edit.slideId, owner === "notes" ? null : edit.elementId, paragraph, run]);
    if (targets.has(key)) throw new CliError("text-edits-invalid", "Duplicate text target in manifest.", 3);
    targets.add(key);
    const binding = bindings.get(key);
    if (!binding || binding.count !== 1 || !binding.bound) throw new CliError("text-edits-target-unbound", "Text target is missing, ambiguous or not source-bound.", 3);
    if (owner === "notes") {
      let item = notes.get(edit.slideId);
      if (!item) {
        const slide = slides.get(edit.slideId);
        if (!slide || !Array.isArray(slide.notes)) throw new CliError("text-edits-target-unbound", "Notes body is unavailable.", 3);
        item = { slide, content: structuredClone(slide.notes) }; notes.set(edit.slideId, item);
      }
      const text = item.content[paragraph]?.runs[run];
      if (!text || text.text !== edit.expected) throw new CliError("text-edits-baseline-mismatch", "Expected text does not match the imported baseline.", 3);
      text.text = edit.replacement; continue;
    }
    const target = owners.get(JSON.stringify([edit.slideId, edit.elementId]));
    if (!target) throw new CliError("text-edits-target-unbound", "Text owner is missing or ambiguous in the imported slide.", 3);
    const rootKey = JSON.stringify([edit.slideId, target.root.id]); let item = staged.get(rootKey);
    if (!item) {
      item = { slide: target.slide, rootId: target.root.id, root: structuredClone(target.root) as PresentationElementInput }; staged.set(rootKey, item);
    }
    let element: PresentationElementInput = item.root;
    for (const index of target.path) {
      if (element.type !== "group" || !element.children[index]) throw new CliError("text-edits-target-unbound", "Imported text ancestry is inconsistent.", 3);
      element = element.children[index]!;
    }
    if (element.id !== edit.elementId || (owner === "text-box" ? element.type !== "text" : owner === "shape-text" ? element.type !== "shape" || !element.text : element.type !== "table")) throw new CliError("text-edits-target-unbound", "Imported text ancestry is inconsistent.", 3);
    const content = element.type === "text" ? element.content : element.type === "shape" ? element.text?.content : element.type === "table" ? element.rows[edit.rowIndex as number]?.cells[edit.cellIndex as number]?.content : undefined;
    const text = !Array.isArray(content) ? undefined : content[paragraph]?.runs[run];
    if (!text || text.text !== edit.expected) throw new CliError("text-edits-baseline-mismatch", "Expected text does not match the imported baseline.", 3);
    text.text = edit.replacement;
  }
  checkAbort(signal);
  for (const item of staged.values()) { checkAbort(signal); item.slide.replaceElement(item.rootId, item.root); }
  for (const item of notes.values()) { checkAbort(signal); item.slide.setNotes(item.content); }
  return plan.edits.length;
}
