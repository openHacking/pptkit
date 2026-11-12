import type { ImportedPresentation } from "@pptkit/pptx-parser";
import type { PresentationElement, PresentationElementInput, PresentationSlide } from "@pptkit/core";
import { CliError } from "./arguments.js";
import { checkAbort } from "./read-file.js";

const record = (v: unknown): v is Record<string, unknown> => typeof v === "object" && v !== null && !Array.isArray(v);
const key = (slideId: string, elementId: string) => JSON.stringify([slideId, elementId]);
type PropertyKind = "description" | "name" | "decorative";
type Owner = { slide: PresentationSlide; root: PresentationElement; path: number[]; element: PresentationElement };
function owners(session: ImportedPresentation, signal?: AbortSignal): Map<string, Owner | null> {
  const result = new Map<string, Owner | null>();
  for (const slide of session.document.slides) for (const root of slide.elements) {
    const pending: { element: PresentationElement; path: number[] }[] = [{ element: root, path: [] }];
    while (pending.length) {
      checkAbort(signal); const { element, path } = pending.pop()!, id = key(slide.id, element.id);
      result.set(id, result.has(id) ? null : { slide, root, path, element });
      if (element.type === "group") for (let i = element.children.length - 1; i >= 0; i--) pending.push({ element: element.children[i]!, path: [...path, i] });
    }
  }
  return result;
}
function bindings(session: ImportedPresentation, signal?: AbortSignal, kind: PropertyKind = "description"): Map<string, boolean> {
  const result = new Map<string, boolean>();
  for (const finding of session.report.findings) {
    checkAbort(signal); if (finding.feature !== (kind === "name" ? "drawing.name-existing" : kind === "decorative" ? "drawing.decorative-existing" : "drawing.description-attribute") || !finding.slideId || !finding.elementId) continue;
    const id = key(finding.slideId, finding.elementId);
    if (result.has(id)) throw new CliError(`${kind}-candidates-inconsistent`, "Drawing property target identity is ambiguous.", 3);
    result.set(id, finding.tier === "editable");
  }
  return result;
}
function propertyCandidates(session: ImportedPresentation, sourceSha256: string, signal: AbortSignal | undefined, kind: PropertyKind) {
  const found = owners(session, signal), targets: { slideId: string; elementId: string; expected: string | boolean | null }[] = []; let unboundCount = 0;
  for (const [id, bound] of bindings(session, signal, kind)) {
    checkAbort(signal); if (!bound) { unboundCount++; continue; }
    const owner = found.get(id); if (!owner) throw new CliError(`${kind}-candidates-inconsistent`, "Bound drawing property owner is missing or ambiguous.", 3);
    const expected = kind === "name" ? owner.element.name : kind === "decorative" ? owner.element.accessibility?.decorative : owner.element.accessibility?.description ?? null;
    if (expected === undefined) throw new CliError(`${kind}-candidates-inconsistent`, "Bound drawing property is unavailable.", 3);
    targets.push({ slideId: owner.slide.id, elementId: owner.element.id, expected });
  }
  return { version: `pptkit-${kind}-edits-v1`, sourceSha256, targets, unboundCount };
}
function parsePropertyEdits(bytes: Uint8Array, sourceSha256: string, kind: PropertyKind): readonly unknown[] {
  let value: unknown; try { value = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)); }
  catch { throw new CliError(`${kind}-edits-invalid`, "Drawing property edits must be valid UTF-8 JSON.", 3); }
  if (!record(value) || value.version !== `pptkit-${kind}-edits-v1` || Object.keys(value).some(k => !["version", "sourceSha256", "edits"].includes(k)) || typeof value.sourceSha256 !== "string" || !/^[0-9a-f]{64}$/.test(value.sourceSha256) || !Array.isArray(value.edits) || !value.edits.length || value.edits.length > 10000) throw new CliError(`${kind}-edits-invalid`, "Expected a bounded drawing property edit manifest.", 3);
  if (value.sourceSha256 !== sourceSha256) throw new CliError(`${kind}-edits-source-mismatch`, "Drawing property edits do not match the captured source SHA-256.", 3);
  return value.edits as unknown[];
}
function applyPropertyEdits(session: ImportedPresentation, edits: readonly unknown[], signal: AbortSignal | undefined, kind: PropertyKind): void {
  const found = owners(session, signal), available = bindings(session, signal, kind), seen = new Set<string>();
  const staged = new Map<string, { slide: PresentationSlide; rootId: string; root: PresentationElementInput }>();
  for (const edit of edits) {
    checkAbort(signal);
    if (!record(edit) || Object.keys(edit).some(k => !["slideId", "elementId", "expected", "replacement"].includes(k)) || typeof edit.slideId !== "string" || !edit.slideId || typeof edit.elementId !== "string" || !edit.elementId || !(kind === "decorative" ? typeof edit.expected === "boolean" : typeof edit.expected === "string" || kind === "description" && edit.expected === null) || !(kind === "decorative" ? typeof edit.replacement === "boolean" : typeof edit.replacement === "string" || kind === "description" && edit.replacement === null)) throw new CliError(`${kind}-edits-invalid`, kind === "name" ? "Name edits require an identity and string values." : kind === "decorative" ? "Decorative edits require an identity and boolean values." : "Description edits require an identity and string or null values.", 3);
    const id = key(edit.slideId, edit.elementId); if (seen.has(id)) throw new CliError(`${kind}-edits-invalid`, "Duplicate drawing property target in manifest.", 3); seen.add(id);
    const owner = found.get(id); if (!owner || !available.get(id)) throw new CliError(`${kind}-edits-target-unbound`, "Drawing property owner is missing, ambiguous or not source-bound.", 3);
    if ((kind === "name" ? owner.element.name : kind === "decorative" ? owner.element.accessibility?.decorative : owner.element.accessibility?.description ?? null) !== edit.expected) throw new CliError(`${kind}-edits-baseline-mismatch`, "Expected drawing property does not match the imported baseline.", 3);
    const rootKey = key(edit.slideId, owner.root.id); let item = staged.get(rootKey);
    if (!item) { item = { slide: owner.slide, rootId: owner.root.id, root: structuredClone(owner.root) as PresentationElementInput }; staged.set(rootKey, item); }
    let element = item.root; for (const index of owner.path) { if (element.type !== "group" || !element.children[index]) throw new CliError(`${kind}-edits-target-unbound`, "Drawing property ancestry is inconsistent.", 3); element = element.children[index]!; }
    if (element.id !== edit.elementId) throw new CliError(`${kind}-edits-target-unbound`, "Drawing property ancestry is inconsistent.", 3);
    if (kind === "name") element.name = edit.replacement as string;
    else if (kind === "decorative") element.accessibility = { ...element.accessibility, decorative: edit.replacement as boolean };
    else if (edit.replacement === null) { if (element.accessibility) delete element.accessibility.description; }
    else element.accessibility = { ...element.accessibility, description: edit.replacement as string };
  }
  for (const item of staged.values()) { checkAbort(signal); item.slide.replaceElement(item.rootId, item.root); }
}

export const descriptionCandidates = (session: ImportedPresentation, hash: string, signal?: AbortSignal) => propertyCandidates(session, hash, signal, "description");
export const nameCandidates = (session: ImportedPresentation, hash: string, signal?: AbortSignal) => propertyCandidates(session, hash, signal, "name");
export const parseDescriptionEdits = (bytes: Uint8Array, hash: string) => parsePropertyEdits(bytes, hash, "description");
export const parseNameEdits = (bytes: Uint8Array, hash: string) => parsePropertyEdits(bytes, hash, "name");
export const applyDescriptionEdits = (session: ImportedPresentation, edits: readonly unknown[], signal?: AbortSignal) => applyPropertyEdits(session, edits, signal, "description");
export const applyNameEdits = (session: ImportedPresentation, edits: readonly unknown[], signal?: AbortSignal) => applyPropertyEdits(session, edits, signal, "name");
export const decorativeCandidates = (session: ImportedPresentation, hash: string, signal?: AbortSignal) => propertyCandidates(session, hash, signal, "decorative");
export const parseDecorativeEdits = (bytes: Uint8Array, hash: string) => parsePropertyEdits(bytes, hash, "decorative");
export const applyDecorativeEdits = (session: ImportedPresentation, edits: readonly unknown[], signal?: AbortSignal) => applyPropertyEdits(session, edits, signal, "decorative");
