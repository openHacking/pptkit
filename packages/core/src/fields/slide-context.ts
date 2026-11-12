import { PresentationFieldError, type FieldEvaluationContext, type FieldSlideOrderOptions, type FieldSlideSequence } from "../types/field.js";
import { captureDefinitionInput } from "../document/definition-input.js";

/** Snapshot explicit document slide order; duplicate custom-show occurrences require caller-supplied occurrence IDs instead. */
export function fieldContextForSlide(document: FieldSlideSequence, options: FieldSlideOrderOptions): FieldEvaluationContext {
  const fail = (message: string): never => { throw new PresentationFieldError("invalid-field-context", message); };
  const slides = document?.slides;
  let captured: FieldSlideOrderOptions;
  try { captured = captureDefinitionInput(options, ["slideId", "firstPageNumber", "includeHidden", "order"]); }
  catch { return fail("Slide numbering options must be bounded plain data."); }
  if (!Array.isArray(slides) || !captured || typeof captured.slideId !== "string" || typeof captured.includeHidden !== "boolean" || !Number.isSafeInteger(captured.firstPageNumber) || captured.firstPageNumber < 0) return fail("Invalid document or slide numbering options.");
  if (slides.length > 100000 || (captured.order !== undefined && (!Array.isArray(captured.order) || captured.order.length > 100000))) return fail("Slide order exceeds the supported size.");
  const byId = new Map<string, boolean>();
  for (const slide of slides) {
    if (!slide || typeof slide.id !== "string" || !slide.id.trim() || slide.id.length > 512 || typeof slide.hidden !== "boolean" || byId.has(slide.id)) return fail("Document slide identities must be nonempty and unique.");
    byId.set(slide.id, slide.hidden);
  }
  const order = captured.order ?? [...byId.keys()], seen = new Set<string>(), selected: string[] = [];
  for (const id of order) {
    if (typeof id !== "string" || !byId.has(id) || seen.has(id)) return fail("The requested order contains an unknown or duplicate slide.");
    seen.add(id); if (captured.includeHidden || !byId.get(id)) selected.push(id);
  }
  const index = selected.indexOf(captured.slideId);
  if (index < 0 || !Number.isSafeInteger(captured.firstPageNumber + index)) return fail("The target slide is excluded, absent or its number overflows.");
  return Object.freeze({ pageOrder: Object.freeze(selected), pageId: captured.slideId, firstPageNumber: captured.firstPageNumber });
}
