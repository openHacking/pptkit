import { captureDefinitionInput } from "../document/definition-input.js";
import { PresentationFieldError, type FieldEvaluationContext, type FieldOccurrenceOptions, type FieldSlideSequence } from "../types/field.js";
import { fieldContextForSlide } from "./slide-context.js";

/** Number explicit occurrences without treating repeated visits as a single slide. */
export function fieldContextForOccurrences(sequence: FieldSlideSequence, options: FieldOccurrenceOptions): FieldEvaluationContext {
  const fail = (message: string): never => { throw new PresentationFieldError("invalid-field-context", message); };
  let captured: FieldOccurrenceOptions;
  try { captured = captureDefinitionInput(options, ["occurrences", "occurrenceId", "firstPageNumber", "includeHidden"]); }
  catch { return fail("Occurrence options must be bounded plain data."); }
  const sourceSlides = sequence?.slides;
  if (!Array.isArray(sourceSlides) || sourceSlides.length > 100000 || !Array.isArray(captured.occurrences) || captured.occurrences.length > 100000) return fail("Invalid slide sequence or occurrence list.");
  const slides = new Map<string, boolean>();
  for (const slide of sourceSlides) {
    if (!slide || typeof slide.id !== "string" || !slide.id.trim() || slide.id.length > 512 || typeof slide.hidden !== "boolean" || slides.has(slide.id)) return fail("Slide identities must be nonempty and unique.");
    slides.set(slide.id, slide.hidden);
  }
  const mapped = captured.occurrences.map(occurrence => {
    if (!occurrence || typeof occurrence !== "object" || Object.keys(occurrence).some(key => key !== "id" && key !== "slideId") || typeof occurrence.slideId !== "string" || !slides.has(occurrence.slideId)) return fail("An occurrence references an unknown slide.");
    return { id: occurrence.id, hidden: slides.get(occurrence.slideId)! };
  });
  return fieldContextForSlide({ slides: mapped }, { slideId: captured.occurrenceId, firstPageNumber: captured.firstPageNumber, includeHidden: captured.includeHidden });
}
