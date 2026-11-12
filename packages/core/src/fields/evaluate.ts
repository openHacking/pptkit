import { PresentationFieldError, type EvaluatedPresentationField, type FieldEvaluationContext, type PresentationField } from "../types/field.js";
import { isDynamicDateUpdateFormat } from "./date-format.js";

/** Pure explicit-context evaluation; does not mutate or attach a field to a document. */
export function evaluatePresentationField(field: PresentationField, context: FieldEvaluationContext = {}): EvaluatedPresentationField {
  const fail = (code: "invalid-field" | "missing-field-context" | "invalid-field-context", message: string): never => { throw new PresentationFieldError(code, message); };
  if (!field || typeof field !== "object" || typeof field.id !== "string" || !field.id.trim() || field.id.length > 512) return fail("invalid-field", "A field requires a nonempty identity of at most 512 characters.");
  if (!context || typeof context !== "object") return fail("invalid-field-context", "Field context must be an object.");
  let value: string;
  if (field.kind === "footer" || field.kind === "static-date") {
    if (typeof field.text !== "string") return fail("invalid-field", "Static fields require text.");
    value = field.text;
  } else if (field.kind === "slide-number") {
    if (context.pageOrder === undefined || context.pageId === undefined || context.firstPageNumber === undefined) return fail("missing-field-context", "Slide number requires a final page order, page identity and first page number.");
    if (!Array.isArray(context.pageOrder) || context.pageOrder.length > 100000 || typeof context.pageId !== "string" || !Number.isSafeInteger(context.firstPageNumber) || context.firstPageNumber < 0) return fail("invalid-field-context", "Invalid page order, page identity or first page number.");
    const seen = new Set<string>(); let index = -1;
    for (let i = 0; i < context.pageOrder.length; i++) {
      const id = context.pageOrder[i];
      if (typeof id !== "string" || !id.trim() || id.length > 512 || seen.has(id)) return fail("invalid-field-context", "Page occurrence identities must be nonempty and unique.");
      seen.add(id); if (id === context.pageId) index = i;
    }
    if (index < 0 || !Number.isSafeInteger(context.firstPageNumber + index)) return fail("invalid-field-context", "The page is absent from the order or its number overflows.");
    value = String(context.firstPageNumber + index);
  } else if (field.kind === "dynamic-date") {
    const validStyle = (style: unknown) => typeof style === "string" && ["short", "medium", "long", "full"].includes(style);
    if ((field.dateStyle === undefined && field.timeStyle === undefined) || (field.dateStyle !== undefined && !validStyle(field.dateStyle)) || (field.timeStyle !== undefined && !validStyle(field.timeStyle))) return fail("invalid-field", "Dynamic date requires a supported date or time style.");
    if (field.updateFormat !== undefined && !isDynamicDateUpdateFormat(field.updateFormat)) return fail("invalid-field", "Invalid native date-update format.");
    if (context.instantMs === undefined || context.locale === undefined || context.timeZone === undefined) return fail("missing-field-context", "Dynamic date requires an instant, locale and time zone.");
    if (typeof context.instantMs !== "number" || !Number.isFinite(context.instantMs) || !Number.isInteger(context.instantMs) || Math.abs(context.instantMs) > 8640000000000000 || typeof context.locale !== "string" || !context.locale || typeof context.timeZone !== "string" || !context.timeZone) return fail("invalid-field-context", "Invalid date context.");
    try {
      if (Intl.DateTimeFormat.supportedLocalesOf([context.locale]).length !== 1) return fail("invalid-field-context", "Requested locale is unavailable.");
      value = new Intl.DateTimeFormat(context.locale, { ...(field.dateStyle !== undefined ? { dateStyle: field.dateStyle } : {}), ...(field.timeStyle !== undefined ? { timeStyle: field.timeStyle } : {}), timeZone: context.timeZone }).format(context.instantMs);
    } catch (error) {
      if (error instanceof PresentationFieldError) throw error;
      return fail("invalid-field-context", "The locale or time zone is unavailable or invalid.");
    }
  } else return fail("invalid-field", "Unknown field kind.");
  return Object.freeze({ id: field.id, kind: field.kind, value });
}
