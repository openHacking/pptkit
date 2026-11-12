import type { TextParagraphInput } from "../types/element.js";
import { PresentationFieldError, type FieldEvaluationContext, type FieldTextParagraph, type FieldTextRun, type MaterializedFieldText } from "../types/field.js";
import { captureDefinitionInput } from "../document/definition-input.js";
import { evaluatePresentationField } from "./evaluate.js";

/** Produce detached text and explicit native page-number runs with a separate evaluation ledger. */
export function materializeFieldText(paragraphs: readonly FieldTextParagraph[], context: FieldEvaluationContext = {}): MaterializedFieldText {
  const captured = captureDefinitionInput({ paragraphs, context }, ["paragraphs", "context"]);
  if (!Array.isArray(captured.paragraphs)) throw new PresentationFieldError("invalid-field", "Field text must contain paragraphs.");
  const fields: (MaterializedFieldText["fields"][number])[] = [];
  let pageNumber: string | undefined;
  const content: TextParagraphInput[] = captured.paragraphs.map((paragraph, paragraphIndex) => {
    if (!paragraph || typeof paragraph !== "object" || !Array.isArray(paragraph.runs) || Object.keys(paragraph).some(key => key !== "runs" && key !== "style")) throw new PresentationFieldError("invalid-field", "Invalid field text paragraph.");
    const runs = paragraph.runs.map((run: FieldTextRun, runIndex: number) => {
      if (!run || typeof run !== "object" || Object.keys(run).some(key => !["text", "field", "output", "style", "action"].includes(key)) || (run.text === undefined) === (run.field === undefined) || (run.field === undefined && run.output !== undefined) || (run.output !== undefined && run.output !== "static" && run.output !== "native")) throw new PresentationFieldError("invalid-field", "A field text run requires exactly one of text or field and a valid output mode.");
      let text: string;
      if (run.field !== undefined) {
        let value;
        if (run.field && run.field.kind === "slide-number" && pageNumber !== undefined) {
          if (typeof run.field.id !== "string" || !run.field.id.trim() || run.field.id.length > 512) throw new PresentationFieldError("invalid-field", "A field requires a nonempty identity of at most 512 characters.");
          value = { id: run.field.id, kind: run.field.kind, value: pageNumber };
        } else {
          value = evaluatePresentationField(run.field, captured.context);
          if (value.kind === "slide-number") pageNumber = value.value;
        }
        fields.push(Object.freeze({ ...value, paragraphIndex, runIndex })); text = value.value;
      } else {
        if (typeof run.text !== "string") throw new PresentationFieldError("invalid-field", "A text run requires a string.");
        text = run.text;
      }
      if (run.output === "native" && ((run.field?.kind !== "slide-number" && run.field?.kind !== "dynamic-date") || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(run.field.id) || (run.field.kind === "dynamic-date" && !run.field.updateFormat))) throw new PresentationFieldError("invalid-field", "Native output requires a UUID and a slide-number or explicit date-update format.");
      const nativeField = run.output === "native" && run.field?.kind === "slide-number" ? { id: run.field.id, kind: "slide-number" as const } : run.output === "native" && run.field?.kind === "dynamic-date" && run.field.updateFormat ? { id: run.field.id, kind: "dynamic-date" as const, updateFormat: run.field.updateFormat } : undefined;
      return { text, ...(nativeField ? { field: nativeField } : {}), ...(run.style === undefined ? {} : { style: run.style }), ...(run.action === undefined ? {} : { action: run.action }) };
    });
    return { runs, ...(paragraph.style === undefined ? {} : { style: paragraph.style }) };
  });
  return { content, fields: Object.freeze(fields) };
}
