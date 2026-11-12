import { evaluatePresentationField, PresentationFieldError, type NormalizedElement, type NormalizedPresentation, type NormalizedTextParagraph, type PresentationField } from "@pptkit/core";
import type { LayoutElement, LayoutResult } from "../types/layout.js";
import type { LayoutOptions } from "../types/analysis.js";

/** Update caches on a detached document or layout, retaining shared definitions. */
export function refreshLayoutFields(layout: LayoutResult | NormalizedPresentation, options: NonNullable<LayoutOptions["fieldRefresh"]>): void {
  if (!options || typeof options !== "object" || !Number.isSafeInteger(options.firstPageNumber) || options.firstPageNumber < 0 || typeof options.includeHidden !== "boolean" || Object.keys(options).some(key => !["firstPageNumber", "includeHidden", "instantMs", "locale", "timeZone", "dateStyle", "timeStyle"].includes(key))) throw new PresentationFieldError("invalid-field-context", "Field refresh requires an explicit first number and hidden-slide policy.");
  if ((options.instantMs !== undefined && (typeof options.instantMs !== "number" || !Number.isInteger(options.instantMs) || Math.abs(options.instantMs) > 8640000000000000)) || (options.locale !== undefined && (typeof options.locale !== "string" || !options.locale)) || (options.timeZone !== undefined && (typeof options.timeZone !== "string" || !options.timeZone)) || (options.dateStyle !== undefined && !["short", "medium", "long", "full"].includes(options.dateStyle)) || (options.timeStyle !== undefined && !["short", "medium", "long", "full"].includes(options.timeStyle))) throw new PresentationFieldError("invalid-field-context", "Invalid optional field refresh date/time configuration.");
  const plain = (paragraphs: readonly NormalizedTextParagraph[]) => paragraphs.map(paragraph => paragraph.runs.map(run => run.text).join("")).join("\n");
  const visit = (elements: readonly (LayoutElement | NormalizedElement)[], action: (paragraphs: NormalizedTextParagraph[]) => void, visibleOnly = false) => {
    const stack = [...elements];
    while (stack.length) {
      const element = stack.pop()!;
      if (visibleOnly && element.hidden) continue;
      if (element.type === "group") for (const child of element.children) stack.push(child);
      else if (element.type === "text") { action(element.content); element.plainText = plain(element.content); }
      else if (element.type === "shape" && element.text) { action(element.text.content); element.text.plainText = plain(element.text.content); }
      else if (element.type === "table") for (const row of element.rows) for (const cell of row.cells) action(cell.content);
    }
  };
  const refreshDate = (field: NonNullable<NormalizedTextParagraph["runs"][number]["field"]>): string => {
    if (field.kind !== "dynamic-date") throw new PresentationFieldError("invalid-field", "Unsupported dynamic field kind.");
    const style = options.dateStyle !== undefined ? { dateStyle: options.dateStyle, ...(options.timeStyle !== undefined ? { timeStyle: options.timeStyle } : {}) } : options.timeStyle !== undefined ? { timeStyle: options.timeStyle } : undefined;
    if (!style) throw new PresentationFieldError("missing-field-context", "Refreshing a date field requires a date or time cache style.");
    return evaluatePresentationField({ id: field.id, kind: "dynamic-date", updateFormat: field.updateFormat, ...style } as PresentationField, { ...(options.instantMs !== undefined ? { instantMs: options.instantMs } : {}), ...(options.locale !== undefined ? { locale: options.locale } : {}), ...(options.timeZone !== undefined ? { timeZone: options.timeZone } : {}) }).value;
  };
  const checkId = (id: string) => { if (typeof id !== "string" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) throw new PresentationFieldError("invalid-field", "A refreshed dynamic field requires a UUID."); };
  const refreshShared = (paragraphs: NormalizedTextParagraph[]) => {
    for (const paragraph of paragraphs) for (const run of paragraph.runs) {
      if (!run.field) continue;
      checkId(run.field.id);
      if (run.field.kind === "slide-number") throw new PresentationFieldError("invalid-field-context", "Refreshing page numbers in shared master or layout artwork requires per-slide overrides.");
      run.text = refreshDate(run.field);
    }
  };
  const layouts = new Map(layout.layouts.map(item => [item.id, item])), masters = new Map((layout.masters ?? []).map(item => [item.id, item])), seenLayouts = new Set<string>(), seenMasters = new Set<string>();
  for (const slide of layout.slides) {
    if ((!options.includeHidden && slide.hidden) || slide.showInheritedShapes === false) continue;
    const definition = layouts.get(slide.layoutId); if (!definition) continue;
    if (!seenLayouts.has(definition.id)) { seenLayouts.add(definition.id); visit(definition.elements, refreshShared, true); }
    if (definition.showMasterShapes !== false && definition.masterId && !seenMasters.has(definition.masterId)) {
      const master = masters.get(definition.masterId); if (master) { seenMasters.add(master.id); visit(master.elements ?? [], refreshShared, true); }
    }
  }
  let number = options.firstPageNumber;
  for (const slide of layout.slides) {
    if (!options.includeHidden && slide.hidden) continue;
    if (!Number.isSafeInteger(number)) throw new PresentationFieldError("invalid-field-context", "Final slide numbering exceeds the safe integer range.");
    const pageNumber = String(number++);
    const refresh = (paragraphs: NormalizedTextParagraph[]) => {
      for (const paragraph of paragraphs) for (const run of paragraph.runs) {
        if (!run.field) continue;
        checkId(run.field.id);
        if (run.field.kind === "slide-number") run.text = pageNumber;
        else run.text = refreshDate(run.field);
      }
    };
    visit(slide.elements, refresh); refresh(slide.notes);
  }
}
