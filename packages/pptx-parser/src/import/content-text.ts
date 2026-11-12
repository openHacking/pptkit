import type { DynamicDateUpdateFormat, ElementAction, TextFrameStyleInput, TextParagraphInput, TextRunInput } from "@pptkit/core";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { ContentReader } from "./content-reader.js";
import { effectiveTextFonts, frameTextStyles, frameTextValues, listTextStyles, mergeTextStyles, paragraphTextStyles, paragraphTextValues, runTextStyles, runTextValues, type TextStyleContext, type TextStylePatch } from "./content-text-style.js";

const dateFieldTypes: Readonly<Record<string, DynamicDateUpdateFormat>> = { datetime: "application-default", datetime1: "month-day-year", datetime2: "weekday-month-day-year", datetime3: "day-month-year", datetime4: "month-day-year-long", datetimeFigureOut: "automatic-month-day-year" };
const timeFieldTypes: Readonly<Record<string, DynamicDateUpdateFormat>> = { datetime5: "day-short-month-year", datetime6: "month-year", datetime7: "short-month-year", datetime8: "month-day-year-time-12", datetime9: "month-day-year-time-12-seconds", datetime10: "time-24", datetime11: "time-24-seconds", datetime12: "time-12", datetime13: "time-12-seconds" };

export interface ContentTextInheritance {
  local: TextStyleContext;
  frame: TextStylePatch;
  paragraph: (level: number) => TextStylePatch;
  record: (node: XmlElement, patch: TextStylePatch, prefix: "frame." | "paragraph." | "run.") => void;
}
export interface ContentTextRun {
  readonly paragraphIndex: number;
  readonly runIndex: number;
  readonly text: string;
  readonly node?: XmlElement;
  readonly editable?: boolean;
}

export async function contentText(
  body: XmlElement, r: ContentReader, checkpoint: () => Promise<void>,
  readAction: (node: XmlElement) => ElementAction | undefined,
  includeFrame = true, inheritance?: ContentTextInheritance,
  recordRun?: (run: ContentTextRun) => void,
  fieldTextProfile?: "slide-number-v1" | "date-and-number-v1" | "date-time-and-number-v1",
): Promise<{ content: TextParagraphInput[]; frame: TextFrameStyleInput }> {
  r.known(body);
  const local: TextStyleContext = inheritance?.local ?? { r, level: "slide-list-style", at: () => ({ part: r.part }) };
  const frameContext = { ...local, level: "slide-frame" as const };
  // Notes and table callers expose content only and keep frame semantics in
  // their owning reader's audit; source placeholder inheritance is opt-in.
  const framePatch = includeFrame ? mergeTextStyles(inheritance?.frame ?? new Map(), frameTextStyles(r.one(body, r.a, "bodyPr"), frameContext)) : new Map();
  if (includeFrame) inheritance?.record(body, framePatch, "frame.");
  const list = await listTextStyles(r.one(body, r.a, "lstStyle"), local);
  const paragraphs: TextParagraphInput[] = [];
  for (const paragraph of children(body).filter(n => n.namespaceUri === r.a && n.local === "p")) {
    await checkpoint(); r.known(paragraph);
    const properties = r.one(paragraph, r.a, "pPr");
    const level = properties ? r.number(properties, "lvl", 0, 8, 0) : 0;
    const defaults = mergeTextStyles(
      inheritance?.paragraph(level) ?? new Map(), list.defaults, list.levels[level]!,
      await paragraphTextStyles(properties, { ...local, level: "paragraph" }),
    );
    inheritance?.record(paragraph, defaults, "paragraph.");
    const runs: TextRunInput[] = [];
    for (const node of children(paragraph)) {
      await checkpoint();
      if (node.namespaceUri !== r.a || !["r", "br", ...(fieldTextProfile ? ["fld"] : [])].includes(node.local)) continue;
      let field: TextRunInput["field"];
      if (node.local === "fld") {
        const id = attribute(node, "id"), type = attribute(node, "type"), items = children(node);
        const dateFormat = type && fieldTextProfile !== "slide-number-v1" && Object.hasOwn(dateFieldTypes, type) ? dateFieldTypes[type] : type && fieldTextProfile === "date-time-and-number-v1" && Object.hasOwn(timeFieldTypes, type) ? timeFieldTypes[type] : undefined;
        if ((type !== "slidenum" && !dateFormat) || !id || !/^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$/i.test(id) || node.attributes.some(a => a.namespaceUri !== "http://www.w3.org/2000/xmlns/" && (a.namespaceUri !== "" || !["id", "type"].includes(a.local))) || node.children.some(child => (child.kind === "text" || child.kind === "cdata") && !/^[\t\r\n ]*$/.test(child.value)) || items.some(item => item.namespaceUri !== r.a || !["rPr", "t"].includes(item.local)) || items.filter(item => item.local === "t").length !== 1 || items.filter(item => item.local === "rPr").length > 1 || items.some(item => item.local === "t" && children(item).length > 0)) continue;
        r.attr(node, "id"); r.attr(node, "type"); field = type === "slidenum" ? { id: id.slice(1, -1), kind: "slide-number" } : { id: id.slice(1, -1), kind: "dynamic-date", updateFormat: dateFormat! };
      }
      r.known(node);
      const props = r.one(node, r.a, "rPr"), action = props ? readAction(props) : undefined;
      const patch = effectiveTextFonts(mergeTextStyles(defaults, await runTextStyles(props, { ...local, level: "run" })), local);
      inheritance?.record(node, patch, "run.");
      const textNode = node.local === "r" || node.local === "fld" ? r.one(node, r.a, "t", true)! : undefined;
      const value = textNode ? r.text(textNode) : "\n";
      recordRun?.({ paragraphIndex: paragraphs.length, runIndex: runs.length, text: value, ...(textNode ? { node: textNode } : {}), ...(field ? { editable: false } : {}) });
      runs.push({ text: value, style: runTextValues(patch), ...(field ? { field } : {}), ...(action ? { action } : {}) });
    }
    // End properties style empty insertion text; never the previous real run.
    const end = r.one(paragraph, r.a, "endParaRPr");
    if (end) r.known(end);
    if (!runs.length) {
      const patch = effectiveTextFonts(mergeTextStyles(defaults, await runTextStyles(end, { ...local, level: "end-paragraph" })), local);
      inheritance?.record(end ?? paragraph, patch, "run.");
      runs.push({ text: "", style: runTextValues(patch) });
      recordRun?.({ paragraphIndex: paragraphs.length, runIndex: 0, text: "" });
    }
    paragraphs.push({ runs, style: paragraphTextValues(defaults) });
  }
  if (!paragraphs.length) invalidStructure("Text body requires at least one paragraph.", r.part);
  return { content: paragraphs, frame: frameTextValues(framePatch) };
}
