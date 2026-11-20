import type { structureContext } from "./structure-context.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { PptxDiagramCacheMapping, PptxDiagramDataModelInfo, PptxDiagramTextComparison } from "../types/smartart.js";
import { children } from "./structure-xml.js";

type Context = Awaited<ReturnType<typeof structureContext>>;
type TextResult = { paragraphs: string[] } | { reason: string };
const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";

/** Observed direct text only. Explicit associations never become semantic ownership/freshness. */
export async function compareDiagramText(c: Context, models: readonly PptxDiagramDataModelInfo[], mappings: readonly PptxDiagramCacheMapping[],
  checkpoint: (part: string, units?: number) => Promise<void>, chargeFinding: (part: string) => void,
): Promise<PptxDiagramTextComparison[]> {
  const indices = new Map<string, Map<number, XmlElement>>(), sourceTexts = new Map<string, Map<number, TextResult>>(), cacheTexts = new Map<string, Map<number, TextResult>>();
  const index = async (part: string) => {
    const previous = indices.get(part); if (previous) return previous;
    const result = new Map<number, XmlElement>(), root = c.rawXml.get(part), stack = root ? [root] : []; let ordinal = 0;
    while (stack.length) { await checkpoint(part); const node = stack.pop()!; result.set(++ordinal, node); for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === "element") stack.push(child); } }
    indices.set(part, result); return result;
  };
  const hasOpaqueCharacters = (node: XmlElement) => node.children.some(child => (child.kind === "text" || child.kind === "cdata") && /[^\t\r\n ]/.test(child.value));
  const readText = async (part: string, node: XmlElement | undefined, bodyNamespace: string, bodyLocal: string): Promise<TextResult> => {
    if (!node) return { reason: "Source XML identity is unavailable." };
    const bodies = children(node).filter(child => child.namespaceUri === bodyNamespace && child.local === bodyLocal);
    if (bodies.length !== 1) return { reason: "Expected exactly one direct text body; absent/ambiguous bodies are not inferred as empty text." };
    const body = bodies[0]!, paragraphs: string[] = []; let bodyPr = 0, listStyle = 0;
    if (hasOpaqueCharacters(body)) return { reason: "Text body has uninterpreted direct character content." };
    for (const paragraph of children(body)) {
      await checkpoint(part);
      if (paragraph.namespaceUri !== c.a) return { reason: "Text body contains an unrecognized namespace." };
      if (paragraph.local === "bodyPr") { if (++bodyPr > 1) return { reason: "Duplicate text body properties." }; continue; }
      if (paragraph.local === "lstStyle") { if (++listStyle > 1) return { reason: "Duplicate text list properties." }; continue; }
      if (paragraph.local !== "p" || hasOpaqueCharacters(paragraph)) return { reason: "Text body is outside direct paragraph coverage." };
      let text = "", pPr = 0, endPr = 0;
      for (const run of children(paragraph)) {
        await checkpoint(part);
        if (run.namespaceUri !== c.a) return { reason: "Paragraph contains an unrecognized namespace." };
        if (run.local === "pPr") { if (++pPr > 1) return { reason: "Duplicate paragraph properties." }; continue; }
        if (run.local === "endParaRPr") { if (++endPr > 1) return { reason: "Duplicate paragraph end properties." }; continue; }
        if (!["r", "br"].includes(run.local) || hasOpaqueCharacters(run)) return { reason: "Fields or uninterpreted run content are outside direct character coverage." };
        const parts = children(run); let rPr = 0, leaves = 0, value = "";
        for (const child of parts) {
          await checkpoint(part);
          if (child.namespaceUri !== c.a) return { reason: "Run contains an unrecognized namespace." };
          if (child.local === "rPr") { if (++rPr > 1) return { reason: "Duplicate run properties." }; continue; }
          if (child.local !== "t" || run.local === "br" || ++leaves > 1 || child.children.some(item => item.kind !== "text" && item.kind !== "cdata")) return { reason: "Run text is not one direct scalar text leaf." };
          for (const item of child.children) if (item.kind === "text" || item.kind === "cdata") { await checkpoint(part, Math.max(1, item.value.length)); value += item.value; }
        }
        if (run.local === "r" && leaves !== 1) return { reason: "Text run lacks one scalar text leaf." };
        text += run.local === "br" ? "\n" : value;
      }
      paragraphs.push(text);
    }
    return paragraphs.length ? { paragraphs } : { reason: "Text body has no direct paragraphs." };
  };
  const rememberedText = async (part: string, elementIndex: number, cache: boolean) => {
    const store = cache ? cacheTexts : sourceTexts; let byIndex = store.get(part); if (!byIndex) store.set(part, byIndex = new Map());
    const previous = byIndex.get(elementIndex); if (previous) return previous;
    const value = await readText(part, (await index(part)).get(elementIndex), cache ? DSP : c.a.replace(/\/main$/, "/diagram"), cache ? "txBody" : "t");
    byIndex.set(elementIndex, value); return value;
  };
  const byPart = new Map<string, { model: PptxDiagramDataModelInfo; points: Map<number, PptxDiagramDataModelInfo["points"][number]>; presentations: Map<number, NonNullable<PptxDiagramDataModelInfo["presentationMappings"]>[number]> }>();
  for (const model of models) {
    await checkpoint(model.part);
    const points = new Map<number, PptxDiagramDataModelInfo["points"][number]>(), presentations = new Map<number, NonNullable<PptxDiagramDataModelInfo["presentationMappings"]>[number]>();
    for (const point of model.points) { await checkpoint(model.part); points.set(point.elementIndex, point); }
    for (const point of model.presentationMappings ?? []) { await checkpoint(model.part); presentations.set(point.elementIndex, point); }
    byPart.set(model.part, { model, points, presentations });
  }
  const result: PptxDiagramTextComparison[] = [];
  for (const mapping of mappings) {
    await checkpoint(mapping.cachePart); const data = byPart.get(mapping.dataPart);
    const reason = !data || data.model.status !== "inspected" || data.model.findings.length ? "Data graph is incomplete or has uninterpreted findings; no character equivalence is inferred."
      : mapping.status === "source-only" ? "Cache mapping is source-only; no character equivalence is inferred." : undefined;
    if (reason) { chargeFinding(mapping.cachePart); result.push({ dataPart: mapping.dataPart, cachePart: mapping.cachePart, coverage: "direct-paragraph-vectors", freshness: "unknown", status: "source-only", reason, shapes: [] }); continue; }
    const shapes: PptxDiagramTextComparison["shapes"][number][] = [];
    for (const shape of mapping.shapes) {
      await checkpoint(mapping.cachePart);
      let point = shape.model.status === "resolved" ? data!.points.get(shape.model.targetElementIndex!) : undefined;
      const presentation = point?.kind === "pres" ? point.elementIndex : undefined;
      if (presentation !== undefined) {
        const association = data!.presentations.get(presentation);
        point = association?.agreement === "same-target" ? data!.points.get(association.targetElementIndex!) : undefined;
      }
      const location = { cacheElementIndex: shape.elementIndex, ...(presentation !== undefined ? { presentationElementIndex: presentation } : {}), ...(point && ["node", "asst"].includes(point.kind) ? { sourcePointElementIndex: point.elementIndex } : {}) };
      if (!point || !["node", "asst"].includes(point.kind)) {
        chargeFinding(mapping.cachePart); shapes.push({ ...location, status: "unavailable", reason: "Cache reference does not name content/assistant directly or through agreeing explicit presentation evidence." }); continue;
      }
      const source = await rememberedText(mapping.dataPart, point.elementIndex, false), cache = await rememberedText(mapping.cachePart, shape.elementIndex, true);
      if (!("paragraphs" in source) || !("paragraphs" in cache)) {
        chargeFinding(mapping.cachePart); shapes.push({ ...location, status: "unavailable", reason: "reason" in source ? source.reason : "reason" in cache ? cache.reason : "Unavailable direct text." }); continue;
      }
      let units = source.paragraphs.length + cache.paragraphs.length;
      for (const paragraph of source.paragraphs) units += paragraph.length;
      for (const paragraph of cache.paragraphs) units += paragraph.length;
      await checkpoint(mapping.cachePart, Math.max(1, units));
      const equal = source.paragraphs.length === cache.paragraphs.length && source.paragraphs.every((text, i) => text === cache.paragraphs[i]);
      if (!equal) chargeFinding(mapping.cachePart);
      shapes.push({ ...location, status: equal ? "equal" : "different", sourceParagraphs: [...source.paragraphs], cacheParagraphs: [...cache.paragraphs], nonemptyText: [...source.paragraphs, ...cache.paragraphs].some(text => text.trim().length > 0) });
    }
    result.push({ dataPart: mapping.dataPart, cachePart: mapping.cachePart, coverage: "direct-paragraph-vectors", freshness: "unknown", status: "observed", shapes });
  }
  return result;
}
