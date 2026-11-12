import type { XmlElement } from "../xml/read-xml.js";
import type { PptxExtensionInfo, PptxSectionInfo, PptxSectionListInfo, PptxSlideInfo, PptxStructureFinding } from "../types/structure.js";
import { XMLNS } from "../compatibility/configuration.js";
import { attribute, children, integer, invalidStructure, singleChild } from "./structure-xml.js";

export const SECTION_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";
export const SECTION_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const extensionUri = (node: XmlElement): string | undefined => attribute(node, "uri")?.replace(/[\t\r\n ]+/g, " ").replace(/^ | $/g, "");
export const isSectionExtension = (node: XmlElement, p: string): boolean => node.namespaceUri === p && node.local === "ext" && extensionUri(node) === SECTION_URI;
const attrs = (node: XmlElement, names: readonly string[], part: string): void => {
  if (node.attributes.some((a) => a.namespaceUri !== XMLNS && (a.namespaceUri !== "" || !names.includes(a.local)))) invalidStructure(`Unexpected attribute on section ${node.local}.`, part);
};
const noText = (node: XmlElement, part: string): void => {
  if (node.children.some((n) => (n.kind === "text" || n.kind === "cdata") && !/^[\t\r\n ]*$/.test(n.value))) invalidStructure(`Unexpected text in section ${node.local}.`, part);
};
const descriptor = (node: XmlElement, interpreted = false): PptxExtensionInfo => {
  const uri = extensionUri(node);
  return { ...(uri !== undefined ? { uri } : {}), status: interpreted ? "interpreted" : "uninterpreted", children: children(node).map(({ name, local, prefix, namespaceUri }) => ({ name, local, prefix, namespaceUri })) };
};

/** Reads only the named MS-PPTX section extension; never infers membership from names. */
export async function readSections(presentation: XmlElement, p: string, part: string, slides: readonly PptxSlideInfo[], checkpoint: () => Promise<void>): Promise<{ sectionList?: PptxSectionListInfo; presentationExtensions: PptxExtensionInfo[]; findings: PptxStructureFinding[] }> {
  const extList = singleChild(presentation, p, "extLst", part), presentationExtensions: PptxExtensionInfo[] = [], findings: PptxStructureFinding[] = [];
  if (!extList) return { presentationExtensions, findings };
  noText(extList, part); attrs(extList, [], part);
  let sectionRoot: XmlElement | undefined;
  for (const ext of children(extList)) {
    await checkpoint();
    if (ext.namespaceUri !== p || ext.local !== "ext") invalidStructure("Unexpected presentation extension-list child.", part);
    const known = isSectionExtension(ext, p);
    presentationExtensions.push(descriptor(ext, known));
    if (!known) continue;
    if (sectionRoot) invalidStructure("Duplicate native section extension.", part);
    attrs(ext, ["uri"], part); noText(ext, part);
    const payload = children(ext);
    if (payload.length !== 1 || payload[0]!.namespaceUri !== SECTION_NS || payload[0]!.local !== "sectionLst") invalidStructure("Native section extension must contain one p14:sectionLst.", part);
    sectionRoot = payload[0]!;
  }
  if (!sectionRoot) return { presentationExtensions, findings };
  attrs(sectionRoot, [], part); noText(sectionRoot, part);
  const nodes = children(sectionRoot);
  if (!nodes.length) invalidStructure("Native section list requires at least one section; an empty section may have no slides.", part);
  const byId = new Map(slides.map((slide, index) => [slide.id, { slide, index }])), usedSlides = new Set<number>(), identities = new Set<string>(), sections: PptxSectionInfo[] = [];
  let previous = -1;
  for (const [sectionIndex, node] of nodes.entries()) {
    await checkpoint();
    if (node.namespaceUri !== SECTION_NS || node.local !== "section") invalidStructure("Unexpected native section-list child.", part);
    attrs(node, ["id", "name"], part); noText(node, part);
    const rawId = attribute(node, "id"), id = rawId?.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, ""), name = attribute(node, "name");
    if (id !== undefined) {
      if (!/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/.test(id) || identities.has(id)) invalidStructure("Invalid or duplicate section GUID.", part);
      identities.add(id);
    }
    const contents = children(node), members = contents[0], nestedExtensions = contents[1];
    if (!members || members.namespaceUri !== SECTION_NS || members.local !== "sldIdLst" || contents.length > 2 || (nestedExtensions && (nestedExtensions.namespaceUri !== SECTION_NS || nestedExtensions.local !== "extLst"))) invalidStructure("Expected section slide list followed by an optional extension list.", part);
    attrs(members!, [], part); noText(members!, part);
    const references: (PptxSlideInfo & { id: number })[] = [];
    for (const item of children(members!)) {
      await checkpoint();
      if (item.namespaceUri !== SECTION_NS || item.local !== "sldId") invalidStructure("Unexpected section slide-list child.", part);
      attrs(item, ["id"], part); noText(item, part);
      if (children(item).length) invalidStructure("Section slide ID cannot contain elements.", part);
      const slideId = integer(attribute(item, "id"), 256, 2147483647, "section slide ID", part), resolved = byId.get(slideId);
      if (!resolved || usedSlides.has(slideId)) invalidStructure("Section slide ID is missing from the presentation order or belongs to multiple section entries.", part);
      usedSlides.add(slideId);
      if (resolved!.index !== previous + 1) findings.push({ code: "section-order", part, sectionIndex, slideId, message: "Section membership does not form a contiguous partition of presentation order; source order is retained." });
      previous = resolved!.index;
      references.push(resolved!.slide);
    }
    const extensions: PptxExtensionInfo[] = [];
    if (nestedExtensions) {
      attrs(nestedExtensions, [], part); noText(nestedExtensions, part);
      for (const ext of children(nestedExtensions)) {
        await checkpoint();
        // p14:extLst uses the Transitional p:CT_ExtensionList type in MS-PPTX,
        // even when embedded in a Strict presentation.
        if (ext.namespaceUri !== "http://schemas.openxmlformats.org/presentationml/2006/main" || ext.local !== "ext") invalidStructure("Unexpected section extension-list child.", part);
        extensions.push(descriptor(ext));
      }
    }
    sections.push({ ...(id !== undefined ? { id } : {}), ...(name !== undefined ? { name } : {}), slides: references.map(({ id: slideId, part: slidePart, relationshipId }) => ({ id: slideId, part: slidePart, relationshipId })), extensions });
  }
  for (const slide of slides) {
    await checkpoint();
    if (!usedSlides.has(slide.id)) findings.push({ code: "unsectioned-slide", part, slideId: slide.id, message: "Slide is absent from the native section list; no section is inferred." });
  }
  return { sectionList: { part, extensionUri: SECTION_URI, sections }, presentationExtensions, findings };
}
