import type { PptxInspectionOptions } from "../types/structure.js";
import type { PptxInspection, PptxMasterInfo, PptxPartReference, PptxSlideInfo, PptxStructureFinding } from "../types/structure.js";
import type { XmlElement } from "../xml/read-xml.js";
import { structureContext } from "./structure-context.js";
import { readTransition } from "./transition.js";
import { smartArtReader } from "./smartart.js";
import { inspectDiagramResources } from "./diagram-resources.js";
import { readSections } from "./sections.js";
import { attribute, children, hiddenSlide, integer, invalidStructure, singleChild, size } from "./structure-xml.js";

export async function readPresentationStructure(bytes: Uint8Array, options?: PptxInspectionOptions): Promise<PptxInspection> {
  return readStructureFromContext(await structureContext(bytes, options));
}

export async function readStructureFromContext(c: Awaited<ReturnType<typeof structureContext>>): Promise<PptxInspection> {
  const main = c.entry.mainPart;
  const presentation = await c.root(main, "presentation"), findings: PptxStructureFinding[] = [];
  const missing = (part: string, role: string): void => { findings.push({ code: "missing-inheritance", part, message: `No ${role} relationship; inheritance is unresolved.` }); };
  const named = (root: XmlElement, part: string): { name?: string } => {
    const common = singleChild(root, c.p, "cSld", part);
    const name = common ? attribute(common, "name") : undefined;
    return name === undefined ? {} : { name };
  };
  const list = (root: XmlElement, part: string, listName: string, itemName: string): XmlElement[] => {
    const container = singleChild(root, c.p, listName, part);
    if (!container) return [];
    const values = children(container);
    if (values.some((node) => node.namespaceUri !== c.p || node.local !== itemName)) invalidStructure(`Unexpected child in ${listName}.`, part);
    if (container.children.some((node) => (node.kind === "text" || node.kind === "cdata") && !/^[\t\r\n ]*$/.test(node.value))) invalidStructure(`Unexpected character content in ${listName}.`, part);
    return values;
  };
  const references = (nodes: readonly XmlElement[], part: string, role: string): PptxPartReference[] => {
    const seen = new Set<string>();
    return nodes.map((node) => {
      const ref = c.reference(part, attribute(node, "id", c.r), role);
      if (seen.has(ref.part)) invalidStructure(`Duplicate ${role} target in an identity list.`, part);
      seen.add(ref.part); return ref;
    });
  };
  const identities = (nodes: readonly XmlElement[], part: string, min: number, max: number): number[] => {
    const seen = new Set<number>();
    return nodes.map((node) => {
      const id = integer(attribute(node, "id"), min, max, `${node.local}/@id`, part);
      if (seen.has(id)) invalidStructure(`Duplicate ${node.local} numeric identity.`, part);
      seen.add(id); return id;
    });
  };
  const slideSizeNode = singleChild(presentation, c.p, "sldSz", main), notesSizeNode = singleChild(presentation, c.p, "notesSz", main, true)!;
  const slideSize = slideSizeNode ? size(slideSizeNode, 914400, 51206400, main) : undefined;
  const notesSize = size(notesSizeNode, 0, 27273042316900, main);
  const first = attribute(presentation, "firstSlideNum"), firstSlideNumber = first === undefined ? 1 : integer(first, -2147483648, 2147483647, "firstSlideNum", main);

  const themeFor = async (part: string): Promise<string | undefined> => {
    const theme = c.optionalTarget(part, "theme");
    if (theme) await c.root(theme, "theme", "theme", c.a);
    else missing(part, "theme");
    return theme;
  };
  const masterNodes = list(presentation, main, "sldMasterIdLst", "sldMasterId"), masterRefs = references(masterNodes, main, "slideMaster"), masterIds = identities(masterNodes, main, 2147483648, 4294967295);
  const slideMasters: PptxMasterInfo[] = [], masterByPart = new Map<string, PptxMasterInfo>(), masterLayoutParts = new Map<string, Set<string>>();
  for (let i = 0; i < masterRefs.length; i++) {
    const ref = masterRefs[i]!, root = await c.root(ref.part, "sldMaster", "presentationml.slideMaster");
    const layoutNodes = list(root, ref.part, "sldLayoutIdLst", "sldLayoutId"), layoutRefs = references(layoutNodes, ref.part, "slideLayout"), layoutIds = identities(layoutNodes, ref.part, 2147483648, 4294967295);
    const themePart = await themeFor(ref.part);
    const master: PptxMasterInfo = { ...ref, id: masterIds[i]!, layouts: layoutRefs.map((layout, j) => ({ ...layout, id: layoutIds[j]! })), ...(themePart ? { themePart } : {}) };
    slideMasters.push(master); masterByPart.set(ref.part, master); masterLayoutParts.set(ref.part, new Set(layoutRefs.map((layout) => layout.part)));
  }
  const layouts = new Map<string, { part: string; masterPart?: string; name?: string }>();
  const readLayout = async (part: string, expectedMaster?: string) => {
    let layout = layouts.get(part);
    if (!layout) {
      const root = await c.root(part, "sldLayout", "presentationml.slideLayout"), masterPart = c.optionalTarget(part, "slideMaster");
      layout = { part, ...named(root, part), ...(masterPart ? { masterPart } : {}) }; layouts.set(part, layout);
      if (masterPart) {
        await c.root(masterPart, "sldMaster", "presentationml.slideMaster");
        const master = masterByPart.get(masterPart);
        if (!master) findings.push({ code: "unlisted-master", part, message: "Layout points to a master absent from the presentation master list." });
        else if (!masterLayoutParts.get(masterPart)!.has(part)) findings.push({ code: "unlisted-layout", part, message: "Layout is absent from its master's layout identity list." });
      } else missing(part, "slideMaster");
    }
    if (expectedMaster && layout.masterPart && layout.masterPart !== expectedMaster) invalidStructure("Listed layout points back to a different slide master.", part);
    return layout;
  };
  for (const master of slideMasters) for (const layout of master.layouts) await readLayout(layout.part, master.part);

  const auxiliaryMasters = async (role: "notesMaster" | "handoutMaster"): Promise<PptxPartReference[]> => {
    const nodes = list(presentation, main, `${role}IdLst`, `${role}Id`);
    if (nodes.length > 1) invalidStructure(`At most one ${role} is permitted.`, main);
    const refs = references(nodes, main, role);
    for (const ref of refs) { await c.root(ref.part, role, `presentationml.${role}`); await themeFor(ref.part); }
    return refs;
  };
  const notesMasters = await auxiliaryMasters("notesMaster"), handoutMasters = await auxiliaryMasters("handoutMaster");
  const slideNodes = list(presentation, main, "sldIdLst", "sldId"), slideRefs = references(slideNodes, main, "slide"), slideIds = identities(slideNodes, main, 256, 2147483647);
  const slides: PptxSlideInfo[] = [], usedNotes = new Set<string>(), readDiagrams = smartArtReader(c);
  for (let i = 0; i < slideRefs.length; i++) {
    const ref = slideRefs[i]!, root = await c.root(ref.part, "sld", "presentationml.slide");
    const layoutPart = c.optionalTarget(ref.part, "slideLayout"), notesPart = c.optionalTarget(ref.part, "notesSlide");
    if (layoutPart) await readLayout(layoutPart); else missing(ref.part, "slideLayout");
    if (notesPart) {
      if (usedNotes.has(notesPart)) invalidStructure("Multiple slides share the same notes slide.", ref.part);
      usedNotes.add(notesPart);
      const notesRoot = await c.root(notesPart, "notes", "presentationml.notesSlide");
      // Notes-page actions may also use slide relationships. Distinguish those
      // explicit links from the implicit owning-slide backlink; a self-link can
      // share that same relationship without manufacturing a second owner.
      const linkIds = new Set<string>(), stack = [notesRoot];
      while (stack.length) {
        await c.checkpoint();
        const node = stack.pop()!;
        if ((node.namespaceUri === c.p || node.namespaceUri === c.a) && node.local === "ext") continue;
        if (node.namespaceUri === c.a && ["hlinkClick", "hlinkHover"].includes(node.local)) {
          const id = attribute(node, "id", c.r); if (id !== undefined) linkIds.add(id);
        }
        for (const child of children(node)) stack.push(child);
      }
      const slideRelations = c.relationships(notesPart, "slide"), implicit = slideRelations.filter(r => !linkIds.has(r.id));
      if (implicit.length > 1) invalidStructure("Multiple owning-slide backlinks in a notes part.", notesPart);
      const owner = implicit[0] ?? slideRelations.find(r => r.targetMode === "Internal" && r.fragment === undefined && r.targetPart === ref.part);
      const back = owner ? c.target(owner, "slide") : undefined;
      if (back && back !== ref.part) invalidStructure("Notes slide points back to a different slide.", notesPart);
      if (!back) missing(notesPart, "slide backlink");
      const notesMaster = c.optionalTarget(notesPart, "notesMaster");
      if (notesMaster) {
        await c.root(notesMaster, "notesMaster", "presentationml.notesMaster");
        if (!notesMasters.some((master) => master.part === notesMaster)) findings.push({ code: "unlisted-master", part: notesPart, message: "Notes slide points to an unlisted notes master." });
      } else missing(notesPart, "notesMaster");
    }
    const transitionNode = singleChild(root, c.p, "transition", ref.part);
    const transition = transitionNode ? readTransition(transitionNode, (await c.sourceElement(ref.part, transitionNode)).node, c.p, ref.part) : undefined;
    const diagrams = await readDiagrams(ref.part, root);
    slides.push({ ...(transition ? { transition } : {}), ...(diagrams.length ? { diagrams } : {}), ...ref, id: slideIds[i]!, hidden: hiddenSlide(root, ref.part), ...named(root, ref.part), ...(layoutPart ? { layoutPart } : {}), ...(notesPart ? { notesPart } : {}) });
  }
  const listedSlides = new Set(slides.map((slide) => slide.part));
  for (const relation of c.relationships(main, "slide")) {
    const target = c.target(relation, "slide");
    if (!listedSlides.has(target)) {
      await c.root(target, "sld", "presentationml.slide");
      findings.push({ code: "unlisted-slide", part: target, message: "Presentation has a slide relationship not referenced by sldIdLst; it does not determine page order." });
    }
  }
  const sectionInfo = await readSections(presentation, c.p, main, slides, c.checkpoint);
  for (const finding of sectionInfo.findings) findings.push(finding);
  const diagramResources = await inspectDiagramResources(c, slides);
  c.abort();
  return { ...c.entry, ...(diagramResources.inventory ? { diagramInventory: diagramResources.inventory } : {}), metadata: c.metadata, ...(sectionInfo.sectionList ? { sectionList: sectionInfo.sectionList } : {}), presentationExtensions: sectionInfo.presentationExtensions, ...(slideSize ? { slideSize } : {}), notesSize, firstSlideNumber, firstSlideNumberDeclared: first !== undefined, slides: diagramResources.slides, slideMasters, notesMasters, handoutMasters, layouts: [...layouts.values()], findings, markupCompatibility: structuredClone(c.reports) };
}
