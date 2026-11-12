import type { structureContext } from "./structure-context.js";
import type { PptxSlideInfo } from "../types/structure.js";
import type { PptxDiagramCacheInfo, PptxDiagramCacheReference, PptxDiagramCacheCandidate, PptxDiagramResourceInfo, PptxDiagramResourceInventory, PptxDiagramImageRelationship } from "../types/smartart.js";
import type { OpcRelationshipInfo } from "../types/opc.js";
import { attribute, children } from "./structure-xml.js";
import { readDiagramData } from "./diagram-data.js";
import { mapDiagramModels } from "./diagram-mapping.js";
import { compareDiagramText } from "./diagram-text-comparison.js";
import { PackageInspectionError } from "../types/archive.js";

type Context = Awaited<ReturnType<typeof structureContext>>;
const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const CACHE_REL = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
const CACHE_TYPE = "application/vnd.ms-office.drawingml.diagramDrawing+xml";
const MCE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
type CacheTarget = Pick<PptxDiagramCacheInfo, "status" | "relationshipId" | "target" | "part" | "contentType">;

/** Inspect source associations; do not render, infer freshness, authorize edits or collect garbage. */
export async function inspectDiagramResources(c: Context, slides: readonly PptxSlideInfo[]): Promise<{ slides: PptxSlideInfo[]; inventory?: PptxDiagramResourceInventory }> {
  const parts = new Map(c.entry.package.parts.map(p => [p.name, p]));
  const bySource = new Map<string, OpcRelationshipInfo[]>(), incoming = new Map<string, { source: string; relationshipId: string; type: string }[]>();
  for (const relation of c.entry.package.relationships) {
    await c.checkpoint();
    let list = bySource.get(relation.source); if (!list) bySource.set(relation.source, list = []); list.push(relation);
    if (relation.targetPart) {
      let edges = incoming.get(relation.targetPart); if (!edges) incoming.set(relation.targetPart, edges = []);
      edges.push({ source: relation.source, relationshipId: relation.id, type: relation.type });
    }
  }
  const cacheTarget = (slide: string, id: string): CacheTarget => {
    const relation = c.findRelationship(slide, id), base = { relationshipId: id };
    if (!relation) return { ...base, status: "missing-relationship" };
    const target = { ...base, target: relation.target, ...(relation.targetPart ? { part: relation.targetPart } : {}) };
    if (relation.type !== CACHE_REL) return { ...target, status: "wrong-relationship-type" };
    if (relation.targetMode === "External") return { ...target, status: "external" };
    if (relation.fragment !== undefined) return { ...target, status: "fragment-target" };
    const part = relation.targetPart ? parts.get(relation.targetPart) : undefined;
    const result = { ...target, ...(part ? { contentType: part.contentType } : {}) };
    if (part?.contentType.toLowerCase() !== CACHE_TYPE.toLowerCase()) return { ...result, status: "wrong-content-type" };
    const root = c.rawXml.get(part.name);
    return { ...result, status: root?.namespaceUri === DSP && root.local === "drawing" ? "located" : "wrong-root" };
  };
  const dataCache = new Map<string, { status?: "source-only" | "ambiguous-reference"; references: PptxDiagramCacheReference[] }>();
  const referencesFor = async (part: string) => {
    let cached = dataCache.get(part); if (cached) return cached;
    // No global claim to understand DSP/MCE is added to the projection engine.
    if (parts.get(part)?.xml?.namespaceUris.includes(MCE)) {
      cached = { status: "source-only", references: [] }; dataCache.set(part, cached); return cached;
    }
    const root = c.rawXml.get(part)!, dgm = c.a.replace(/\/main$/, "/diagram"), references: PptxDiagramCacheReference[] = [];
    const lists = children(root).filter(n => n.namespaceUri === dgm && n.local === "extLst");
    for (const list of lists) for (const ext of children(list)) {
      await c.checkpoint();
      if (ext.namespaceUri !== c.a || ext.local !== "ext" || attribute(ext, "uri") !== DSP) continue;
      for (const block of children(ext)) {
        await c.checkpoint();
        if (block.namespaceUri !== DSP || block.local !== "dataModelExt") continue;
        const id = attribute(block, "relId"), minimumVersion = attribute(block, "minVer");
        references.push({ elementIndex: (await c.sourceElement(part, block)).elementIndex,
          ...(id !== undefined ? { relationshipId: id } : {}), ...(minimumVersion !== undefined ? { minimumVersion } : {}) });
      }
    }
    cached = { ...(lists.length > 1 || references.length > 1 ? { status: "ambiguous-reference" as const } : {}), references }; dataCache.set(part, cached); return cached;
  };
  type MutableResource = { part: string; contentType: string; roles: Set<PptxDiagramResourceInfo["roles"][number]>; owners: PptxDiagramResourceInfo["owners"][number][]; imageRelationships: PptxDiagramImageRelationship[]; modelReferences?: { elementIndex: number; modelId?: string }[] };
  const resources = new Map<string, MutableResource>();
  const resource = (part: string, role: PptxDiagramResourceInfo["roles"][number]) => {
    let result = resources.get(part);
    if (!result) { result = { part, contentType: parts.get(part)!.contentType, roles: new Set(), owners: [], imageRelationships: [] }; resources.set(part, result); }
    result.roles.add(role); return result;
  };
  const updated: PptxSlideInfo[] = [], cacheCandidates: PptxDiagramCacheCandidate[] = [];
  for (const slide of slides) {
    const diagrams = [], candidateOwners = new Map<string, number[]>();
    for (const diagram of slide.diagrams ?? []) {
      await c.checkpoint();
      for (const component of diagram.components) if (component.status === "located" && component.part) {
        resource(component.part, component.role).owners.push({ slidePart: slide.part, elementIndex: diagram.elementIndex, role: component.role });
      }
      const data = diagram.components.find(component => component.role === "data");
      let cache: PptxDiagramCacheInfo;
      if (data?.status !== "located" || !data.part) cache = { status: "unresolved-data", freshness: "unknown", references: [] };
      else {
        const found = await referencesFor(data.part), id = found.references[0]?.relationshipId;
        cache = { dataPart: data.part, freshness: "unknown", references: found.references,
          ...(found.status ? { status: found.status } : id === undefined ? { status: "missing-reference" } : cacheTarget(slide.part, id)) };
      }
      if (cache.status === "located" && cache.part && cache.relationshipId) {
        resource(cache.part, "cache").owners.push({ slidePart: slide.part, elementIndex: diagram.elementIndex, role: "cache" });
        let owners = candidateOwners.get(cache.relationshipId); if (!owners) candidateOwners.set(cache.relationshipId, owners = []); owners.push(diagram.elementIndex);
      }
      diagrams.push({ ...diagram, cache });
    }
    for (const relation of bySource.get(slide.part) ?? []) if (relation.type === CACHE_REL) {
      await c.checkpoint();
      const result = cacheTarget(slide.part, relation.id);
      // A declared cache relationship always has an ID/target and the role is already known.
      cacheCandidates.push({ ...result, slidePart: slide.part, relationshipId: relation.id, target: relation.target,
        status: result.status as PptxDiagramCacheCandidate["status"], referencedBy: candidateOwners.get(relation.id) ?? [] });
      if (result.status === "located" && result.part) resource(result.part, "cache");
    }
    updated.push(diagrams.length ? { ...slide, diagrams } : slide);
  }
  for (const entry of resources.values()) {
    if (![...entry.roles].some(role => role !== "image")) continue;
    const images = new Map<string, { elementIndex: number; attribute: "embed" | "link" }[]>();
    for (const relation of bySource.get(entry.part) ?? []) if (relation.type === `${c.r}/image`) images.set(relation.id, []);
    const root = c.rawXml.get(entry.part), stack = root ? [root] : [];
    while (stack.length) {
      await c.checkpoint();
      const node = stack.pop()!;
      if (entry.roles.has("cache") && node.namespaceUri === DSP && node.local === "sp") {
        const modelId = attribute(node, "modelId");
        (entry.modelReferences ??= []).push({ elementIndex: (await c.sourceElement(entry.part, node)).elementIndex, ...(modelId !== undefined ? { modelId } : {}) });
      }
      if (node.namespaceUri === c.a && node.local === "blip") for (const attr of ["embed", "link"] as const) {
        const id = attribute(node, attr, c.r);
        if (id !== undefined) {
          let refs = images.get(id); if (!refs) images.set(id, refs = []);
          refs.push({ elementIndex: (await c.sourceElement(entry.part, node)).elementIndex, attribute: attr });
        }
      }
      const nodes = children(node); for (let i = nodes.length - 1; i >= 0; i--) stack.push(nodes[i]!);
    }
    for (const [id, references] of images) {
      await c.checkpoint();
      const relation = c.findRelationship(entry.part, id), target = relation?.targetPart ? parts.get(relation.targetPart) : undefined;
      const status = !relation ? "missing-relationship" : relation.type !== `${c.r}/image` ? "wrong-relationship-type"
        : relation.targetMode === "External" ? "external" : relation.fragment !== undefined ? "fragment-target"
        : !target?.contentType.toLowerCase().startsWith("image/") ? "wrong-content-type" : "located";
      entry.imageRelationships.push({ relationshipId: id, status, references, ...(relation ? { target: relation.target } : {}),
        ...(target ? { part: target.name, contentType: target.contentType } : {}) });
      if (status === "located" && target) resource(target.name, "image");
    }
  }
  const dataModels = [];
  let graphWork = 0, graphFindings = 0;
  const graphCheckpoint = async (part: string, units = 1) => {
    graphWork += units;
    if (graphWork > c.diagramLimits.maxWorkUnits) throw new PackageInspectionError("package-limit-exceeded", "Diagram graph/text inspection exceeds maxWorkUnits.", part);
    await c.checkpoint();
  };
  const graphFinding = (part: string) => { if (++graphFindings > c.diagramLimits.maxFindings) throw new PackageInspectionError("package-limit-exceeded", "Diagram graph inspection exceeds maxFindings.", part); };
  for (const entry of resources.values()) if (entry.roles.has("data")) {
    const checkpoint = () => graphCheckpoint(entry.part), finding = () => graphFinding(entry.part);
    dataModels.push(await readDiagramData(c, entry.part, parts.get(entry.part)!.xml!.namespaceUris.includes(MCE), checkpoint, finding));
  }
  const mapped = await mapDiagramModels(c, dataModels, updated, graphCheckpoint, graphFinding);
  const textComparisons = c.diagramTextComparison ? await compareDiagramText(c, mapped.dataModels, mapped.cacheMappings, graphCheckpoint, graphFinding) : undefined;
  const inventory: PptxDiagramResourceInventory = { coverage: "direct-source-associations", cacheCandidates, ...(dataModels.length ? { dataModels: mapped.dataModels } : {}), ...(mapped.cacheMappings.length ? { cacheMappings: mapped.cacheMappings } : {}),
    ...(textComparisons?.length ? { textComparisons } : {}), parts: [...resources.values()].map(entry => ({ ...entry, roles: [...entry.roles], incomingRelationships: incoming.get(entry.part) ?? [] })) };
  return { slides: updated, ...(inventory.parts.length || cacheCandidates.length ? { inventory } : {}) };
}
