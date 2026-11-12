import type { structureContext } from "./structure-context.js";
import type { PptxSlideInfo } from "../types/structure.js";
import type { PptxDiagramDataModelInfo, PptxDiagramCacheMapping, PptxDiagramPresentationMapping, PptxDiagramModelReference } from "../types/smartart.js";
import type { XmlElement } from "../xml/read-xml.js";
import { attribute, children } from "./structure-xml.js";
import { diagramModelIdentity } from "./diagram-model-id.js";

type Context = Awaited<ReturnType<typeof structureContext>>;
const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const MCE = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** Explicit source evidence, normalized by data part and data/cache pair. No ownership inference. */
export async function mapDiagramModels(c: Context, models: readonly PptxDiagramDataModelInfo[], slides: readonly PptxSlideInfo[],
  checkpoint: (part: string) => Promise<void>, chargeFinding: (part: string) => void,
): Promise<{ dataModels: PptxDiagramDataModelInfo[]; cacheMappings: PptxDiagramCacheMapping[] }> {
  const dataModels: PptxDiagramDataModelInfo[] = [], cacheMappings: PptxDiagramCacheMapping[] = [];
  const indexed = new Map<string, { model: PptxDiagramDataModelInfo; identities: Map<string, { kind: "point" | "connection"; index: number; count: number }> }>();
  for (const model of models) {
    const identities = new Map<string, { kind: "point" | "connection"; index: number; count: number }>();
    const presentations = new Map<number, { sources: PptxDiagramPresentationMapping["sources"][number][] }>();
    for (const [kind, entities] of [["point", model.points], ["connection", model.connections]] as const) for (const entity of entities) {
      await checkpoint(model.part);
      const key = diagramModelIdentity(entity.modelId);
      if (key !== undefined) {
        const known = identities.get(key); if (known) known.count++; else identities.set(key, { kind, index: entity.elementIndex, count: 1 });
      }
    }
    for (const point of model.points) {
      await checkpoint(model.part);
      if (point.kind !== "pres") continue;
      presentations.set(point.elementIndex, { sources: point.presentationAssociation ? [{ kind: "presAssocID", elementIndex: point.elementIndex, reference: point.presentationAssociation }] : [] });
    }
    for (const connection of model.connections) {
      await checkpoint(model.part);
      if (connection.kind !== "presOf" || connection.destination.status !== "resolved") continue;
      presentations.get(connection.destination.targetElementIndex!)?.sources.push({ kind: "presOf", elementIndex: connection.elementIndex, reference: connection.source });
    }
    const mappings: PptxDiagramPresentationMapping[] = [];
    for (const [elementIndex, { sources }] of presentations) {
      await checkpoint(model.part);
      const targets = new Set<number>(); let unresolved = false;
      for (const source of sources) {
        await checkpoint(model.part);
        if (source.reference.status === "resolved") targets.add(source.reference.targetElementIndex!); else unresolved = true;
      }
      const agreement = !sources.length ? "none" : unresolved ? "unresolved" : targets.size > 1 ? "different-targets" : "same-target";
      mappings.push({ elementIndex, sources, agreement, ...(agreement === "same-target" ? { targetElementIndex: [...targets][0]! } : {}) });
    }
    const updated = mappings.length ? { ...model, presentationMappings: mappings } : model;
    indexed.set(model.part, { model: updated, identities }); dataModels.push(updated);
  }
  // Reuse a single source-tree scan for shared caches. Only the recognized DSP
  // shape tree participates; raw modelReferences elsewhere remain raw inventory.
  const caches = new Map<string, { root: XmlElement; reason?: string; shapes: { elementIndex: number; modelId?: string }[] }>();
  const partInfo = new Map(c.entry.package.parts.map(part => [part.name, part]));
  const seen = new Map<string, Set<string>>();
  for (const slide of slides) for (const diagram of slide.diagrams ?? []) {
    await checkpoint(slide.part);
    const cache = diagram.cache;
    if (cache?.status !== "located" || !cache.dataPart || !cache.part) continue;
    let pairs = seen.get(cache.dataPart); if (!pairs) seen.set(cache.dataPart, pairs = new Set());
    if (pairs.has(cache.part)) continue; pairs.add(cache.part);
    const data = indexed.get(cache.dataPart); if (!data) continue;
    let source = caches.get(cache.part);
    if (!source) {
      const root = c.rawXml.get(cache.part)!;
      const trees = children(root).filter(node => node.namespaceUri === DSP && node.local === "spTree");
      const reason = partInfo.get(cache.part)?.xml?.namespaceUris.includes(MCE) ? "Cache MCE needs an effective semantic branch policy."
        : trees.length !== 1 ? "Expected one direct cache shape tree; ambiguous or missing trees are not merged." : undefined;
      source = { root, ...(reason ? { reason } : {}), shapes: [] }; caches.set(cache.part, source);
      const stack = !reason ? [trees[0]!] : [];
      while (stack.length) {
        await checkpoint(cache.part);
        const node = stack.pop()!;
        if (node.namespaceUri === DSP && node.local === "sp") {
          const modelId = attribute(node, "modelId");
          source.shapes.push({ elementIndex: (await c.sourceElement(cache.part, node)).elementIndex, ...(modelId !== undefined ? { modelId } : {}) });
        } else if (node.namespaceUri === DSP && ["spTree", "grpSp"].includes(node.local)) {
          const nodes = children(node); for (let i = nodes.length - 1; i >= 0; i--) {
            await checkpoint(cache.part);
            const child = nodes[i]!;
            if (child.namespaceUri === DSP && ["sp", "grpSp"].includes(child.local)) stack.push(child);
          }
        }
      }
    }
    const findings: PptxDiagramCacheMapping["findings"][number][] = [], shapes: PptxDiagramCacheMapping["shapes"][number][] = [];
    const reason = data.model.status === "source-only" ? "Data graph is source-only; cache identities are not resolved against an empty inferred graph." : source.reason;
    if (reason) {
      chargeFinding(cache.part);
      findings.push({ code: "source-only", elementIndex: (await c.sourceElement(cache.part, source.root)).elementIndex, reason });
    } else for (const shape of source.shapes) {
      await checkpoint(cache.part);
      const key = diagramModelIdentity(shape.modelId), entity = key === undefined ? undefined : data.identities.get(key);
      const status = shape.modelId === undefined ? "absent" : key === undefined ? "unrecognized" : !entity ? "missing" : entity.count > 1 ? "ambiguous" : entity.kind !== "point" ? "wrong-kind" : "resolved";
      const model: PptxDiagramModelReference = { ...(shape.modelId !== undefined ? { value: shape.modelId } : {}), status,
        ...(status === "resolved" ? { targetElementIndex: entity!.index } : {}) };
      shapes.push({ elementIndex: shape.elementIndex, model });
      if (status !== "resolved") {
        chargeFinding(cache.part);
        findings.push({ code: "unresolved-reference", elementIndex: shape.elementIndex, attribute: "modelId", reason: `Cache modelId is ${status}; expected a unique data-model point.` });
      }
    }
    cacheMappings.push({ dataPart: cache.dataPart, cachePart: cache.part, coverage: "explicit-model-references", status: reason ? "source-only" : findings.length ? "partial" : "inspected", shapes, findings });
  }
  return { dataModels, cacheMappings };
}
