import type { structureContext } from "./structure-context.js";
import type { PptxDiagramComponentInfo, PptxDiagramComponentRole, PptxDiagramInfo } from "../types/smartart.js";
import type { XmlElement } from "../xml/read-xml.js";
import { attribute, children, singleChild } from "./structure-xml.js";

type Context = Awaited<ReturnType<typeof structureContext>>;
const roles: readonly [PptxDiagramComponentRole, string, string, string][] = [
  ["data", "dm", "diagramData", "dataModel"],
  ["layout", "lo", "diagramLayout", "layoutDef"],
  ["colors", "cs", "diagramColors", "colorsDef"],
  ["style", "qs", "diagramQuickStyle", "styleDef"],
];
const nonVisual: Readonly<Record<string, string>> = { spTree: "nvGrpSpPr", grpSp: "nvGrpSpPr", sp: "nvSpPr", pic: "nvPicPr", cxnSp: "nvCxnSpPr", graphicFrame: "nvGraphicFramePr" };
const identityKey = (id: string | undefined): string | undefined => {
  const value = id?.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
  if (!value || !/^\+?\d+$/.test(value)) return undefined;
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 0 && number <= 4294967295 ? String(number) : undefined;
};

/** Effective slide objects only. Does not walk arbitrary extensions or inactive MCE branches. */
export function smartArtReader(c: Context) {
  const parts = new Map(c.entry.package.parts.map(part => [part.name, part]));
  const dgm = c.a.replace(/\/main$/, "/diagram");
  const component = (part: string, refs: XmlElement | undefined, spec: typeof roles[number]): PptxDiagramComponentInfo => {
    const [role, attr, relationType, rootName] = spec;
    const relationshipId = refs ? attribute(refs, attr, c.r) : undefined;
    if (relationshipId === undefined) return { role, status: "missing-reference" };
    const base = { role, relationshipId }, relation = c.findRelationship(part, relationshipId);
    if (!relation) return { ...base, status: "missing-relationship" };
    const target = { ...base, target: relation.target, ...(relation.targetPart ? { part: relation.targetPart } : {}) };
    if (relation.type !== `${c.r}/${relationType}`) return { ...target, status: "wrong-relationship-type" };
    if (relation.targetMode === "External") return { ...target, status: "external" };
    if (relation.fragment !== undefined) return { ...target, status: "fragment-target" };
    const inventory = relation.targetPart ? parts.get(relation.targetPart) : undefined;
    const result = { ...target, ...(inventory ? { contentType: inventory.contentType } : {}) };
    // QuickStyle relationships point to a diagramStyle content type.
    const contentName = role === "style" ? "diagramStyle" : relationType;
    if (inventory?.contentType.toLowerCase() !== `application/vnd.openxmlformats-officedocument.drawingml.${contentName}+xml`.toLowerCase()) return { ...result, status: "wrong-content-type" };
    const root = c.rawXml.get(inventory.name);
    if (root?.namespaceUri !== dgm || root.local !== rootName) return { ...result, status: "wrong-root" };
    const layoutIdentifier = role === "layout" ? attribute(root, "uniqueId") : undefined;
    return { ...result, status: "located", ...(layoutIdentifier !== undefined ? { layoutIdentifier } : {}) };
  };
  return async (part: string, root: XmlElement): Promise<PptxDiagramInfo[]> => {
    const common = singleChild(root, c.p, "cSld", part);
    const tree = common ? singleChild(common, c.p, "spTree", part) : undefined;
    if (!tree) return [];
    const counts = new Map<string, number>(), diagrams: { node: XmlElement; groups: number[]; id?: string; name?: string; refs?: XmlElement }[] = [];
    const stack: { node: XmlElement; groups: number[] }[] = [{ node: tree, groups: [] }];
    while (stack.length) {
      await c.checkpoint();
      const { node, groups } = stack.pop()!;
      const nvName = nonVisual[node.local];
      if (!nvName || node.namespaceUri !== c.p) continue;
      const nv = singleChild(node, c.p, nvName, part), props = nv ? singleChild(nv, c.p, "cNvPr", part) : undefined;
      const id = props ? attribute(props, "id") : undefined, key = identityKey(id);
      if (key !== undefined) counts.set(key, (counts.get(key) ?? 0) + 1);
      if (node.local === "graphicFrame") {
        const graphic = singleChild(node, c.a, "graphic", part);
        const data = graphic ? singleChild(graphic, c.a, "graphicData", part) : undefined;
        if (data && attribute(data, "uri") === dgm) {
          const refs = singleChild(data, dgm, "relIds", part), name = props ? attribute(props, "name") : undefined;
          diagrams.push({ node, groups, ...(id !== undefined ? { id } : {}), ...(name !== undefined ? { name } : {}), ...(refs ? { refs } : {}) });
        }
      }
      if (node.local === "spTree" || node.local === "grpSp") {
        const nested = node.local === "grpSp" ? [...groups, (await c.sourceElement(part, node)).elementIndex] : groups;
        const nodes = children(node);
        for (let i = nodes.length - 1; i >= 0; i--) if (nodes[i]!.namespaceUri === c.p && nonVisual[nodes[i]!.local]) stack.push({ node: nodes[i]!, groups: nested });
      }
    }
    const result: PptxDiagramInfo[] = [];
    for (const item of diagrams) {
      await c.checkpoint();
      const key = identityKey(item.id);
      result.push({ elementIndex: (await c.sourceElement(part, item.node)).elementIndex, groupElementIndices: item.groups,
        ...(item.id !== undefined ? { drawingId: item.id } : {}), ...(item.name !== undefined ? { name: item.name } : {}),
        identityStatus: item.id === undefined ? "missing" : key === undefined ? "invalid" : counts.get(key)! > 1 ? "ambiguous" : "unique",
        components: roles.map(spec => component(part, item.refs, spec)) });
    }
    return result;
  };
}
