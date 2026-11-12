import { attribute, children, integer, invalidStructure } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { ContentReader } from "./content-reader.js";

export interface NotesMasterBody { node: XmlElement; placeholder: XmlElement; sourceId: string; grouped: boolean }
const containers: Readonly<Record<string, string>> = { spTree: "nvGrpSpPr", sp: "nvSpPr", pic: "nvPicPr", graphicFrame: "nvGraphicFramePr", cxnSp: "nvCxnSpPr", grpSp: "nvGrpSpPr" };

/** Notes inherit by placeholder type, not slide/layout idx. Never inspect extension islands. */
export async function notesMasterBody(root: XmlElement, r: ContentReader, checkpoint: () => Promise<void>): Promise<{ count: number; candidate?: NotesMasterBody }> {
  const common = r.one(root, r.p, "cSld"), tree = common ? r.one(common, r.p, "spTree") : undefined;
  const stack = tree ? [{ node: tree, grouped: false }] : [], ids = new Set<number>(), indices = new Set<number>();
  let count = 0, candidate: NotesMasterBody | undefined;
  while (stack.length) {
    await checkpoint(); const { node, grouped } = stack.pop()!;
    if (node.namespaceUri !== r.p || !containers[node.local]) continue;
    const nv = r.one(node, r.p, containers[node.local]!), identity = nv ? r.one(nv, r.p, "cNvPr") : undefined;
    if (!identity) invalidStructure("Missing notes-master drawing identity.", r.part);
    const id = integer(attribute(identity!, "id"), 0, 4294967295, "notes master cNvPr/@id", r.part);
    if (ids.has(id)) invalidStructure("Duplicate notes-master drawing identity.", r.part); ids.add(id);
    const nvPr = nv ? r.one(nv, r.p, "nvPr") : undefined, placeholder = nvPr ? r.one(nvPr, r.p, "ph") : undefined;
    if (placeholder) {
      const index = attribute(placeholder, "idx") === undefined ? 0 : integer(attribute(placeholder, "idx"), 0, 4294967294, "notes master ph/@idx", r.part);
      if (indices.has(index)) invalidStructure("Duplicate notes-master placeholder index.", r.part); indices.add(index);
      if (attribute(placeholder, "type")?.trim() === "body") {
        count++; candidate = { node, placeholder, sourceId: String(id), grouped: grouped || node.local !== "sp" };
      }
    }
    if (node.local === "spTree" || node.local === "grpSp") for (const child of children(node)) if (child.namespaceUri === r.p && containers[child.local]) stack.push({ node: child, grouped: grouped || node.local === "grpSp" });
  }
  return { count, ...(candidate ? { candidate } : {}) };
}
