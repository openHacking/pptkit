import type { TextParagraphInput } from "@pptkit/core";
import type { structureContext } from "../presentation/structure-context.js";
import { children, integer, invalidStructure } from "../presentation/structure-xml.js";
import type { ImportDiagnostic, ImportSourceIdentity, PptxSourceLocation } from "../types/import.js";
import type { XmlElement } from "../xml/read-xml.js";
import { contentAction } from "./content-actions.js";
import { contentPaths, contentReader, type ContentReader } from "./content-reader.js";
import { contentText, type ContentTextRun, type ContentTextInheritance } from "./content-text.js";

type NotesCapabilities = (feature: string, source: PptxSourceLocation, reason: string, mapping: "editable" | "fallback" | "unsupported") => void;

/** Notes text is distinct from notes-page artwork and master inheritance. */
export async function contentNotes(
  part: string,
  slideId: string,
  context: Awaited<ReturnType<typeof structureContext>>,
  slideIds: ReadonlyMap<string, string>,
  checkpoint: () => Promise<void>,
  capabilities: NotesCapabilities,
  recordRun?: (run: ContentTextRun) => void,
  styleInheritance?: (r: ContentReader, paths: WeakMap<XmlElement, string>, source: PptxSourceLocation, placeholder: XmlElement) => ContentTextInheritance | Promise<ContentTextInheritance>,
  notesTextProfile?: "master-style-v1",
  fieldTextProfile?: "slide-number-v1" | "date-and-number-v1" | "date-time-and-number-v1",
): Promise<{ content?: TextParagraphInput[]; identity?: ImportSourceIdentity; diagnostics: ImportDiagnostic[] }> {
  const root = await context.root(part, "notes"), r = contentReader(part, context.p, context.a);
  const paths = await contentPaths(root, checkpoint), diagnostics: ImportDiagnostic[] = [];
  r.known(root);
  const common = r.one(root, r.p, "cSld"), tree = common ? r.one(common, r.p, "spTree") : undefined;
  if (common) r.known(common);
  if (tree) r.known(tree);
  const directDrawings = new Set(tree ? children(tree) : []);
  const seen = new Set<number>(), owners = new Map<string, string>();
  const bodies: { node: XmlElement; sourceId: string }[] = [];
  const drawings = tree ? [tree] : [];
  while (drawings.length) {
    await checkpoint();
    const node = drawings.pop()!;
    if (node.namespaceUri !== r.p) continue;
    const kind = { spTree: "nvGrpSpPr", grpSp: "nvGrpSpPr", sp: "nvSpPr", pic: "nvPicPr", cxnSp: "nvCxnSpPr", graphicFrame: "nvGraphicFramePr" }[node.local];
    if (!kind) continue;
    const nv = r.one(node, r.p, kind), cnv = nv ? r.one(nv, r.p, "cNvPr") : undefined;
    if (!cnv) invalidStructure("Missing notes drawing identity.", part);
    const sourceId = r.number(cnv!, "id", 0, 4294967295);
    if (seen.has(sourceId)) invalidStructure("Duplicate drawing identity in a notes part.", part);
    seen.add(sourceId);
    owners.set(paths.get(node)!, String(sourceId));
    if (node === tree) {
      r.known(nv!);
      // Root bookkeeping has no notes-text semantics. Its attributes and group
      // geometry remain visible to the source-property audit.
      const group = r.one(nv!, r.p, "cNvGrpSpPr"), nvPr = r.one(nv!, r.p, "nvPr");
      if (group) r.known(group);
      if (nvPr) r.known(nvPr);
    }
    if (node.local === "sp") {
      const nvPr = r.one(nv!, r.p, "nvPr"), ph = nvPr ? r.one(nvPr, r.p, "ph") : undefined;
      // Never infer speaker notes from a shape name, order or arbitrary txBody.
      // Nested placeholder shapes require notes-page/group semantics first.
      if (ph && directDrawings.has(node) && r.attr(ph, "type") === "body") bodies.push({ node, sourceId: String(sourceId) });
    }
    if (node.local === "spTree" || node.local === "grpSp") for (const child of children(node)) drawings.push(child);
  }
  let content: TextParagraphInput[] | undefined, identity: ImportSourceIdentity | undefined;
  const candidate = bodies.length === 1 ? bodies[0] : undefined;
  const body = candidate ? r.one(candidate.node, r.p, "txBody") : undefined;
  if (candidate && body) {
    const at: PptxSourceLocation = { part, path: paths.get(candidate.node)!, sourceObjectId: candidate.sourceId };
    r.known(candidate.node);
    const nv = r.one(candidate.node, r.p, "nvSpPr")!, nvPr = r.one(nv, r.p, "nvPr")!, ph = r.one(nvPr, r.p, "ph")!;
    r.known(nv); r.known(nvPr); r.attr(ph, "type");
    if (r.attr(ph, "idx") !== undefined) integer(r.attr(ph, "idx"), 0, 4294967295, "notes placeholder/@idx", part);
    content = (await contentText(body, r, checkpoint, (owner) => {
      const click = contentAction(owner, r, context, slideIds);
      if (!click) return undefined;
      const source = { ...at, path: paths.get(click.node)!, ...(click.relationshipId !== undefined ? { relationshipId: click.relationshipId } : {}) };
      capabilities("notes.click-action", source, click.reason, click.action ? "editable" : "unsupported");
      if (!click.action) diagnostics.push({ code: "notes-click-action-unmapped", severity: "warning", source, slideId, message: click.reason });
      return click.action;
    }, false, await styleInheritance?.(r, paths, at, ph), recordRun, fieldTextProfile)).content;
    identity = { kind: "notes", id: slideId, source: at };
    capabilities("notes.body.mapped-fields", at, notesTextProfile ? "Direct notes-body paragraphs, runs and supported clicks map with the explicit bounded notes-master text profile. Notes-page geometry and other placeholders are not projected; scalar-save permissions remain separate. Slide SVG does not render speaker notes." : "Direct notes-body paragraphs, runs and supported clicks map to Core speaker notes. Notes-page geometry, other placeholders, master/theme inheritance and edit-save are not mapped. Slide SVG does not render speaker notes.", "editable");
    capabilities("notes.inheritance-and-defaults", at, notesTextProfile ? "Supported notesStyle, type-matched body-placeholder list defaults and local text properties resolve under the independent notes theme. Unsupported or absent values still use reported Core projection defaults; full notes-page appearance is not certified." : "Notes master/theme/list-style inheritance is unresolved. Omitted styles use Core projection defaults, not inferred source appearance.", "fallback");
  } else if (bodies.length > 0) {
    const reason = bodies.length > 1 ? "Multiple direct notes-body placeholders cannot be unambiguously mapped to one Core notes body." : "Notes-body placeholder has no local text body; inherited notes text is not synthesized.";
    capabilities("notes.body", { part, path: paths.get(tree!)! }, reason, "unsupported");
    diagnostics.push({ code: "notes-body-unmapped", severity: "warning", source: { part }, slideId, message: reason });
  }
  await r.audit(root, paths.get(root)!, (path, reason) => {
    let ownerPath = path;
    while (ownerPath && !owners.has(ownerPath)) ownerPath = ownerPath.slice(0, ownerPath.lastIndexOf("/"));
    const sourceId = owners.get(ownerPath);
    capabilities("notes.unmapped-property", { part, path, ...(sourceId !== undefined ? { sourceObjectId: sourceId } : {}) }, reason, "unsupported");
  }, checkpoint);
  return { ...(content ? { content } : {}), ...(identity ? { identity } : {}), diagnostics };
}
