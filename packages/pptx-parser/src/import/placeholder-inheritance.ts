import type { Box, TransformInput } from "@pptkit/core";
import type { structureContext } from "../presentation/structure-context.js";
import { attribute, children, integer, invalidStructure, singleChild } from "../presentation/structure-xml.js";
import type { PptxInspection } from "../types/structure.js";
import { PptxImportError, type PptxSessionLimits, type PptxSourceLocation } from "../types/import.js";
import type { PptxInheritedDimension, PptxPlaceholderDimension, PptxPlaceholderInfo, PptxPlaceholderLevel, PptxPlaceholderMatch } from "../types/inheritance.js";
import type { XmlElement } from "../xml/read-xml.js";
import { contentPaths, contentReader } from "./content-reader.js";

const containers: Readonly<Record<string, string>> = { sp: "nvSpPr", pic: "nvPicPr", graphicFrame: "nvGraphicFramePr", cxnSp: "nvCxnSpPr", grpSp: "nvGrpSpPr" };
// Source enum tokens; no inference from names, order or Core placeholder keys.
const masterTypes: Readonly<Record<string, string>> = {
  title: "title", ctrTitle: "title", body: "body", subTitle: "body", obj: "body",
  chart: "body", tbl: "body", dgm: "body", media: "body", clipArt: "body", pic: "body",
  dt: "dt", ftr: "ftr", sldNum: "sldNum",
};
const placeholderTypes = new Set([...Object.keys(masterTypes), "hdr", "sldImg"]);
type Dimensions = Partial<Record<PptxPlaceholderDimension, PptxInheritedDimension>>;
interface Entry {
  node: XmlElement;
  properties?: XmlElement;
  xfrm?: XmlElement;
  info: PptxPlaceholderInfo;
  local: Dimensions;
  transform: TransformInput;
  parent?: Entry;
}
interface Candidate { count: number; entry: Entry }
interface Part {
  entries: Entry[];
  byNode: WeakMap<XmlElement, Entry>;
  byIndex: Map<number, Candidate>;
  byType: Map<string, Candidate>;
  paths: WeakMap<XmlElement, string>;
  resolved: boolean;
}

/** One import owns caches and budgets. No XML is cloned, edited or exposed. */
export function placeholderInheritance(context: Awaited<ReturnType<typeof structureContext>>, inspection: PptxInspection, limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>) {
  const parts = new Map<string, Part>(), entries: Entry[] = [];
  const layouts = new Map(inspection.layouts.map(layout => [layout.part, layout]));
  const addCandidate = <T>(map: Map<T, Candidate>, key: T, entry: Entry) => {
    const previous = map.get(key);
    if (previous) previous.count++; else map.set(key, { count: 1, entry });
  };
  const read = async (part: string, level: PptxPlaceholderLevel): Promise<Part> => {
    const cached = parts.get(part); if (cached) return cached;
    const localName = level === "slide" ? "sld" : level === "layout" ? "sldLayout" : "sldMaster";
    const root = await context.root(part, localName, level === "slide" ? "presentationml.slide" : level === "layout" ? "presentationml.slideLayout" : "presentationml.slideMaster");
    const paths = await contentPaths(root, checkpoint);
    const result: Part = { entries: [], byNode: new WeakMap(), byIndex: new Map(), byType: new Map(), paths, resolved: false };
    const r = contentReader(part, context.p, context.a);
    const common = singleChild(root, context.p, "cSld", part), tree = common ? singleChild(common, context.p, "spTree", part) : undefined;
    const stack = tree ? [{ node: tree, grouped: false }] : [], ids = new Set<number>();
    while (stack.length) {
      await checkpoint();
      const { node, grouped } = stack.pop()!;
      if (node.namespaceUri !== context.p) continue;
      const container = node.local === "spTree" ? "nvGrpSpPr" : containers[node.local];
      if (!container) continue;
      const nv = singleChild(node, context.p, container, part);
      const identity = nv ? singleChild(nv, context.p, "cNvPr", part) : undefined;
      let sourceObjectId: string | undefined;
      if (identity) {
        const id = integer(attribute(identity, "id"), 0, 4294967295, "cNvPr/@id", part);
        if (ids.has(id)) invalidStructure("Duplicate drawing identity in a placeholder source part.", part);
        ids.add(id); sourceObjectId = String(id);
      }
      const nvPr = nv ? singleChild(nv, context.p, "nvPr", part) : undefined;
      const ph = nvPr ? singleChild(nvPr, context.p, "ph", part) : undefined;
      if (ph) {
        if (!identity) invalidStructure("Placeholder is missing its drawing identity.", part);
        const source = { part, path: paths.get(node)!, sourceObjectId: sourceObjectId! };
        if (entries.length >= limits.maxPlaceholderEntries) throw new PptxImportError("import-limit-exceeded", "Placeholder provenance exceeds maxPlaceholderEntries.", source);
        const typeValue = attribute(ph, "type"), indexValue = attribute(ph, "idx");
        const type = typeValue === undefined ? "obj" : typeValue.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
        if (!placeholderTypes.has(type)) invalidStructure("Invalid placeholder type.", part);
        const index = indexValue === undefined ? 0 : integer(indexValue, 0, 4294967295, "ph/@idx", part);
        const properties = node.local === "graphicFrame" ? node : singleChild(node, context.p, node.local === "grpSp" ? "grpSpPr" : "spPr", part);
        const xfrm = properties ? singleChild(properties, node.local === "graphicFrame" ? context.p : context.a, "xfrm", part) : undefined;
        const local: Dimensions = {};
        const transform: TransformInput = { rotation: 0, flipH: false, flipV: false };
        if (xfrm) {
          const off = singleChild(xfrm, context.a, "off", part), ext = singleChild(xfrm, context.a, "ext", part);
          const dimension = (owner: XmlElement, attr: string, key: PptxPlaceholderDimension, min: number, max: number) => {
            const valueEmu = r.number(owner, attr, min, max);
            local[key] = { valueEmu, level, source: { ...source, path: `${paths.get(owner)!}/@${attr}` } };
          };
          if (off) { dimension(off, "x", "x", -27273042329600, 27273042316900); dimension(off, "y", "y", -27273042329600, 27273042316900); }
          if (ext) { dimension(ext, "cx", "width", 0, 27273042316900); dimension(ext, "cy", "height", 0, 27273042316900); }
          transform.rotation = r.number(xfrm, "rot", -2147483648, 2147483647, 0) / 60000;
          transform.flipH = r.bool(xfrm, "flipH"); transform.flipV = r.bool(xfrm, "flipV");
        }
        const info: PptxPlaceholderInfo = { source, declaration: { ...source, path: paths.get(ph)! }, level, type, index, explicitType: typeValue !== undefined, explicitIndex: indexValue !== undefined, grouped, dimensions: local, geometryStatus: "incomplete" };
        const entry: Entry = { node, ...(properties ? { properties } : {}), ...(xfrm ? { xfrm } : {}), info, local, transform };
        result.entries.push(entry); result.byNode.set(node, entry); entries.push(entry);
        addCandidate(result.byIndex, index, entry); addCandidate(result.byType, type, entry);
      }
      // Follow drawing containers only; extension islands and unknown wrappers
      // must never manufacture placeholder declarations or collide with IDs.
      if (node.local === "grpSp" || node.local === "spTree") {
        const nodes = children(node);
        for (let i = nodes.length - 1; i >= 0; i--) if (nodes[i]!.namespaceUri === context.p && containers[nodes[i]!.local]) stack.push({ node: nodes[i]!, grouped: grouped || node.local === "grpSp" });
      }
    }
    parts.set(part, result); return result;
  };
  const bind = (entry: Entry, parent: Part | undefined, parentPart: string | undefined, rule: "index" | "type") => {
    const info = entry.info, key = rule === "index" ? info.index : masterTypes[info.type];
    const candidate = key === undefined ? undefined : rule === "index" ? parent?.byIndex.get(key as number) : parent?.byType.get(key as string);
    let match: PptxPlaceholderMatch = { rule, key: key ?? info.type, candidateCount: candidate?.count ?? 0, status: candidate ? candidate.count === 1 ? "matched" : "ambiguous" : "missing", ...(parentPart ? { part: parentPart } : {}) };
    if (info.grouped || entry.node.local === "grpSp" || (candidate?.count === 1 && (candidate.entry.info.grouped || candidate.entry.node.local === "grpSp")) || key === undefined) {
      match = { ...match, status: "unsupported", reason: key === undefined ? "This placeholder type has no supported master matching rule." : "Grouped placeholders have no validated inheritance coordinate policy." };
    } else if (candidate?.count === 1) {
      entry.parent = candidate.entry;
      match = { ...match, source: candidate.entry.info.source };
    }
    entry.info = { ...info, parent: match };
  };
  const resolve = (entry: Entry) => {
    const dimensions: Dimensions = { ...entry.parent?.info.dimensions, ...entry.local };
    let reason: string | undefined;
    if (entry.info.grouped || entry.node.local === "grpSp") reason = "Grouped placeholder geometry inheritance is not mapped.";
    const inherited = Object.values(dimensions).some(value => value.level !== entry.info.level);
    // Dimensions are independently inherited. Rotated/reflected ancestor
    // transforms need native evidence before combining with partial local xfrm.
    if (inherited) for (let parent = entry.parent; parent; parent = parent.parent) {
      if (parent.transform.rotation || parent.transform.flipH || parent.transform.flipV) reason = "Inherited non-neutral rotation/reflection requires a separately validated transform cascade.";
    }
    const complete = ["x", "y", "width", "height"].every(key => Object.hasOwn(dimensions, key));
    entry.info = { ...entry.info, dimensions, geometryStatus: reason ? "unsupported" : complete ? "resolved" : "incomplete", ...(reason ? { reason } : {}) };
  };
  const slide = async (part: string, layoutPart?: string) => {
    const local = await read(part, "slide"), layout = layoutPart ? await read(layoutPart, "layout") : undefined;
    const masterPart = layoutPart ? layouts.get(layoutPart)?.masterPart : undefined;
    const master = masterPart ? await read(masterPart, "master") : undefined;
    if (master && !master.resolved) {
      for (const entry of master.entries) { await checkpoint(); resolve(entry); }
      master.resolved = true;
    }
    if (layout && !layout.resolved) {
      for (const entry of layout.entries) { await checkpoint(); bind(entry, master, masterPart, "type"); resolve(entry); }
      layout.resolved = true;
    }
    for (const entry of local.entries) { await checkpoint(); bind(entry, layout, layoutPart, "index"); resolve(entry); }
    return {
      paths: local.paths,
      get: (node: XmlElement) => local.byNode.get(node)?.info,
      ancestors(node: XmlElement) {
        const result: { node: XmlElement; part: string; level: PptxPlaceholderLevel; paths: WeakMap<XmlElement, string>; source: PptxSourceLocation }[] = [];
        for (let entry = local.byNode.get(node)?.parent; entry; entry = entry.parent) {
          result.unshift({ node: entry.node, part: entry.info.source.part, level: entry.info.level, paths: parts.get(entry.info.source.part)!.paths, source: entry.info.source });
        }
        return result;
      },
      geometry(node: XmlElement): { box: Box; transform: TransformInput } | undefined {
        const entry = local.byNode.get(node);
        if (entry?.info.geometryStatus !== "resolved") return undefined;
        const d = entry.info.dimensions;
        return { box: { x: d.x!.valueEmu / 12700, y: d.y!.valueEmu / 12700, width: d.width!.valueEmu / 12700, height: d.height!.valueEmu / 12700 }, transform: { ...entry.transform } };
      },
      /** Nearest geometry choice wins, including an unsupported custom choice. */
      preset(node: XmlElement): { node: XmlElement; source: PptxSourceLocation } | undefined {
        for (let entry = local.byNode.get(node); entry; entry = entry.parent) {
          if (!entry.properties) continue;
          const choices = children(entry.properties).filter(child => child.namespaceUri === context.a && ["prstGeom", "custGeom"].includes(child.local));
          if (choices.length > 1) invalidStructure("Conflicting placeholder geometry choices.", entry.info.source.part);
          if (choices[0]) return { node: choices[0], source: { ...entry.info.source, path: parts.get(entry.info.source.part)!.paths.get(choices[0])! } };
        }
        return undefined;
      },
    };
  };
  return { slide, paths: (part: string) => parts.get(part)?.paths, report: () => ({ placeholders: entries.map(entry => entry.info) }) };
}
