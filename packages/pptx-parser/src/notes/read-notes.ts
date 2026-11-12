import { snapshotOptions } from "../opc/read-package.js";
import { structureContext } from "../presentation/structure-context.js";
import { readStructureFromContext } from "../presentation/read-structure.js";
import { attribute, children, singleChild } from "../presentation/structure-xml.js";
import { MCE } from "../compatibility/configuration.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { PptxSourceLocation } from "../types/source.js";
import type { PptxNotesDrawingInfo, PptxNotesFieldInfo, PptxNotesFinding, PptxNotesGeometryInfo, PptxNotesInspection, PptxNotesInspectionOptions, PptxNotesNodeInfo, PptxNotesPageInfo, PptxNotesPartInfo, PptxNotesPlaceholderMatch, PptxNotesRole, PptxNotesTextInfo } from "../types/notes.js";
import { NotesBudget } from "./budget.js";

const owners: Readonly<Record<string, string>> = { spTree: "nvGrpSpPr", grpSp: "nvGrpSpPr", sp: "nvSpPr", pic: "nvPicPr", cxnSp: "nvCxnSpPr", graphicFrame: "nvGraphicFramePr" };
const roles: Readonly<Record<string, PptxNotesRole>> = { body: "body", sldImg: "slide-image", hdr: "header", ftr: "footer", dt: "date", sldNum: "slide-number" };
const bookkeeping = new Set(["nvGrpSpPr", "grpSpPr", "extLst"]);
interface Meta { node: XmlElement; index: number; parent?: Meta; rootChain: boolean; partRoot?: boolean; tree: boolean; opaque: boolean; container?: number; owner?: number; parentOwner?: number; field?: number; paragraph?: number; drawing?: PptxNotesDrawingInfo["kind"] }
function numericId(raw: string): string | undefined {
  const value = raw.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, ""), number = Number(value);
  return /^[+-]?\d+$/.test(value) && Number.isSafeInteger(number) && number >= 0 && number <= 4294967295 ? String(number) : undefined;
}
function freeze<T>(value: T): T {
  const stack: unknown[] = [value];
  while (stack.length) { const next = stack.pop(); if (next && typeof next === "object" && !Object.isFrozen(next)) { Object.freeze(next); for (const child of Object.values(next)) stack.push(child); } }
  return value;
}

/** Read-only source inventory. No notes geometry, cached field or artwork becomes authored state. */
export async function inspectNotes(bytes: Uint8Array, options: PptxNotesInspectionOptions = {}): Promise<PptxNotesInspection> {
  const budget = new NotesBudget(options.notesLimits), captured = snapshotOptions(options);
  const c = await structureContext(bytes, { limits: captured.limits, zipLimits: captured.zipLimits, xmlLimits: captured.xmlLimits, ...(captured.signal ? { signal: captured.signal } : {}) });
  const structure = await readStructureFromContext(c), parts = new Map<string, PptxNotesPartInfo>(), findings: PptxNotesFinding[] = [];
  const step = async () => { budget.work(); await c.checkpoint(); };
  const source = (part: string, index: number): PptxSourceLocation => { budget.record(); return { part: budget.text(part), elementIndex: index }; };
  const copySource = (value: PptxSourceLocation) => source(value.part, value.elementIndex!);
  const expanded = (node: Pick<XmlElement, "namespaceUri" | "local" | "name" | "prefix">) => ({ namespaceUri: budget.text(node.namespaceUri), local: budget.text(node.local), name: budget.text(node.name), prefix: budget.text(node.prefix) });
  const attrs = (node: XmlElement): PptxNotesNodeInfo["attributes"] => node.attributes.map(a => { budget.work(); budget.record(); return { ...expanded(a), value: budget.text(a.value) }; });
  const finding = (code: PptxNotesFinding["code"], at: PptxSourceLocation, message: string) => { budget.finding(); budget.record(); findings.push({ code, source: copySource(at), message: budget.text(message) }); };
  // a:ext under xfrm is an extent, not an extension island.
  const extension = (node: XmlElement, parent?: XmlElement) => [c.p, c.a].includes(node.namespaceUri) && node.local === "ext" && parent?.local === "extLst" && [c.p, c.a].includes(parent.namespaceUri);
  const kind = (node: XmlElement, tree: boolean, container?: number): PptxNotesDrawingInfo["kind"] | undefined => {
    if (node.namespaceUri === c.p && node.local === "spTree" && tree) return "spTree";
    if (container === undefined || node.namespaceUri === MCE || (node.namespaceUri === c.p && bookkeeping.has(node.local))) return undefined;
    return node.namespaceUri === c.p && node.local !== "spTree" && Object.hasOwn(owners, node.local) ? node.local as PptxNotesDrawingInfo["kind"] : "uninterpreted";
  };
  const readPart = async (part: string, partKind: "notes" | "notesMaster"): Promise<PptxNotesPartInfo> => {
    const existing = parts.get(part); if (existing) return existing;
    const effective = await c.root(part, partKind, `presentationml.${partKind === "notes" ? "notesSlide" : "notesMaster"}`), raw = c.rawXml.get(part)!;
    const all: Meta[] = [], byNode = new WeakMap<XmlElement, Meta>();
    const pending: Meta[] = [{ node: raw, index: 0, rootChain: true, tree: false, opaque: false }];
    while (pending.length) {
      await step(); const m = pending.pop()!, n = m.node; m.index = all.length + 1; all.push(m); byNode.set(n, m);
      m.partRoot = m.rootChain && n.namespaceUri === c.p && n.local === partKind;
      const rootTree = n.namespaceUri === c.p && n.local === "spTree" && m.parent?.node.namespaceUri === c.p && m.parent.node.local === "cSld" && m.parent.parent?.partRoot === true;
      m.tree ||= rootTree;
      if (!m.opaque && m.tree) { const drawing = kind(n, rootTree, m.container); if (drawing) m.drawing = drawing; }
      if (m.drawing) { if (m.owner !== undefined) m.parentOwner = m.owner; m.owner = m.index; }
      if (!m.opaque && n.namespaceUri === c.a && n.local === "fld") m.field = m.index;
      if (!m.opaque && n.namespaceUri === c.a && n.local === "p") m.paragraph = m.index;
      const island = extension(n, m.parent?.node) || m.drawing === "uninterpreted";
      const container = m.drawing === "spTree" || m.drawing === "grpSp" ? m.index : n.namespaceUri === MCE ? m.container : undefined;
      const nested = children(n);
      for (let i = nested.length - 1; i >= 0; i--) pending.push({ node: nested[i]!, index: 0, parent: m, rootChain: m.rootChain && n.namespaceUri === MCE, tree: m.tree, opaque: m.opaque || island, ...(container !== undefined ? { container } : {}), ...(m.owner !== undefined ? { owner: m.owner } : {}), ...(m.field !== undefined ? { field: m.field } : {}), ...(m.paragraph !== undefined ? { paragraph: m.paragraph } : {}) });
    }
    const effectiveByIndex = new Map<number, XmlElement>(), projectedMeta = new WeakMap<XmlElement, Meta>();
    const active = new Set<number>(), activeDrawing = new Map<number, { direct: boolean; parent?: number }>();
    const common = singleChild(effective, c.p, "cSld", part);
    const effectiveStack: { node: XmlElement; parent?: XmlElement; tree: boolean; owner?: number; container?: number; field?: number; paragraph?: number }[] = [{ node: effective, tree: false }];
    while (effectiveStack.length) {
      await step(); const item = effectiveStack.pop()!, n = item.node, at = await c.sourceElement(part, n); active.add(at.elementIndex);
      const rootTree = n.namespaceUri === c.p && n.local === "spTree" && item.parent === common;
      const tree = item.tree || rootTree;
      const drawing = tree ? kind(n, rootTree, item.container) : undefined;
      if (drawing) activeDrawing.set(at.elementIndex, { direct: item.parent?.namespaceUri === c.p && item.parent.local === "spTree", ...(item.owner !== undefined ? { parent: item.owner } : {}) });
      // ProcessContent may unwrap an otherwise opaque raw drawing root. Use
      // the projected ownership for actual effective nodes, retaining raw-only
      // wrappers separately rather than hiding their now-effective children.
      const rawMeta = all[at.elementIndex - 1]!;
      effectiveByIndex.set(at.elementIndex, n); projectedMeta.set(n, rawMeta);
      rawMeta.tree = tree; rawMeta.opaque = false;
      if (drawing) rawMeta.drawing = drawing; else delete rawMeta.drawing;
      const owner = drawing ? at.elementIndex : item.owner, field = n.namespaceUri === c.a && n.local === "fld" ? at.elementIndex : item.field;
      const paragraph = n.namespaceUri === c.a && n.local === "p" ? at.elementIndex : item.paragraph;
      if (owner !== undefined) rawMeta.owner = owner; else delete rawMeta.owner;
      if (field !== undefined) rawMeta.field = field; else delete rawMeta.field;
      if (paragraph !== undefined) rawMeta.paragraph = paragraph; else delete rawMeta.paragraph;
      if (extension(n, item.parent) || drawing === "uninterpreted") continue;
      const container = drawing === "spTree" || drawing === "grpSp" ? at.elementIndex : undefined;
      for (const child of children(n)) effectiveStack.push({ node: child, parent: n, tree, ...(owner !== undefined ? { owner } : {}), ...(container !== undefined ? { container } : {}), ...(field !== undefined ? { field } : {}), ...(paragraph !== undefined ? { paragraph } : {}) });
    }
    const nodeInfo = (m: Meta): PptxNotesNodeInfo => { budget.record(); return { source: source(part, m.index), name: expanded(m.node), view: active.has(m.index) ? "effective" : "raw-only", attributes: attrs(m.node) }; };
    const texts: PptxNotesTextInfo[] = [], fields: PptxNotesFieldInfo[] = [], geometries: PptxNotesGeometryInfo[] = [], drawings: PptxNotesDrawingInfo[] = [], headerFooter: PptxNotesNodeInfo[] = [], extensions: PptxNotesPartInfo["extensions"][number][] = [];
    const text = (m: Meta): PptxNotesTextInfo => {
      budget.record(); const fragments: string[] = [];
      for (const child of m.node.children) { budget.work(); if (child.kind === "text" || child.kind === "cdata") fragments.push(budget.text(child.value)); }
      return { source: source(part, m.index), view: active.has(m.index) ? "effective" : "raw-only", ...(m.owner !== undefined ? { ownerDrawingElementIndex: m.owner } : {}), ...(m.paragraph !== undefined ? { paragraphElementIndex: m.paragraph } : {}), value: fragments.join("") };
    };
    const ids = new Set<string>();
    for (const m of all) {
      await step(); const n = m.node; if (m.opaque) continue;
      if (extension(n, m.parent?.node)) {
        const names = children(n).map(x => { budget.work(); budget.record(); return expanded(x); });
        extensions.push({ ...nodeInfo(m), children: names });
      }
      if (n.namespaceUri === c.p && n.local === "hf") headerFooter.push(nodeInfo(m));
      if (m.drawing) {
        const observed = effectiveByIndex.get(m.index) ?? n;
        // Inactive MCE content is observed, not admitted as active structure.
        // Do not reject its ambiguous children or silently choose one.
        const one = (parent: XmlElement | undefined, ns: string, local: string): XmlElement | undefined => {
          if (!parent) return undefined;
          if (active.has(m.index)) return singleChild(parent, ns, local, part);
          const candidates = children(parent).filter(child => child.namespaceUri === ns && child.local === local);
          return candidates.length === 1 ? candidates[0] : undefined;
        };
        const nv = one(observed, c.p, owners[m.drawing] ?? ""), identity = one(nv, c.p, "cNvPr"), ph = one(one(nv, c.p, "nvPr"), c.p, "ph");
        const rawId = identity ? attribute(identity, "id") : undefined, drawingName = identity ? attribute(identity, "name") : undefined;
        const type = ph ? attribute(ph, "type") : undefined, index = ph ? attribute(ph, "idx") : undefined, phMeta = ph ? (projectedMeta.get(ph) ?? byNode.get(ph))! : undefined;
        const role: PptxNotesRole = m.drawing === "spTree" ? "group-root" : ph ? Object.hasOwn(roles, type ?? "") ? roles[type!]! : "other-placeholder" : "artwork";
        const effectiveOwner = activeDrawing.get(m.index);
        const parentOwner = effectiveOwner ? effectiveOwner.parent : m.parentOwner;
        const info: PptxNotesDrawingInfo = { ...nodeInfo(m), kind: m.drawing, role, direct: effectiveOwner?.direct ?? false,
          ...(parentOwner !== undefined ? { parentDrawingElementIndex: parentOwner } : {}),
          ...(rawId !== undefined ? { rawId: budget.text(rawId) } : {}), ...(drawingName !== undefined ? { drawingName: budget.text(drawingName) } : {}),
          ...(identity ? { nonVisual: nodeInfo((projectedMeta.get(identity) ?? byNode.get(identity))!) } : {}),
          textBodyCount: children(observed).filter(x => x.namespaceUri === c.p && x.local === "txBody").length,
          ...(phMeta ? { placeholder: { ...nodeInfo(phMeta), ...(type !== undefined ? { type: budget.text(type) } : {}), ...(index !== undefined ? { index: budget.text(index) } : {}) } } : {}) };
        drawings.push(info);
        if (info.view === "effective") {
          if (m.drawing === "uninterpreted") finding("unknown-drawing", info.source, "Uninterpreted drawing root; its subtree is not mapped.");
          else if (rawId === undefined) finding("missing-drawing-id", info.source, "Drawing has no source id.");
          else {
            const key = numericId(rawId);
            if (key === undefined) finding("uninterpreted-drawing-id", info.source, "Drawing id is not an admitted unsigned integer.");
            else if (ids.has(key)) finding("duplicate-drawing-id", info.source, "Numeric drawing identity is ambiguous within this effective part.");
            else ids.add(key);
          }
          if (ph && role === "other-placeholder") finding("unknown-placeholder", info.source, "Placeholder has an absent or unclassified explicit type; no default type is guessed.");
        }
      }
      if (m.tree && n.namespaceUri === c.a && n.local === "fld") {
        const cachedTexts: PptxNotesTextInfo[] = [];
        for (const child of children(effectiveByIndex.get(m.index) ?? n)) if (child.namespaceUri === c.a && child.local === "t") cachedTexts.push(text((projectedMeta.get(child) ?? byNode.get(child))!));
        const id = attribute(n, "id"), type = attribute(n, "type");
        fields.push({ ...nodeInfo(m), ...(m.owner !== undefined ? { ownerDrawingElementIndex: m.owner } : {}), ...(m.paragraph !== undefined ? { paragraphElementIndex: m.paragraph } : {}), ...(id !== undefined ? { id: budget.text(id) } : {}), ...(type !== undefined ? { type: budget.text(type) } : {}), cachedTexts });
      } else if (m.tree && n.namespaceUri === c.a && n.local === "t" && m.field === undefined) texts.push(text(m));
      if (m.tree && m.owner !== undefined && [c.p, c.a].includes(n.namespaceUri)) {
        const geometryKind = n.local === "xfrm" ? "transform" : n.namespaceUri === c.a && n.local === "prstGeom" ? "preset" : n.namespaceUri === c.a && n.local === "custGeom" ? "custom" : undefined;
        if (geometryKind && m.parent && ["spPr", "grpSpPr", "graphicFrame"].includes(m.parent.node.local) && m.parent.node.namespaceUri === c.p) geometries.push({ ...nodeInfo(m), ownerDrawingElementIndex: m.owner, kind: geometryKind, components: children(n).map(x => nodeInfo(byNode.get(x)!)) });
      }
    }
    const rawRoot = await c.sourceElement(part, effective);
    const info: PptxNotesPartInfo = { part: budget.text(part), source: source(part, rawRoot.elementIndex), kind: partKind, attributes: attrs(rawRoot.node), drawings, geometries, texts, fields, headerFooter, extensions };
    budget.record(); parts.set(part, info); return info;
  };
  const masterTypes = new Map<string, Map<string, PptxNotesDrawingInfo[]>>();
  const master = async (part: string) => {
    const info = await readPart(part, "notesMaster");
    let types = masterTypes.get(part);
    if (!types) {
      types = new Map(); masterTypes.set(part, types);
      for (const drawing of info.drawings) { await step(); if (drawing.view !== "effective" || drawing.placeholder?.type === undefined) continue; const type = drawing.placeholder.type, list = types.get(type) ?? []; list.push(drawing); types.set(type, list); }
    }
    return types;
  };
  for (const ref of structure.notesMasters) await master(ref.part);
  const pages: PptxNotesPageInfo[] = [];
  for (const slide of structure.slides) {
    await step(); budget.record();
    const base = { slideId: slide.id, slidePart: budget.text(slide.part) };
    if (!slide.notesPart) { pages.push({ ...base, body: { status: "absent", candidates: [] }, placeholderMatches: [] }); continue; }
    const part = slide.notesPart, local = await readPart(part, "notes"), masterPart = c.optionalTarget(part, "notesMaster"), overridePart = c.optionalTarget(part, "themeOverride");
    const types = masterPart ? await master(masterPart) : undefined, masterThemePart = masterPart ? c.optionalTarget(masterPart, "theme") : undefined;
    if (masterThemePart) await c.root(masterThemePart, "theme", "theme", c.a);
    if (overridePart) await c.root(overridePart, "themeOverride", "themeOverride", c.a);
    const bodies = local.drawings.filter(d => d.view === "effective" && d.direct && d.kind === "sp" && d.role === "body");
    const status: PptxNotesPageInfo["body"]["status"] = bodies.length === 0 ? "absent" : bodies.length > 1 ? "ambiguous-direct" : bodies[0]!.textBodyCount === 0 ? "missing-local-text" : bodies[0]!.textBodyCount > 1 ? "ambiguous-local-text" : "unique-direct";
    const placeholderMatches: PptxNotesPlaceholderMatch[] = [];
    for (const drawing of local.drawings) {
      await step(); if (drawing.view !== "effective" || !drawing.placeholder) continue;
      budget.record(); const type = drawing.placeholder.type, candidates = type === undefined ? [] : types?.get(type) ?? [];
      const refs: PptxSourceLocation[] = [];
      for (const candidate of candidates) { await step(); refs.push(copySource(candidate.source)); }
      placeholderMatches.push({ source: copySource(drawing.placeholder.source), status: !masterPart ? "missing-master" : type === undefined ? "missing-type" : candidates.length === 0 ? "missing-candidate" : candidates.length > 1 ? "ambiguous" : !drawing.direct || !candidates[0]!.direct ? "grouped" : "matched-direct-type", candidates: refs });
    }
    pages.push({ ...base, notesPart: budget.text(part), ...(masterPart ? { masterPart: budget.text(masterPart) } : {}), ...(masterThemePart ? { masterThemePart: budget.text(masterThemePart) } : {}), ...(overridePart ? { overridePart: budget.text(overridePart) } : {}), body: { status, candidates: bodies.map(d => copySource(d.source)) }, placeholderMatches });
  }
  const structureFindings = [];
  for (const finding of structure.findings) { await step(); budget.finding(); budget.record(); structureFindings.push({ ...finding, part: budget.text(finding.part), message: budget.text(finding.message) }); }
  c.abort();
  return freeze({ schemaVersion: "pptkit-notes-source-inventory-v1", dialect: c.entry.dialect, notesSize: { ...structure.notesSize }, pages, parts: [...parts.values()], findings, structureFindings, workUnits: budget.workUnits, recordCount: budget.recordCount, textCharacters: budget.textCharacters, geometryPolicy: "unresolved-source-only", fieldEvaluation: "not-performed", authoredMapping: "not-performed" });
}
