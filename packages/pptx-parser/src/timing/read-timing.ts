import { snapshotOptions } from "../opc/read-package.js";
import { structureContext } from "../presentation/structure-context.js";
import { readStructureFromContext } from "../presentation/read-structure.js";
import { attribute, children } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { PptxSourceLocation } from "../types/source.js";
import type { PptxTimingDrawingInfo, PptxTimingFinding, PptxTimingInspection, PptxTimingInspectionOptions, PptxTimingNodeInfo, PptxTimingReference } from "../types/timing.js";
import { TimingBudget } from "./budget.js";

const behaviors = new Set(["anim", "animClr", "animEffect", "animMotion", "animRot", "animScale", "cmd", "set", "audio", "video"]);
const containers = new Set("tnLst par seq excl childTnLst subTnLst cBhvr cMediaNode stCondLst endCondLst prevCondLst nextCondLst endSync tgtEl sldTgt attrNameLst attrName sndTgt inkTgt txEl charRg pRg graphicEl oleChartEl subSp bg bldLst bldP bldDgm bldOleChart bldGraphic bldSub tmplLst tmpl rtn to from by progress tavLst tav val strVal boolVal intVal fltVal clrVal".split(" "));
const nonVisual = new Set(["nvSpPr", "nvGrpSpPr", "nvGraphicFramePr", "nvPicPr", "nvCxnSpPr"]);
const owners = new Set(["sp", "grpSp", "graphicFrame", "pic", "cxnSp", "spTree"]);
interface Meta { node: XmlElement; index: number; ordinal: number; parent?: Meta; root?: Meta; parentTime?: Meta; inDrawings: boolean }
function numericKey(value: string | undefined): string | undefined {
  if (value === undefined || !/^\d+$/.test(value)) return undefined;
  const text = value.replace(/^0+(?=\d)/, "");
  return text.length <= 10 && Number(text) <= 4294967295 ? text : undefined;
}
function freeze<T>(value: T): T {
  const pending: unknown[] = [value];
  while (pending.length) { const next = pending.pop(); if (next && typeof next === "object" && !Object.isFrozen(next)) { Object.freeze(next); for (const child of Object.values(next)) pending.push(child); } }
  return value;
}

/** Read-only source inventory. No native behaviors become authored Core effects. */
export async function inspectNativeTiming(bytes: Uint8Array, options: PptxTimingInspectionOptions = {}): Promise<PptxTimingInspection> {
  const budget = new TimingBudget(options.timingLimits), captured = snapshotOptions(options);
  const c = await structureContext(bytes, { limits: captured.limits, zipLimits: captured.zipLimits, xmlLimits: captured.xmlLimits, ...(captured.signal ? { signal: captured.signal } : {}) }), structure = await readStructureFromContext(c);
  const findings: PptxTimingFinding[] = [], slides: PptxTimingInspection["slides"][number][] = [];
  const step = async () => { budget.work(); await c.checkpoint(); };
  function location(part: string, meta: Meta): PptxSourceLocation {
    budget.record(); budget.text(part);
    const components: string[] = [];
    for (let current: Meta | undefined = meta; current; current = current.parent) {
      const ordinal = String(current.ordinal); budget.charge(current.node.namespaceUri.length + current.node.local.length + ordinal.length + 6);
      components.push(`Q{${current.node.namespaceUri}}${current.node.local}[${ordinal}]`);
    }
    return { part, elementIndex: meta.index, path: "/" + components.reverse().join("/") };
  }
  function finding(code: PptxTimingFinding["code"], source: PptxSourceLocation, message: string): void {
    budget.finding(); budget.record(); budget.text(message); findings.push({ code, source, message });
  }
  for (const slide of structure.slides) {
    await step(); const effective = await c.root(slide.part, "sld", "presentationml.slide"), raw = c.rawXml.get(slide.part)!;
    const all: Meta[] = [], pending: Meta[] = [{ node: raw, index: 0, ordinal: 1, inDrawings: false }];
    while (pending.length) {
      await step(); const meta = pending.pop()!; meta.index = all.length + 1; const node = meta.node;
      if (node.namespaceUri === c.p && node.local === "timing" && !meta.root) meta.root = meta;
      meta.inDrawings ||= node.namespaceUri === c.p && node.local === "spTree" && meta.parent?.node.namespaceUri === c.p && meta.parent.node.local === "cSld";
      all.push(meta);
      const counts = new Map<string, number>(), nested: Meta[] = [];
      for (const child of children(node)) {
        await step(); const key = JSON.stringify([child.namespaceUri, child.local]), ordinal = (counts.get(key) ?? 0) + 1; counts.set(key, ordinal);
        nested.push({ node: child, index: 0, ordinal, parent: meta, ...(meta.root ? { root: meta.root } : {}), ...(node.namespaceUri === c.p && node.local === "cTn" ? { parentTime: meta } : meta.parentTime ? { parentTime: meta.parentTime } : {}), inDrawings: meta.inDrawings });
      }
      for (let i = nested.length - 1; i >= 0; i--) pending.push(nested[i]!);
    }
    const activeTiming = new Set<number>(), activeDrawings = new Set<number>();
    const active: { node: XmlElement; parent?: XmlElement; grandparent?: XmlElement; timing: boolean; drawing: boolean }[] = [{ node: effective, timing: false, drawing: false }];
    while (active.length) {
      await step(); const item = active.pop()!, node = item.node;
      if ((node.namespaceUri === c.p || node.namespaceUri === c.a) && node.local === "ext") continue;
      const timing = item.timing || (item.parent === effective && node.namespaceUri === c.p && node.local === "timing");
      const drawing = item.drawing || (node.namespaceUri === c.p && node.local === "spTree" && item.parent?.namespaceUri === c.p && item.parent.local === "cSld");
      if (timing || (drawing && node.namespaceUri === c.p && node.local === "cNvPr" && item.parent?.namespaceUri === c.p && nonVisual.has(item.parent.local) && item.grandparent?.namespaceUri === c.p && owners.has(item.grandparent.local))) {
        const source = await c.sourceElement(slide.part, node);
        if (timing) activeTiming.add(source.elementIndex); else activeDrawings.add(source.elementIndex);
      }
      for (const child of children(node)) { await step(); active.push({ node: child, parent: node, ...(item.parent ? { grandparent: item.parent } : {}), timing, drawing }); }
    }
    const drawings: PptxTimingDrawingInfo[] = [], nodes: PptxTimingNodeInfo[] = [];
    for (const meta of all) {
      const n = meta.node;
      if (!meta.inDrawings || n.namespaceUri !== c.p || n.local !== "cNvPr" || meta.parent?.node.namespaceUri !== c.p || !nonVisual.has(meta.parent.node.local) || meta.parent.parent?.node.namespaceUri !== c.p || !owners.has(meta.parent.parent.node.local)) continue;
      await step(); budget.record(); const rawId = attribute(n, "id"), name = attribute(n, "name");
      const info: PptxTimingDrawingInfo = { source: location(slide.part, meta), view: activeDrawings.has(meta.index) ? "effective" : "raw-only", ownerKind: budget.text(meta.parent.parent.node.local), ...(rawId === undefined ? {} : { rawId: budget.text(rawId) }), ...(name === undefined ? {} : { name: budget.text(name) }) };
      drawings.push(info);
    }
    for (const meta of all) {
      if (!meta.root) continue;
      await step(); budget.record(); const n = meta.node, native = n.namespaceUri === c.p;
      const role: PptxTimingNodeInfo["role"] = !native ? "unknown" : n.local === "timing" ? "timing" : n.local === "cTn" ? "time-node" : behaviors.has(n.local) ? "behavior" : n.local === "cond" ? "condition" : n.local === "spTgt" ? "drawing-target" : n.local === "tn" ? "time-reference" : containers.has(n.local) ? "container" : "unknown";
      const attrs = n.attributes.map(a => { budget.record(); return { namespaceUri: budget.text(a.namespaceUri), local: budget.text(a.local), value: budget.text(a.value) }; });
      let text = "";
      for (const child of n.children) if (child.kind === "text" || child.kind === "cdata") { budget.text(child.value); text += child.value; }
      const rawId = role === "time-node" ? attribute(n, "id") : undefined;
      const info: PptxTimingNodeInfo = { source: location(slide.part, meta), namespaceUri: budget.text(n.namespaceUri), local: budget.text(n.local), view: activeTiming.has(meta.index) ? "effective" : "raw-only", role, timingRootElementIndex: meta.root.index,
        ...(meta.parent ? { parentElementIndex: meta.parent.index } : {}), ...(meta.parentTime ? { parentTimeNodeElementIndex: meta.parentTime.index } : {}), attributes: attrs, ...(text ? { text } : {}), ...(rawId === undefined ? {} : { rawId: budget.text(rawId) }) };
      if (role === "unknown") finding("unknown-timing-node", info.source, "Unclassified source timing node; no execution or rewrite semantics are inferred.");
      if (role === "timing" && info.view === "raw-only") finding("raw-only-timing", info.source, "Not a direct effective slide timing container; raw content is inventoried without binding.");
      nodes.push(info);
    }
    const timeNodes = nodes.filter(n => n.role === "time-node" && n.view === "effective");
    const effectiveDrawings = drawings.filter(n => n.view === "effective");
    type Identity = { rawId?: string; source: PptxSourceLocation };
    async function identities(values: readonly Identity[], duplicateCode: "duplicate-time-id" | "duplicate-drawing-id") {
      const exact = new Map<string, Identity[]>(), numeric = new Map<string, Identity[]>();
      for (const item of values) {
        await step();
        if (item.rawId === undefined || !item.rawId.length) { finding("missing-identity", item.source, "Source identity is absent/empty; no ID is synthesized."); continue; }
        const list = exact.get(item.rawId) ?? []; list.push(item); exact.set(item.rawId, list);
        const key = numericKey(item.rawId);
        if (key !== undefined) { const aliases = numeric.get(key) ?? []; aliases.push(item); numeric.set(key, aliases); }
        if (key !== item.rawId) finding("uninterpreted-identity", item.source, "Noncanonical source identity is retained without semantic normalization.");
      }
      for (const list of exact.values()) { await step(); if (list.length > 1) finding(duplicateCode, list[0]!.source, "Multiple source records share the same lexical identity."); }
      return { exact, numeric };
    }
    const timeIds = await identities(timeNodes, "duplicate-time-id"), drawingIds = await identities(effectiveDrawings, "duplicate-drawing-id");
    const reference = async (rawId: string | undefined, info: PptxTimingNodeInfo, index: Awaited<ReturnType<typeof identities>>): Promise<PptxTimingReference> => {
      await step(); budget.record(); if (rawId !== undefined) budget.text(rawId);
      const key = numericKey(rawId), candidates = info.view === "raw-only" ? [] : key === undefined ? rawId === undefined ? [] : index.exact.get(rawId) ?? [] : index.numeric.get(key) ?? [];
      const locations: PptxTimingReference["candidates"][number][] = [];
      for (const candidate of candidates) { await step(); budget.record(); const meta = all[candidate.source.elementIndex! - 1]!; locations.push({ source: location(slide.part, meta), match: candidate.rawId === rawId ? "exact" : "numeric-alias" }); }
      const exactCount = locations.filter(l => l.match === "exact").length;
      const status = info.view !== "effective" || key !== rawId || (locations.length && !exactCount) ? "source-only" : locations.length > 1 ? "ambiguous" : exactCount === 1 ? "located" : "missing";
      if (status === "missing" || status === "ambiguous") finding(status === "missing" ? "missing-reference" : "ambiguous-reference", info.source, "Source reference has no unique supported lexical identity candidate.");
      return { ...(rawId === undefined ? {} : { rawId }), status, candidates: locations, ...(status === "source-only" ? { reason: budget.text("Raw-only or noncanonical/alias identity is not bound by this inventory profile.") } : {}) };
    };
    for (const info of nodes) {
      const meta = all[info.source.elementIndex! - 1]!, n = meta.node;
      const spid = n.namespaceUri === c.p ? attribute(n, "spid") : undefined;
      if (spid !== undefined || info.role === "drawing-target") {
        const targetKind = info.role !== "drawing-target" ? "uninterpreted" : children(n).length ? "subobject" : n.attributes.some(a => !(a.namespaceUri === "" && a.local === "spid") && a.namespaceUri !== "http://www.w3.org/2000/xmlns/") ? "uninterpreted" : "whole-object";
        Object.assign(info, { drawingReference: await reference(spid, info, drawingIds), targetKind });
        if (targetKind !== "whole-object") finding("subobject-target", info.source, "Target/build details remain source-only; whole-object edit authority is not inferred.");
      }
      if (info.role === "time-reference") Object.assign(info, { timeReference: await reference(attribute(n, "val"), info, timeIds) });
    }
    const roots = nodes.filter(n => n.role === "timing" && n.view === "effective");
    if (roots.length > 1) finding("multiple-effective-timing", roots[0]!.source, "Multiple effective timing containers require a separate interpretation policy.");
    budget.record(); budget.text(slide.part); slides.push({ slideId: slide.id, part: slide.part, nodes, drawings });
  }
  c.abort();
  return freeze({ schemaVersion: "pptkit-native-timing-inventory-v1", dialect: structure.dialect, slides, findings, workUnits: budget.workUnits, recordCount: budget.recordCount, textCharacters: budget.textCharacters, nativePlayback: "not-performed", authoredMapping: "not-performed" });
}
