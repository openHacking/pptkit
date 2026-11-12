import { snapshotOptions } from "../opc/read-package.js";
import { byteViewInfo } from "../utils/byte-view.js";
import { structureContext } from "../presentation/structure-context.js";
import { readStructureFromContext } from "../presentation/read-structure.js";
import { attribute, children, singleChild } from "../presentation/structure-xml.js";
import { PackageInspectionError } from "../types/archive.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { MasterPictureInventory, MasterPictureInventoryItem, MasterPictureInventoryLimits, MasterPictureInventoryOptions, MasterPictureRelationship } from "../types/master-pictures.js";

export const DEFAULT_MASTER_PICTURE_LIMITS: Readonly<MasterPictureInventoryLimits> = Object.freeze({ maxWorkUnits: 1_000_000, maxRecords: 100_000, maxTextCharacters: 1_000_000 });
function limits(input?: Partial<MasterPictureInventoryLimits>): MasterPictureInventoryLimits {
  if (input !== undefined && (!input || typeof input !== "object" || Array.isArray(input) || ![Object.prototype, null].includes(Object.getPrototypeOf(input)))) throw new TypeError("Invalid master picture limits.");
  const result: Record<string, number> = { ...DEFAULT_MASTER_PICTURE_LIMITS };
  if (input) for (const key of Reflect.ownKeys(input)) {
    const desc = Object.getOwnPropertyDescriptor(input, key)!;
    if (typeof key !== "string" || !Object.hasOwn(result, key) || !("value" in desc)) throw new TypeError("Unknown/accessor master picture limit.");
    if (desc.value !== undefined) result[key] = desc.value;
  }
  for (const [key, value] of Object.entries(result)) if (!Number.isSafeInteger(value) || value <= 0) throw new RangeError(`Invalid master picture limit ${key}.`);
  return result as unknown as MasterPictureInventoryLimits;
}
function freeze<T>(value: T): T {
  const stack: unknown[] = [value];
  while (stack.length) { const item = stack.pop(); if (item && typeof item === "object" && !Object.isFrozen(item)) { Object.freeze(item); for (const child of Object.values(item)) stack.push(child); } }
  return value;
}
/** Inventory effective pictures in listed slide masters; it does not project or prove write safety. */
export async function inspectMasterPictures(bytes: Uint8Array, options: MasterPictureInventoryOptions = {}): Promise<MasterPictureInventory> {
  if (!(bytes instanceof Uint8Array)) throw new TypeError("PPTX input must be Uint8Array.");
  if (byteViewInfo(bytes).shared) throw new TypeError("SharedArrayBuffer input is not supported.");
  const cap = limits(options.masterPictureLimits), captured = snapshotOptions(options);
  const c = await structureContext(bytes, { limits: captured.limits, zipLimits: captured.zipLimits, xmlLimits: captured.xmlLimits, ...(captured.signal ? { signal: captured.signal } : {}) });
  const structure = await readStructureFromContext(c); let workUnits = 0, recordCount = 0, textCharacters = 0;
  const fail = (name: string): never => { throw new PackageInspectionError("package-limit-exceeded", `Master picture inventory exceeds ${name}.`); };
  const work = async () => { if (++workUnits > cap.maxWorkUnits) fail("maxWorkUnits"); await c.checkpoint(); };
  const record = () => { if (++recordCount > cap.maxRecords) fail("maxRecords"); };
  const text = (value: string | undefined) => { if (value !== undefined && (textCharacters += value.length) > cap.maxTextCharacters) fail("maxTextCharacters"); return value; };
  const listed = new Map<string, string | null>();
  for (const master of structure.slideMasters) for (const layout of master.layouts) {
    await work(); const prior = listed.get(layout.part);
    listed.set(layout.part, prior !== undefined && prior !== master.part ? null : master.part);
  }
  const layoutMaster = new Map<string, string>();
  for (const layout of structure.layouts) { await work(); if (layout.masterPart && listed.get(layout.part) === layout.masterPart) layoutMaster.set(layout.part, layout.masterPart); }
  const slidesByMaster = new Map<string, number[]>();
  const usesByMaster = new Map<string, MasterPictureInventory["masters"][number]["slideUses"][number][]>();
  for (const slide of structure.slides) {
    await work(); const master = slide.layoutPart ? layoutMaster.get(slide.layoutPart) : undefined;
    if (master && slide.layoutPart) {
      const list = slidesByMaster.get(master) ?? []; list.push(slide.id); slidesByMaster.set(master, list);
      const slideRoot = await c.root(slide.part, "sld"), layoutRoot = await c.root(slide.layoutPart, "sldLayout");
      const slideShowMasterSp = text(attribute(slideRoot, "showMasterSp")), layoutShowMasterSp = text(attribute(layoutRoot, "showMasterSp"));
      record(); const uses = usesByMaster.get(master) ?? [];
      uses.push({ slideId: slide.id, slidePart: text(slide.part)!, layoutPart: text(slide.layoutPart)!, ...(slideShowMasterSp !== undefined ? { slideShowMasterSp } : {}), ...(layoutShowMasterSp !== undefined ? { layoutShowMasterSp } : {}) });
      usesByMaster.set(master, uses);
    }
  }
  const masters: MasterPictureInventory["masters"][number][] = [];
  for (const master of structure.slideMasters) {
    await work(); record(); const part = text(master.part)!; const root = await c.root(part, "sldMaster");
    const common = singleChild(root, c.p, "cSld", part), tree = common ? singleChild(common, c.p, "spTree", part) : undefined;
    const pictures: MasterPictureInventoryItem[] = []; let otherDrawingCount = 0;
    const stack: { node: XmlElement; groups: readonly number[] }[] = tree ? children(tree).reverse().map(node => ({ node, groups: [] })) : [];
    while (stack.length) {
      await work(); const { node, groups } = stack.pop()!;
      if (node.namespaceUri !== c.p) continue;
      if (node.local === "grpSp") {
        const group = await c.sourceElement(part, node), nested = children(node);
        for (const _ of groups) await work();
        const next = [...groups, group.elementIndex];
        for (let i = nested.length - 1; i >= 0; i--) stack.push({ node: nested[i]!, groups: next });
        continue;
      }
      if (node.local !== "pic") { if (["sp", "cxnSp", "graphicFrame"].includes(node.local)) otherDrawingCount++; continue; }
      record(); for (const _ of groups) record(); const source = await c.sourceElement(part, node);
      const nv = singleChild(node, c.p, "nvPicPr", part), cnv = nv ? singleChild(nv, c.p, "cNvPr", part) : undefined, nvPr = nv ? singleChild(nv, c.p, "nvPr", part) : undefined;
      const ph = nvPr ? singleChild(nvPr, c.p, "ph", part) : undefined;
      const fill = singleChild(node, c.p, "blipFill", part), blip = fill ? singleChild(fill, c.a, "blip", part) : undefined, rect = fill ? singleChild(fill, c.a, "srcRect", part) : undefined;
      const properties = singleChild(node, c.p, "spPr", part), xfrm = properties ? singleChild(properties, c.a, "xfrm", part) : undefined;
      const off = xfrm ? singleChild(xfrm, c.a, "off", part) : undefined, ext = xfrm ? singleChild(xfrm, c.a, "ext", part) : undefined;
      const attr = (node: XmlElement | undefined, name: string) => text(node ? attribute(node, name) : undefined);
      const rel = (id: string | undefined): MasterPictureRelationship | undefined => {
        if (id === undefined) return undefined; record(); const relation = c.findRelationship(part, id);
        const status = !relation ? "missing" : relation.type !== `${c.r}/image` ? "wrong-type" : relation.fragment !== undefined ? "fragmented" : relation.targetMode === "External" ? "external-image" : relation.targetPart ? "internal-image" : "missing";
        return { id: text(id)!, status, ...(status === "internal-image" && relation?.targetPart ? { targetPart: text(relation.targetPart)! } : {}) };
      };
      const id = attr(cnv, "id"), name = attr(cnv, "name"), type = attr(ph, "type"), index = attr(ph, "idx");
      const embed = rel(blip ? attribute(blip, "embed", c.r) : undefined), link = rel(blip ? attribute(blip, "link", c.r) : undefined);
      const x=attr(off,"x"),y=attr(off,"y"),cx=attr(ext,"cx"),cy=attr(ext,"cy"),rotation=attr(xfrm,"rot"),flipH=attr(xfrm,"flipH"),flipV=attr(xfrm,"flipV");
      const left=attr(rect,"l"),top=attr(rect,"t"),right=attr(rect,"r"),bottom=attr(rect,"b");
      pictures.push({ source: { part, elementIndex: source.elementIndex }, ...(id !== undefined ? { sourceObjectId:id } : {}), ...(name !== undefined ? { name } : {}), groupElementIndices:[...groups],
        ...(ph ? { placeholder:{ ...(type !== undefined ? { type } : {}), ...(index !== undefined ? { index } : {}) } } : {}), ...(embed ? { embed } : {}), ...(link ? { link } : {}),
        ...(xfrm ? { transform:{ ...(x !== undefined?{x}:{}),...(y !== undefined?{y}:{}),...(cx !== undefined?{cx}:{}),...(cy !== undefined?{cy}:{}),...(rotation !== undefined?{rotation}:{}),...(flipH !== undefined?{flipH}:{}),...(flipV !== undefined?{flipV}:{}) } } : {}),
        ...(rect ? { crop:{ ...(left !== undefined?{left}:{}),...(top !== undefined?{top}:{}),...(right !== undefined?{right}:{}),...(bottom !== undefined?{bottom}:{}) } } : {}) });
    }
    const layoutParts = master.layouts.map(l => text(l.part)!); for (const _ of layoutParts) record();
    const slideIds = slidesByMaster.get(master.part) ?? []; for (const _ of slideIds) record();
    masters.push({ part, sourceMasterId: master.id, layoutParts, slideIds, slideUses: usesByMaster.get(master.part) ?? [], pictures, otherDrawingCount });
  }
  return freeze({ schemaVersion:"pptkit-master-picture-inventory-v1", dialect:c.entry.dialect, masters, workUnits, recordCount, textCharacters, coverage:"effective-listed-masters-only", mapping:"not-performed" });
}
