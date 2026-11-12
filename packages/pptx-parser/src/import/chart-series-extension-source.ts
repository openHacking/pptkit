import { readXml, type XmlElement } from "../xml/read-xml.js";
import type { XmlInspectionOptions } from "../types/xml.js";

/** Recognizes one observed series extension shape; this is not package closure. */
export async function inspectColumnSeriesExtensions(bytes: Uint8Array, options: XmlInspectionOptions, maxWork: number, checkpoint: () => Promise<void>): Promise<readonly { readonly seriesId: number; readonly uniqueId: string }[]> {
  if (!Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError("Series extension inspection requires a positive work limit.");
  if (!(bytes instanceof Uint8Array) || bytes.length * 2 > maxWork) throw new RangeError("Series extension inspection exceeds its work limit.");
  let work = bytes.length * 2;
  const step = async () => { if (++work > maxWork) throw new RangeError("Series extension inspection exceeds its work limit."); await checkpoint(); };
  const { root } = await readXml(bytes, options, true);
  const C = "http://schemas.openxmlformats.org/drawingml/2006/chart", A = "http://schemas.openxmlformats.org/drawingml/2006/main", C14 = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart", C16 = "http://schemas.microsoft.com/office/drawing/2014/chart", XMLNS = "http://www.w3.org/2000/xmlns/";
  if (!root || root.namespaceUri !== C || root.local !== "chartSpace") throw new TypeError("Series extension inspection requires a transitional chart space.");
  const all: { node: XmlElement; parent?: XmlElement }[] = [], stack: { node: XmlElement; parent?: XmlElement }[] = [{ node: root }];
  while (stack.length) { await step(); const item = stack.pop()!; all.push(item); for (let i = item.node.children.length - 1; i >= 0; i--) { const child = item.node.children[i]!; if (child.kind === "element") stack.push({ node: child, parent: item.node }); } }
  const direct = (node: XmlElement, namespace: string, local: string) => node.children.filter((child): child is XmlElement => child.kind === "element" && child.namespaceUri === namespace && child.local === local);
  const one = (node: XmlElement, namespace: string, local: string) => { const found = direct(node, namespace, local); if (found.length !== 1) throw new TypeError(`Expected one ${local} child.`); return found[0]!; };
  const attrs = (node: XmlElement, keys: readonly string[]) => { const ordinary = node.attributes.filter(attr => attr.namespaceUri !== XMLNS); if (ordinary.length !== keys.length || ordinary.some(attr => attr.namespaceUri !== "" || !keys.includes(attr.local))) throw new TypeError("Unsupported series extension attributes."); return ordinary; };
  const elements = (node: XmlElement, count: number) => { if (node.children.some(child => child.kind !== "element" && (child.kind !== "text" || !/^[\t\r\n ]*$/u.test(child.value))) || node.children.filter(child => child.kind === "element").length !== count) throw new TypeError("Unsupported series extension content."); };
  const chart = one(root, C, "chart"), plot = one(chart, C, "plotArea"), family = one(plot, C, "barChart"), series = direct(family, C, "ser");
  const direction = one(family, C, "barDir"); if (attrs(direction, ["val"])[0]!.value !== "col") throw new TypeError("Series extension inspection requires a column chart."); elements(direction, 0);
  const groups = plot.children.filter((child): child is XmlElement => child.kind === "element" && child.namespaceUri === C && /Chart$/u.test(child.local));
  if (groups.length !== 1 || groups[0] !== family || !series.length) throw new TypeError("Series extensions require one chart family.");
  const familySet = new Set(series), extensionLists = all.filter(item => item.node.namespaceUri === C && item.node.local === "extLst");
  if (extensionLists.length !== series.length || extensionLists.some(item => !item.parent || !familySet.has(item.parent))) throw new TypeError("Series extensions must be directly owned by every series.");
  const result: { seriesId: number; uniqueId: string }[] = [], ids = new Set<number>(), uids = new Set<string>(), allowedForeign = new Set<XmlElement>();
  for (const ser of series) {
    await step(); const idx = one(ser, C, "idx"), idText = attrs(idx, ["val"])[0]!.value; elements(idx, 0);
    const seriesId = Number(idText); if (!/^\d+$/u.test(idText) || !Number.isSafeInteger(seriesId) || seriesId < 0 || seriesId > 4294967295 || ids.has(seriesId)) throw new TypeError("Invalid series extension owner identity."); ids.add(seriesId);
    const list = one(ser, C, "extLst"); attrs(list, []); elements(list, 2);
    const extensions = direct(list, C, "ext"); if (extensions.length !== 2) throw new TypeError("Unknown series extensions.");
    const byUri = new Map(extensions.map(ext => [attrs(ext, ["uri"])[0]!.value.toUpperCase(), ext]));
    const fill = byUri.get("{6F2FDCE9-48DA-4B69-8628-5D25D57E5C99}"), unique = byUri.get("{C3380CC4-5D6E-409C-BE32-E72D297353CC}");
    if (!fill || !unique || byUri.size !== 2) throw new TypeError("Unknown or duplicate series extension URI.");
    elements(fill, 1); const invert = one(fill, C14, "invertSolidFillFmt"); attrs(invert, []); elements(invert, 1); allowedForeign.add(invert);
    const properties = one(invert, C14, "spPr"); attrs(properties, []); elements(properties, 2); allowedForeign.add(properties);
    const solid = one(properties, A, "solidFill"), line = one(properties, A, "ln"); attrs(solid, []); elements(solid, 1); attrs(line, []); elements(line, 1);
    const color = one(solid, A, "srgbClr"), noFill = one(line, A, "noFill"); const rgb = attrs(color, ["val"])[0]!.value; elements(color, 0); attrs(noFill, []); elements(noFill, 0);
    if (!/^[0-9a-fA-F]{6}$/u.test(rgb)) throw new TypeError("Unsupported negative series color.");
    elements(unique, 1); const uid = one(unique, C16, "uniqueId"); attrs(uid, ["val"]); elements(uid, 0); allowedForeign.add(uid);
    const uniqueId = uid.attributes.find(attr => attr.namespaceUri === "" && attr.local === "val")!.value.toUpperCase();
    if (!/^\{[0-9A-F]{8}-(?:[0-9A-F]{4}-){3}[0-9A-F]{12}\}$/u.test(uniqueId) || uids.has(uniqueId)) throw new TypeError("Invalid or duplicated series unique ID.");
    uids.add(uniqueId); result.push({ seriesId, uniqueId });
  }
  if (all.some(({ node }) => (node.namespaceUri === C14 || node.namespaceUri === C16) && !allowedForeign.has(node))) throw new TypeError("Unrecognized chart extension content.");
  return result;
}
