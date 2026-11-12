import { readXml, type XmlElement } from "../xml/read-xml.js";
import type { XmlInspectionOptions } from "../types/xml.js";
import type { ChartLegendIndexSource } from "./chart-legend-order.js";

/** Private eligibility only: no dependency closure or write permission. */
export async function captureColumnLegendSources(bytes: Uint8Array, options: XmlInspectionOptions, maxWork: number, checkpoint: () => Promise<void>): Promise<{ entries: readonly ChartLegendIndexSource[]; physicalIds: readonly number[]; workUnits: number }> {
  if (!Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError("Legend source capture requires a positive work limit.");
  if (!(bytes instanceof Uint8Array) || bytes.length * 2 > maxWork) throw new RangeError("Legend source capture exceeds its work limit.");
  let work = bytes.length * 2;
  const charge = async () => { if (++work > maxWork) throw new RangeError("Legend source capture exceeds its work limit."); await checkpoint(); };
  const { root } = await readXml(bytes, options, true);
  const c = root?.namespaceUri;
  if (!root || root.local !== "chartSpace" || !["http://schemas.openxmlformats.org/drawingml/2006/chart", "http://purl.oclc.org/ooxml/drawingml/chart"].includes(c!)) throw new TypeError("Expected a supported chart source.");
  const index = new Map<XmlElement, number>(), all: XmlElement[] = [], stack = [root];
  while (stack.length) { await charge(); const node = stack.pop()!; index.set(node, all.length + 1); all.push(node); for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === "element") stack.push(child); } }
  const direct = (node: XmlElement, local: string) => node.children.filter((child): child is XmlElement => child.kind === "element" && child.namespaceUri === c && child.local === local);
  const only = (node: XmlElement, local: string) => { const found = direct(node, local); if (found.length !== 1) throw new TypeError(`Legend source requires one direct ${local}.`); return found[0]!; };
  const simpleValue = (node: XmlElement) => {
    if (node.attributes.length !== 1 || node.attributes[0]!.namespaceUri !== "" || node.attributes[0]!.local !== "val" || node.children.some(child => child.kind !== "text" || !/^[\t\r\n ]*$/u.test(child.value))) throw new TypeError("Legend source contains unsupported index metadata.");
    return node.attributes[0]!.value;
  };
  const chart = only(root, "chart"), plot = only(chart, "plotArea"), family = only(plot, "barChart"), legend = only(chart, "legend");
  const allowedNamespaces = new Set([c, "http://schemas.openxmlformats.org/drawingml/2006/main", "http://purl.oclc.org/ooxml/drawingml/main", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"]);
  if (direct(family, "barDir").length !== 1 || simpleValue(only(family, "barDir")) !== "col" || all.some(node => !allowedNamespaces.has(node.namespaceUri) || node.namespaceUri === c && ["trendline", "extLst", "pivotFmt", "bandFmt"].includes(node.local))) throw new TypeError("Legend source contains unsupported chart semantics.");
  const groups = plot.children.filter((node): node is XmlElement => node.kind === "element" && node.namespaceUri === c && /Chart$/u.test(node.local));
  if (groups.length !== 1 || groups[0] !== family) throw new TypeError("Legend source requires one column-chart family.");
  const series = direct(family, "ser"), idLexemes = series.map(item => simpleValue(only(item, "idx"))), physicalIds = idLexemes.map(Number);
  if (!series.length || idLexemes.some(value => !/^\d+$/u.test(value)) || physicalIds.some(id => !Number.isSafeInteger(id) || id < 0 || id > 4294967295) || new Set(physicalIds).size !== series.length) throw new TypeError("Legend source has invalid series identities.");
  const legendEntries = direct(legend, "legendEntry");
  if (all.filter(node => node.namespaceUri === c && node.local === "legendEntry").length !== legendEntries.length) throw new TypeError("Legend entry ancestry is ambiguous.");
  const entries: ChartLegendIndexSource[] = [], positions = new Set<number>();
  for (const entry of legendEntries) {
    await charge();
    if (entry.attributes.length || entry.children.some(child => child.kind !== "element" && (child.kind !== "text" || !/^[\t\r\n ]*$/u.test(child.value))) || entry.children.some(child => child.kind === "element" && (child.namespaceUri !== c || !["idx", "delete"].includes(child.local)))) throw new TypeError("Legend entry contains unsupported markup.");
    const idx = only(entry, "idx"), lexical = simpleValue(idx), deletion = direct(entry, "delete");
    if (deletion.length > 1 || deletion.length && !["0", "1", "true", "false"].includes(simpleValue(deletion[0]!))) throw new TypeError("Legend entry deletion is ambiguous.");
    const position = Number(lexical);
    if (!/^\d+$/u.test(lexical) || !Number.isSafeInteger(position) || position < 0 || position >= series.length || positions.has(position)) throw new TypeError("Legend entry index is duplicated or outside the series collection.");
    positions.add(position); entries.push({ elementIndex: index.get(idx)!, namespaceUri: c!, position, lexical });
  }
  return { entries, physicalIds, workUnits: work };
}
