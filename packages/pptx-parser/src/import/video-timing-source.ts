import { attribute, children } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";

/** Reject video metadata that could change the timing represented by the movie clock. */
export function plainVideoTimingSource(nvPr: XmlElement | undefined, p: string, a: string, r: string): boolean {
  const attrs = (node: XmlElement) => node.attributes.filter(x => x.namespaceUri !== "http://www.w3.org/2000/xmlns/");
  if (!nvPr || attrs(nvPr).length || nvPr.children.some(x => (x.kind === "text" || x.kind === "cdata") && x.value.trim())) return false;
  const nodes = children(nvPr);
  const direct = nodes.filter(n => n.namespaceUri === a && n.local === "videoFile");
  const lists = nodes.filter(n => n.namespaceUri === p && n.local === "extLst");
  if (direct.length !== 1 || lists.length > 1 || nodes.length !== 1 + lists.length) return false;
  const file = direct[0]!;
  if (attrs(file).length !== 1 || attrs(file)[0]!.namespaceUri !== r || attrs(file)[0]!.local !== "link" || children(file).length || file.children.some(x => (x.kind === "text" || x.kind === "cdata") && x.value.trim())) return false;
  if (!lists.length) return true;
  const list = lists[0]!, extensions = children(list);
  if (attrs(list).length || extensions.length !== 1 || list.children.some(x => (x.kind === "text" || x.kind === "cdata") && x.value.trim())) return false;
  const ext = extensions[0]!, media = children(ext);
  if (ext.namespaceUri !== p || ext.local !== "ext" || attrs(ext).length !== 1 || attribute(ext, "uri")?.toUpperCase() !== "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" || media.length !== 1 || ext.children.some(x => (x.kind === "text" || x.kind === "cdata") && x.value.trim())) return false;
  const item = media[0]!;
  return item.namespaceUri === "http://schemas.microsoft.com/office/powerpoint/2010/main" && item.local === "media" && attrs(item).length === 1 && attrs(item)[0]!.namespaceUri === r && attrs(item)[0]!.local === "embed" && children(item).length === 0 && !item.children.some(x => (x.kind === "text" || x.kind === "cdata") && x.value.trim());
}
