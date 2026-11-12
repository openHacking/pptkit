import type { Box, ConnectorElementInput, TransformInput } from "@pptkit/core";
import type { XmlElement } from "../xml/read-xml.js";
import { attribute, children } from "../presentation/structure-xml.js";
import type { ContentReader } from "./content-reader.js";

/** Maps only the bounded, literal custom geometry emitted for round connectors. */
export function contentRoundPath(node: XmlElement, properties: XmlElement, geometry: { box: Box; transform: TransformInput }, r: ContentReader): { fields: Pick<ConnectorElementInput, "type" | "start" | "end" | "route" | "pathMode" | "style" | "box" | "transform"> } | { reason: string } {
  const reject = () => ({ reason: "Custom geometry is outside the bounded round linear-path profile." });
  const only = (element: XmlElement, names: readonly string[]) => element.attributes.every(attr => attr.namespaceUri === "http://www.w3.org/2000/xmlns/" || (attr.namespaceUri === "" && names.includes(attr.local))) && !element.children.some(child => (child.kind === "text" || child.kind === "cdata") && !/^[\t\r\n ]*$/.test(child.value));
  const mark = (element: XmlElement) => { r.known(element); for (const attr of element.attributes) if (attr.namespaceUri !== "http://www.w3.org/2000/xmlns/") r.attr(element, attr.local, attr.namespaceUri); };
  if (children(node).some(child => child.namespaceUri === r.p && child.local === "txBody")) return reject();
  if (children(properties).some(child => child.namespaceUri !== r.a || !["xfrm", "custGeom", "ln"].includes(child.local))) return reject();
  const custom = r.one(properties, r.a, "custGeom"), line = r.one(properties, r.a, "ln");
  if (!custom || !line || !only(custom, []) || !only(line, ["w", "cap"]) || attribute(line, "cap") !== "rnd") return reject();
  const parts = children(custom), allowed = ["avLst", "gdLst", "ahLst", "cxnLst", "rect", "pathLst"];
  if (parts.length !== allowed.length || parts.some((part, i) => part.namespaceUri !== r.a || part.local !== allowed[i])) return reject();
  if (parts.slice(0, 4).some(part => !only(part, []) || children(part).length)) return reject();
  const rect = parts[4]!;
  if (!only(rect, ["l", "t", "r", "b"]) || children(rect).length || ["l", "t", "r", "b"].some(key => attribute(rect, key) !== key)) return reject();
  const list = parts[5]!, paths = children(list);
  if (!only(list, []) || paths.length !== 1 || paths[0]!.namespaceUri !== r.a || paths[0]!.local !== "path") return reject();
  const path = paths[0]!;
  if (!only(path, ["w", "h", "fill"]) || attribute(path, "fill") !== "none") return reject();
  const unsigned = (value: string | undefined) => value && /^\d{1,14}$/.test(value) && Number.isSafeInteger(Number(value)) ? Number(value) : undefined;
  const width = unsigned(attribute(path, "w")), height = unsigned(attribute(path, "h"));
  if (!width || !height) return reject();
  const commands = children(path); if (commands.length < 2 || commands.length > 257) return reject();
  const closed = commands[commands.length - 1]?.local === "close";
  const drawn = closed ? commands.slice(0, -1) : commands;
  if (drawn.length < (closed ? 3 : 2)) return reject();
  const points: { x: number; y: number }[] = [];
  for (let i = 0; i < commands.length; i++) {
    const command = commands[i]!;
    if (command.namespaceUri !== r.a || !only(command, []) || command.local !== (i === 0 ? "moveTo" : closed && i === commands.length - 1 ? "close" : "lnTo")) return reject();
    const pts = children(command);
    if (command.local === "close") { if (pts.length) return reject(); continue; }
    if (pts.length !== 1 || pts[0]!.namespaceUri !== r.a || pts[0]!.local !== "pt" || !only(pts[0]!, ["x", "y"]) || children(pts[0]!).length) return reject();
    const x = unsigned(attribute(pts[0]!, "x")), y = unsigned(attribute(pts[0]!, "y"));
    if (x === undefined || y === undefined || x > width || y > height) return reject();
    const px = geometry.box.x + geometry.box.width * x / width, py = geometry.box.y + geometry.box.height * y / height;
    const cx = geometry.box.x + geometry.box.width / 2, cy = geometry.box.y + geometry.box.height / 2;
    const dx = (px - cx) * (geometry.transform.flipH ? -1 : 1), dy = (py - cy) * (geometry.transform.flipV ? -1 : 1);
    const angle = (geometry.transform.rotation ?? 0) * Math.PI / 180;
    points.push({ x: cx + dx * Math.cos(angle) - dy * Math.sin(angle), y: cy + dx * Math.sin(angle) + dy * Math.cos(angle) });
  }
  const lineParts = children(line);
  if (lineParts.length !== 3 || lineParts.some((child, i) => child.namespaceUri !== r.a || child.local !== ["solidFill", "prstDash", "round"][i])) return reject();
  const [solid, dash, round] = lineParts, rgb = solid && children(solid);
  if (!solid || !dash || !round || !only(solid, []) || !rgb || rgb.length !== 1 || rgb[0]!.namespaceUri !== r.a || rgb[0]!.local !== "srgbClr" || !only(rgb[0]!, ["val"]) || !/^[0-9a-fA-F]{6}$/.test(attribute(rgb[0]!, "val") ?? "") || !only(dash, ["val"]) || attribute(dash, "val") !== "solid" || children(dash).length || !only(round, []) || children(round).length) return reject();
  const colorChildren = children(rgb[0]!);
  if (colorChildren.length > 1 || colorChildren.some(child => child.namespaceUri !== r.a || child.local !== "alpha" || !only(child, ["val"]) || children(child).length !== 0 || unsigned(attribute(child, "val")) === undefined || unsigned(attribute(child, "val"))! > 100000)) return reject();
  const style = r.stroke(properties); if (!style || !style.width || style.paint?.type !== "solid") return reject();
  if (closed) points.push({ ...points[0]! });
  for (const item of [custom, ...parts, path, ...commands, ...commands.flatMap(children), line, ...lineParts, ...rgb, ...colorChildren]) mark(item);
  const nv = r.one(node, r.p, "nvSpPr"), shape = nv && r.one(nv, r.p, "cNvSpPr"); if (shape) r.known(shape);
  return { fields: { type: "connector", start: points[0]!, end: points[points.length - 1]!, route: points.slice(1, -1), pathMode: closed ? "closed-round" : "open-round", style, box: geometry.box, transform: { rotation: 0, flipH: false, flipV: false } } };
}
