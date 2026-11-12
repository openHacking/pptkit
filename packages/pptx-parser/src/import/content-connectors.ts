import type { Box, ConnectorAnchor, ConnectorElementInput, ConnectorEndpointInput, Point, PresentationElementInput, TransformInput } from "@pptkit/core";
import type { XmlElement } from "../xml/read-xml.js";
import { children } from "../presentation/structure-xml.js";
import type { ContentReader } from "./content-reader.js";

export interface ContentConnection { side: "start" | "end"; node: XmlElement; sourceId: number; site: number }
export interface ContentDrawing { input: PresentationElementInput; scope: string }
type Fields = Pick<ConnectorElementInput, "type" | "start" | "end" | "route" | "style" | "transform" | "box">;

export function contentConnector(node: XmlElement, properties: XmlElement, geometry: { box: Box; transform: TransformInput }, r: ContentReader): { fields: Fields; connections: ContentConnection[] } | { reason: string } {
  const preset = r.one(properties, r.a, "prstGeom");
  const kind = preset ? r.attr(preset, "prst") : undefined;
  if (!preset || (kind !== "straightConnector1" && kind !== "line" && kind !== "bentConnector2" && kind !== "bentConnector3") || r.one(properties, r.a, "custGeom")) return { reason: "Only straight, unadjusted bentConnector2 and bounded bentConnector3 connectors are mapped; other routes remain source-only." };
  const adjustments = r.one(preset, r.a, "avLst");
  if (kind === "bentConnector2" && adjustments && children(adjustments).length) return { reason: "bentConnector2 with nonempty adjustments remains source-only." };
  let adjustment = 50000;
  if (kind === "bentConnector3" && adjustments && children(adjustments).length) {
    const guides = children(adjustments), guide = guides[0];
    if (guides.length !== 1 || !guide || guide.namespaceUri !== r.a || guide.local !== "gd" || children(guide).length || guide.attributes.some(a => a.namespaceUri !== "http://www.w3.org/2000/xmlns/" && (a.namespaceUri !== "" || !["name", "fmla"].includes(a.local))) || guide.children.some(child => (child.kind === "text" || child.kind === "cdata") && !/^[\t\r\n ]*$/.test(child.value))) return { reason: "bentConnector3 has unsupported adjustment content." };
    const name = r.attr(guide, "name"), formula = r.attr(guide, "fmla"), match = formula?.trim().match(/^val[\t\r\n ]+([+-]?\d+)$/);
    const value = match ? Number(match[1]) : NaN;
    if (name !== "adj1" || !Number.isSafeInteger(value) || value < 0 || value > 100000) return { reason: "bentConnector3 requires one bounded literal adj1 adjustment." };
    adjustment = value;
  }
  if (adjustments) r.known(adjustments);
  const nv = r.one(node, r.p, "nvCxnSpPr")!, connectionsNode = r.one(nv, r.p, "cNvCxnSpPr");
  const connections: ContentConnection[] = [];
  if (connectionsNode) {
    r.known(connectionsNode);
    for (const [name, side] of [["stCxn", "start"], ["endCxn", "end"]] as const) {
      const connection = r.one(connectionsNode, r.a, name);
      if (connection) connections.push({ node: connection, side, sourceId: r.number(connection, "id", 0, 4294967295), site: r.number(connection, "idx", 0, 4294967295) });
    }
  }
  const { box, transform } = geometry;
  const point = (x: number, y: number): Point => {
    const cx = box.x + box.width / 2, cy = box.y + box.height / 2, angle = (transform.rotation ?? 0) * Math.PI / 180;
    const dx = (x - cx) * (transform.flipH ? -1 : 1), dy = (y - cy) * (transform.flipV ? -1 : 1);
    return { x: cx + dx * Math.cos(angle) - dy * Math.sin(angle), y: cy + dx * Math.sin(angle) + dy * Math.cos(angle) };
  };
  const start = point(box.x, box.y), end = point(box.x + box.width, box.y + box.height);
  const route = kind === "bentConnector2" ? [point(box.x + box.width, box.y)] : kind === "bentConnector3" ? [point(box.x + box.width * adjustment / 100000, box.y), point(box.x + box.width * adjustment / 100000, box.y + box.height)] : [];
  const points = [start, ...route, end], x = Math.min(...points.map(p => p.x)), y = Math.min(...points.map(p => p.y));
  // Core/Layout resolves endpoint coordinates. Keeping the source transform too
  // would rotate/flip the same line again in preview and export.
  return { fields: { type: "connector", start, end, ...(route.length ? { route } : {}), box: { x, y, width: Math.max(...points.map(p => p.x)) - x, height: Math.max(...points.map(p => p.y)) - y }, transform: { rotation: 0, flipH: false, flipV: false }, style: r.stroke(properties) ?? { paint: { type: "none" } } }, connections };
}

/** Resolve after the whole slide forest is mapped, so forward references work. */
export function connectorAttachment(connection: ContentConnection, scope: string, drawings: ReadonlyMap<number, ContentDrawing>): { endpoint: ConnectorEndpointInput } | { reason: string } {
  const target = drawings.get(connection.sourceId);
  if (!target) return { reason: `Connector target ${connection.sourceId} is missing or not projected; its cached endpoint remains a free point.` };
  if (target.scope !== scope) return { reason: `Connector target ${connection.sourceId} is outside its immediate group scope; its cached endpoint remains a free point.` };
  const kind = target.input.type === "text" ? "rect" : target.input.type === "shape" ? target.input.shape : undefined;
  if (kind !== "rect" && kind !== "roundRect" && kind !== "ellipse") return { reason: `Connector target ${connection.sourceId} has no mapped native site table; its cached endpoint remains a free point.` };
  const site = kind === "ellipse" ? connection.site / 2 : connection.site;
  const anchor = (["top", "left", "bottom", "right"] as const)[site] as ConnectorAnchor | undefined;
  if (!Number.isInteger(site) || anchor === undefined) return { reason: `Connector site ${connection.site} on target ${connection.sourceId} has no mapped cardinal anchor; its cached endpoint remains a free point.` };
  return { endpoint: { elementId: target.input.id!, anchor } };
}
