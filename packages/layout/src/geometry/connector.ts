import type { Box, ConnectorEndpointInput, NormalizedConnectorElement, NormalizedTransform, Point } from "@pptkit/core";

function endpoint(point: ConnectorEndpointInput, boxes: ReadonlyMap<string, Box>, transforms?: ReadonlyMap<string, NormalizedTransform>): Point {
  if (!("elementId" in point)) return { ...point };
  const box = boxes.get(point.elementId);
  if (!box) throw new Error(`Connector target "${point.elementId}" is outside the current coordinate scope.`);
  const local = (() => { switch (point.anchor) {
    case "top": return { x: box.x + box.width / 2, y: box.y };
    case "right": return { x: box.x + box.width, y: box.y + box.height / 2 };
    case "bottom": return { x: box.x + box.width / 2, y: box.y + box.height };
    case "left": return { x: box.x, y: box.y + box.height / 2 };
    default: return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
  } })();
  const transform = transforms?.get(point.elementId);
  if (!transform || (!transform.rotation && !transform.flipH && !transform.flipV)) return local;
  const cx = box.x + box.width / 2, cy = box.y + box.height / 2;
  const x = (local.x - cx) * (transform.flipH ? -1 : 1), y = (local.y - cy) * (transform.flipV ? -1 : 1);
  const angle = transform.rotation * Math.PI / 180;
  return { x: cx + x * Math.cos(angle) - y * Math.sin(angle), y: cy + x * Math.sin(angle) + y * Math.cos(angle) };
}

export function connectorGeometry(element: NormalizedConnectorElement, boxes: ReadonlyMap<string, Box>, transforms?: ReadonlyMap<string, NormalizedTransform>): { start: Point; end: Point; route: Point[]; box: Box } {
  const start = endpoint(element.start, boxes, transforms);
  const end = endpoint(element.end, boxes, transforms);
  const route = element.route.map((point) => ({ ...point }));
  let x = Math.min(start.x, end.x), right = Math.max(start.x, end.x);
  let y = Math.min(start.y, end.y), bottom = Math.max(start.y, end.y);
  for (const point of route) {
    x = Math.min(x, point.x); right = Math.max(right, point.x);
    y = Math.min(y, point.y); bottom = Math.max(bottom, point.y);
  }
  return { start, end, route, box: { x, y, width: right - x, height: bottom - y } };
}
