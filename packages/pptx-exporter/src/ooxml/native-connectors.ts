import type { NormalizedTransform } from "@pptkit/core";
import type { LayoutConnectorElement, LayoutElement } from "@pptkit/layout";

export type NativeConnectorResult = { supported: true; connections: string; transform: NormalizedTransform; preset: "straightConnector1" | "bentConnector2" | "bentConnector3"; adjustment?: number } | { supported: false; reason: string };
interface DrawingObject { id: number; scope: string | undefined; element: LayoutElement }
const untransformed = (value: NormalizedTransform) => value.rotation === 0 && !value.flipH && !value.flipV;

/** Part-local IDs are allocated before serialization to resolve forward connections. */
export function nativeDrawingContext(elements: readonly LayoutElement[]) {
  const objects = new Map<string, DrawingObject>(); let next = 2;
  const stack: { elements: readonly LayoutElement[]; scope: string | undefined; index: number }[] = [{ elements, scope: undefined, index: 0 }];
  while (stack.length) {
    const frame = stack[stack.length - 1]!;
    if (frame.index === frame.elements.length) { stack.pop(); continue; }
    const element = frame.elements[frame.index++]!;
    if (element.hidden) continue;
    objects.set(element.id, { id: next++, scope: frame.scope, element });
    if (element.type === "group") stack.push({ elements: element.children, scope: element.id, index: 0 });
  }
  return {
    nextObjectId: () => next++,
    objectId: (id: string) => {
      const object = objects.get(id);
      if (!object) throw new Error(`Missing drawing identity for ${id}.`);
      return object.id;
    },
    connector: (connector: LayoutConnectorElement): NativeConnectorResult => {
      const elbow = connector.route.length === 1 && connector.start.x !== connector.end.x && connector.start.y !== connector.end.y && connector.route[0]!.x === connector.end.x && connector.route[0]!.y === connector.start.y;
      const [first, second] = connector.route, dx = connector.end.x - connector.start.x;
      let adjustment: number | undefined;
      if (connector.route.length === 2 && dx !== 0 && connector.end.y !== connector.start.y && first!.x === second!.x && first!.y === connector.start.y && second!.y === connector.end.y) {
        const ratio = (first!.x - connector.start.x) / dx;
        if (ratio > 0 && ratio < 1) {
          const value = Math.round(ratio * 100000), width = Math.round(Math.abs(dx) * 12700), left = Math.round(Math.min(connector.start.x, connector.end.x) * 12700);
          const projected = left + (dx < 0 ? width * (1 - value / 100000) : width * value / 100000);
          if (value > 0 && value < 100000 && Math.abs(projected - Math.round(first!.x * 12700)) <= 0.5) adjustment = value;
        }
      }
      if (connector.route.length && !elbow && adjustment === undefined) return { supported: false, reason: "This explicit polyline route cannot be represented by the supported native connector presets without changing its geometry." };
      if (!untransformed(connector.transform)) return { supported: false, reason: "Connector-local rotation/flips are not supported by the native straight profile." };
      const owner = objects.get(connector.id); let connections = "";
      for (const side of ["start", "end"] as const) {
        const endpoint = connector.connections?.[side];
        if (!endpoint) continue;
        const target = objects.get(endpoint.elementId);
        if (!target || target.scope !== owner?.scope) return { supported: false, reason: "A target is hidden, inherited from another part, or outside the connector's immediate group scope." };
        const shape = target.element.type === "text" ? "rect" : target.element.type === "shape" ? target.element.shape : undefined;
        if (shape !== "rect" && shape !== "roundRect" && shape !== "ellipse") return { supported: false, reason: "Native attachment currently supports text rectangles, rect, roundRect and ellipse targets." };
        const site = endpoint.anchor === "top" ? 0 : endpoint.anchor === "left" ? 1 : endpoint.anchor === "bottom" ? 2 : endpoint.anchor === "right" ? 3 : undefined;
        if (site === undefined) return { supported: false, reason: "Center anchors have no matching site in the supported native preset connection tables." };
        // ECMA-376 presetShapeDefinitions: ellipse has intermediate diagonal sites.
        const index = site * (shape === "ellipse" ? 2 : 1);
        connections += `<a:${side === "start" ? "stCxn" : "endCxn"} id="${target.id}" idx="${index}"/>`;
      }
      return { supported: true, connections, preset: elbow ? "bentConnector2" : adjustment === undefined ? "straightConnector1" : "bentConnector3", ...(adjustment === undefined ? {} : { adjustment }), transform: { rotation: 0, flipH: connector.end.x < connector.start.x, flipV: connector.end.y < connector.start.y } };
    },
  };
}
