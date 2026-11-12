import type { Box, Point, ShapeKind } from "@pptkit/core";

/** Points used by the SVG polygon presets, in the element's parent coordinate system. */
export function shapePolygonPoints(box: Box, shape: ShapeKind, slant?: number): Point[] {
  const { x, y, width, height } = box, inset = Math.min(width, height) * (slant ?? 0.25);
  const pairs = shape === "parallelogram"
    ? [[x, y + height], [x + inset, y], [x + width, y], [x + width - inset, y + height]]
    : shape === "triangle"
    ? [[x + width / 2, y], [x + width, y + height], [x, y + height]]
    : shape === "diamond"
    ? [[x + width / 2, y], [x + width, y + height / 2], [x + width / 2, y + height], [x, y + height / 2]]
    : shape === "arrow"
    ? [[x, y + height * 0.25], [x + width * 0.62, y + height * 0.25], [x + width * 0.62, y], [x + width, y + height / 2], [x + width * 0.62, y + height], [x + width * 0.62, y + height * 0.75], [x, y + height * 0.75]]
    : [[x, y], [x + width * 0.65, y], [x + width, y + height / 2], [x + width * 0.65, y + height], [x, y + height], [x + width * 0.35, y + height / 2]];
  return pairs.map(([px, py]) => ({ x: px!, y: py! }));
}
