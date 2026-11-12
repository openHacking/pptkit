import type { Box, NormalizedShapeElement } from "@pptkit/core";

/** Resolve the local, untransformed text rectangle for a normalized shape. */
export function resolveShapeTextBox(element: Pick<NormalizedShapeElement, "shape" | "slant" | "box">): Box {
  const { x, y, width, height } = element.box;
  if (element.shape !== "parallelogram") return { ...element.box };
  // The preset uses one dimensionless inset on both axes; slant is relative
  // to the shorter side. A zero-width shape uses its degenerate frame.
  const fraction = width === 0 ? 0 : (1 + 5 * (element.slant ?? 0.25) * (Math.min(width, height) / width)) / 12;
  const box = { x: x + width * fraction, y: y + height * fraction, width: width * (1 - 2 * fraction), height: height * (1 - 2 * fraction) };
  if (![box.x, box.y, box.width, box.height].every(Number.isFinite)) throw new RangeError("Shape text frame cannot be represented as finite coordinates.");
  return box;
}
