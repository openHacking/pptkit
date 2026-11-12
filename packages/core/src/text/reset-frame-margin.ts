import type { Insets } from "../types/geometry.js";
import type { TextFrameStyleInput } from "../types/style.js";

/** Remove one declaration while retaining the other sides of a scalar shorthand. */
export function resetFrameMarginSide(frame: TextFrameStyleInput | undefined, side: keyof Insets): void {
  if (frame?.margin === undefined) return;
  const margin = typeof frame.margin === "number"
    ? { top: frame.margin, right: frame.margin, bottom: frame.margin, left: frame.margin }
    : { ...frame.margin };
  delete margin[side];
  frame.margin = margin;
}
