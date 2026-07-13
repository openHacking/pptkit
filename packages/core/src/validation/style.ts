import type { ColorValue, PaintInput } from "../types/style.js";
import { isFiniteNumber } from "./geometry.js";

export function isValidColor(value: ColorValue): boolean {
  return typeof value === "string" ? /^#?[0-9a-f]{6}$/i.test(value) : typeof value.theme === "string";
}

export function isValidOpacity(value: unknown): value is number {
  return isFiniteNumber(value) && value >= 0 && value <= 1;
}

export function isValidPaint(paint: PaintInput | undefined): boolean {
  if (paint === undefined || paint.type === "none") return true;
  return isValidColor(paint.color) && (paint.opacity === undefined || isValidOpacity(paint.opacity));
}
