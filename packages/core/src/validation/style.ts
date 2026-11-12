import type { BackgroundPaintInput, ColorValue, PaintInput } from "../types/style.js";
import { isFiniteNumber } from "./geometry.js";
import { TEXT_FONT_SCRIPT_CODES } from "../text/font-scripts.js";
const colorRoles = new Set(["background1", "background2", "text1", "text2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink", "followedHyperlink"]);

export function isValidFontFamily(value: unknown): boolean {
  if (typeof value === "string") return value.trim().length > 0 && !/[\u0000-\u001F\u007F]/u.test(value);
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  if (![Object.prototype, null].includes(Object.getPrototypeOf(value))) return false;
  return Object.keys(value).length === 1 && "theme" in value && (value.theme === "heading" || value.theme === "body");
}

export function scriptFontFamilyErrors(value: unknown): { key?: string; message: string }[] {
  if (value === undefined || value === null) return [];
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) return [{ message: "Expected a plain script-font map or null reset." }];
  const known = new Set<string>(TEXT_FONT_SCRIPT_CODES), errors: { key?: string; message: string }[] = [];
  for (const [key, family] of Object.entries(value)) {
    if (!known.has(key)) errors.push({ key, message: "Unsupported font script code." });
    else if (family !== undefined && family !== null && !isValidFontFamily(family)) errors.push({ key, message: "Expected a nonempty font family, theme role or null reset." });
  }
  return errors;
}

export function isValidColor(value: ColorValue): boolean {
  return typeof value === "string" ? /^#?[0-9a-f]{6}$/i.test(value) : Boolean(value && typeof value === "object" && colorRoles.has(value.theme));
}

export function isValidOpacity(value: unknown): value is number {
  return isFiniteNumber(value) && value >= 0 && value <= 1;
}

export function isValidPaint(paint: PaintInput | undefined): boolean {
  if (paint === undefined) return true;
  if (!paint || typeof paint !== "object" || Array.isArray(paint)) return false;
  if (paint.type === "none") return true;
  if (paint.type !== "solid") return false;
  return isValidColor(paint.color) && (paint.opacity === undefined || isValidOpacity(paint.opacity));
}

export function isValidShapeFill(paint: import("../types/style.js").ShapeFillPaintInput | undefined): boolean {
  if (paint?.type !== "pattern") return isValidBackground(paint);
  return ["horizontal", "vertical", "cross", "downwardDiagonal", "upwardDiagonal", "diagonalCross", "percent5", "percent10", "percent20", "percent25", "percent30", "percent40", "percent50", "percent60", "percent70", "percent75", "percent80", "percent90"].includes(paint.pattern) && isValidColor(paint.foreground) && isValidColor(paint.background)
    && (paint.foregroundOpacity === undefined || isValidOpacity(paint.foregroundOpacity)) && (paint.backgroundOpacity === undefined || isValidOpacity(paint.backgroundOpacity));
}

export function isValidBackground(paint: BackgroundPaintInput | undefined): boolean {
  if (paint?.type !== "linearGradient") return isValidPaint(paint as PaintInput | undefined);
  if (!isFiniteNumber(paint.angle) || paint.angle < 0 || paint.angle >= 360 || (paint.scaled !== undefined && typeof paint.scaled !== "boolean") || !Array.isArray(paint.stops) || paint.stops.length < 2 || paint.stops.length > 256) return false;
  let previous = -1;
  for (const stop of paint.stops) {
    if (!stop || typeof stop !== "object" || !isFiniteNumber(stop.position) || stop.position < 0 || stop.position > 1 || stop.position < previous || !isValidColor(stop.color) || (stop.opacity !== undefined && !isValidOpacity(stop.opacity))) return false;
    previous = stop.position;
  }
  return true;
}
