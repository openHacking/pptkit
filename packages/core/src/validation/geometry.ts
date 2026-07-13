import type { Box, Point, PresentationSize, Size } from "../types/geometry.js";

export function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

export function isValidPoint(point: Point): boolean {
  return isFiniteNumber(point.x) && isFiniteNumber(point.y);
}

export function isValidSize(size: Size): boolean {
  return isFiniteNumber(size.width) && isFiniteNumber(size.height) && size.width >= 0 && size.height >= 0;
}

export function isValidBox(box: Box): boolean {
  return isValidPoint(box) && isValidSize(box);
}

export function assertBox(box: Box, context: string): void {
  if (!isValidBox(box)) throw new Error(`${context} must contain finite coordinates and non-negative dimensions.`);
}

export function assertOptionalDimension(value: number | undefined, label: string): void {
  if (value !== undefined && (!isFiniteNumber(value) || value < 0)) {
    throw new Error(`${label} must be a non-negative finite number.`);
  }
}

export function normalizeSize(size: Partial<PresentationSize> | undefined, defaults: PresentationSize): PresentationSize {
  const normalized: PresentationSize = {
    width: size?.width ?? defaults.width,
    height: size?.height ?? defaults.height,
    unit: size?.unit ?? defaults.unit,
  };
  if (!isFiniteNumber(normalized.width) || !isFiniteNumber(normalized.height) || normalized.width <= 0 || normalized.height <= 0) {
    throw new Error("Presentation width and height must be positive finite numbers.");
  }
  return normalized;
}
