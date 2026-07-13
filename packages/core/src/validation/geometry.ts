import type { Box, PresentationSize } from "../types/geometry.js";

export function assertBox(box: Box, context: string): void {
  for (const [key, value] of Object.entries(box)) {
    if (!Number.isFinite(value)) {
      throw new Error(`${context} has a non-finite ${key} value.`);
    }
  }
  if (box.width < 0 || box.height < 0) {
    throw new Error(`${context} has a negative width or height.`);
  }
}

export function assertOptionalDimension(value: number | undefined, label: string): void {
  if (value !== undefined && (!Number.isFinite(value) || value < 0)) {
    throw new Error(`${label} must be a non-negative finite number.`);
  }
}

export function normalizeSize(size: Partial<PresentationSize> | undefined, defaults: PresentationSize): PresentationSize {
  const normalized = {
    width: size?.width ?? defaults.width,
    height: size?.height ?? defaults.height,
    unit: size?.unit ?? defaults.unit,
  };
  assertOptionalDimension(normalized.width, "Presentation width");
  assertOptionalDimension(normalized.height, "Presentation height");
  if (normalized.width === 0 || normalized.height === 0) {
    throw new Error("Presentation width and height must be greater than zero.");
  }
  return normalized;
}
