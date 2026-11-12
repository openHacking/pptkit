import type { PptxDiagramInspectionLimits } from "../types/structure.js";

export function captureDiagramTextComparison(value: unknown): "direct-paragraphs" | undefined {
  if (value !== undefined && value !== "direct-paragraphs") throw new TypeError("Unknown diagram text comparison profile.");
  return value;
}

export const DEFAULT_DIAGRAM_LIMITS: Readonly<PptxDiagramInspectionLimits> = Object.freeze({ maxWorkUnits: 1000000, maxFindings: 10000 });
export function captureDiagramLimits(value?: Partial<PptxDiagramInspectionLimits>): PptxDiagramInspectionLimits {
  if (value !== undefined && (!value || typeof value !== "object" || Array.isArray(value))) throw new TypeError("Invalid diagram inspection limits.");
  if (value && Object.keys(value).some(key => !Object.hasOwn(DEFAULT_DIAGRAM_LIMITS, key))) throw new TypeError("Unknown diagram inspection limit.");
  const limits = { ...DEFAULT_DIAGRAM_LIMITS, ...value };
  for (const [key, number] of Object.entries(limits)) if (!Number.isSafeInteger(number) || number <= 0) throw new RangeError(`Invalid diagram inspection limit ${key}.`);
  return limits;
}
