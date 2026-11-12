import type { PptxInspectionOptions } from "../types/structure.js";
import { planPropertyCreation } from "./plan-property-creation.js";

/** Plan creation of a custom-properties part against the complete source graph. */
export function planCustomPropertyCreation(input: Uint8Array, dialect: "strict" | "transitional", options: PptxInspectionOptions, maxOutputBytes: number, maxFindings: number): Promise<{ part: string; relationshipId: string; relationships: Uint8Array; contentTypes: Uint8Array }> {
  return planPropertyCreation(input, "custom", dialect, options, maxOutputBytes, maxFindings);
}
