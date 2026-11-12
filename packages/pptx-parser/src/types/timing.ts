import type { OpcInspectionOptions } from "./opc.js";
import type { PptxSourceLocation } from "./source.js";
import type { PresentationDialect } from "./presentation.js";
export interface PptxTimingInspectionLimits { readonly maxWorkUnits: number; readonly maxRecords: number; readonly maxFindings: number; readonly maxTextCharacters: number }
export interface PptxTimingInspectionOptions extends OpcInspectionOptions { readonly timingLimits?: Partial<PptxTimingInspectionLimits> }
export type PptxTimingView = "effective" | "raw-only";
export interface PptxTimingAttribute { readonly namespaceUri: string; readonly local: string; readonly value: string }
export interface PptxTimingReference {
  readonly rawId?: string;
  readonly status: "located" | "missing" | "ambiguous" | "source-only";
  readonly candidates: readonly { readonly source: PptxSourceLocation; readonly match: "exact" | "numeric-alias" }[];
  readonly reason?: string;
}
export interface PptxTimingNodeInfo {
  readonly source: PptxSourceLocation;
  readonly namespaceUri: string;
  readonly local: string;
  readonly view: PptxTimingView;
  readonly role: "timing" | "time-node" | "behavior" | "condition" | "drawing-target" | "time-reference" | "container" | "unknown";
  readonly timingRootElementIndex: number;
  readonly parentElementIndex?: number;
  /** Exact nearest raw cTn ancestor, not an inferred event/activation parent. */
  readonly parentTimeNodeElementIndex?: number;
  readonly attributes: readonly PptxTimingAttribute[];
  /** Direct decoded character content only, not descendant text or XML bytes. */
  readonly text?: string;
  readonly rawId?: string;
  readonly drawingReference?: PptxTimingReference;
  readonly timeReference?: PptxTimingReference;
  readonly targetKind?: "whole-object" | "subobject" | "uninterpreted";
}
export interface PptxTimingDrawingInfo {
  readonly source: PptxSourceLocation;
  readonly view: PptxTimingView;
  readonly ownerKind: string;
  readonly rawId?: string;
  readonly name?: string;
}
export interface PptxTimingFinding {
  readonly code: "duplicate-time-id" | "duplicate-drawing-id" | "uninterpreted-identity" | "missing-identity" | "missing-reference" | "ambiguous-reference" | "raw-only-timing" | "unknown-timing-node" | "subobject-target" | "multiple-effective-timing";
  readonly source: PptxSourceLocation;
  readonly message: string;
}
export interface PptxSlideTimingInventory {
  readonly slideId: number;
  readonly part: string;
  readonly nodes: readonly PptxTimingNodeInfo[];
  readonly drawings: readonly PptxTimingDrawingInfo[];
}
export interface PptxTimingInspection {
  readonly schemaVersion: "pptkit-native-timing-inventory-v1";
  readonly dialect: PresentationDialect;
  readonly slides: readonly PptxSlideTimingInventory[];
  readonly findings: readonly PptxTimingFinding[];
  readonly workUnits: number;
  readonly recordCount: number;
  readonly textCharacters: number;
  readonly nativePlayback: "not-performed";
  readonly authoredMapping: "not-performed";
}
