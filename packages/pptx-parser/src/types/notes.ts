import type { OpcInspectionOptions } from "./opc.js";
import type { PresentationDialect } from "./presentation.js";
import type { PptxSourceLocation } from "./source.js";
import type { PptxSize, PptxStructureFinding } from "./structure.js";
import type { XmlExpandedName } from "./xml.js";

export interface PptxNotesInspectionLimits { readonly maxWorkUnits: number; readonly maxRecords: number; readonly maxFindings: number; readonly maxTextCharacters: number }
export interface PptxNotesInspectionOptions extends OpcInspectionOptions { readonly notesLimits?: Partial<PptxNotesInspectionLimits> }
export type PptxNotesView = "effective" | "raw-only";
/** Literal source attributes, without schema defaults or value evaluation. */
export interface PptxNotesNodeInfo {
  readonly source: PptxSourceLocation;
  readonly name: XmlExpandedName;
  readonly view: PptxNotesView;
  readonly attributes: readonly (XmlExpandedName & { readonly value: string })[];
}
export type PptxNotesRole = "group-root" | "body" | "slide-image" | "header" | "footer" | "date" | "slide-number" | "other-placeholder" | "artwork";
export interface PptxNotesDrawingInfo extends PptxNotesNodeInfo {
  readonly kind: "spTree" | "sp" | "grpSp" | "pic" | "cxnSp" | "graphicFrame" | "uninterpreted";
  readonly role: PptxNotesRole;
  /** True only for direct effective spTree children; raw-only records are false. */
  readonly direct: boolean;
  readonly parentDrawingElementIndex?: number;
  readonly rawId?: string;
  readonly drawingName?: string;
  readonly nonVisual?: PptxNotesNodeInfo;
  readonly textBodyCount: number;
  readonly placeholder?: PptxNotesNodeInfo & { readonly type?: string; readonly index?: string };
}
export interface PptxNotesGeometryInfo extends PptxNotesNodeInfo {
  readonly ownerDrawingElementIndex: number;
  readonly kind: "transform" | "preset" | "custom";
  /** Direct components only; custom path interpretation remains source-only. */
  readonly components: readonly PptxNotesNodeInfo[];
}
export interface PptxNotesTextInfo { readonly source: PptxSourceLocation; readonly view: PptxNotesView; readonly ownerDrawingElementIndex?: number; readonly paragraphElementIndex?: number; readonly value: string }
export interface PptxNotesFieldInfo extends PptxNotesNodeInfo {
  readonly ownerDrawingElementIndex?: number;
  /** Raw element index of the enclosing DrawingML paragraph, when present. */
  readonly paragraphElementIndex?: number;
  readonly id?: string;
  readonly type?: string;
  /** Stored cache segments, not a current date, page number or evaluated value. */
  readonly cachedTexts: readonly PptxNotesTextInfo[];
}
export interface PptxNotesPartInfo {
  readonly part: string;
  readonly source: PptxSourceLocation;
  readonly kind: "notes" | "notesMaster";
  readonly attributes: readonly (XmlExpandedName & { readonly value: string })[];
  readonly drawings: readonly PptxNotesDrawingInfo[];
  readonly geometries: readonly PptxNotesGeometryInfo[];
  readonly texts: readonly PptxNotesTextInfo[];
  readonly fields: readonly PptxNotesFieldInfo[];
  readonly headerFooter: readonly PptxNotesNodeInfo[];
  readonly extensions: readonly (PptxNotesNodeInfo & { readonly children: readonly XmlExpandedName[] })[];
}
export interface PptxNotesPlaceholderMatch {
  readonly source: PptxSourceLocation;
  readonly status: "matched-direct-type" | "missing-master" | "missing-type" | "missing-candidate" | "ambiguous" | "grouped";
  /** Exact explicit-type candidates, never selected by idx or name. */
  readonly candidates: readonly PptxSourceLocation[];
}
export interface PptxNotesPageInfo {
  readonly slideId: number;
  readonly slidePart: string;
  readonly notesPart?: string;
  readonly masterPart?: string;
  readonly masterThemePart?: string;
  readonly overridePart?: string;
  readonly body: { readonly status: "absent" | "unique-direct" | "ambiguous-direct" | "missing-local-text" | "ambiguous-local-text"; readonly candidates: readonly PptxSourceLocation[] };
  readonly placeholderMatches: readonly PptxNotesPlaceholderMatch[];
}
export interface PptxNotesFinding { readonly code: "missing-drawing-id" | "uninterpreted-drawing-id" | "duplicate-drawing-id" | "unknown-drawing" | "unknown-placeholder"; readonly source: PptxSourceLocation; readonly message: string }
export interface PptxNotesInspection {
  readonly schemaVersion: "pptkit-notes-source-inventory-v1";
  readonly dialect: PresentationDialect;
  readonly notesSize: PptxSize;
  readonly pages: readonly PptxNotesPageInfo[];
  /** Actual distinct parts, including listed or referenced masters, scanned once. */
  readonly parts: readonly PptxNotesPartInfo[];
  readonly findings: readonly PptxNotesFinding[];
  readonly structureFindings: readonly PptxStructureFinding[];
  readonly workUnits: number;
  readonly recordCount: number;
  readonly textCharacters: number;
  readonly geometryPolicy: "unresolved-source-only";
  readonly fieldEvaluation: "not-performed";
  readonly authoredMapping: "not-performed";
}
