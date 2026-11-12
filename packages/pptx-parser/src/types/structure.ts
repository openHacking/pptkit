import type { MarkupCompatibilityReport } from "./compatibility.js";
import type { PresentationPackageInspection } from "./presentation.js";
import type { PptxDocumentMetadata } from "./metadata.js";
import type { XmlExpandedName } from "./xml.js";
import type { PptxDiagramInfo, PptxDiagramResourceInventory } from "./smartart.js";
import type { OpcInspectionOptions } from "./opc.js";

/** Per-inspection graph-processing budgets, shared across distinct data parts. */
export interface PptxDiagramInspectionLimits { readonly maxWorkUnits: number; readonly maxFindings: number }
export interface PptxInspectionOptions extends OpcInspectionOptions {
  readonly diagramLimits?: Partial<PptxDiagramInspectionLimits>;
  /** Explicit character/paragraph observations only; never cache freshness or ownership proof. */
  readonly diagramTextComparison?: "direct-paragraphs";
}

export interface PptxSize { readonly widthEmu: number; readonly heightEmu: number }
export interface PptxPartReference { readonly relationshipId: string; readonly part: string }
export interface PptxTransitionInfo {
  readonly status: "mapped" | "source-only";
  readonly speed: "fast" | "med" | "slow";
  readonly type?: "fade" | "push" | "wipe";
  readonly direction?: "left" | "right" | "up" | "down";
  readonly durationMs?: number;
  readonly onClick: boolean;
  readonly afterMs?: number;
  readonly reason?: string;
}
export interface PptxSlideInfo extends PptxPartReference {
  readonly id: number;
  readonly hidden: boolean;
  readonly transition?: PptxTransitionInfo;
  /** Direct/effective slide diagrams; no master inheritance, cache interpretation or node editing. */
  readonly diagrams?: readonly PptxDiagramInfo[];
  /** cSld name, not inferred title text. */
  readonly name?: string;
  readonly layoutPart?: string;
  readonly notesPart?: string;
}
export interface PptxMasterInfo extends PptxPartReference {
  readonly id: number;
  readonly layouts: readonly (PptxPartReference & { readonly id: number })[];
  readonly themePart?: string;
}
export interface PptxExtensionInfo {
  readonly uri?: string;
  readonly status: "interpreted" | "uninterpreted";
  /** Direct payload roots only, not raw XML or preservation handles. */
  readonly children: readonly XmlExpandedName[];
}
export interface PptxSectionInfo {
  /** Optional in the source schema; never synthesized from a name or index. */
  readonly id?: string;
  readonly name?: string;
  readonly slides: readonly (PptxPartReference & { readonly id: number })[];
  readonly extensions: readonly PptxExtensionInfo[];
}
export interface PptxSectionListInfo {
  readonly part: string;
  readonly extensionUri: string;
  readonly sections: readonly PptxSectionInfo[];
}
export interface PptxStructureFinding {
  readonly code: "missing-inheritance" | "unlisted-master" | "unlisted-layout" | "unlisted-slide" | "section-order" | "unsectioned-slide";
  readonly part: string;
  readonly message: string;
  readonly sectionIndex?: number;
  readonly slideId?: number;
}
/** Structural inventory; no Core document, semantic editability or preservation guarantee. */
export interface PptxInspection extends PresentationPackageInspection {
  readonly diagramInventory?: PptxDiagramResourceInventory;
  readonly metadata: PptxDocumentMetadata;
  readonly sectionList?: PptxSectionListInfo;
  readonly presentationExtensions: readonly PptxExtensionInfo[];
  readonly slideSize?: PptxSize;
  readonly notesSize: PptxSize;
  readonly firstSlideNumber: number;
  /** Whether firstSlideNum was explicitly present; omitted by older producers of this interface. */
  readonly firstSlideNumberDeclared?: boolean;
  /** Authoritative presentation order, independent of ZIP/relationship/name ordering. */
  readonly slides: readonly PptxSlideInfo[];
  readonly slideMasters: readonly PptxMasterInfo[];
  readonly notesMasters: readonly PptxPartReference[];
  readonly handoutMasters: readonly PptxPartReference[];
  readonly layouts: readonly { readonly part: string; readonly masterPart?: string; readonly name?: string }[];
  readonly findings: readonly PptxStructureFinding[];
  readonly markupCompatibility: readonly { readonly part: string; readonly report: MarkupCompatibilityReport }[];
}
