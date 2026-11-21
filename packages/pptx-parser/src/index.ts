export { inspectMasterPictures as inspectPptxMasterPictures, DEFAULT_MASTER_PICTURE_LIMITS } from "./master-pictures/inventory.js";
export type { MasterPictureInventory, MasterPictureInventoryItem, MasterPictureInventoryLimits, MasterPictureInventoryOptions, MasterPictureRelationship, MasterPicturePartInventory } from "./types/master-pictures.js";
import { readArchive } from "./archive/read-archive.js";
import type { ZipArchiveInspection, ZipInspectionOptions } from "./types/archive.js";
export { DEFAULT_ZIP_LIMITS } from "./archive/read-archive.js";
export { PackageInspectionError } from "./types/archive.js";
export type { PackageSourceLocation, PackageInspectionErrorCode, ZipArchiveInspection, ZipEntryInfo, ZipInspectionLimits, ZipInspectionOptions } from "./types/archive.js";

/** ZIP envelope inspection only: does not establish OPC/PPTX validity or import semantics. */
export async function inspectZipArchive(bytes: Uint8Array, options?: ZipInspectionOptions): Promise<ZipArchiveInspection> {
  return (await readArchive(bytes, options)).inspection;
}

export { DEFAULT_XML_LIMITS } from "./xml/read-xml.js";
export type { XmlExpandedName, XmlInspectionLimits, XmlInspectionOptions, XmlPartInspection } from "./types/xml.js";
import { readXml } from "./xml/read-xml.js";
import type { XmlInspectionOptions, XmlPartInspection } from "./types/xml.js";

/** Bounded XML well-formedness inspection only; does not validate OOXML schemas. */
export async function inspectXmlPart(bytes: Uint8Array, options?: XmlInspectionOptions): Promise<XmlPartInspection> {
  return (await readXml(bytes, options)).inspection;
}

export { DEFAULT_OPC_LIMITS } from "./opc/read-package.js";
export type { OpcInspectionLimits, OpcInspectionOptions, OpcPackageInspection, OpcPartInfo, OpcRelationshipInfo } from "./types/opc.js";
import { readOpcPackage } from "./opc/read-package.js";
import type { OpcInspectionOptions, OpcPackageInspection } from "./types/opc.js";

/** OPC package profile inspection; not PPTX semantic import or an OOXML schema guarantee. */
export async function inspectOpcPackage(bytes: Uint8Array, options?: OpcInspectionOptions): Promise<OpcPackageInspection> {
  return readOpcPackage(bytes, options);
}

export { DEFAULT_COMPARISON_LIMITS } from "./comparison/compare-packages.js";
export type { OpcComparisonProfile, OpcComparisonLimits, OpcComparisonOptions, OpcDifferenceCode, OpcPackageDifference, OpcComparisonSummary, OpcPackageComparison } from "./types/comparison.js";
export { comparePackages as compareOpcPackages } from "./comparison/compare-packages.js";

export { DEFAULT_CLOSURE_LIMITS } from "./preservation/inspect-closure.js";
export { inspectClosure as inspectOpcClosure } from "./preservation/inspect-closure.js";
export type { OpcClosureLimits, OpcClosureOptions, OpcSourceRelationship, OpcClosureFinding, OpcClosurePart, OpcClosureInspection } from "./types/closure.js";

export type { PresentationDialect, PresentationFileKind, PresentationPackageInspection } from "./types/presentation.js";
import { readPresentationPackage } from "./presentation/inspect.js";
import type { PresentationPackageInspection } from "./types/presentation.js";

/** Identify the PresentationML entry point and dialect; does not import or validate its full schema. */
export async function inspectPresentationPackage(bytes: Uint8Array, options?: OpcInspectionOptions): Promise<PresentationPackageInspection> {
  return readPresentationPackage(bytes, options);
}

export type { MarkupCompatibilityOptions, MarkupCompatibilityDecision, MarkupCompatibilityReport, MarkupCompatibilityInspection } from "./types/compatibility.js";
import { inspectCompatibility } from "./compatibility/inspect.js";
import type { MarkupCompatibilityInspection, MarkupCompatibilityOptions } from "./types/compatibility.js";

/** Apply an explicit consumer's MCE configuration; returns inventory, not a preservation serializer. */
export async function inspectMarkupCompatibility(bytes: Uint8Array, options: MarkupCompatibilityOptions): Promise<MarkupCompatibilityInspection> {
  return inspectCompatibility(bytes, options);
}

export type { PptxInspection, PptxSize, PptxPartReference, PptxSlideInfo, PptxMasterInfo, PptxStructureFinding } from "./types/structure.js";
export type { PptxSectionInfo, PptxSectionListInfo, PptxExtensionInfo } from "./types/structure.js";
export type { PptxDocumentMetadata, MetadataPartInfo, CorePropertyInfo, MetadataKeywordSegment, ExtendedPropertyInfo, CustomPropertyInfo, MetadataPropertyValue, MetadataScalarKind } from "./types/metadata.js";
import type { PptxInspection, PptxInspectionOptions } from "./types/structure.js";
import { readPresentationStructure } from "./presentation/read-structure.js";

/** Read presentation ordering/dimensions and structural relationships; not semantic import or a save session. */
export async function inspectPptx(bytes: Uint8Array, options?: PptxInspectionOptions): Promise<PptxInspection> {
  return readPresentationStructure(bytes, options);
}

export { PptxImportError } from "./types/import.js";
export type { ImportedPresentation, PptxImportOptions, PptxSessionLimits, PptxImportErrorCode, PptxSourceLocation, ImportOperation, ImportCapabilityTier, ImportFinding, ImportDiagnostic, ImportSourceIdentity, ImportReport, ImportedTextRunSource, ImportedNotesTextRunSource, ImportedTableTextRunSource, PptxSaveChange, SaveLoss, PptxSaveMode, PptxSaveReport, PptxSavePlan, PptxSavePreparation } from "./types/import.js";
export type { PptxPlaceholderLevel, PptxPlaceholderDimension, PptxPlaceholderMatch, PptxInheritedDimension, PptxPlaceholderInfo, PptxPlaceholderInheritance } from "./types/inheritance.js";
export type { PptxTextStyleLevel, PptxTextStyleProperty, PptxTextStyleValue, PptxTextStylePropertyInfo, PptxTextStyleInheritance } from "./types/inheritance.js";
export { DEFAULT_SESSION_LIMITS } from "./import/import-pptx.js";
import { importPresentation } from "./import/import-pptx.js";
import type { ImportedPresentation, PptxImportOptions } from "./types/import.js";

/** Explicit projection/no-op session, with opt-in fixed-topology direct text-run edits. */
export async function importPptx(bytes: Uint8Array, options: PptxImportOptions): Promise<ImportedPresentation> {
  return importPresentation(bytes, options);
}

export type { PptxTransitionInfo } from "./types/structure.js";
export type { PptxDiagramInfo, PptxDiagramComponentInfo, PptxDiagramComponentRole } from "./types/smartart.js";
export type { PptxDiagramCacheReference, PptxDiagramCacheInfo, PptxDiagramCacheCandidate, PptxDiagramImageRelationship, PptxDiagramResourceInfo, PptxDiagramResourceInventory } from "./types/smartart.js";
export type { ImportedDiagramPreview } from "./types/import.js";
export type { PptxDiagramModelReference, PptxDiagramTextParagraph, PptxDiagramPointInfo, PptxDiagramConnectionInfo, PptxDiagramModelFinding, PptxDiagramDataModelInfo } from "./types/smartart.js";
export type { PptxInspectionOptions, PptxDiagramInspectionLimits } from "./types/structure.js";
export { DEFAULT_DIAGRAM_LIMITS } from "./presentation/diagram-limits.js";
export type { PptxDiagramPresentationMapping, PptxDiagramCacheMapping } from "./types/smartart.js";
export type { PptxDiagramTextComparison } from "./types/smartart.js";

export { describeImportCapabilities, DEFAULT_IMPORT_CAPABILITY_LIMITS } from "./import/capability-report.js";
export type { ImportCapabilityLimits, ImportCapabilityOptions, ImportCapabilityReferences, ImportCapabilityAssessment, ImportCapabilitySubject, ImportCapabilityReport } from "./types/capability.js";

export type { ImportedChartValueSource, ImportedChartLabelSource, ImportedChartSeriesNameSource, ImportedChartSeriesOrderSource, ImportedPictureSource, ImportedAudioSource, ImportedVideoSource, ImportedSlideVisibilitySource, ImportedCoreMetadataSource, ImportedCustomPropertySource } from "./types/import.js";

export { inspectNativeTiming as inspectPptxTiming } from "./timing/read-timing.js";
export { inspectPptxCaptionTracks } from "./captions/inspect.js";
export type { PptxCaptionInspection, PptxCaptionInspectionOptions, PptxCaptionInspectionLimits, PptxCaptionTrackSource } from "./types/captions.js";
export type { PptxTimingInspectionLimits, PptxTimingInspectionOptions, PptxTimingView, PptxTimingAttribute, PptxTimingReference, PptxTimingNodeInfo, PptxTimingDrawingInfo, PptxTimingFinding, PptxSlideTimingInventory, PptxTimingInspection } from "./types/timing.js";

export type { PptxTableStyleInheritance, PptxTableStyleSelection, PptxTableStylePropertyInfo } from "./types/inheritance.js";

export { inspectNotes as inspectPptxNotes } from "./notes/read-notes.js";
export { DEFAULT_NOTES_INSPECTION_LIMITS } from "./notes/budget.js";
export type { PptxNotesInspectionLimits, PptxNotesInspectionOptions, PptxNotesView, PptxNotesNodeInfo, PptxNotesRole, PptxNotesDrawingInfo, PptxNotesGeometryInfo, PptxNotesTextInfo, PptxNotesFieldInfo, PptxNotesPartInfo, PptxNotesPlaceholderMatch, PptxNotesPageInfo, PptxNotesFinding, PptxNotesInspection } from "./types/notes.js";
