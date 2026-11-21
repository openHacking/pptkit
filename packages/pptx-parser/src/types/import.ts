import type { PresentationDocument } from "@pptkit/core";
import type { PptxInspection } from "./structure.js";
import type { PptxInspectionOptions } from "./structure.js";
import type { PptxPlaceholderInheritance, PptxTextStyleInheritance, PptxTableStyleInheritance } from "./inheritance.js";
import type { PptxSourceLocation } from "./source.js";
export type { PptxSourceLocation } from "./source.js";

export type ImportOperation = "create" | "import" | "edit" | "preview" | "export" | "no-op-save" | "edit-save";
export type ImportCapabilityTier = "editable" | "preserved" | "fallback" | "unsupported";
export interface ImportFinding {
  readonly id: string;
  readonly feature: string;
  readonly operation: ImportOperation;
  readonly tier: ImportCapabilityTier;
  readonly source: PptxSourceLocation;
  readonly slideId?: string;
  readonly elementId?: string;
  readonly reason: string;
  /** Source dependencies when an explicit profile provides a bounded resolution trace. */
  readonly resolvedFrom?: readonly PptxSourceLocation[];
}
export interface ImportDiagnostic {
  readonly code: string;
  readonly severity: "info" | "warning" | "error";
  readonly message: string;
  readonly source?: PptxSourceLocation;
  readonly slideId?: string;
  readonly elementId?: string;
  /** Core validation path, distinct from an OOXML source path. */
  readonly documentPath?: string;
}
export interface ImportSourceIdentity {
  readonly kind: "presentation" | "slide" | "master" | "layout" | "element" | "asset" | "notes";
  readonly id: string;
  readonly source: PptxSourceLocation;
}
export interface ImportedChartValueSource {
  readonly slideId: string;
  readonly elementId: string;
  readonly seriesIndex: number;
  readonly pointIndex: number;
  /** Scatter coordinates are independently addressed; absent for category charts. */
  readonly coordinate?: "x" | "y" | "size";
  readonly source: PptxSourceLocation;
  readonly sourceElementIndex?: number;
  readonly workbook?: { readonly part: string; readonly worksheetPart: string; readonly cellReference: string; readonly sourceElementIndex: number };
  readonly status: "bound" | "unmapped";
  readonly reason?: string;
}
/** Source eligibility for one Core category; final save checks still apply. */
export interface ImportedChartLabelSource {
  readonly slideId: string;
  readonly elementId: string;
  readonly categoryIndex: number;
  readonly status: "bound" | "unmapped";
  readonly reason?: string;
  /** Observed cache/cell pairs; an unmapped record may be incomplete. */
  readonly sources: readonly { readonly cache: PptxSourceLocation; readonly workbook: { readonly part: string; readonly worksheetPart: string; readonly cellReference: string; readonly sourceElementIndex: number; readonly preservesWhitespace: boolean } }[];
}
/** Source eligibility for one chart series name; final save checks still apply. */
export interface ImportedChartSeriesNameSource extends Omit<ImportedChartLabelSource, "categoryIndex"> { readonly seriesIndex: number }
/** Source identity only; does not grant permission to edit the image. */
export interface ImportedPictureSource {
  readonly slideId: string;
  readonly elementId: string;
  readonly assetId: string;
  readonly source: PptxSourceLocation;
  readonly sourceElementIndex: number;
  readonly relationshipId: string;
  readonly mediaPart: string;
  /** The original blip also advertises an SVG representation. */
  readonly hasSvgExtension: boolean;
  /** The original blip also has an external relationship; it is never fetched. */
  readonly hasExternalLink: boolean;
}
/** Read-only source inventory; audio bytes and playback are not projected. */
export interface ImportedAudioSource {
  readonly slideId: string;
  readonly declaration?: "audioFile" | "wavAudioFile" | "p14:media";
  readonly relationshipAttribute?: "link" | "embed";
  /** For Office media declarations, relation-target comparison with one direct audio declaration. */
  readonly companionStatus?: "same-part" | "different-part" | "unverified";
  readonly source: PptxSourceLocation;
  readonly sourceElementIndex: number;
  readonly relationshipId?: string;
  readonly status: "wave-pcm" | "external" | "unsupported";
  readonly mediaPart?: string;
  /** Unresolved relationship target for external references; never fetched. */
  readonly externalTarget?: string;
  readonly contentType?: string;
  readonly reason?: string;
  readonly pcm?: { readonly channels: 1 | 2; readonly sampleRate: number; readonly bitsPerSample: 8 | 16; readonly frameCount: number; readonly durationSeconds: number };
}
/** Read-only slide-local video and poster source references; no decoding or write permission. */
export interface ImportedVideoSource {
  readonly slideId: string;
  readonly declaration: "videoFile" | "p14:media" | "poster";
  readonly relationshipAttribute: "link" | "embed";
  readonly companionStatus?: "same-part" | "different-part" | "unverified";
  readonly source: PptxSourceLocation;
  readonly sourceElementIndex: number;
  readonly relationshipId?: string;
  readonly status: "iso-bmff-framing" | "internal" | "external" | "unsupported";
  readonly mediaPart?: string;
  /** Unresolved relationship target for external references; never fetched. */
  readonly externalTarget?: string;
  readonly contentType?: string;
  readonly reason?: string;
  /** Bounded movie-header clock, when available; not track duration or playback validation. */
  readonly movieDurationMs?: number;
  readonly framing?: { readonly majorBrand: string; readonly minorVersion: number; readonly compatibleBrands: readonly string[]; readonly boxCount: number };
}
/** Opt-in metadata eligibility; insertable does not guarantee a valid save plan. */
export interface ImportedCoreMetadataSource {
  readonly field: "title" | "author" | "subject" | "description" | "language" | "revision" | "keywords" | "company" | "createdAt" | "modifiedAt";
  readonly source?: PptxSourceLocation;
  readonly status: "bound" | "unmapped" | "insertable";
  /** Present for bound values under removal profiles; source guards may still block a save. */
  readonly removable?: boolean;
  readonly reason?: string;
}
/** Raw custom-property identity inventory; does not grant edit-save permission. */
export interface ImportedCustomPropertySource {
  readonly source: PptxSourceLocation;
  readonly valueElementIndex?: number;
  /** Present only when the save profile binds this source to a Core array entry. */
  readonly projectedIndex?: number;
  /** Present with the rename profile when the source name is eligible for a rename. */
  readonly nameEditable?: boolean;
  readonly status: "bound" | "unmapped";
  readonly reason?: string;
}
/** Opt-in source eligibility only; does not guarantee a save plan. */
export interface ImportedSlideVisibilitySource {
  readonly slideId: string;
  readonly source: PptxSourceLocation;
  readonly status: "bound" | "unmapped";
  readonly reason?: string;
}
/** Source identities for explicit series permutations; eligibility is not a save guarantee. */
export interface ImportedChartSeriesOrderSource {
  readonly slideId: string;
  readonly elementId: string;
  readonly source: { readonly part: string };
  readonly status: "bound" | "unmapped";
  readonly reason?: string;
  readonly series: readonly { readonly seriesIndex: number; readonly sourceId: number }[];
}
export interface ImportReport {
  /** Explicit number/date cache projection; not permission to edit source fields. */
  readonly fieldTextProfile?: "slide-number-v1" | "date-and-number-v1" | "date-time-and-number-v1";
  /** Opt-in slide-local projected picture references; no write permission. */
  readonly pictureSources?: readonly ImportedPictureSource[];
  readonly audioSources?: readonly ImportedAudioSource[];
  readonly videoSources?: readonly ImportedVideoSource[];
  /** Explicit bounded notes-master text inheritance; not notes-page layout or write permission. */
  readonly notesTextProfile?: "master-style-v1";
  /** Explicit font content profile, independent of save permissions. */
  readonly textFontProfile?: "script-families-v1" | "script-families-v2" | "script-families-v3";
  /** Explicit core metadata source eligibility, not a save guarantee. */
  readonly coreMetadataSources?: readonly ImportedCoreMetadataSource[];
  /** Opt-in raw custom-property identities, never edit-save permission. */
  readonly customPropertySources?: readonly ImportedCustomPropertySource[];
  /** V4 source availability only; not authorization for an insertion. */
  readonly customPropertyInsertion?: { readonly status: "available" | "unavailable"; readonly part?: string; readonly reason?: string };
  /** Explicit slide visibility source eligibility, not a save guarantee. */
  readonly slideVisibilitySources?: readonly ImportedSlideVisibilitySource[];
  /** Explicit chart value profile only; eligibility, not a ready save plan. */
  readonly chartValues?: readonly ImportedChartValueSource[];
  /** V4 chart category sources only; eligibility is not a ready save plan. */
  readonly chartLabels?: readonly ImportedChartLabelSource[];
  /** V5 series-name sources only; eligibility is not a ready save plan. */
  readonly chartSeriesNames?: readonly ImportedChartSeriesNameSource[];
  readonly chartSeriesOrders?: readonly ImportedChartSeriesOrderSource[];
  /** Opt-in cached drawing approximations, never semantic SmartArt nodes. */
  readonly diagramPreviews?: readonly ImportedDiagramPreview[];
  readonly projection: "structure" | "content";
  readonly inspection: PptxInspection;
  readonly findings: readonly ImportFinding[];
  readonly identities: readonly ImportSourceIdentity[];
  readonly diagnostics: readonly ImportDiagnostic[];
  /** Content projection only. Bounded placeholder matching/dimension provenance. */
  readonly placeholderInheritance?: PptxPlaceholderInheritance;
  readonly textStyleInheritance?: PptxTextStyleInheritance;
  readonly tableStyleInheritance?: PptxTableStyleInheritance;
  /** Opt-in drawing-run provenance, not a blanket editability guarantee. */
  readonly textRuns?: readonly ImportedTextRunSource[];
  /** text-runs-v2/v3. Notes belong to slides, not fabricated drawing IDs. */
  readonly notesTextRuns?: readonly ImportedNotesTextRunSource[];
  /** text-runs-v3 only. Compact Core cell identity plus physical source column. */
  readonly tableTextRuns?: readonly ImportedTableTextRunSource[];
}
export interface ImportedDiagramPreview {
  readonly slideId: string;
  readonly elementId: string;
  readonly source: PptxSourceLocation;
  readonly cachePart?: string;
  readonly status: "cache" | "partial-cache" | "placeholder";
  readonly freshness: "unknown";
  readonly reason?: string;
}
export interface ImportedTextRunSource {
  readonly slideId: string;
  readonly elementId: string;
  /** Present in text-runs-v2/v3; v1 retains its original record shape. */
  readonly owner?: "text-box" | "shape-text";
  readonly paragraphIndex: number;
  readonly runIndex: number;
  readonly source: PptxSourceLocation;
  readonly sourceElementIndex?: number;
  readonly status: "bound" | "unmapped";
  readonly reason?: string;
}
export interface ImportedNotesTextRunSource {
  readonly slideId: string;
  readonly paragraphIndex: number;
  readonly runIndex: number;
  readonly source: PptxSourceLocation;
  readonly sourceElementIndex?: number;
  readonly status: "bound" | "unmapped";
  readonly reason?: string;
}
export interface ImportedTableTextRunSource {
  readonly slideId: string;
  readonly elementId: string;
  /** Zero-based Core/source row and compact Core anchor cell. */
  readonly rowIndex: number;
  readonly cellIndex: number;
  /** Zero-based physical grid column of the source anchor cell. */
  readonly columnIndex: number;
  readonly paragraphIndex: number;
  readonly runIndex: number;
  readonly source: PptxSourceLocation;
  readonly sourceElementIndex?: number;
  readonly status: "bound" | "unmapped";
  readonly reason?: string;
}
export interface PptxSaveChange {
  readonly kind: "text" | "style" | "geometry" | "asset" | "structure" | "metadata" | "other";
  readonly documentPath: string;
  readonly slideId?: string;
  readonly elementId?: string;
  readonly source?: PptxSourceLocation;
  /** Scalar text changes only; no mutable authoring state is retained. */
  readonly text?: { readonly before: string; readonly after: string };
  /** Existing numeric chart values are classified as other, with precise source provenance. */
  readonly chartValue?: { readonly before: number; readonly after: number; readonly seriesIndex: number; readonly pointIndex: number; readonly coordinate?: "x" | "y" | "size" };
  /** Existing category label change; a label can own several physical cache points. */
  readonly chartLabel?: { readonly before: string; readonly after: string; readonly categoryIndex: number };
  readonly chartSeriesName?: { readonly before: string; readonly after: string; readonly seriesIndex: number };
  /** Explicit core metadata string change; values are not inferred from an application or clock. */
  readonly metadata?: { readonly before: string; readonly after: string };
  /** Existing presentation numbering change, when both values are present. */
  readonly firstSlideNumber?: { readonly before?: number; readonly after?: number };
}
export interface SaveLoss {
  readonly id: string;
  readonly feature: string;
  readonly source?: PptxSourceLocation;
  readonly slideId?: string;
  readonly elementId?: string;
  readonly reason: string;
  readonly impact: string;
  readonly alternative?: string;
}
export type PptxSaveMode = "no-op" | "edit" | "rebuild";
export interface PptxSaveReport {
  readonly mode: PptxSaveMode;
  readonly losses: readonly SaveLoss[];
  readonly diagnostics: readonly ImportDiagnostic[];
  /** Blocked reports have no output assurance. */
  readonly assurance: "none" | "original-bytes" | "scoped-preservation" | "semantic-rebuild";
  readonly changes?: readonly PptxSaveChange[];
}
declare const planIdentity: unique symbol;
export interface PptxSavePlan {
  readonly [planIdentity]: true;
  readonly report: PptxSaveReport;
  readonly byteLength: number;
  dispose(): void;
}
export type PptxSavePreparation = { readonly status: "ready"; readonly plan: PptxSavePlan } | { readonly status: "blocked"; readonly report: PptxSaveReport };
export interface PptxSessionLimits {
  /** Cumulative opt-in notes-master matching, style/theme evaluation and audits. */
  maxNotesStyleWork: number;
  /** Session-wide table-style declarations and selected property provenance. */
  maxTableStyleProperties?: number;
  maxTableStyleDependencies?: number;
  maxTableStyleWork?: number;
  /** Opt-in numeric chart binding scan/work across all raw source references. */
  maxChartBindingWork?: number;
  /** Cache-preview traversal, including repeated uses of shared cache XML. */
  maxDiagramPreviewWork?: number;
  /** Distinct placeholder declarations visited across slides, layouts and masters. */
  maxPlaceholderEntries?: number;
  /** Selected frame/paragraph/run property provenance records in content mode. */
  maxTextStyleProperties?: number;
  /** Sum of resolvedFrom locations across selected property records. */
  maxTextStyleDependencies?: number;
  /** Charged normalized diff values and path characters during edit preparation. */
  maxSaveAnalysisWork?: number;
  /** Complete edit-change/diagnostic reports are never truncated. */
  maxSaveChanges?: number;
  /** Combined source/candidate expanded bytes checked by preservation validation. */
  maxSaveValidationBytes?: number;
  /** Exact source/candidate payload comparison work. */
  maxSaveValidationWork?: number;
  /** Expanded ZIP bytes per embedded chart workbook, including control streams. */
  maxExpandedChartWorkbookBytes?: number;
  /** Sum across distinct inspected chart workbooks in a content import. */
  maxTotalExpandedChartWorkbookBytes?: number;
  /** Total cached/literal points, workbook cells/shared strings and referenced cell reads. */
  maxChartDataPoints?: number;
  /** Expanded bytes per mapped image part, before Base64 registration. */
  maxImportedAssetBytes?: number;
  /** Sum of distinct mapped image parts; repeated relationships count once. */
  maxTotalImportedAssetBytes?: number;
  maxReportFindings: number;
  /** Normalized values visited during snapshot validation, including primitives. */
  maxSnapshotValues: number;
  maxPreparedOutputBytes: number;
  /** Includes preparations in flight; released by disposing plans/session. */
  maxRetainedOutputBytes: number;
  /** Includes reservations held by preparations in flight. */
  maxLivePlans: number;
}
export interface PptxImportOptions extends PptxInspectionOptions {
  /** Explicit exact currency projection; exact-save-v2 additionally requires scalar-v1 source saves. */
  customCurrencyProjection?: "exact-v1" | "exact-save-v2";
  /** Content only. Preserve existing directly bound drawing names; may be combined with description mixed-v5. */
  drawingNameSaveProfile?: "existing-v1";
  /** Content only. Preserve existing directly bound decorative values; may be combined with description mixed-v5. */
  decorativeSaveProfile?: "existing-v1";
  /** Content only. Preserve explicitly edited slide visibility through source show attributes. */
  slideVisibilitySaveProfile?: "attribute-v1";
  /** Content only. V1 edits values; V2 adds names; V3 removals; V4 appends; V5 can create a missing part. */
  customPropertySaveProfile?: "scalar-v1" | "scalar-and-name-v2" | "remove-v3" | "append-and-remove-v4" | "create-and-append-v5";
  /** Content only. Inventory raw custom-property element identities without granting save authority. */
  customPropertySourceProfile?: "raw-v1";
  /** Content only. V5 adds missing core strings inside an existing core-properties part. */
  metadataSaveProfile?: "core-strings-v1" | "core-values-v2" | "core-keywords-v3" | "extended-company-v4" | "core-insert-v5" | "core-date-values-v6" | "core-date-insert-v7" | "core-create-v8" | "core-revision-insert-v9" | "core-keywords-insert-v10" | "extended-company-insert-v11" | "extended-create-v12" | "metadata-remove-v13" | "metadata-date-remove-v14" | "metadata-keywords-remove-v15" | "metadata-revision-remove-v16" | "metadata-default-remove-v17";
  /** Content only. V2 admits insertion/removal; v3 batches; v4 combines text; v5 adds chart/number. */
  descriptionEditSaveProfile?: "existing-v1" | "attribute-v2" | "batch-v3" | "mixed-v4" | "mixed-v5";
  /** Content only. Map simple slide-number and optional date fields; only adjacent bound text is eligible for text-run saves. */
  fieldTextProfile?: "slide-number-v1" | "date-and-number-v1" | "date-time-and-number-v1";
  /** Content only. One source-bound picture retarget with a new byte-backed raster asset. */
  pictureEditSaveProfile?: "replace-v1";
  /** Content only. Inventory projected slide-local picture reference identities. */
  pictureSourceProfile?: "embedded-v1";
  /** Content only. Inspect slide-local picture audio and companion Office media sources as bounded PCM WAVE. */
  audioSourceProfile?: "wave-pcm-v1";
  /** Read-only inventory of videoFile, its Office companion, and poster source references. */
  videoSourceProfile?: "iso-bmff-v1";
  /** Explicit partial video-object projection; does not validate codecs or playback. */
  videoProjection?: "iso-bmff-v1";
  /** Content only. Project supported embedded PCM WAVE as Core audio; no source edit-save. */
  audioProjection?: "wave-pcm-v1";
  /** Content only. artwork-v10 adds shape patterns; artwork-v11 adds explicit bounded outer shadows. */
  masterArtwork?: "pictures-v1" | "artwork-v1" | "artwork-v2" | "artwork-v3" | "artwork-v4" | "artwork-v5" | "artwork-v6" | "artwork-v7" | "artwork-v8" | "artwork-v9" | "artwork-v10" | "artwork-v11";
  /** Content only. Resolve notesStyle/body-placeholder defaults and independent notes theme. */
  notesTextProfile?: "master-style-v1";
  /** Content only. V1 maps named slots; v2 adds supplemental rules; v3 requires their effective native bucket. */
  textFontProfile?: "script-families-v1" | "script-families-v2" | "script-families-v3";
  /** Content only. V4 adds category labels; V5 adds series names with proven simple workbook sources. */
  chartEditSaveProfile?: "existing-values-v1" | "existing-values-v2" | "existing-values-v3" | "existing-values-and-labels-v4" | "existing-values-and-series-names-v5" | "existing-series-order-v6";
  /** Content maps reported subsets of drawings, notes, tables and bar/line/pie/doughnut/area charts. */
  projection: "structure" | "content";
  /** Fixed-topology characters: v1 boxes; v2 adds shapes/notes; v3 adds table anchors. */
  editSaveProfile?: "text-runs-v1" | "text-runs-v2" | "text-runs-v3";
  /** Content only. V1 replaces an existing firstSlideNum; v2 also inserts/clears it and can combine with text-run and existing chart-value profiles. */
  firstSlideNumberSaveProfile?: "existing-v1" | "attribute-v2";
  /** Requires a text-run profile. Preserve proven direct shape paragraph-index targets during scalar text edits. */
  textTargetSaveProfile?: "paragraph-stable-v1";
  /** Content only. Explicitly use cached drawing approximations of unknown freshness. */
  diagramPreview?: "cache";
  /** Content only. Show a static placeholder for media without a mapped poster. */
  mediaPreview?: "placeholder-v1";
  sessionLimits?: Partial<PptxSessionLimits>;
}
export interface ImportedPresentation {
  readonly document: PresentationDocument;
  readonly report: ImportReport;
  prepareSave(options: { readonly mode: PptxSaveMode; readonly signal?: AbortSignal; readonly customPropertyRemovals?: readonly string[]; readonly customPropertyAdditions?: readonly number[]; readonly coreMetadataInsertions?: readonly ("title" | "author" | "subject" | "description" | "language" | "revision")[]; readonly coreMetadataRemovals?: readonly ("title" | "author" | "language" | "revision")[]; readonly chartSeriesOrders?: readonly { readonly slideId: string; readonly elementId: string; readonly identities: readonly number[] }[] }): Promise<PptxSavePreparation>;
  /** Synchronous ownership/lifecycle/consent/freshness guard; does not copy output. */
  assertSavePlanCurrent(plan: PptxSavePlan, options?: { readonly acceptLossIds?: readonly string[] }): void;
  save(plan: PptxSavePlan, options?: { readonly acceptLossIds?: readonly string[] }): Promise<{ readonly bytes: Uint8Array; readonly report: PptxSaveReport }>;
  dispose(): void;
}
export type PptxImportErrorCode = "import-invalid-package" | "import-unsupported" | "import-limit-exceeded" | "operation-aborted" | "session-disposed" | "plan-disposed" | "plan-session-mismatch" | "save-snapshot-stale" | "save-loss-consent-mismatch" | "save-limit-exceeded" | "save-invalid-document" | "save-candidate-invalid" | "session-not-serializable";
export class PptxImportError extends Error {
  readonly source: PptxSourceLocation | undefined;
  constructor(readonly code: PptxImportErrorCode, message: string, source?: PptxSourceLocation, options?: ErrorOptions) {
    super(message, options); this.name = "PptxImportError"; this.source = source ? Object.freeze({ ...source }) : undefined;
  }
}
