/** Serializable, value-checked operations on factory-managed presentations. */
export type PresentationEditBox = { readonly x: number; readonly y: number; readonly width: number; readonly height: number | null } | null;
/** Null properties are absent; a null transform is distinct from an empty transform. */
export type PresentationEditTransform = { readonly rotation: number | null; readonly flipH: boolean | null; readonly flipV: boolean | null } | null;
export type PresentationEditAction = { readonly type: "url"; readonly url: string; readonly tooltip: string | null } | { readonly type: "slide"; readonly slideId: string; readonly tooltip: string | null } | null;
export type PresentationEditReadingOrder = readonly { readonly groupId?: string; readonly elementIds: readonly string[] }[] | null;
export type PresentationEditAccessibility = { readonly description: string | null; readonly decorative: boolean | null } | null;
/** String token preserves negative zero across JSON serialization. */
export type PresentationEditChartValue = number | null | "-0";
export type PresentationEditChange =
  | { readonly type: "chart-category"; readonly slideId: string; readonly elementId: string; readonly categoryIndex: number; readonly expected: string; readonly value: string }
  | { readonly type: "chart-series-name"; readonly slideId: string; readonly elementId: string; readonly seriesIndex: number; readonly expected: string; readonly value: string }
  | { readonly type: "chart-point"; readonly slideId: string; readonly elementId: string; readonly seriesIndex: number; readonly pointIndex: number; readonly coordinate: "values" | "xValues" | "yValues" | "sizes"; readonly expected: PresentationEditChartValue; readonly value: PresentationEditChartValue }
  | { readonly type: "notes-string"; readonly slideId: string; readonly expected: string; readonly value: string }
  | { readonly type: "notes-run"; readonly slideId: string; readonly paragraphIndex: number; readonly runIndex: number; readonly expected: string; readonly value: string }
  | { readonly type: "notes-run-action"; readonly slideId: string; readonly paragraphIndex: number; readonly runIndex: number; readonly expected: PresentationEditAction; readonly value: PresentationEditAction }
  | { readonly type: "element-accessibility"; readonly slideId: string; readonly elementId: string; readonly expected: PresentationEditAccessibility; readonly value: PresentationEditAccessibility }
  | { readonly type: "slide-reading-order"; readonly slideId: string; readonly expected: PresentationEditReadingOrder; readonly value: PresentationEditReadingOrder }
  | { readonly type: "text-run-action"; readonly slideId: string; readonly elementId: string; readonly paragraphIndex: number; readonly runIndex: number; readonly expected: PresentationEditAction; readonly value: PresentationEditAction }
  | { readonly type: "table-cell-text-run-action"; readonly slideId: string; readonly elementId: string; readonly rowIndex: number; readonly cellIndex: number; readonly paragraphIndex: number; readonly runIndex: number; readonly expected: PresentationEditAction; readonly value: PresentationEditAction }
  | { readonly type: "element-action"; readonly slideId: string; readonly elementId: string; readonly expected: PresentationEditAction; readonly value: PresentationEditAction }
  | { readonly type: "element-transform"; readonly slideId: string; readonly elementId: string; readonly expected: PresentationEditTransform; readonly value: PresentationEditTransform }
  | { readonly type: "element-name"; readonly slideId: string; readonly elementId: string; readonly expected: string | null; readonly value: string | null }
  | { readonly type: "element-box"; readonly slideId: string; readonly elementId: string; readonly expected: PresentationEditBox; readonly value: PresentationEditBox }
  | { readonly type: "element-hidden"; readonly slideId: string; readonly elementId: string; readonly expected: boolean | null; readonly value: boolean | null }
  | { readonly type: "element-opacity"; readonly slideId: string; readonly elementId: string; readonly expected: number | null; readonly value: number | null }
  | { readonly type: "text-run"; readonly slideId: string; readonly elementId: string; readonly paragraphIndex: number; readonly runIndex: number; readonly expected: string; readonly value: string }
  | { readonly type: "table-cell-text-run"; readonly slideId: string; readonly elementId: string; readonly rowIndex: number; readonly cellIndex: number; readonly paragraphIndex: number; readonly runIndex: number; readonly expected: string; readonly value: string }
  | { readonly type: "text-string"; readonly slideId: string; readonly elementId: string; readonly expected: string; readonly value: string }
  | { readonly type: "table-cell-text-string"; readonly slideId: string; readonly elementId: string; readonly rowIndex: number; readonly cellIndex: number; readonly expected: string; readonly value: string }
  | { readonly type: "slide-hidden"; readonly slideId: string; readonly expected: boolean; readonly value: boolean }
  | { readonly type: "first-slide-number"; readonly expected: number | null; readonly value: number | null }
  | { readonly type: "slide-order"; readonly expected: readonly string[]; readonly value: readonly string[] };

export interface PresentationEditOperation {
  readonly schemaVersion: "pptkit-edit-v1";
  readonly documentId: string;
  readonly changes: readonly PresentationEditChange[];
}

export interface PresentationEditResult {
  readonly inverse: PresentationEditOperation;
  readonly changes: readonly PresentationEditChange[];
}

export type PresentationEditErrorCode = "invalid-operation" | "wrong-document" | "missing-target" | "edit-conflict";
export class PresentationEditError extends Error {
  constructor(readonly code: PresentationEditErrorCode, message: string) { super(message); this.name = "PresentationEditError"; }
}
