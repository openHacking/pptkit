import type { ElementAction, TextParagraphInput } from "./element.js";
import type { TextParagraphStyleInput, TextRunStyleInput } from "./style.js";

/** Native update intent, distinct from locale-dependent formatting of the cached value. */
export type DynamicDateUpdateFormat = "application-default" | "month-day-year" | "weekday-month-day-year" | "day-month-year" | "month-day-year-long" | "automatic-month-day-year" | "day-short-month-year" | "month-year" | "short-month-year" | "month-day-year-time-12" | "month-day-year-time-12-seconds" | "time-24" | "time-24-seconds" | "time-12" | "time-12-seconds";
export type DateTimeCacheStyle = "short" | "medium" | "long" | "full";
/** Authored identity and semantics; no presentation-format IDs or cached display values. */
export type PresentationField =
  | { readonly id: string; readonly kind: "slide-number" }
  | { readonly id: string; readonly kind: "static-date" | "footer"; readonly text: string }
  | ({ readonly id: string; readonly kind: "dynamic-date"; readonly updateFormat?: DynamicDateUpdateFormat } & ({ readonly dateStyle: DateTimeCacheStyle; readonly timeStyle?: DateTimeCacheStyle } | { readonly dateStyle?: never; readonly timeStyle: DateTimeCacheStyle }));

/** The caller supplies the final ordered occurrence IDs, including any hidden pages it wishes to count. */
export interface FieldEvaluationContext {
  readonly pageOrder?: readonly string[];
  readonly pageId?: string;
  readonly firstPageNumber?: number;
  /** Milliseconds since the Unix epoch; no implicit clock is read. */
  readonly instantMs?: number;
  readonly locale?: string;
  readonly timeZone?: string;
}
/** Explicit numbering policy for document slides; custom shows can select/reorder unique slide IDs. */
export interface FieldSlideOrderOptions {
  readonly slideId: string;
  readonly firstPageNumber: number;
  readonly includeHidden: boolean;
  readonly order?: readonly string[];
}
/** A final slide sequence, including detached Layout results after pagination. */
export interface FieldSlideSequence { readonly slides: readonly { readonly id: string; readonly hidden: boolean }[] }
export interface FieldSlideOccurrence { readonly id: string; readonly slideId: string }
export interface FieldOccurrenceOptions {
  readonly occurrences: readonly FieldSlideOccurrence[];
  readonly occurrenceId: string;
  readonly firstPageNumber: number;
  readonly includeHidden: boolean;
}
export interface EvaluatedPresentationField { readonly id: string; readonly kind: PresentationField["kind"]; readonly value: string }
/** Field runs default to static text; native output requires an explicitly supported update format. */
export type FieldTextRun = ({ readonly text: string; readonly field?: never; readonly output?: never } | { readonly field: PresentationField; readonly text?: never; readonly output?: "static" | "native" }) & { readonly style?: TextRunStyleInput; readonly action?: ElementAction };
export interface FieldTextParagraph { readonly runs: readonly FieldTextRun[]; readonly style?: TextParagraphStyleInput }
export interface MaterializedFieldText {
  readonly content: TextParagraphInput[];
  readonly fields: readonly (EvaluatedPresentationField & { readonly paragraphIndex: number; readonly runIndex: number })[];
}
export type PresentationFieldErrorCode = "invalid-field" | "missing-field-context" | "invalid-field-context";
export class PresentationFieldError extends Error {
  constructor(readonly code: PresentationFieldErrorCode, message: string) { super(message); this.name = "PresentationFieldError"; }
}
