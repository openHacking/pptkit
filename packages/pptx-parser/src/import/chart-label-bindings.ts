import type { ChartLabelClosure } from "./chart-label-closure.js";

/** Private identity for associating a source closure with a Core chart string. */
export type ChartLabelBinding = {
  readonly slideId: string;
  readonly elementId: string;
  readonly baselineValue: string;
  readonly closure: ChartLabelClosure;
} & ({ readonly categoryIndex: number; readonly seriesIndex?: never } | { readonly seriesIndex: number; readonly categoryIndex?: never });
export interface ChartLabelEditProposal { readonly binding: ChartLabelBinding; readonly value: string }
