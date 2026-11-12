import type { ImportedChartValueSource } from "../types/import.js";
export interface ChartValueBinding extends ImportedChartValueSource {
  /** Reapply the opt-in sparse profile during candidate rebinding. */
  readonly sparseProfile?: boolean;
  /** Reapply the bubble profile during candidate rebinding. */
  readonly v3Profile?: boolean;
  readonly baselineValue: number;
  readonly cacheLexical?: string;
  readonly cacheNamespace?: string;
  readonly workbookLexical?: string;
  readonly workbookNamespace?: string;
}
export interface ChartValueProposal { readonly binding: ChartValueBinding; readonly value: number }
