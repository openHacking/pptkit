import type { TextStylePresetResetProperty } from "./style.js";
import type { JsonValue } from "./presentation.js";

export interface LayoutRebindRequest {
  readonly slideId: string;
  readonly layoutId: string;
  /** Explicit logical element ID → target placeholder key; null detaches. */
  readonly bindings?: Readonly<Record<string, string | null>>;
  /** Explicit old inherited object ID → target inherited object ID for connectors. */
  readonly references?: Readonly<Record<string, string>>;
}
export interface LayoutRebindOptions {
  /** Preserve effective local foreground values, or retain declarations under the target theme. */
  readonly localAppearance: "preserve" | "adopt";
  readonly unmatched?: "block" | "detach";
  /** Explicitly resume target geometry/text/background inheritance; adopt mode only. */
  readonly reset?: readonly ("geometry" | "text" | "background" | `text.${TextStylePresetResetProperty}`)[];
}
export interface LayoutRebindFinding {
  readonly code: string;
  readonly severity: "error" | "warning" | "info";
  readonly message: string;
  readonly slideId: string;
  readonly elementId?: string;
  readonly path?: string;
}
export interface LayoutRebindMatch {
  readonly slideId: string;
  readonly elementId: string;
  readonly sourceKey?: string;
  readonly targetKey?: string;
  readonly sourceIndex?: number;
  readonly targetIndex?: number;
  readonly method: "key" | "index" | "kind" | "explicit" | "ordinary" | "unmatched";
  readonly candidates: readonly string[];
}
export interface LayoutRebindPropertyChange {
  readonly slideId: string;
  readonly path: string;
  readonly before?: JsonValue;
  readonly after?: JsonValue;
}
export interface LayoutRebindPlan {
  readonly status: "ready" | "needs-review";
  /** Exact serialized authoring value snapshot, not a security token or a digest. */
  readonly sourceSnapshot: string;
  readonly requests: readonly LayoutRebindRequest[];
  readonly options: LayoutRebindOptions;
  readonly matches: readonly LayoutRebindMatch[];
  readonly findings: readonly LayoutRebindFinding[];
  readonly changes: readonly LayoutRebindPropertyChange[];
  readonly changedSlideIds: readonly string[];
}
export type LayoutRebindApplyResult =
  | { readonly status: "applied"; readonly changedSlideIds: readonly string[] }
  | { readonly status: "needs-review"; readonly plan: LayoutRebindPlan };
export class LayoutRebindError extends Error {
  constructor(readonly code: "stale-plan" | "invalid-plan", message: string) { super(message); this.name = "LayoutRebindError"; }
}
