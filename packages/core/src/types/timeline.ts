/** Abstract time only; a span does not define an effect or a drawing target. */
export interface FiniteTimelineSpan {
  readonly kind: "span";
  readonly id: string;
  readonly delayMs?: number;
  readonly durationMs: number;
  readonly repeat?: number;
  readonly fill?: "remove" | "hold";
}
export interface FiniteTimelineGroup {
  readonly kind: "sequence" | "parallel";
  readonly id: string;
  readonly delayMs?: number;
  readonly children: readonly FiniteTimelineNode[];
}
export type FiniteTimelineNode = FiniteTimelineSpan | FiniteTimelineGroup;
export interface FiniteTimelineLimits {
  readonly maxNodes?: number;
  readonly maxDepth?: number;
  readonly maxTimeMs?: number;
  readonly maxRepeat?: number;
  readonly maxIdLength?: number;
  readonly maxTotalIdLength?: number;
}
export interface FiniteTimelineLeaf {
  readonly id: string;
  readonly order: number;
  readonly startMs: number;
  readonly endMs: number;
  readonly durationMs: number;
  readonly repeat: number;
  readonly fill: "remove" | "hold";
}
export interface FiniteTimelineBoundary {
  readonly id: string;
  readonly order: number;
  readonly timeMs: number;
  readonly type: "start" | "end";
}
/** Serializable schedule, not authored effects or mutable playback state. */
export interface FiniteTimelinePlan {
  readonly schemaVersion: "pptkit-finite-timeline-v1";
  readonly nodeCount: number;
  readonly depth: number;
  readonly durationMs: number;
  readonly leaves: readonly FiniteTimelineLeaf[];
  readonly boundaries: readonly FiniteTimelineBoundary[];
}
export type FiniteTimelineSample =
  | { readonly id: string; readonly phase: "before"; readonly progress: null; readonly iteration: null }
  | { readonly id: string; readonly phase: "active"; readonly progress: number; readonly iteration: number }
  | { readonly id: string; readonly phase: "ended"; readonly progress: 1 | null; readonly iteration: number };
export type FiniteTimelineErrorCode = "invalid-timeline-input" | "unsupported-timeline-feature" | "timeline-limit-exceeded" | "timeline-time-overflow" | "invalid-timeline-plan";
export class FiniteTimelineError extends Error {
  constructor(readonly code: FiniteTimelineErrorCode, message: string, readonly path: string) {
    super(message);
    this.name = "FiniteTimelineError";
  }
}
