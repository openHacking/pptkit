import type { FiniteTimelineLimits, FiniteTimelineNode, FiniteTimelinePlan, FiniteTimelineSample } from "./timeline.js";
/** Caller-supplied whole-object identity; not a raw format ID or subtarget. */
export interface EventTimelineTarget { readonly pageId: string; readonly elementId: string }
export type EventTimelineTrigger =
  | { readonly type: "enter" | "with-previous" | "after-previous" | "page-click" }
  | { readonly type: "object-click"; readonly elementId: string };
export interface EventTimelineStageInput {
  readonly id: string;
  readonly target: EventTimelineTarget;
  readonly trigger: EventTimelineTrigger;
  readonly timing: FiniteTimelineNode;
}
export interface SlideEventTimelineInput { readonly stages: readonly EventTimelineStageInput[] }
/** Retains authored group identities as well as deterministic compiled intervals. */
export interface NormalizedSlideEventTimeline {
  readonly schemaVersion: "pptkit-slide-event-timeline-v1";
  readonly source: SlideEventTimelineInput;
  readonly plan: EventTimelinePlan;
}
export interface EventTimelineInput {
  readonly pageId: string;
  readonly elementIds: readonly string[];
  readonly stages: readonly EventTimelineStageInput[];
}
export interface EventTimelineStage {
  readonly id: string;
  readonly target: EventTimelineTarget;
  readonly trigger: EventTimelineTrigger;
  readonly timing: FiniteTimelinePlan;
}
export interface EventTimelinePlan {
  readonly schemaVersion: "pptkit-event-timeline-v1";
  readonly pageId: string;
  readonly elementIds: readonly string[];
  readonly timingNodeCount: number;
  readonly stages: readonly EventTimelineStage[];
}
export interface EventTimelineLimits extends FiniteTimelineLimits {
  readonly maxElements?: number;
  readonly maxStages?: number;
  readonly maxEvents?: number;
}
interface EventIdentity {
  readonly id: string;
  readonly pageId: string;
  readonly visitId: string;
  readonly timeMs: number;
}
export type EventTimelineEvent =
  | (EventIdentity & { readonly type: "page-click" })
  | (EventIdentity & { readonly type: "object-click"; readonly elementId: string });
export interface EventTimelineReplayRequest {
  readonly visitId: string;
  readonly timeMs: number;
  /** Current inventory is mandatory to reject stale/deleted targets. */
  readonly elementIds: readonly string[];
  readonly events: readonly EventTimelineEvent[];
}
export type EventTimelineIgnoreReason = "other-page" | "other-visit" | "unknown-source" | "before-ready" | "wrong-trigger" | "wrong-source" | "no-gate";
interface ObservationIdentity {
  readonly index: number;
  readonly eventId: string;
  readonly pageId: string;
  readonly visitId: string;
  readonly timeMs: number;
}
export type EventTimelineObservation =
  | (ObservationIdentity & { readonly status: "consumed"; readonly stageId: string })
  | (ObservationIdentity & { readonly status: "ignored"; readonly reason: EventTimelineIgnoreReason })
  | (ObservationIdentity & { readonly status: "future" });
export interface EventTimelineStageSample {
  readonly id: string;
  readonly target: EventTimelineTarget;
  readonly startMs: number;
  readonly endMs: number;
  readonly eventId: string | null;
  readonly eventIndex: number | null;
  readonly samples: readonly FiniteTimelineSample[];
}
export interface EventTimelineReplay {
  readonly schemaVersion: "pptkit-event-timeline-replay-v1";
  readonly pageId: string;
  readonly visitId: string;
  readonly timeMs: number;
  readonly pendingStageId: string | null;
  readonly stages: readonly EventTimelineStageSample[];
  readonly events: readonly EventTimelineObservation[];
}
export type EventTimelineErrorCode = "invalid-event-timeline-input" | "unsupported-event-timeline-feature" | "event-timeline-limit-exceeded" | "event-timeline-time-overflow" | "invalid-event-timeline-plan";
export class EventTimelineError extends Error {
  constructor(readonly code: EventTimelineErrorCode, message: string, readonly path: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "EventTimelineError";
  }
}
