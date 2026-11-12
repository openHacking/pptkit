# Experimental foreground event timeline replay

Core exports `compileEventTimeline(input, limits?)`, `replayEventTimeline(plan, request, limits?)` and `EventTimelineError` through its root/browser entries. The functions compile a detached foreground gate sequence and deterministically recompute activations/progress from an explicit ordered event log. They extend [finite timeline arithmetic](finite-timeline.md) for ADV-06.1.2.1.

They do not dispatch events, install listeners, choose wall clocks, create animation effects, mutate documents/Layout or serialize native timing. Whole-object identity is caller-supplied; authoring/IR/format transfer and structural copy/delete remapping are separate tasks. This foreground policy is not claimed to reproduce PowerPoint's click-to-complete or independent trigger-track behavior.

## Program and identity

`EventTimelineInput` has `pageId`, `elementIds` and ordered `stages`. Each stage supplies a unique `id`, a whole-object `target: { pageId, elementId }`, a `trigger` and an authored `FiniteTimelineNode` tree. Targets must refer to the program's page and declared inventory. Object-click source IDs must be in that inventory too. Paragraph/series/SmartArt subtarget fields and unknown trigger kinds reject.

Stage IDs are unique across the sequence. Finite timing node IDs are unique within their stage tree and may repeat across stages; the pair `(stage ID, leaf ID)` identifies a sampled time span. These IDs are not raw OOXML IDs or a new shared animation identity registry.

| Trigger | Activation time |
| --- | --- |
| `enter` | Zero, allowed only on the first stage |
| `with-previous` | Immediate predecessor's activation time |
| `after-previous` | Immediate predecessor's complete finite end |
| `page-click` | First eligible matching page click |
| `object-click` with `elementId` | First eligible matching object click from that source |

With/after require a predecessor. The first stage may be enter or click-gated. Gate readiness is the immediate predecessor's end (zero for the first gate). A click earlier than readiness is ignored, never queued. Exact equality is eligible. The immediate predecessor rule does not wait for a longer, overlapping earlier sibling. Stage activation is distinct from its timing tree's delayed leaf start; the complete delayed/repeated tree span determines stage end.

`EventTimelinePlan` freezes `schemaVersion: "pptkit-event-timeline-v1"`, `pageId`, copied `elementIds`, ordered stage snapshots and aggregate `timingNodeCount` including containers. Each compiled stage retains the frozen target/trigger and finite schedule. A plan adds no property to a presentation or Canonical IR.

## Request, visit and current inventory

Every `EventTimelineReplayRequest` explicitly supplies:

- `visitId`: caller-owned page-visit identity.
- `timeMs`: current nonnegative integer active time in that visit.
- `elementIds`: **current** whole-object inventory, required on every call.
- `events`: the complete ordered log to evaluate.

All declared stage targets and trigger sources must still exist in the current inventory, including future stages. Deletion rejects before output; the evaluator never retargets by name, order or visible content. Removing an unreferenced object or adding an unrelated object is allowed. Event sources absent from the current inventory are ignored as unknown-source when they do not already invalidate a stage binding.

Inventory checks prove nominal ID presence. The caller must not recycle an object ID for a different logical object while retaining its program/log; use a fresh ID or explicitly recompile/reset the epoch. This utility cannot infer delete-and-recreate generations from the same supplied strings.

A fresh page entry/re-entry uses a fresh `visitId`, resets the caller's active clock/log, and can reuse a still-valid program. Old-visit events cannot activate it even if accidentally retained. Automatic enter stages start at zero for the selected visit. Passing elapsed time 50 for a new visit means evaluating 50 active milliseconds after that visit's start; the function does not infer navigation or reset wall time.

## Ordered log and consumption

An event has `id`, `pageId`, `visitId`, integer `timeMs` and `type: "page-click" | "object-click"`; object clicks require `elementId`. Page clicks reject object-source fields, including explicitly undefined ones. The tuple `(pageId, visitId, id)` must be unique in the complete log. Equal IDs in different scopes are allowed, but a duplicate physical-event identity within one visit rejects rather than triggering twice.

Timestamps must be nondecreasing across the supplied log. Every retained row uses the caller's common active-time basis for that request; array order defines equal-time order. Logs are never silently sorted. Wrong-scope entries still undergo complete type/time/field/identity validation, as do future entries.

A monotonic cursor visits each log row at most once while advancing gates. Each eligible click consumes exactly one row for one stage. Separate equal-time events with distinct identities can start separate zero-duration gates. Skipped early/wrong-type/wrong-source events are not reused by later gates.

## Replay output and observations

`EventTimelineReplay` freezes `schemaVersion: "pptkit-event-timeline-replay-v1"`, page/visit/time, `stages`, `events` and `pendingStageId` (null once all stages have started). A pending ID can also name an automatic stage whose start is later than the requested time. It does not mean that all earlier active spans have ended.

Started stages expose `id`, copied target, activation `startMs`, complete `endMs`, consumed `eventId`/`eventIndex` (both null for automatic triggers), and immutable finite span `samples`. They remain in program order, including ended stages. Hold/remove and zero/repeat boundaries retain the finite utility's semantics.

Every log row has a frozen observation carrying its index, event ID, page/visit and time:

| Status | Additional information |
| --- | --- |
| `consumed` | `stageId` |
| `ignored` | `reason`: other-page, other-visit, unknown-source, before-ready, wrong-trigger, wrong-source or no-gate |
| `future` | Valid current-scope/source event later than the sample time |

Scope classification precedes time classification: other-page wins over other-visit, then unknown-source, then future. A matching past event initially has no-gate; a scanned gate can classify it as before-ready, wrong-trigger, wrong-source or consumed. Unneeded past clicks retain no-gate. This is a report of deterministic evaluation, not instructions to dispatch browser or application actions.

## Seek, pause and replay

Identical plans, current inventories, ordered logs, visits and active times produce identical results. Seeking backward removes future activations/consumptions from the result; seeking forward against the same log recomputes the same ones. Calling replay repeatedly never fires a click again because it fires no click at all.

For pause, the caller holds active time **and log snapshot** fixed, yielding identical output. Appending a click at the current active timestamp changes the input and may activate a gate. Whether a paused UI captures/ignores clicks, how wall time is converted to active time, and event dispatch deduplication belong to the later playback-clock/controller contract. No hidden listeners or timers enforce that policy here.

## Budgets, validation and errors

All [finite limits](finite-timeline.md#budgets-and-errors) apply. maxNodes is shared by every stage's timing tree, counting containers as well as leaves. maxDepth/repeat/time constrain each tree, with maxTimeMs additionally bounding request/log timestamps and activated absolute stage ends. Dynamic end overflow or budget failure publishes no partial replay. A sample later than the configured horizon requires a larger maxTimeMs; no clamping occurs.

| Additional limit | Default | Supported maximum |
| --- | --- | --- |
| `maxElements` | 10,000 per inventory | 100,000 |
| `maxStages` | 1,000 | 10,000 |
| `maxEvents` | 10,000 | 100,000 |

Compile counts all occurrence IDs in page/inventory/stage/target/trigger metadata plus every authored timing node ID toward maxTotalIdLength. Replay jointly counts its transported plan metadata/leaf IDs, request visit/current inventory and full log IDs; discarded source group IDs are absent from transported schedules. Consequently a request can need more identifier budget than compilation of the same program. Unknown options and invalid/over-cap limits reject. Repeats are algebraic, and total sample leaves remain bounded by the shared node budget.

Only plain enumerable data records and dense data arrays are accepted; descriptor checks reject accessors/custom iterator fields without executing them. Serialized plans undergo complete header/target/trigger and finite schedule validation, including future stages and aggregate counters. Schedule validation proves coherent stored intervals, not original source topology or authenticity. Ordinary-data validation is not a JavaScript Proxy sandbox.

`EventTimelineError` has readonly `code` and JSON-pointer-style `path`, with the underlying finite error retained as `cause` when applicable. Codes are invalid-event-timeline-input, unsupported-event-timeline-feature, event-timeline-limit-exceeded, event-timeline-time-overflow and invalid-event-timeline-plan. Errors leave inputs unchanged and retries independent. Replay costs O(elements + stages + total timing data + events), with no per-gate log rescan, repeat expansion, I/O or shared mutable registry.

## Checked example

<!-- doc-test: docs/examples/event-timeline.ts -->
```ts
import { compileEventTimeline, replayEventTimeline, type EventTimelineEvent } from "@pptkit/core";

export function replayClickExample(timeMs: number, visitId = "visit-1") {
  const plan = compileEventTimeline({
    pageId: "page", elementIds: ["shape", "button"],
    stages: [
      { id: "initial", target: { pageId: "page", elementId: "shape" },
        trigger: { type: "enter" }, timing: { kind: "span", id: "time", durationMs: 100 } },
      { id: "clicked", target: { pageId: "page", elementId: "shape" },
        trigger: { type: "object-click", elementId: "button" },
        timing: { kind: "span", id: "time", durationMs: 200, fill: "hold" } },
    ],
  });
  const events: readonly EventTimelineEvent[] = [
    { id: "early", pageId: "page", visitId: "visit-1", type: "object-click", elementId: "button", timeMs: 50 },
    { id: "ready", pageId: "page", visitId: "visit-1", type: "object-click", elementId: "button", timeMs: 100 },
  ];
  // Supply the current inventory on every evaluation. Removing button or
  // shape rejects the stale program until the caller explicitly recompiles.
  return replayEventTimeline(plan, { visitId, timeMs, elementIds: ["shape", "button"], events });
}
```

At 150 ms, early is ignored before-ready, ready is consumed once for clicked, and that span has progress 0.25. Changing the selected visit to visit-2 makes both clicks other-visit. A fresh entry at zero starts only initial. These are progress observations, not a native effect playback claim.

## Evidence and remaining gates

[Twelve Core groups](../../packages/core/test/event-timeline.test.mjs) cover hand-computed with/after/delay and immediate-predecessor semantics, exact scope/reason state tables, same-time one-time consumption, repeated seeks/pause snapshots/new visits, deleted targets/sources, nested finite trees/fills, malformed/future/accessor data, serialized/corrupt plans, shared node/text/event budgets and safe-integer overflow. Three hundred generated programs/logs compare with a separate event-first state machine; 1,000 zero-duration gates and equal-time events verify monotonic consumption and quota edges.

[Real browser tests](../../scripts/browser-bundles.test.mjs) compare Node/ESM/global/JSON transport, visits and frozen output, with network blocked and matching rejection codes. Installed [runtime](../../tests/consumers/node.mjs) and [types](../../tests/consumers/types.ts) verify root/browser APIs, mandatory current inventories, source discriminants and readonly observations.

[ADR0036](../architecture/decisions/0036-event-timeline-replay.md) keeps this utility separate from authored effect/target/native contracts. ADV-06.1.2.2 still owns authoring/IR/Layout and copy/delete transport, ADV-06.1.2.3 native mapping, and ADV-08 clock/controller behavior. Independent trigger tracks, native click-to-complete, retrigger/restart, subtargets, media and actual playback remain unimplemented. Parent tasks remain open.

The separate [slide-bound profile](slide-event-timeline.md) now derives inventories from real owned Core objects and transports compiled plans through IR/Layout. Its copy/delete/replace guards do not grant native source mapping or effect playback.
