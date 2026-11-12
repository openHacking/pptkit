# Experimental finite timeline arithmetic

Core exports `compileFiniteTimeline(root, limits?)`, `sampleFiniteTimeline(plan, timeMs, limits?)` and `FiniteTimelineError` through its root and browser entries. This is a detached, format-independent finite-time utility advancing ADV-06.1.1.1. It compiles abstract spans into intervals and samples their progress at an absolute time.

A span is not an animation effect or a drawing target. No animation field is added to presentation authoring, Canonical IR or Layout. This API does not play or export Appear/Fade/Disappear, dispatch events, navigate slides or interpret native timing. The existing [slide-transition profile](slide-transitions.md) remains separate. Public target/event/effect integration and its native mapping gates remain open.

## Input tree

`FiniteTimelineNode` is the readonly union of a span and sequence/parallel groups. Every node has a nonempty unique `id`, a `kind` and optional nonnegative integer `delayMs`. IDs belong to the timing tree; they are not element IDs or OOXML IDs.

| Kind | Required fields | Optional fields | Timing |
| --- | --- | --- | --- |
| `span` | `durationMs` | `repeat` (default 1), `fill` (`remove` by default or `hold`) | Own delay, then duration × repeat |
| `sequence` | `children` | — | Own delay, then each child's complete delayed/repeated span in input order |
| `parallel` | `children` | — | Own delay, then all children share the same origin; total is the maximum child span |

An empty group retains its own delay. Delays on parents and children accumulate. Group repetition is not admitted. The complete input is validated before any plan is returned: only plain enumerable data records and dense data arrays are accepted. Accessors, unknown keys/kinds, sparse arrays, custom array fields/iterators, duplicate IDs, shared nodes and ancestor cycles reject. Null-prototype records work. This validates ordinary data; it is not a sandbox for arbitrary JavaScript proxies or runtime side effects.

All times are nonnegative safe integer milliseconds. Negative zero normalizes to zero for time fields. Fractional/infinite repeats, `autoReverse` (including false), easing, triggers, targets, end/restart conditions, group repetition and unknown fields reject explicitly, without silently choosing finite/linear behavior. Negative delays and fractional/nonfinite durations reject. Integer repeats have their own budget and are never expanded into per-iteration nodes or events.

## Schedule and boundaries

`FiniteTimelinePlan` is a frozen serializable record:

- `schemaVersion: "pptkit-finite-timeline-v1"`, `nodeCount` (including groups), maximum root-inclusive `depth`, and total `durationMs`.
- Frozen `leaves` in depth-first input order: `id`, zero-based `order`, absolute `startMs`/`endMs`, per-iteration `durationMs`, `repeat` and `fill`.
- Frozen `boundaries`, two per leaf: `id`, `order`, `timeMs`, `type: "start" | "end"`. They sort by time, leaf preorder, then start before end for the same zero-duration leaf.

At one timestamp, an earlier leaf's end can precede a later leaf's start; zero-duration start precedes its own end. Boundary records describe abstract intervals. They are not slideshow callbacks and do not define click consumption or event dispatch.

Arithmetic is checked exactly with integer operations before applying the time ceiling. Overflow throws rather than rounds, clamps or invents an end. Compilation costs O(nodes + leaves log leaves) and retains O(nodes + leaves) bounded state. Repetition does not increase either count.

## Absolute-time sampling

`sampleFiniteTimeline` returns a frozen array of frozen discriminated `FiniteTimelineSample` records in leaf order:

| Absolute time | `phase` | `progress` | `iteration` |
| --- | --- | --- | --- |
| Before leaf start | `before` | null | null |
| `[startMs, endMs)` | `active` | Linear fraction in `[0,1)` | Zero-based repeat iteration |
| At/after final end | `ended` | 1 for hold; null for remove | repeat − 1 |

Each internal repeat boundary starts the next iteration at progress zero. A zero-duration leaf is ended at its start, including when repeated; there is no division by zero. `hold` concerns abstract progress only, without implying visibility, hit testing, opacity or a target property. Sample time may be any nonnegative safe integer, including later than the plan's end or configured plan-duration ceiling.

Sampling is pure random access: seeking backward, forward or to the same point yields the same result. There is no clock, mutable playhead, listener, cancellation handle or hidden global registry. It validates the complete detached schedule on every call, including future leaves and the full boundary list, then computes exact iteration/remainder and floating progress. Sampling costs O(leaves), with no sorting or repeat expansion.

JSON-transported plans work after validation. Unknown schemas, invalid records/counters/intervals, duplicate leaf IDs, inconsistent repeat ends, missing/duplicate/reordered/mismatched boundaries and accessors reject. Validation proves the stored schedule's coherence and bounded sampling work; it does not authenticate its source or reconstruct discarded group topology. Keep/recompile the authored tree for semantic edits. Plans compiled under larger limits need corresponding sampling limits too.

## Budgets and errors

| Limit | Default | Supported maximum |
| --- | --- | --- |
| `maxNodes` | 10,000 | 100,000 |
| `maxDepth` | 128 | 256 |
| `maxTimeMs` | 86,400,000 | Number.MAX_SAFE_INTEGER |
| `maxRepeat` | 1,000,000 | Number.MAX_SAFE_INTEGER |
| `maxIdLength` | 256 UTF-16 code units | 65,536 |
| `maxTotalIdLength` | 1,000,000 UTF-16 code units | 16,000,000 |

Node/depth/repeat/ID limits are positive integers; maxTimeMs may be zero. Compilation counts every node ID, including groups, toward total ID characters. Schedule sampling counts transported leaf IDs; other group IDs are absent from the schedule. A span with a large repeat still counts as one node. Unknown options and values above the supported caps reject.

`FiniteTimelineError` exposes readonly `code` and JSON-pointer-style `path` (empty for root), plus a human message. Codes are `invalid-timeline-input`, `unsupported-timeline-feature`, `timeline-limit-exceeded`, `timeline-time-overflow` and `invalid-timeline-plan`. Input and plan errors publish no partial result, mutate no caller data and poison no retry state. This utility owns no resources to dispose.

## Checked example

<!-- doc-test: docs/examples/finite-timeline.ts -->
```ts
import { compileFiniteTimeline, sampleFiniteTimeline, type FiniteTimelineNode } from "@pptkit/core";

export function finiteTimingExample(timeMs: number) {
  const timing: FiniteTimelineNode = {
    kind: "sequence", id: "sequence", delayMs: 50,
    children: [
      { kind: "span", id: "first", durationMs: 100, repeat: 2, fill: "hold" },
      { kind: "parallel", id: "parallel", children: [
        { kind: "span", id: "second", durationMs: 80 },
        { kind: "span", id: "instant", delayMs: 20, durationMs: 0, fill: "hold" },
      ] },
    ],
  };
  const plan = compileFiniteTimeline(timing);
  // Keep the authored tree for editing. A JSON schedule can be transported
  // and sampled, but it does not encode drawing targets or native effects.
  const transported = JSON.parse(JSON.stringify(plan)) as typeof plan;
  const samples = sampleFiniteTimeline(transported, timeMs);
  return { plan, samples };
}
```

At 250 ms the first span is ended/held, the second starts at progress zero, and the delayed zero span is still before. Total duration is 330 ms. Nothing in this example creates native animation.

## Verification and next gates

[Core tests](../../packages/core/test/timeline.test.mjs) cover independently calculated nested timestamps, both fills, repeat/zero boundaries, shuffled seeks, exact arithmetic near the safe-integer ceiling, all quotas and retries, 10,000-node width/256 depth, malformed/unknown/cyclic/accessor inputs, frozen snapshots and corrupted JSON schedules. Four hundred seeded trees compare against a separate relative-interval and expanded-per-millisecond sample oracle.

[Browser tests](../../scripts/browser-bundles.test.mjs) compare Node, real Chromium ESM/global and JSON transport, including large integer repeats and matching error codes with networking blocked. [Installed Node](../../tests/consumers/node.mjs) and [type consumers](../../tests/consumers/types.ts) check package root/browser exports, readonly/discriminated output, supported node kinds and unsupported effect/target fields.

[ADR0035](../architecture/decisions/0035-public-finite-timeline-arithmetic.md) records this numerical boundary. ADV-06.1.1's fractional/infinite/autoReverse/end/restart semantics and parent gates remain open. ADV-06.1.2.1 now has the separate [foreground event utility](event-timeline.md), while target/IR integration remains pending; native effects still require the separate mapping and actual-playback evidence. These tests do not close ADV-06 or ADV-07.

The separate [foreground event utility](event-timeline.md) now composes finite plans with explicit page/visit logs and current inventories. It defines arithmetic replay only and does not change finite spans into native effects or install a playback clock.
