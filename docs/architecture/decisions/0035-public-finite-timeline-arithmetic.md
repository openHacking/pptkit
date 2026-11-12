# ADR 0035: Public detached finite timeline arithmetic

- Status: accepted experimental numerical utility; not animation authoring/native admission
- Date: 2026-10-26

## Problem

ADR0013's private timing experiment establishes finite ordering but has no typed public consumer. Reimplementing time arithmetic in future event, preview and format work would introduce incompatible zero/repeat/fill/seek boundaries. Promoting effect/target semantics before the native mapping gates would instead overstate what the experiment proves.

## Decision

Core owns two pure public utilities, compileFiniteTimeline/sampleFiniteTimeline, with readonly span/sequence/parallel inputs, a versioned detached schedule, discriminated progress samples and located typed errors. Spans define abstract time only. No target, effect, event, authoring/IR/Layout field or playback state is added. This is the finite arithmetic slice ADV-06.1.1.1; broader public animation integration still follows ADV-06/06.1.2.3's mapping prerequisites.

Keep integer milliseconds, complete parent/child delays, finite positive integer repeat and remove/hold semantics from ADR0013. Reject fractional/infinite repeats, autoReverse, easing, group repeats, end/restart, triggers/targets and unknown fields explicitly. Check exact arithmetic before the caller's time budget. Unlike the prototype, bound repeat independently of node count because evaluation is algebraic rather than expanded.

Compile immutable intervals and deterministic start/end boundaries. Sample absolute time without a clock or shared global state. Accept serialized schedules only after complete data/identity/interval/boundary validation; validate sorted boundaries linearly rather than sorting on each sample. Source group topology is not reconstructed/authenticated by schedule validation. Keep the source tree for semantic edits.

Use ordinary plain data/descriptor checks to reject accessors, sparse/custom arrays and unknown properties without running getters/iterators. Limit nodes, depth, time, repeat and identifier characters; captured schedules do not share caller containers. This is not a JavaScript proxy sandbox. No package, dependency or reverse format-layer edge is introduced.

## Evidence and remaining work

The [API contract](../../api/finite-timeline.md) links Core interval/boundary/malformed/quota/transport tests, 400 independent seeded interval/per-millisecond oracle checks, exact large-repeat arithmetic, 10,000-node and 256-depth edges, real Node/ESM/global browser parity and installed runtime/declaration consumers.

The original research modules remain private historical evidence, not production imports. Public event logs/re-entry, target identity and copy/delete reconciliation, authoring/IR transport, native source preservation/mapping, effects and actual playback remain distinct unclosed tasks. No native animation support or full ADV-06.1.1/ADV-06 completion is claimed.
