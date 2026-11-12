# ADR 0014 — Stable targets and candidate event gates

Status: research candidate for ADV-06.1.2. Public model and native mapping are not admitted.

## Target identities

The executable [research binder](../../../scripts/research/event-timing.mjs) binds a detached list of stable object IDs in one explicit page to abstract timing stages. Every target is the tuple `{ pageId, elementId }`, not a name, page position, array index or cached drawing order. Group descendants can be addressed by their existing object identities when included in the caller's inventory. Duplicates, missing IDs and cross-page references reject before a plan exists. Targets are copied/frozen; editing an input object or returning a sample cannot retarget the plan.

Paragraph/word/glyph targets require stable semantic identities and source-range reconciliation. Chart series/category and SmartArt data nodes require their separately admitted identities, not drawing-cache indices. This experiment rejects subtarget keys rather than guessing an object. It does not resolve raw OOXML IDs or acquire structural-copy/delete/remapping permission. Those remain ADV-07.3/07.5/07.6/07.10 gates.

## Candidate gate semantics

One research plan represents one ordered foreground sequence for a single page. Each stage holds a finite ADR0013 timing tree and an explicit trigger: enter, with-previous, after-previous, page-click or object-click. Enter is allowed only for the first stage; with/after require a predecessor. With-previous shares the immediate predecessor's activation time. After-previous starts at its finite end. Stage delay remains inside its timing tree. With-previous does not mean after delay, and after-previous refers to the immediate predecessor, not the longest active sibling.

Page/object clicks are distinct events. An object click must match the exact source object ID; neither wrong-page nor wrong-object events activate the gate. The candidate policy ignores clicks earlier than the predecessor's finite end and consumes the first matching eligible click; it never silently queues an early click. Equality at the end is eligible. A consumed click cannot activate two later gates; separate equal-time events retain input log order. Each gate activates at most once per replay. This conservative foreground policy is **not claimed to reproduce PowerPoint click-to-complete behavior** or independent object-trigger sequences. Such behavior requires native evidence and a separately chosen contract.

The log is explicit, ordered and captured by the caller. Invalid negative/noninteger/nonfinite timestamps, unsupported event kinds/fields and descending times reject, including invalid future entries. Replay validates the complete log, then recomputes starts/samples at absolute time. It dispatches no events and performs no navigation. Re-entry means a fresh replay log; mutable playheads/listeners/timers/media are absent. Late logs, independent tracks, retrigger/restart, cancellation, concurrent animation/link actions, media events and page exit are pending designs.

## Bounds and evidence

The private prototype permits at most 10,000 object IDs, 1,000 stages, 10,000 combined timing nodes and 10,000 events. Container nodes count against the shared timing budget even when they contain no effects. Each finite tree retains depth/time limits. Trigger scans use one monotonic event cursor, avoiding rescanning the whole log per gate; timing samples use the captured finite schedules. Binding and replay do not mutate input or persist application state. The prototype is tooling research over ordinary data, not a sandbox for hostile runtime objects.

[Five event research groups](../../../scripts/event-timing-research.test.mjs), run with `pnpm test:timing-research`, verify independently calculated activation/end times, with/after alignment, early/wrong-page/wrong-source rejection, exact-end and equal-time consumption, zero-duration stages, random-access replay, target snapshots, result/input mutation independence, dangling/duplicate/unsupported targets/triggers and malformed logs. A cross-stage 12,002-container counterexample verifies the aggregate node budget. The existing five finite timing groups retain their seeded/interval evidence.

This advances whole-object identity and foreground-event research only. Public Core animation semantics, independent trigger tracks, text/chart/SmartArt subtargets, copy/delete mapping, native timing serialization/import, actual on-click playback, link/media coexistence and full ADV-06.1.2 remain incomplete. Research success does not unlock Appear/Fade/Disappear element effects without ADV-06.2.1's native experiment and the corresponding task gates.

Follow-up: [ADR0036](0036-event-timeline-replay.md) makes the detached foreground numerical policy public, adding explicit visits/current inventories, complete observations, aggregate bounds and transported-plan validation. The original prototype remains historical research; authored/IR/native target/effect admission and actual playback remain open.
