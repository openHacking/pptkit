# ADR 0013 — Finite animation time research contract

Status: experimental research decision for ADV-06.1.1. No public animation API or native effect admission.

## Ownership

Core will eventually hold authored semantic timing/targets. Format adapters will own OOXML timing and unknown-source retention. Consumers/render adapters will evaluate detached state; clocks, navigation, event subscriptions and temporary playback state must not be written into Core or static Layout. This ADR precedes a public model. The executable [reference experiment](../../../scripts/research/timing-model.mjs) lives under tooling research and is not imported by packages or published.

## Finite reference profile

Time is a nonnegative safe integer count of milliseconds. Negative delays, fractional/nonfinite values and arithmetic beyond the captured time budget reject rather than clamp. Parent delay precedes every child; a sequence begins each child after the previous child's complete delayed/repeated span, while parallel children share the parent's delayed origin. Parallel span is the maximum child span. Empty containers retain their own delay. Root delay contributes to total duration. Node IDs are unique in the whole tree; they identify research timing nodes, not drawing/paragraph targets.

Effect leaves specify duration, optional nonnegative delay, positive integer repeat and fill (`remove` by default or `hold`). Repetition is evaluated algebraically without expanding cycles into nodes/events. The active interval is `[start,end)`. Internal repeat boundaries begin the next iteration at progress zero. At the final boundary the phase is ended; hold exposes progress one, remove exposes no value. A zero-duration leaf is ended at its start, with no division by zero. This is abstract progress only, not visibility, opacity or hit-testing semantics.

Groups are sequence/parallel trees with children and optional delay. Group repetition, fractional/infinite repeats, auto-reverse, easing, explicit end/restart conditions, click/event activation and same-target composition are outside this reference profile and reject as unknown properties. Their design/acceptance remains ADV-06.1.2/ADV-07.7; absence of implementation does not silently reduce them to finite linear effects. An indefinite node cannot be assigned a fabricated finite end to unblock a sequence.

Absolute-time sampling has no mutable playhead. Forward/backward/repeated seeks produce the same result for the same captured plan/time. Start/end boundary records sort by time, then leaf preorder; start precedes end on one zero-duration leaf. The ordering describes the reference's abstract boundary list only; it does not authorize dispatch of clicks, media or native slideshow events. Those ordering rules need a separate event contract.

## Validation and bounded work

Only plain data records and the named properties are accepted. Record accessors and unknown nodes/properties reject. Compilation detects ancestor cycles and duplicate IDs, creates detached frozen primitive schedules and leaves caller state untouched. Limits are explicit: 10,000 visited nodes, depth 128 and 86,400,000 ms by default, with bounded overrides (node limit up to 100,000, depth up to 256 and safe integer time). Integer repeat is capped by the configured node budget even though repeats are not expanded; this is research policy subject to native/corpus evidence before a public contract. Sampling costs O(leaves); this is not a real-time performance or hostile-object sandbox guarantee.

## Evidence and remaining gates

Run `pnpm test:timing-research`. [Five test groups](../../../scripts/timing-research.test.mjs) verify independently hand-computed nested timestamps and same-time zero-duration ordering, exact start/repeat/end boundaries, random-access sampling, source snapshot/freeze behavior, unknown/malformed/overflow/cyclic input, limits/retry and 200 seeded trees compared with a separate relative-interval oracle.

This advances finite time/order semantics only. Public target/event identity, finite/infinite mixed sequences, fractional repeat/auto-reverse/fill/end/restart, easing, groups with repetition, target property composition, native mapping and actual playback remain unproven. ADV-06.1.1 and related parents stay incomplete until those contracts and inherited gates are addressed. No Appear/Fade/Disappear element animation support is claimed by this prototype; the existing page transition profile is a different capability.

Follow-up: [ADR0035](0035-public-finite-timeline-arithmetic.md) promotes the bounded numerical subset to a standalone Core utility with span nodes, independent repeat/ID budgets, typed errors and validated serialized schedules. The private effect-shaped prototype remains historical research; this does not admit animation targets/effects or native playback.
