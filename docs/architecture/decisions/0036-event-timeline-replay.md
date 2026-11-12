# ADR 0036: Detached foreground event timeline replay

- Status: accepted experimental event arithmetic utility; not authored/native animation admission
- Date: 2026-10-26

## Problem

ADR0014's private event prototype leaves re-entry identity, stale target checks, complete event observations and public transport unimplemented. A native/browser callback order must not silently become timeline semantics. Pure time arithmetic also must not acquire listeners or mutable playback state.

## Decision

Core provides compileEventTimeline/replayEventTimeline over explicit whole-object inventories, foreground stages and ordered event logs. Reuse the public finite compiler/validator/sampler through private accounting views, rather than import research code or duplicate finite arithmetic. No authoring/Canonical IR/Layout/format field or effect is added.

Specify immediate-predecessor with/after and readiness, ignore early clicks without queuing, consume each log row once, and preserve explicit equal-time array order. Scope event identity by page/visit/event ID. Require a caller visit ID and current element inventory on every replay, reject dangling targets/trigger sources including future stages, and isolate old visits. Replay returns frozen activations, finite samples and complete consumed/ignored/future observations, never dispatches events or mutates input.

Pause/seek/replay are defined over caller-owned active time/log snapshots, not wall-clock callbacks. A new visit resets the caller's epoch/log; old-visit events cannot activate it. Re-evaluating a consumption report does not authorize redispatch. UI paused-click policy, clock conversion and controller state remain ADV-08 work.

Validate complete transported programs and logs, including future/wrong-scope rows, unknown fields, duplicate identities and malformed finite schedules. Bound stages/inventories/events and aggregate timing nodes/identifier text; exact absolute stage ends fail atomically on overflow. Keep one log cursor and linear schedule validation. No global plan registry, package, dependency or source XML enters Core.

## Evidence and open gates

[The contract](../../api/event-timeline.md) links explicit hand-calculated state tables, malformed/transport/deletion/budget/zero/nested tests, 300 generated event-first oracle programs, 1,000 same-time zero gates, Node/Chromium ESM/global parity and installed runtime/type consumers.

This is an explicit foreground policy, not evidence of native click-to-complete or independent trigger sequences. Whole-object IDs are caller-bound numerical inputs; their authoring/IR/format/copy/delete integration remains ADV-06.1.2.2. Native mapping/playback, effects, subtargets, media, retrigger/restart and full ADV-06/08 gates remain open. Source plan validation does not authenticate discarded topology, and ordinary data validation does not sandbox JavaScript Proxies.
