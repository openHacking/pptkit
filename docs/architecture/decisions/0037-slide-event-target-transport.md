# ADR 0037: Slide-bound event target transport

- Status: accepted experimental whole-object authoring/IR/Layout identity slice; native output rejected
- Date: 2026-10-26

## Problem

Detached event arithmetic cannot know the real slide inventory or follow object IDs across authoring copies/IR/Layout. Storing caller-provided inventory snapshots would permit stale targets. A new timing field also risks silent loss in existing static/format outputs and stale references on table continuation pages.

## Decision

Core stores only a frozen authored `{stages}` snapshot in the optional slide eventTimeline field. Setter/init compile against the actual owned page/object forest using fixed integrated budgets, excluding shared layout objects. Normalization retains a versioned optional IR v2 wrapper containing both the complete authored source tree and detached compiled EventTimelinePlan; Layout validates owner/current targets, recompiles the source and compares the complete stored plan before retaining detached metadata. Group identities/topology are not discarded into intervals. No effect or native timing mapping is invented.

Expose setEventTimeline/getEventTimelinePlan and atomic replaceElement. Prepare replacement trees in the existing ElementStore, validate the candidate whole-object timing scope before ledger/collection commit, and preserve root identity. Timed target/source removal rejects until an explicit program change. The existing object registry additionally tracks registered root instances: released snapshots cannot replace or unregister later objects reusing the same nominal IDs. Whole-slide copies consume the allocator's complete descendant map and remap page/target/trigger references together; object-only copies leave the page program intact. Stage/time IDs retain their local namespaces.

Avoid an append-time regression for ordinary slides: no candidate forest copy/compile occurs without an attached program. Bound inventory collection before expanding large group arrays; bound ID length/text before sorting. Do not add a package/dependency, global plan registry or raw source IDs to Core.

A table split of a timed slide remains unapplied with a located diagnostic; a single-page transform keeps identity. Default PPTX output throws a typed immutable located rejection before assets/files. Only explicit eventTimelineMode: omit discards the normalized candidate timing, reports per-stage loss and permits static pagination/output. Static SVG warns without applying progress. Parser source timing remains source-only, and authored timeline mutations block preservation saves.

## Evidence and limits

[The public contract](../../api/slide-event-timeline.md) links Core transaction/copy/identity tests, multi-package IR/JSON/Layout/format/Parser tests, browser parity and installed consumers. Existing untimed documents omit the field and preserve previous output behavior; the experimental optional IR field and required slide methods are additive engine interfaces in the unreleased lockstep packages.

This does not close native mapping/playback gates. Source structural saves, cross-document transport, subtargets, independent tracks, object-only gate cloning, multi-page timing ownership and retained external-plan generations still require their own contracts/evidence. No ordinary group, static preview or omission output is counted as native animation.
