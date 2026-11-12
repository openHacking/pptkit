# ADR0058 — Reviewed value plans and atomic layout rebinding

- Date: 2026-10-27
- Status: accepted for the bounded authoring foundation; complete template/reset/native/source gates remain open

## Problem

Slides had immutable creation-time layout selection. Replacing a shared layout definition did not switch selected pages to a different template. A mutable layout setter alone would leave dangling placeholders/connectors, ambiguous matching and partially changed multi-page selections. Copying inherited artwork onto slides would break native ownership.

## Decision

Core owns a pure `prepareLayoutRebind(PresentationData, requests, options)` transformation and the factory document owns `applyLayoutRebind(plan)`. The legacy document/slide interfaces gain no required mutation method. Read-only authoring data aliases permit normalizing/validating candidates without fake document classes, method stubs or Proxies. Existing documents remain structurally compatible.

Plans contain frozen captured requests/options, matches, findings, effective property changes, changed page IDs and an exact serialized authoring snapshot. The snapshot is explicitly a value string, not a cryptographic digest or an opaque permission token. Applying a plan reads data descriptors, captures current values, checks freshness before normalization and recomputes decisions. Caller-provided status/reports are not trusted. Changes to any authoring data conservatively stale the plan. Discarding a plan cancels without state or allocator changes.

Matching uses a logical key with the same role, then a unique same-role candidate. Multiple candidates, conflicting roles and missing candidates need review; explicit mappings can choose a key, change role with a finding or detach with old resolved geometry. Ordinary unbound objects are not assigned by names or traversal position. Core has no native placeholder index identity: native indices remain Parser provenance. This foundation does not guess a third identity from array order or claim complete index-based native template rebinding.

`localAppearance: preserve` materializes effective normalized foreground properties, old theme colors/fonts/scripts and background. It retains IDs, text/actions, image assets and local connector references. Regional table rules become effective cell declarations with a finding. Target master/layout artwork remains shared and changes with the template; it is not copied to simulate preservation. This preserves authoring values, not every possible host measurement/composition result. `adopt` retains local declarations under the target theme; explicit geometry/text/background reset resumes supported inheritance. Text reset only concerns matched text/shape text, retains strings/actions and never invents content or fitting. Core estimates possible overflow and reports changed font families/target shrink policy without reading host fonts or changing font size to fit.

Inherited connector IDs require explicit target maps when owners disappear. Slide-level references must stay in the inherited root coordinate scope. Preserving a connector while changing its inherited target is blocked rather than duplicating Layout's anchor/transform solver in Core. Existing local identities survive. Complete resolved geometric preservation and arbitrary transformed/grouped inherited references remain separate work.

The document stages all changed roots through one `ElementStore.replaceMany` call. Candidate slides are plain data validated together with every shared definition/resource. Only after validation do global IDs/allocator and the existing managed slide handles commit. Empty pages and idempotent default-layout operations remain valid; unselected pages and shared definitions retain their objects. There is no new package, dependency, I/O, native-format field or hidden plan registry.

## Verification and limits

[ADR0059](0059-explicit-placeholder-indices.md) subsequently adds optional authored numeric indices between key and unique-role matching. The original unindexed behavior and output remain compatible; generated native ordinals are still not inferred Core identities.

[API](../../api/layout-rebind.md), focused Core tests, public Node/ESM/global comparisons, independent native XML reads and installed consumers cover the bounded contract. Parser source-only no-op remains exact; new definitions/rebinding do not gain preservation-write authority. Fresh native acceptance must have its own named artifact evidence, never be inferred from another shared-edit fixture.

Full cross-document/cross-size template reuse, field inheritance, arbitrary per-property paint/style reset, native-index matching, general host-measured preservation and source import/writes remain original parent requirements. The plan/apply foundation advances AUT-07.2.1–.3; it does not close those parent gates.
