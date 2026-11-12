# Architecture

These documents describe contributor-facing system boundaries, implemented contracts, and explicit future directions.

## Reading order

1. [Architecture Overview](overview.md)
2. [Core Authoring Model](core-authoring-model.md)
3. [Canonical Presentation IR v2](canonical-ir-v2.md)
4. [Document Models](document-models.md)
5. [Package Boundaries](package-boundaries.md)
6. [Rendering Pipeline](rendering-pipeline.md)
7. [Layout Engine](layout-engine.md)
8. [Import and Export Boundaries](import-export-boundaries.md)

Supporting decisions:

- [Design Principles](design-principles.md)
- [Editable vs Visual](editable-vs-visual.md)
- [Source Organization](source-organization.md)
- [Preview and Editing](preview-and-editing.md)
- [Testing Strategy](testing-strategy.md)
- [Monorepo Bootstrap](monorepo-bootstrap.md)

Current behavior and future architecture are labeled separately. Broader Parser preservation remains future work; bounded text-run profiles are documented in ADRs 0008–0010. Preview/rendering documents distinguish implemented SVG behavior from future fidelity work. The [capability audit](../planning/capability-matrix.md) links current public contracts to implementation and tests.

- [ADR 0001: import sessions and preservation saves](decisions/0001-import-preservation-save.md) — accepted ownership/lifecycle design; runtime gates remain open.
- [ADR 0002: chart data projection](decisions/0002-chart-data-projection.md) — bounded cache/workbook snapshot priority.
- [ADR 0003: placeholder dimension projection](decisions/0003-placeholder-dimension-projection.md) — source matching/provenance without a Core template replacement.
- [ADR 0005: OPC source comparison](decisions/0005-opc-source-comparison.md) — exact payload and conservative XML profiles, independent of Core projection and Exporter.
- [ADR 0006: source relationship closure](decisions/0006-source-relationship-closure.md) — inactive MCE dependencies, outside ownership and scoped completeness without cleanup/edit permission.
- [ADR 0007: XML character edit boundaries](decisions/0007-xml-character-edit-boundaries.md) — private source-anchored character replacement that retains surrounding lexical XML.
- [ADR 0008: Direct text preservation save](decisions/0008-direct-text-preservation-save.md) — opt-in source-to-Core run bindings, whole-candidate guards, ZIP patching and validated session plans.
- [ADR 0009: Shape-text and notes preservation save](decisions/0009-text-owner-preservation-save.md) — explicit V2 ownership, separate notes identities and source-part guard scoping.
- [ADR 0010: Table-anchor text preservation save](decisions/0010-table-text-preservation-save.md) — explicit V3 scope, physical/compact anchor identities and validated scalar edits.
- [ADR 0011: Node file publication](decisions/0011-node-file-publication.md) — Node-only stable import, exclusive create, snapshot-guarded replace, synchronous plan freshness and commit warnings.

- [ADR 0012: Slide Fade transition](decisions/0012-slide-fade-transition.md) — page semantics, exact timing, source-only unknown effects and separate playback ownership.

- [ADR 0013: Finite animation time research](decisions/0013-finite-motion-time-research.md) — executable bounded timing experiment before public/native animation APIs.

- [ADR 0014: Animation target/event research](decisions/0014-animation-target-event-research.md) — stable whole-object identities and explicit candidate foreground event gates.

- [ADR 0015: Native animation mapping experiment](decisions/0015-native-animation-mapping-experiment.md) — four behavior specimens, independent target/timing inventory and headless producer differences before effect admission.

- [ADR 0016: SmartArt component inventory](decisions/0016-smartart-component-inventory.md) — effective source identities and four component associations before cache or node semantics.

- [ADR 0017: SmartArt cache/resource associations](decisions/0017-smartart-cache-resource-associations.md) — explicit slide-scoped cache references, unassigned candidates and shared direct resource graphs.

- [ADR 0018: Cached diagram preview](decisions/0018-cached-diagram-preview.md) — explicit ordinary-shape fallback projection with source reports and guarded save reprojection.

- [ADR 0019: Cached diagram text rectangles](decisions/0019-cached-diagram-text-transform.md) — independent parent-space text frames, relative rotation and explicit unavailable cases.

- [ADR 0020: SmartArt source data graphs](decisions/0020-smartart-source-data-graph.md) — format-owned identities/references/direct text, separate cache model references and cumulative graph limits.

- [ADR 0021: Explicit SmartArt mappings](decisions/0021-smartart-explicit-model-mapping.md) — normalized presentation evidence and actual cache pairs with descriptive agreement.

- [ADR 0022: SmartArt content-model research](decisions/0022-smartart-content-model-research.md) — ordered forests, assistants, layout-generated Cycle closure and scoped native source editing evidence.

- [ADR 0023: Minimum native SmartArt generation](decisions/0023-native-smartart-generation-research.md) — fresh packages, exact presentation graphs, text-fitting correction and separate lexical/native-edit gates.

- [ADR 0029: Diagram cache consistency research](decisions/0029-diagram-cache-consistency-research.md) — explicit dependency witnesses, staged/guarded publication and native spacing attribution.
- [ADR 0028: Cycle deterministic layout feasibility](decisions/0028-cycle-layout-feasibility.md) — equal-circle packing, explicit text fitting and measured native cache mismatches.
- [ADR 0027: Scoped SmartArt source bindings](decisions/0027-smartart-source-binding-research.md) — source/managed identity separation, context and retirement guards, independent native-artifact checks.
- [ADR 0026: SmartArt identity-state research](decisions/0026-smartart-identity-state-research.md) — explicit pure reservations, structural-edit references, atomic failure and native admission limits.
- [ADR 0025: Generated SmartArt native edit evidence](decisions/0025-generated-smartart-native-edit-evidence.md) — scoped PowerPoint editing, specimen GUIDs and remaining admission gates.
- [ADR 0024: Native SmartArt geometry observation](decisions/0024-smartart-native-geometry-observation.md) — independent page-point bounds, raw transforms/text areas, repeated-name evidence and long-text failures.
- [ADR 0030: Direct diagram text observations](decisions/0030-diagram-text-observations.md) — opt-in bounded character diagnostics and controlled native cache disagreement evidence.
- [ADR 0031: Native stale-cache policy](decisions/0031-native-stale-cache-policy.md) — cross-application conflict/rebuild observations and synchronized-cache write admission gates.
- [ADR 0032: Read-only CLI inspection](decisions/0032-read-only-cli-inspection.md) — public Parser orchestration, versioned bounded output and stable command statuses.
- [ADR 0034: Chart numeric preservation saves](decisions/0034-chart-numeric-preservation-save.md) — opt-in existing values, synchronized cache/XLSX candidates and independent two-layer validation.
- [ADR 0033: Import capability observations](decisions/0033-import-capability-observations.md) — exact source subjects, explicit missing/conflicting operations and bounded versioned reference views.

- [ADR0035: Public finite timeline arithmetic](decisions/0035-public-finite-timeline-arithmetic.md) — bounded integer schedules and transport-safe random-access samples, separate from animation targets/effects/native admission.

- [ADR0036: Foreground event timeline replay](decisions/0036-event-timeline-replay.md) — scoped logs, current inventory guards and deterministic consumption without event dispatch or document/native effects.

- [ADR0037: Slide-bound event target transport](decisions/0037-slide-event-target-transport.md) — owned timing source, atomic identity updates/copies, additive IR/Layout and strict output boundaries.


- [ADR0038: Native timing source inventory](decisions/0038-native-timing-source-inventory.md) — opt-in source observations, part-local candidates and explicit mapping/playback boundaries.


- [ADR0039: Stable paragraph target preservation](decisions/0039-stable-paragraph-target-preservation.md) — opt-in source index preservation with exact target/owner proofs and separate native limitations.


- [ADR0040: Atomic table structure edits](decisions/0040-atomic-table-structure-edits.md) — pure bounded grid transactions, explicit content/frame ownership, complete anchor mapping and existing commit integration.

- [ADR0041: Shared table borders](decisions/0041-shared-table-borders.md) — direct sparse overrides, request-wide bounded physical edges and explicit native representation limits.

- [ADR0042: Table region style materialization](decisions/0042-table-region-style-materialization.md) — Core resolves bounded format-independent rules before shared Layout and direct-native output.

- [ADR0043: Source table-style projection](decisions/0043-source-table-style-projection.md) — Parser-owned style identity, per-slide theme evaluation, bounded provenance and conditional fallback.
- [ADR0044: Table theme-matrix references](decisions/0044-table-theme-matrix-references.md) — positional format lists, lazy placeholder binding, unrounded color composition and retained native compatibility gaps.
- [ADR0045: Table text appearance and explicit fonts](decisions/0045-table-text-appearance-and-uniform-fonts.md) — reference default colors, uniform script-font collections and independent PDF appearance gaps.
- [ADR0046: Script-font authoring and selection](decisions/0046-script-font-authoring-and-selection.md) — additive Core slots, shared Unicode fragments and separate source/native admission gates.
- [ADR0047: Source script-font profile](decisions/0047-source-script-font-profile.md) — independent font cascade/blockers, notes provenance and profile-preserving candidate reprojection.

- [ADR0048: Script font maps and language selection](decisions/0048-script-font-maps-and-language-selection.md) — per-script transport, explicit Han hints and warned fresh-output materialization; source supplemental admission remains separate.

- [ADR0049: Source supplemental fonts](decisions/0049-source-supplemental-font-projection.md) — private collection metadata until native-slot cascade selection, admitted map projection and source/fresh-output boundaries.

- [ADR0050: Notes-master text inheritance](decisions/0050-notes-master-text-inheritance.md) — independent notes theme/masters, type-based text inheritance and preserved source/fresh-output boundaries.
- [ADR0054: Default fresh notes-page output](decisions/0054-default-notes-page-output.md) — fixed aspect-preserving native image/body layout without new Core geometry, plus explicit empty-style native failure.
- [ADR0055: Empty-paragraph insertion properties](decisions/0055-empty-paragraph-insertion-properties.md) — retain the final empty-run style on the paragraph mark without duplicating actions; source permissions stay separate.

- [ADR0051: Declared-slot supplemental projection](decisions/0051-declared-slot-supplemental-projection.md) — explicit v3 source selection without absent-bucket Latin heuristics, located unresolved rules and compatible captured saves.

- [ADR0052: SVG text boundaries and font names](decisions/0052-svg-text-boundaries-and-font-names.md) — actual wrap flags, grapheme-safe line boundaries, empty-line geometry and literal CSS family serialization.

- [ADR0053: Declared fonts and native appearance](decisions/0053-declared-fonts-and-native-appearance.md) — preserve source-profile compatibility and report native selection differences separately, without host fallback inference.


- [ADR0056: notes source inventory](decisions/0056-notes-source-inventory.md).

- [ADR0057: master theme/background ownership](decisions/0057-master-theme-ownership.md).
- [ADR0058: reviewed layout rebinding](decisions/0058-reviewed-layout-rebinding.md) — frozen value plans and atomic multi-page commits.
- [ADR0059: explicit placeholder indices](decisions/0059-explicit-placeholder-indices.md) — authored numeric identity, role-safe matching and collision-free native serialization.

- [ADR0063: explicit chart blanks](decisions/0063-chart-explicit-blanks.md) — null points, bounded gap projection and fresh-export semantics.
- [ADR0064: preservation-first reproduction](decisions/0064-reproduction-and-redaction.md) — scoped closure evidence, immutable no-op baselines and explicit reduction/redaction limits.
