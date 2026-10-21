# PPTKit Implementation Backlog

This is the actionable companion to the [Roadmap](../../ROADMAP.md). All unchecked tasks below are planned work, not implemented capabilities. Existing foundations are summarized in the roadmap and are not recreated here.

## How to use this backlog

- IDs are stable. Check a task only after its acceptance criteria and the shared delivery gate pass; add links to implementation and verification evidence beside the completed task.
- Dependencies are task IDs. They gate completion, not exploratory design. A milestone does not require every task in preceding milestones unless explicitly listed.
- Priority: **P0** establishes trustworthy contracts/evidence; **P1** improves everyday generation; **P2** expands capability; **P3** is specialized or research work. Priority does not override dependencies.
- Owners are package responsibilities, not newly mandated packages. `Core`, `Layout`, `Exporter`, `Renderer`, and `CLI` mean the existing public packages. `Parser` and `SVG parser` mean the planned packages. `Format layer` means parser/exporter-owned package handling; decide any shared public boundary in IMP-01 before creating an abstraction.
- Each row is a bounded deliverable. Before coding, split broad families into linked child tasks for the named supported subset; record excluded cases. Do not interpret “support charts/SmartArt/animation” as universal Office parity.
- Research tasks close with a reviewed decision record: supported subset, feasibility evidence, compatibility/fallback policy, ownership, and follow-up task breakdown. They do not count as shipped features.

## M0 — Reliability and release baseline

| Done | ID / priority | Owner | Depends on | Deliverable and acceptance |
| --- | --- | --- | --- | --- |
| [ ] | REL-01 / P0 | Maintainers | — | Audit public exports, current behavior, architecture, and docs. Publish a capability matrix with implemented/partial/planned states and evidence; include IR v2, chart behavior, minimal CLI, and provisional APIs. |
| [ ] | REL-02 / P0 | Exporter / QA | REL-01, REL-06 | Establish a PowerPoint/LibreOffice open-edit-save matrix. Record exact app/OS versions, repair prompts, editability, text/layout differences, and reopened results for each supported fixture; mark unavailable environments untested. |
| [ ] | REL-03 / P1 | Tooling | REL-06 | Add reproducible small/medium/large benchmarks for normalization, Layout, preview, export, memory, and output bytes. Record runtime/fonts/assets and report baseline variance before setting regression budgets. |
| [ ] | REL-04 / P0 | Maintainers | REL-01 | Reconcile release docs with current preview manifests/workflow. Define API/IR compatibility, version/tag policy, changelog, partial-publish recovery, deprecation and migration rules; verify release-script behavior matches the policy. |
| [ ] | REL-05 / P1 | Tooling | REL-04 | Test packed packages from isolated Node/browser consumers through public exports and declared dependencies. Detect workspace-only imports and missing artifacts; retain existing bundle and release checks. |
| [ ] | REL-06 / P0 | QA / all packages | — | Create a named, small fixture corpus covering rich text, themes, layouts, shapes, connectors, nested transforms, images, tables, charts, and notes. Link XML assertions, preview snapshots, asset licenses, and expected warnings; distinguish SVG evidence from actual PPTX rendering. |

## M1 — Measured layout and composition

| Done | ID / priority | Owner | Depends on | Deliverable and acceptance |
| --- | --- | --- | --- | --- |
| [ ] | LAY-01 / P1 | Layout / Core | REL-01 | Define injectable text measurement and font resolution contracts. Keep the current deterministic estimate/default behavior compatible; verify same input/provider yields the same result and no hidden DOM, filesystem, or network access enters Core. |
| [ ] | LAY-02 / P1 | Layout / runtime adapters | LAY-01 | Implement opt-in font-backed measurement and explicit fallback diagnostics. Cover mixed runs, CJK, explicit breaks, bullets, margins, missing fonts, and unavailable providers with pinned-font fixtures. |
| [ ] | LAY-03 / P1 | Layout | LAY-01 | Diagnose text overflow, clipping, and slide-boundary violations with slide/element IDs and actionable paths. Cover wrapping, margins, rotation, nested transforms, and intentionally off-slide elements; do not silently resize content. |
| [ ] | LAY-04 / P1 | Layout | LAY-01 | Add explicit alignment/distribution and horizontal/vertical stack composition. Verify padding/gaps, fixed versus available space, stable order, nested groups, and unsatisfiable-size diagnostics without mutating input. |
| [ ] | LAY-05 / P2 | Layout | LAY-04 | Add grid and bounded size/position constraints. Specify precedence and reject cycles/unsatisfiable constraints deterministically; cover mixed fixed/intrinsic tracks without promising CSS parity. |
| [ ] | LAY-06 / P1 | Layout | LAY-02, LAY-04 | Resolve intrinsic text/image/group sizing with explicit min/max bounds and aspect ratios. Diagnose unknown image dimensions and cyclic parent/child sizing; preserve explicitly sized elements. |
| [ ] | LAY-07 / P1 | Layout | LAY-02, LAY-03 | Measure table rows/cells including merged spans, padding, and rich text. Define fixed/automatic row sizing and overflow policy; verify exported row geometry and preview agree with resolved Layout. |
| [ ] | LAY-08 / P2 | Layout / Core / consumers | LAY-07, LAY-04 | Add opt-in table pagination and explicit cross-slide continuation. Cover repeated headers, unsplittable merged rows, oversized rows, stable generated IDs, source traceability, and repeat resolution without duplicate slides. |

## M2 — Rich authoring and templates

| Done | ID / priority | Owner | Depends on | Deliverable and acceptance |
| --- | --- | --- | --- | --- |
| [ ] | AUT-01 / P1 | Core / Layout / Exporter / Renderer | LAY-02 | Extend typography in separate slices: numbered/nested lists and tabs; RTL/bidirectional and vertical text; character/paragraph spacing and baseline styles. Each slice specifies font/fallback behavior and proves editability plus preview limitations. |
| [ ] | AUT-02 / P2 | Core / Layout / Exporter / Renderer | REL-01 | Expand preset geometry, then custom paths and connection sites. Start with a documented subset; test bounds, flips/rotation, fills, holes, and connector anchoring with native editable output or explicit fallback. |
| [ ] | AUT-03 / P2 | Core / Exporter / Renderer | REL-01 | Add gradient/pattern fills and selected shadows/effects as separate slices. Specify color/opacity/theme semantics and bounds; verify native OOXML and preview approximation warnings. |
| [ ] | AUT-04 / P1 | Core / Exporter / Renderer | REL-01 | Extend asset inputs with explicit byte-backed sources and resolver contracts. Verify MIME/signature handling, deduplication, load failures, browser/Node parity, and crop/rotation; keep runtime I/O outside Core. |
| [ ] | AUT-05 / P2 | Core / Layout / Exporter / Renderer | LAY-07 | Add table border precedence, cell alignment/text direction, and reusable table styles. Test row/column edits around spans and theme changes without invalid merged cells or silent content loss. |
| [ ] | AUT-06 / P2 | Core / Layout / Exporter / Renderer | REL-06 | Expand charts in slices: scatter/bubble; area/doughnut; combination/secondary axes. Per slice define data/axis semantics, missing values, labels/formatting, editable workbook data, and native-versus-preview parity fixtures. |
| [ ] | AUT-07 / P1 | Core / Exporter / Renderer | REL-01 | Model explicit master/layout/theme inheritance beyond current reusable layouts. Cover multiple masters, placeholder inheritance/overrides, background and color mapping; preserve existing layout APIs with a documented migration if necessary. |
| [ ] | AUT-08 / P2 | Core / Exporter / Renderer | AUT-07 | Add slide numbers, date/footer fields, and theme/layout variants. Verify inherited versus local overrides and deterministic field evaluation without hidden clock/locale dependencies. |
| [ ] | AUT-09 / P2 | Core / Layout | AUT-07, AUT-04 | Support copying slides/layouts/assets between documents with ID remapping and theme adaptation policy. Test connectors, actions, placeholders, asset deduplication, and source detachment. |

## M3 — PPTX import and scoped round-trip

| Done | ID / priority | Owner | Depends on | Deliverable and acceptance |
| --- | --- | --- | --- | --- |
| [ ] | IMP-01 / P0 | Parser / Exporter architecture | REL-01, REL-06 | Define format package graph and public import/save boundaries, capability tiers, diagnostics, and preservation ownership. Decide how a save consumes preserved state without parser/exporter private imports, cycles, or OOXML in Core/IR. |
| [ ] | IMP-02 / P1 | Parser | IMP-01 | Create bounded ZIP/XML inspection with content types, parts, relationships, and resource limits. Test malformed XML, oversized/expanding archives, duplicate/unsafe paths, missing targets, and external relationships; never fetch external links or execute embedded content implicitly. |
| [ ] | IMP-03 / P1 | Parser | IMP-02 | Import deck metadata, dimensions, slide order, theme/master/layout relationships, notes, hidden state, and sections. Preserve source identity mappings and report unresolved inheritance; imported supported semantics pass Core validation. |
| [ ] | IMP-04 / P1 | Parser | IMP-03 | Import a documented subset of text, images, preset shapes, connectors, and groups. Cover units, nested coordinates, style inheritance, asset extraction, and unsupported siblings; compare semantic fixtures from multiple producers. |
| [ ] | IMP-05 / P2 | Parser | IMP-04, AUT-07 | Resolve full supported master/layout/placeholder inheritance into editable semantics while retaining provenance for save. Verify overrides, theme mappings, and multiple masters without duplicating inherited drawing content. |
| [ ] | IMP-06 / P2 | Parser | IMP-04 | Import current supported tables and bar/line/pie charts, including spans and workbook/cached data consistency. Classify unsupported chart/table styles individually; do not imply later AUT-06 families are supported. |
| [ ] | IMP-07 / P1 | Format layer | IMP-02 | Preserve unknown parts, XML subtrees, extensions, and their relationship/resource closure. Test ID collisions, mixed editable/unknown content, missing dependencies, and retained ordering; no opaque fragments leak into common authoring types. |
| [ ] | IMP-08 / P1 | Parser / Exporter | IMP-04, IMP-07 | Implement no-op preservation save for a named fixture subset. Assert semantic package equivalence allowing documented ZIP/XML normalization; unknown parts and references survive, and supported applications reopen without repair. |
| [ ] | IMP-09 / P1 | Parser / Exporter | IMP-08, IMP-05, IMP-06 | Implement supported edit-and-save with dirty-region tracking and ID/relationship reconciliation. Test text edits, slide reorder/copy/delete, asset replacement, and references from preserved objects; reject or diagnose edits that invalidate preserved state. |
| [ ] | IMP-10 / P1 | QA / docs | IMP-09, REL-02 | Publish scoped round-trip guarantees per feature and operation. Test repeated saves, PowerPoint/LibreOffice-authored inputs, no-op versus edits, explicit fallback/loss approval, and every known unsupported fixture; make no universal lossless claim. |

## M4 — SmartArt, media, motion, and review

| Done | ID / priority | Owner | Depends on | Deliverable and acceptance |
| --- | --- | --- | --- | --- |
| [ ] | ADV-01 / P2 | Parser / Format layer | IMP-07, IMP-08 | Recognize and preserve existing SmartArt data, layout, style/color, relationships, and cached drawings. Verify no-op save and unrelated edits retain diagrams; report preserved rather than editable capability. |
| [ ] | ADV-02 / P3 research | Core / Layout / Format layer | ADV-01 | Evaluate editable SmartArt for an explicit subset: process, hierarchy, and cycle. Record data/visual graph semantics, native Office requirements, layout feasibility, and fallback choice using real fixtures; split implementation by supported family. |
| [ ] | ADV-03 / P2 | Core / Layout / Exporter / Renderer / Parser | ADV-02 | Deliver the first approved SmartArt family as a vertical slice with node editing, layout, native export/import where supported, and fallback diagnostics. Editing tests must distinguish native SmartArt from ordinary grouped shapes. |
| [ ] | ADV-04 / P2 | Core / Exporter / Parser / Renderer | AUT-04, IMP-04 | Add audio/video assets, poster frames, embedding/link policy, trim/volume/loop/playback controls in bounded slices. Verify asset relationships and application playback; preview unsupported codecs with explicit placeholders/warnings. |
| [ ] | ADV-05 / P2 | Core / Exporter / Parser | REL-06, IMP-04 | Add a small named set of slide transitions with duration and advance semantics. Verify native editable export/import, no-transition fallback, preserved unknown transitions, and reopen behavior per application. |
| [ ] | ADV-06 / P3 research | Core / Layout / Format layer | IMP-07 | Define format-independent animation timelines: stable targets, effects, sequence/parallel groups, triggers, timing, and unsupported behavior. Validate mappings for entrance/emphasis/exit/motion paths on fixtures before committing public APIs. |
| [ ] | ADV-07 / P2 | Core / Exporter / Parser | ADV-06, ADV-05 | Deliver an initial native animation subset with click/after/with-previous timing. Test references after copy/delete, unknown timeline preservation, invalid timing diagnostics, and application playback; add media synchronization only with ADV-04 complete. |
| [ ] | ADV-08 / P2 | Renderer / consumer adapters | ADV-07 | Provide optional transition/animation playback over detached state with seek/pause/replay and reduced-motion behavior. Test timeline ordering and static fallback; declare preview differences from native slideshow playback. |
| [ ] | ADV-09 / P2 | Core / Parser / Exporter | IMP-04 | Add comment threads, author identity, anchors, replies, and resolved state for a documented Office comment dialect. Preserve unsupported variants; verify IDs and anchors on slide/element deletion and cross-application reopen. |
| [ ] | ADV-10 / P1 | Core / Exporter / Renderer | REL-01 | Extend existing accessibility metadata with reading-order and contrast/missing-description diagnostics. Verify groups, tables, charts, and decorative content; document exported semantics and application limitations. |
| [ ] | ADV-11 / P3 research | Format layer | IMP-07 | Inventory/preserve OLE, controls, macros, and embedded fonts without execution. Decide support/rejection for macro-enabled inputs, font licensing/embedding restrictions, and save-format behavior; add tested preservation or rejection policies before implementation claims. |
| [ ] | ADV-12 / P3 | Core / Format layer | IMP-03 | Add selected presentation delivery settings: custom shows, slide advance/show options, handout/print metadata. Treat signatures/protection/encrypted input as separate research child tasks; validate support/fallback per setting without claiming enforcement the engine cannot provide. |

## M5 — Formats, tooling, and ecosystem

| Done | ID / priority | Owner | Depends on | Deliverable and acceptance |
| --- | --- | --- | --- | --- |
| [ ] | ECO-01 / P1 | CLI | REL-01, REL-05 | Add generation, validation, and preview commands using public engine APIs. Specify inputs/outputs, machine-readable diagnostics, exit codes, asset base paths, and overwrite behavior; test isolated installed usage. |
| [ ] | ECO-02 / P2 | CLI | ECO-01, IMP-04, IMP-08 | Add PPTX inspect/import/save commands with capability and preservation reports. Verify unsupported/error exits and explicit lossy operations; CLI must not reimplement parser/exporter logic. |
| [ ] | ECO-03 / P2 | SVG parser / Core | AUT-02, AUT-03 | Define and implement a directional SVG subset: geometry, paths, text, transforms, and paints. Diagnose filters/external resources/unsupported nodes, bound input complexity, and test editable conversion versus raster fallback. |
| [ ] | ECO-04 / P2 | Renderer / output adapters | LAY-02, REL-06 | Add portable SVG and raster/PDF output as separate adapters. Resolve hybrid foreignObject limitations, font/asset loading, page size, and clipping; compare multi-page output and document editability/runtime tradeoffs. |
| [ ] | ECO-05 / P2 | Renderer / consumer contracts | LAY-03 | Add incremental preview invalidation, hit testing, selection geometry, and coordinate mapping contracts. Verify nested transforms and stable IDs; keep editor UI and authoritative state outside Renderer. |
| [ ] | ECO-06 / P3 | Core / consumer contracts | REL-01, ECO-05 | Define serializable edit operations/change events for undo/redo and collaborative consumers. Test atomic updates, conflict detection, stale references, and replay determinism; network sync, presence, and authorization remain application-owned. |
| [ ] | ECO-07 / P3 research | Maintainers / package owners | REL-01, IMP-01 | Design extension points with capability negotiation, versioning, ownership, failure isolation, and trust boundaries. Demonstrate one consumer extension through public APIs before introducing a registry or new package. |
| [ ] | ECO-08 / P2 | Docs / consumers | LAY-04, AUT-07 | Add engine-level examples for reports, dashboards, diagrams, templates, and framework integrations. Keep examples runnable/typechecked and dependencies explicit; guided authoring themes and product workflow stay in PPTKit Presentation. |
| [ ] | ECO-09 / P2 | Tooling / QA | REL-02, IMP-04 | Add inspection reports for capability coverage, package relationships, asset size, and unsupported structures. Link each finding to a fixture/task and document reproducible import/export/layout debugging procedures. |

## M6 — Stable v1.0 contracts

| Done | ID / priority | Owner | Depends on | Deliverable and acceptance |
| --- | --- | --- | --- | --- |
| [ ] | STB-01 / P0 | Maintainers | REL-01, REL-04 | Select and publish the v1 supported capability set and stability boundary. Classify every remaining task as required-for-v1, experimental, or post-v1; advanced capabilities are not automatically release blockers. |
| [ ] | STB-02 / P0 | Core / all package owners | STB-01 | Review/freeze selected public APIs and IR compatibility/versioning rules. Complete migration examples and public-consumer checks for the selected surface; explicitly label any experimental exports. |
| [ ] | STB-03 / P0 | QA / Tooling | STB-01, REL-02, REL-03, REL-05 | Enforce selected runtime/application matrix, output-size/performance budgets, failure semantics, and reproducible release checks. An environment without evidence remains untested rather than supported. |
| [ ] | STB-04 / P0 | Maintainers / docs | STB-02, STB-03 | Complete user/API/migration/maintainer docs and a release rehearsal against the selected capability set. Every required task has linked evidence; deferred/experimental features and known limitations are published. |

## Recommended execution order

1. Complete REL-01 and REL-06, then establish REL-02/REL-04 and LAY-01.
2. Deliver LAY-02/03/04, consumer packaging checks, and benchmarks; start IMP-01/02 once their prerequisites exist.
3. Extend intrinsic sizing/tables and common authoring features while implementing supported import and preservation independently.
4. Prove no-op then edit-aware round-trip. Begin SmartArt preservation, transitions, media, and timeline research against named fixtures.
5. Deliver approved advanced feature slices and ecosystem adapters; select the v1 boundary explicitly rather than waiting for Office parity.

## Shared acceptance and adversarial review

For every shipped capability, verify the relevant authoring → validation/IR → Layout → preview/export path, plus import/save when claimed. Include examples, public API docs, and architecture updates. Run focused tests and the required workspace build, typecheck, lint, and tests; document any unavailable application verification.

Before closing a task, challenge at least these five failure points and attach evidence:

1. **Compatibility:** does it change existing documents, public exports, defaults, or IR shape without a stated policy?
2. **Hidden loss:** do unknown XML, relationships, timing targets, media, or inherited styles disappear after an edit/save?
3. **Layout divergence:** do browser/font differences, rotation, merged cells, or nested coordinates invalidate the result?
4. **Failure behavior:** are malformed inputs, resource limits, missing assets/fonts, and unsupported features diagnosed at useful paths?
5. **Architecture and complexity:** could the same result use an existing owner/public contract, avoiding private imports, cycles, global state, and unnecessary packages?

A type declaration, copied XML fragment, or screenshot alone does not prove a complete capability. Record supported edits and loss/fallback boundaries alongside the test evidence.
