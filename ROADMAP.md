# PPTKit Roadmap

PPTKit is building a format-independent presentation document engine with editable PPTX output. This roadmap describes capability milestones, not a promise of complete Office parity or a single release containing every feature.

The local [0.2.0 preview candidate](docs/planning/release-0.2.0-preview.md) selects a reviewable provisional release scope without treating the remaining backlog as complete. The v1 stability gates and individually documented native-application limits remain open.

The [2026-10-26 recovery audit](docs/planning/recovery-2026-10-26.md) records the current full automated checks, retained evidence, and explicit gaps after restoring the previous session from plaintext. The [notes slide-image source repair](docs/testing/native-notes-pages.md) now has three native reopen rounds and three additional copied-slide rounds; complete fresh-author notes pages and the remaining backlog stay open.

The [default fresh notes-page layout](docs/testing/authored-notes-pages.md) now exports aspect-preserving native slide-image and separate body placeholders, with three-size native reopen evidence. [Empty insertion-property serialization](docs/testing/notes-insertion-styles.md) now corrects the initial native style loss, verified across three empty and three actual-typing reopen rounds; zero-character run identities/actions and broader typography remain separate. The [37-input notes candidate consolidation](docs/testing/native-notes-consolidated.md) has been rerun across all three consumers with 143 retained corruption controls revalidated; broader notes/typography and round-trip gates stay open.

[B and distinct per-notes override image evidence](docs/testing/notes-image-domains.md) now extends the rect-only source repair with 45 first-open and 135 native reopened pages. Body/artwork and all non-app/core payloads remain exact; complete notes-page roles/geometry/fields/source-only contracts and shared-override native admission remain open.

The [Implementation Backlog](docs/planning/implementation-backlog.md) contains task IDs, dependencies, ownership, and acceptance criteria. This file owns milestone order; the backlog owns task status and delivery details. Update both when scope or milestone order changes.

The [Advanced Capability TODOs](docs/planning/advanced-capability-todos.md) break milestones into stable task IDs with ownership, inherited dependencies, bounded scope, and acceptance evidence. They cover import/save contracts, semantic import and scoped round-trip, inheritance, structural edits and cross-document reuse, SmartArt families, media/captions, transitions, individual animation effects and targets, and controlled preview playback.

The same checklist also covers everyday typography, shapes/images, tables/charts, templates, comments, accessibility, slideshow/print settings, portable outputs and tooling. Formula, 3D, ink, additional SmartArt families and other specialized areas begin with explicit research decisions. A research result does not establish implementation support.

For the starting order, whole-slide coverage map, subsequent queues and planned fixtures, see the [Next Delivery Plan](docs/planning/next-delivery-plan.md). The immediate queue builds on the verified experimental text-run save profile: complete application/reference gates, continue preservation/comparison evidence, inheritance, remaining table styles and chart data/format/native interactive gates, then expand edit operations. SmartArt preservation and timeline research have their own dependency gates. Newly planned slices remain unverified; v1 scope is selected separately under STB-01.

The [next-batch entry](docs/planning/next-delivery-plan.md#本次排期先领取哪一批) separates the text-save profile's completed automated verification from remaining application/reference gates. Nine additional TODOs cover that acceptance, more text owners/topology, source-preserving table/chart/property/action edits, table structural saves and an executable operation matrix. SmartArt inventory/native-editing research and animation timeline research can prepare their own evidence independently; they remain planned.

The detailed TODOs also split multi-master authoring, theme transforms/script fonts, table/chart semantics, independent comparison, modification tracking and save transactions into verifiable leaves. SmartArt has source/data/cache experiments and a separate native-editing admission decision; motion has individual trigger, repetition, composition and Morph/Zoom research tasks. The next-delivery plan records the completed effective-color parity checks and remaining native/theme gates before expanding implementation. Newly planned tasks remain incomplete.

The checklist additionally separates dialect/file-kind admission, MCE source branches, background inheritance and master artwork, template rebinding, deterministic fields, complete notes-page semantics, native section collections, and typed custom properties. These are planned extensions of existing task IDs; their reading/authoring/fresh-output acceptance stays separate from preservation writes under IMP-09.

The next level of tasks covers shared text layout and line breaking, preset/custom geometry, gradients/patterns/shadows, image format policies, and atomic table editing. Advanced acceptance is also split into SmartArt data/cache consistency, slideshow interruption/re-entry, mixed-content edit-save regression and inspect/compare reports. Each task has a bounded deliverable, failure cases and inherited completion gates; these remain planned capabilities.

The [next execution cards](docs/planning/advanced-capability-todos.md#下一批执行卡清单播放证据与操作矩阵) split SmartArt discovery/component/cache reporting, native animation playback admission, and a versioned operation matrix into ten further leaves. These are planned tasks with explicit failure cases and evidence gates; existing headless experiments do not establish playback or native editability.

The [continuation execution cards](docs/planning/advanced-capability-todos.md#接续执行卡数据图图表保存与时间线) add twelve bounded leaves for SmartArt source graphs/cache identity, chart cache/workbook preservation saves, and public timeline/event/target contracts. They specify failure cases, independent evidence and native admission gates; all remain incomplete. The [current handoff](docs/planning/next-delivery-plan.md#接续领取顺序数据图图表保存与时间线) selects the first reviewable deliverable in each track without restarting existing implementations.

The latest [table/notes native-font execution cards](docs/planning/advanced-capability-todos.md#表格与备注原生字体的可领取任务) separate declaration/appearance policy, actual text edits, regular/bold/italic, theme changes, candidate checks, and independent notes baselines/edits. The policy decision, [15-target native append slice](docs/testing/native-table-font-edits.md), and [bold/pure-italic/restored-regular slice](docs/testing/native-table-font-styles.md) now have evidence; theme/notes cards and parent gates remain open. The existing 36-control unedited baseline records 20 matches and 16 differences, which remain explicit after append validation.

## Status legend

For the complete upcoming queue, start with the [capability delivery order](docs/planning/next-delivery-plan.md#规划入口完善幻灯片能力的领取顺序): trustworthy import/round-trip, everyday preservation edits, native SmartArt, motion/media, broader authoring/playback, then extension and release gates. Seven new [SmartArt output/layout cards](docs/planning/advanced-capability-todos.md#下一步执行卡smartart-最小原生输出与布局准入) separate differential diagnosis, minimal generation, geometry/cache decisions, native display and actual node editing. [ADR0023](docs/architecture/decisions/0023-native-smartart-generation-research.md) records corrected text-fitting/observation logic and six standard generated candidates with native headless save/reopen text evidence; interactive node editing and family admission remain open.

- **Implemented** — present in the repository with implementation and verification evidence.
- **Next** — near-term reliability and layout work.
- **Planned** — scheduled behind explicit capability dependencies; not a current API.
- **Research** — requires a feasibility or contract decision before implementation.

No dates or effort estimates are committed. Independent tracks may progress concurrently after their dependencies are complete.

## Implemented foundation

- **Core contract:** method-managed authoring state, readonly collections, document-wide IDs, complete diagnostics, and Canonical Presentation IR v2.
- **Editable content:** rich text, theme/direct colors, paints, strokes, actions, image fitting/crop, transforms, and accessibility metadata.
- **Structured elements:** preset shapes, anchored connectors, nested groups, editable tables with spans, notes, hidden slides, sections, tags, and custom data.
- **Charts:** native bar, line, and pie charts with Layout geometry, SVG preview, editable OOXML export, and chart parity tests.
- **Templates:** themes, reusable layouts, typed placeholders, binding, and explicit inheritance.
- **Layout:** detached resolution for connector anchors/bounds, image contain/cover, nested groups, and chart geometry. Core also provides deterministic text-height estimation. Opt-in Layout analysis accepts synchronous text/font providers and reports text/slide overflow without changing geometry; an explicit `/dom` adapter now measures using loaded browser fonts and a verified coverage inventory.
- **PPTX export:** browser-neutral bytes, Node output, native themes/layouts/tables/notes/media relationships, validation failures, and recoverable asset warnings.
- **Browser preview:** per-slide hybrid SVG, current element families, asset/fidelity warnings, workbench integration, and Chromium visual fixtures.
- **Developer delivery:** package manifests/exports, a local npm release script with tests, API/getting-started documentation, typed examples, browser bundle checks, and CI. CLI commands remain a bootstrap shell.
- **Reliability evidence:** [capability audit](docs/planning/capability-matrix.md), [named fixture corpus](docs/testing/fixture-corpus.md) with XML/relationship/workbook verification and pinned-font snapshots, and [mixed-deck performance baselines](docs/performance/README.md). Concurrent chart export and hyperlink XML escaping have dedicated regression coverage.

These foundations do not establish a formal cross-application compatibility matrix, complete layout fidelity, or blanket round-trip guarantees.

An experimental [independent OPC source comparator](docs/api/package-comparison.md) now checks exact payloads or explicitly selected conservative XML structures, plus graph identities and bounded differences. Independent ZIP/SHA256 and browser DOM checks cover source loss and candidate mutations. Full semantic comparison, native validation and edit-save gates remain open.

Experimental [source relationship closure analysis](docs/api/preservation-closure.md) now includes inactive MCE dependencies, owned relationship parts, shared/cyclic targets, outside incoming owners and selected-byte budgets. Opaque/missing/ambiguous source information produces scoped coverage findings. Unreached parts are not treated as garbage; complete preservation-state and structural reconciliation remain pending.

A private [XML character-edit foundation](docs/architecture/decisions/0007-xml-character-edit-boundaries.md) binds original element identities to lexical character ranges, retains surrounding XML/source encodings and rejects mixed-content targets. The experimental opt-in [`text-runs-v1` profile](docs/api/text-edit-save.md) now adds source-to-Core bindings, whole-plan guards, ZIP patching and actual-candidate validation for existing text-box run characters. Full build/typecheck/test and browser/installed-consumer verification pass, with [execution evidence](docs/planning/execution-record.md#direct-text-run-preservation-saves-2026-10-25). Native headless evidence retains edited/control text but records producer loss of unknown content; PowerPoint, interactive editing and complete reference/fidelity gates remain open. This does not close IMP-09 or establish general edit-save support.

The explicit [`text-runs-v2` profile](docs/api/text-owner-save.md) now extends character-only preservation to mapped shape text and direct speaker notes, with separate notes identities and guards scoped by source part. Mixed-document three-round, browser/installed-consumer and headless producer evidence is linked in the [owner execution record](docs/planning/execution-record.md#shape-text-and-speaker-notes-preservation-saves-2026-10-25). It retains V1/default scope and leaves IMP-09.2.1.2's complete inherited/native gates open.

The explicit [`text-runs-v3` profile](docs/api/table-text-save.md) additionally saves existing table-anchor run characters, keeping compact Core cell indices distinct from physical source columns. Independent merge-partition, mixed-owner three-round, browser/installed-consumer and native headless evidence is linked in the [table execution record](docs/planning/execution-record.md#table-anchor-character-preservation-saves-2026-10-25). Direct formatting/metric/structural writes and complete IMP-09.2.5 native/reference gates remain open; earlier profiles stay compatible.

The Node-only [`pptx-parser/node` file profile](docs/api/file-save.md) now imports stable local sources and publishes prepared candidates through exclusive create or snapshot-guarded same-directory replace. Injected open/write/sync/readback/link/rename/cleanup failures, source/target races, cancellation boundaries and real three-round integration are covered. This advances IMP-09.5.3; cross-process CAS, Windows/filesystem certification and inherited native/rebuild gates remain open.

## M0 — Reliability and release baseline — Next

Backlog: **REL-01–REL-06**.

Audit public contracts and capability status; establish PowerPoint/LibreOffice open-save fixtures, performance/output-size baselines, release/compatibility policy, consumer-package verification, and a traceable regression corpus. Extend existing documentation and automation rather than rebuilding them.

REL-01 and REL-03–REL-06 now have implementation and verification evidence, including release policy and actual installed-tarball consumers. REL-02 has initial LibreOfficeDev headless save/reopen results; interactive/native visual checks remain open. See the [execution record](docs/planning/execution-record.md).

## M1 — Measured layout and composition — Implemented scoped baseline

Backlog: **LAY-01–LAY-08**.

Add pluggable text/font measurement, overflow and boundary diagnostics, alignment/stack/grid composition, intrinsic sizing, table measurement, explicit pagination/continuation, and preview/export layout consistency. Preserve deterministic defaults and isolate environment-backed measurement outside Core.

LAY-01–LAY-04 have implementation and verification evidence: injected provider contracts, an opt-in DOM font adapter, overflow/boundary diagnostics, and explicit alignment/distribution/stacks. LAY-07 adds measured table rows with merged-span constraints and verified preview/export geometry. LAY-06 adds explicit content/parent sizing for text, images and groups with bounded dimensions, aspect ratios and cycle diagnostics. LAY-05 adds bounded fixed/content/fraction grid tracks and affine frame constraints with explicit dependencies, cycle rejection and atomic placement. LAY-08 adds explicit top-level table pagination with repeated headers, merged-block preservation, continuation layouts and source-row mapping. Shared resolved line rendering and general text/story flow remain outside this delivered baseline; see the [execution record](docs/planning/execution-record.md).

## M2 — Rich authoring and templates — Planned

Backlog: **AUT-01–AUT-09**.

Expand text and international typography, geometry/paths, paints/effects, image assets, tables, chart families, masters/layout inheritance, and slide reuse. Features ship as complete vertical slices, each with explicit preview and export behavior. Byte-backed assets now use immutable serializable snapshots; explicit export resolvers add PNG/JPEG/GIF signature classification, per-asset budgets and cancellation. Core asset lifecycle now adds opt-in byte content matching, stable-ID replacement, reference inspection and explicit guarded cleanup. Broader image formats and preservation-aware lifecycle work remain open. Same-document copy now remaps current connector targets and self-slide actions, with cross-package geometry and relationship verification; cross-document reuse remains planned. An opt-in native straight connector export now maps supported cardinal sites and drawing identities with explicit path fallback; broader routes and interactive application evidence remain open.

## M3 — PPTX import and scoped round-trip — Planned

Backlog: **IMP-01–IMP-10**.

Introduce a parser, bounded package inspection, supported semantic import, master/theme inheritance resolution, and preservation-aware saves. Classify every encountered feature as editable, preserved, fallback, or unsupported. Define separate guarantees for no-op saves, supported edits, and destructive operations; do not promise arbitrary PPTX fidelity.

An experimental Parser package now exposes bounded ZIP32 inventory with strict DEFLATE structure and CRC verification. Per-part XML inspection adds namespace, encoding and resource-limit checks. OPC inventory now adds content types, internal/external relationship graphs and cumulative XML budgets. Presentation entry inspection now resolves the main part and identifies its Strict/Transitional dialect and document kind. Explicit MCE inspection now processes alternate branches, ignorable content and opaque extension islands, and is integrated into relationship inspection. `inspectPptx` now reads effective slide ordering, dimensions, hidden state and typed master/layout/notes/theme references, with unresolved-inheritance findings, plus package-linked core/extended/custom metadata with exact scalar strings and explicit uninterpreted complex values. Native MS-PPTX section inspection now retains source identities, empty sections and exact slide references, reports incomplete/out-of-order membership and inventories unknown presentation/section extensions. An explicit structural `importPptx` session maps dimensions/selected metadata/slide ordering into Core, inventories unsupported content, and supports original-byte no-op plans with lifecycle, freshness and resource checks. An opt-in content projection maps direct rich text, seven preset shapes, nested groups, embedded PNG/JPEG/GIF pictures, same-scope straight connector attachments, safe direct click actions, and direct speaker-notes rich text with object/property findings. External images remain unloaded; unsupported routes/actions and full inheritance remain partial. A [direct table profile](docs/api/table-import.md) now maps authored grids, rectangular merges, rich-text cells and bounded direct styles, with independent/browser/native-producer evidence. A [bounded chart profile](docs/api/chart-import.md) now maps single-family 2D bar/line/pie, explicit cache/workbook priority, supported direct styles and approximate preview, with independent three-round cache/worksheet checks, browser/installed consumers and native-producer/headless-reopen evidence. Full chart data/format semantics, native interactive editing, native table styles, notes-page/master/theme semantics, edit-aware preservation and semantic rebuild remain pending.

Import begins after the contract audit and fixture corpus. Initial preservation work can proceed before broader M2 authoring capabilities; unsupported features remain classified until their semantic slices exist.

The content profile also imports [direct speaker-notes rich text](docs/api/notes-import.md), with Core notes replacement, supported run clicks and located source/property reports. Notes-page artwork, fields, master/theme inheritance, preview and preservation edits remain open. Three fresh export/import rounds, public consumers and a producer-loss-aware LibreOffice notes inventory provide scoped evidence. Structural metadata projection now maps the aggregate keyword string and only nonnegative JavaScript-safe revision integers; mixed-language keyword segmentation and unsafe revision lexical values stay in the source report.

The [placeholder dimension profile](docs/api/placeholder-inheritance.md) now matches source index/type within the owning layout/master chain and records property-level EMU provenance. A separate [bounded text style profile](docs/api/text-style-inheritance.md) resolves supported frame/paragraph/list/run properties, theme color maps/overrides and major/minor Latin fonts with source dependencies. Partial theme overrides inherit color/font components independently in the target slide/layout/master scope. Both project local Core values without copying master prompt text. A [named LibreOfficeDev experiment](docs/testing/theme-text-import-evidence.json) records native source/fresh color/font differences, so automated parity does not establish native visual fidelity. Ambiguous/grouped inheritance and non-neutral ancestor transforms remain explicitly limited. A [named text-color transform profile](docs/api/text-color-transforms.md) now adds ordered tint/shade/HSL/alpha and run opacity across Core, preview and fresh output, with independent provenance budgets. Native nine-sample evidence records one gray-shade discrepancy. Remaining transform/native-fidelity work, non-Latin script fonts, backgrounds, Core multi-master authoring, rebinding and preservation writes stay scheduled under AUT-07 and IMP-05/09.

Both projections also map named native section members to Core slide labels by source numeric slide identity. The [section label contract](docs/api/import-session.md#native-section-member-labels) distinguishes labels from native GUID/collection/empty-section semantics. Core native collections, section-boundary editing and native section export remain planned; fresh output currently writes PPTKit label metadata only.

## M4 — SmartArt, media, motion, and review — Planned / Research

[Cache consistency research](docs/architecture/decisions/0029-diagram-cache-consistency-research.md) now captures exact dependency snapshots and stages validated cache bytes with revision/epoch/ownership/cancellation guards. Imported payloads stay unknown; managed-current applies only to its explicit generator/validator scope. An independent SVG geometry-cache integration verifies stale rejection and replacement. Native spacing/stroke/EMU controls retain the earlier layout mismatches; full dependency derivation, native cache generation and whole-package publication remain open.

[Cycle layout feasibility](docs/architecture/decisions/0028-cycle-layout-feasibility.md) now has deterministic equal-circle packing, directed geometry, frame transforms and explicit measurement/overflow with negative-bearing placement. Six pure/two browser groups validate the owned contract. Six retained PowerPoint cache comparisons exceed the 0.1-point research threshold by maximum coordinate deltas of 0.137–0.235 points; native matching, arrow/text behavior, caches and public family admission remain open. Two-node curved edges are an unverified native policy.

[Scoped source-binding research](docs/architecture/decisions/0027-smartart-source-binding-research.md) now maps recognized raw source point IDs into explicit managed identities across reinspection, order/text/index changes and recorded native additions/deletion. Retired source IDs require explicit restoration; context/root changes and producer-wide renumbering reject. Eleven retained source/native versions have independent binding/topology/text checks. This does not authorize source writes, full identity adoption or native admission of stateful fresh output.

[Explicit identity-state research](docs/architecture/decisions/0026-smartart-identity-state-research.md) now preserves semantic-key assignments across generated structural edits with detached state, retired reservations and shared work budgets. Independent DOM verifies 48 structural stages. Public managed identity/source adoption, concurrency and native admission of this stateful option remain open. Separately, the earlier generated Cycle's added blank node was deleted in PowerPoint after specific authorization; save/close/reopen retains four labeled nodes and closure.

[Generated-output native editing research](docs/architecture/decisions/0025-generated-smartart-native-edit-evidence.md) now records PowerPoint recognition, text/add/reorder/layout/theme edits and reopen for standard absent-cache Process/Hierarchy/Cycle. Three GUID Cycle native text cycles retain content IDs. Stateless GUIDs remain specimen-local; the full deletion matrix, native structural identity maintenance, cache/layout policy and public family admission remain open. Original-source Process also overflows in the named LibreOffice controls, so generator-only attribution is unproven.

[Native geometry observation research](docs/architecture/decisions/0024-smartart-native-geometry-observation.md) now covers all 24 generated presentation candidates with independent FODP/PDF artifacts, axis-aligned point bounds, raw transforms/text areas and occurrence-scoped name evidence. Native names can repeat; long-text clipping and PDF text discrepancies remain explicit failures/limitations. These baselines advance layout inputs without admitting a deterministic family layout, cache freshness or interactive native editing.

[SmartArt model research](docs/architecture/decisions/0022-smartart-content-model-research.md) now has an executable ordered-content-forest candidate and named PowerPoint source-edit evidence for Process/Hierarchy. It preserves assistant roles and separates Cycle's generated closure from source parent topology. Public managed identities, full rich text, family geometry/cache generation and minimum native output admission remain open; native source editing is not PPTKit native authoring support.

The [explicit SmartArt mapping extension](docs/api/diagram-mapping.md) now links normalized presentation evidence and recognized cache-shape references for actual data/cache pairs. Independent source checks cover different-target and nonempty-cache cases. Semantic ownership, family model/layout decisions, cache freshness and native editing remain separately gated.

The experimental [SmartArt source graph profile](docs/api/diagram-data.md) now inspects direct point/connection identities, typed references and stored text once per data part, with cumulative graph budgets. Raw cache model references remain source evidence; full data-to-visual mapping, native node editing and family layouts are still gated separately under ADV-02/03.

Backlog: **ADV-01–ADV-15**.

- SmartArt: preserve existing diagrams first, then support an explicit editable subset, then author selected diagram families.
- Media: audio/video assets, poster frames, playback semantics, and portable fallback behavior.
- Motion: slide transitions, element timelines, triggers, sequencing, then interactive preview playback.
- Review: comments, author/thread identity, accessibility coverage, and incremental change events for consumers. Real-time collaboration infrastructure belongs to applications.
- Advanced Office objects: inspect/preserve OLE and controls without executing them; research embedded fonts and advanced presentation settings.
- Specialist content: research native equations, 3D models, and ink as separate capability tracks, with explicit preservation/editing/static-preview boundaries before implementation commitments.

Each feature declares native editability, preservation scope, preview limitations, and cross-application behavior. Unsupported advanced features remain visible in capability reports.

An experimental [SmartArt component inventory](docs/api/smartart-inventory.md) now identifies effective slide/group diagram objects and four source component associations, with raw identities and explicit missing/invalid/external results. Synthetic source, independent DOM/browser and unrelated text-save fixtures establish this inspection profile only. Explicit slide-scoped cache associations, unassigned candidates and direct shared-resource/image relationships now extend this inventory. An explicit [cached preview option](docs/api/diagram-preview.md) now projects bounded drawing shapes/text/groups/images or unavailable placeholders, with separate fallback reports and guarded source-save reprojection. Cache freshness, complete source ownership/closure, full preview fidelity, semantic nodes and native family editing remain open under ADV-01–03.

## M5 — Formats, tooling, and ecosystem — Planned

Backlog: **ECO-01–ECO-09**.

Add practical CLI workflows, directional SVG parsing, portable/raster/PDF output adapters, preview/editor integration contracts, plugin capability negotiation, engine examples, interoperability reporting, and developer inspection tools. Guided authoring products stay in the separate PPTKit Presentation repository.

## M6 — Stable v1.0 contracts — Planned

Backlog: **STB-01–STB-04**.

Stabilize deliberately selected public APIs and IR rules; publish tested runtime/application support, regression budgets, migration guidance, and maintainer procedures. v1.0 requires an explicit supported capability set, not completion of every advanced Office feature.

## Capability delivery rule

A task is complete only when its public semantics, validation/normalization, required Layout behavior, PPTX export or documented fallback, tests, examples, API reference, and architecture documentation are complete where applicable. Documentation-only and research tasks use their own stated acceptance criteria.

No feature is marked implemented solely because its type exists, an XML part is emitted, a preview looks correct, or an unknown part is copied. Compatibility claims require named fixtures and recorded application results.

Experimental [page Fade transitions](docs/api/slide-transitions.md) now connect Core timing, Layout, native fresh output and bounded source import. Independent DOM/three-round/static-preview/preservation evidence advances ADV-05.1.1. Native settings/playback, Push/Wipe and preview playback remain open.

Experimental [Push/Wipe](docs/api/slide-transitions.md#push-and-wipe-extension) now adds all eight cardinal variants through Core, Layout, source import and native output. Playback/interruption and application acceptance remain open; ADV-05.1.2 stays incomplete.

Animation timing research now has an [executable finite-time contract](docs/architecture/decisions/0013-finite-motion-time-research.md) for nested sequence/parallel, integer repeat, boundaries and absolute seek, with hand-calculated and seeded-oracle tests. This is ADV-06.1.1 research progress; public target/events, native effects, general repeat/composition and playback remain planned.

[Target/event research](docs/architecture/decisions/0014-animation-target-event-research.md) now adds stable page/object tuples and explicit foreground gate experiments, with independently calculated activation times. Native playback, independent trigger tracks, subtargets and structural remapping remain planned.

The experimental [direct SmartArt text comparison](docs/api/diagram-text-comparison.md) reports bounded character agreement/disagreement. [Fifteen controlled cache cases](docs/testing/smartart-stale-cache.md) retain LibreOffice stored-cache behavior; [PowerPoint UI evidence](docs/testing/smartart-stale-cache-powerpoint-evidence.json) adds 24 saved outputs with source-driven cache rebuilding, three layout-change/return cases and three conflict rounds per family. Cross-application behavior differs; general freshness, full layout/style/font policies and ADV-03.7 remain open.

The installed [read-only CLI](docs/api/cli-inspection.md) now exposes inspect/compare with versioned JSON, bounded file/output handling and public Parser reports. Complete task/evidence-linked capabilities and authoring/save commands remain open under ECO-09.

The [versioned capability observation view](docs/api/import-capabilities.md) adds complete seven-operation rows with explicit missing/conflicting evidence, exact source queries and bounded task/artifact references. It does not replace complete property inventory or candidate/native validation.

### Existing chart numeric source saves — 2026-10-26

The explicit [existing-values-v1 profile](docs/api/chart-value-save.md) now stages synchronized cache/XLSX existing numeric edits, preserving unrelated payloads and rejecting shared/aliased/formula/external/conflicting sources. Independent three-round Node/browser checks and exact candidate rebinding cover this slice. IMP-09.2.6.1/.2 have partial implementation; structural edits, complete reference/native gates and actual Edit Data remain open.

The [independent/native follow-up](docs/testing/chart-value-save.md) now verifies three source-save rounds with Python cache/XLSX and exact unrelated payload checks, plus named LibreOffice numeric-cache retention and four-page PDF reopen. LibreOffice removes referenced XLSX sources in these resaves, so subsequent numeric writes remain blocked. Interactive Edit Data and PowerPoint gates remain open.

### Finite timeline arithmetic — 2026-10-26

Core now exposes detached [finite timeline compile/sample utilities](docs/api/finite-timeline.md): integer span/group timing, algebraic repeats, exact boundaries/overflow, immutable progress, bounded inputs and validated JSON schedules. Independent Core/browser/installed consumers cover the numerical slice. Drawing targets, authored/native effects, events, IR/format integration and actual playback remain unimplemented under ADV-06/07.

### Foreground event timeline replay — 2026-10-26

Core now exposes [detached event compile/replay utilities](docs/api/event-timeline.md) with page/visit identity, current target guards, early-click rejection, ordered one-row consumption and immutable replay observations. Finite schedules share aggregate budgets and JSON validation. Authored target/IR/format/copy/delete transport, clock/controller behavior and actual native playback remain open.

### Slide-bound foreground targets — 2026-10-26

[Owned stage authoring and target transport](docs/api/slide-event-timeline.md) now spans Core, IR v2, JSON and Layout, including whole-slide remapping, atomic identity-preserving replacement and referenced deletion guards. Native timing output remains rejected unless explicit static omission reports the loss. Static SVG and timed-pagination boundaries are explicit; native adapters/playback and full source structural operations remain pending.


### Native timing source inventory — 2026-10-26

The opt-in [Parser inventory](docs/api/native-timing-inventory.md) now exposes bounded raw/effective timing observations and source target candidates, with browser and independent XML checks. This advances the source-comparison prerequisites of ADV-06.2.1.3 / ADV-06.1.2.3; native mapping, timing edits and actual playback remain open.


### Stable paragraph-target source saves — 2026-10-26

The opt-in [paragraph-stable-v1 profile](docs/api/paragraph-target-save.md) now preserves proven paragraph-index targets during existing scalar text edits, with raw owner/paragraph proofs, independent three-round source checks and browser/installed consumers. Native headless resave changes the sample range even on the unedited control; actual playback and character/structural remapping remain open under ADV-07.8 / ADV-06.


### Atomic table structure edits — 2026-10-26

[Core editTable](docs/api/table-editing.md) now returns atomic row/column and merge/split candidates with explicit content/frame ownership and complete source-anchor mappings, committed through replaceElement. Three fresh output rounds, independent physical-grid and browser/native headless checks cover the current model; wide mixed-merge continuation serialization is fixed. AUT-05.1/05.2 and source structure preservation remain open at full scope.

Direct [table borders](docs/api/table-borders.md) now support sparse cell edges, deterministic shared Layout conflict resolution, exact SVG cells/lines, direct import and fresh native merged-perimeter hints. Independent source/browser checks pass; named native width/alpha/transform observations remain partial. Full theme/table-style precedence, interactive table editing and source style/structure writes remain open under AUT-05 / IMP-06 / IMP-09.

[Regional table authoring](docs/api/table-style.md) now resolves whole-table/banded-row-column/edge/corner/text defaults in Core, including implicit blanks and merged boundaries. Structural edits recompute regional styles and Layout retains effective source-row formatting during pagination. Fresh PPTX materializes direct cells with an explicit warning; named native rule retention, source theme/table-style inheritance, direction/provenance and source style writes remain open.

[Source table-style projection](docs/api/table-style-import.md) now selects inline, registry or default definitions, materializes supported direct region/theme fields and records bounded source dependencies. Two-master scope isolation and three source text-save rounds are verified. Named native fill controls pass after correcting dual-band order; conditional/default appearance, other matrix paints/effects, full native editing/reuse and style-preserving writes remain separate open requirements.

The [matrix extension](docs/architecture/decisions/0044-table-theme-matrix-references.md) adds positional solid/noFill fillRef/lnRef resolution and lazy placeholder color composition. Independent source tests pass; five [native controls](docs/testing/theme-matrix.md) retain background/transform/line gaps and one application failure. Four [concrete continuation cards](docs/planning/advanced-capability-todos.md#表格主题矩阵之后的具体执行卡) separate default/script text, further matrix properties, PowerPoint/region semantics and guarded direct style/metric source writes. Parent tasks remain open.

The [table text extension](docs/architecture/decisions/0045-table-text-appearance-and-uniform-fonts.md) adds reference default colors and equal explicit script-font collections, retaining source dependencies and scalar save boundaries. Native controls expose reference-only color differences and font substitution. IMP-06.1.4.1 continues with omitted defaults, mixed-family/script selection and PowerPoint appearance; this slice does not close the full font/default-appearance tasks.

[Explicit script-font authoring](docs/api/script-fonts.md) now carries separate East Asian/complex-script families through Core, shared Unicode selection, measurement fragments, SVG and distinct native fresh-output slots. Legacy IR remains compatible, and source no-op retains original slots. The default source projection retains its legacy multi-family blocker; the explicit source profile below admits independently mapped slots. [Native observations](docs/testing/script-fonts.md) retain font substitution, run splitting and a missing visual CJK control; full language/shaping/native tasks remain open.

An explicit [source font profile](docs/api/script-font-import.md) now maps Latin/ea/cs fields independently for supported drawings, tables and direct notes, traces missing theme slots and scopes inherited audit state by target theme. Three mixed-owner scalar source saves and three fresh semantic rounds are verified; font writes remain blocked. Legacy default behavior stays compatible. Language/supplemental selection, full notes inheritance, installed-font/shaping evidence and PowerPoint editing remain open.

The [script font map extension](docs/architecture/decisions/0048-script-font-maps-and-language-selection.md) now provides per-script authoring, explicit Han language hints, shared measurement/SVG selection and warned fresh-PPTX materialization. The [new leaf queue](docs/planning/advanced-capability-todos.md#脚本字体映射的接续任务) separately tracks source supplemental transport, independent preservation evidence, real-font shaping and native theme/edit acceptance; parent font tasks remain open.

The experimental [source supplemental v2 profile](docs/api/supplemental-font-import.md) now transports admitted theme/table script rules after native-slot inheritance, with bounded per-key provenance and captured-candidate scalar-save checks. Both dialects, mixed owners, browser/installed consumers and [four-package independent source evidence](docs/testing/supplemental-font-import.md) cover this slice. Full native selection/shaping, notes themes and source font writes remain open.

Experimental [notes-master text inheritance](docs/api/notes-style-import.md) now resolves notesStyle, type-matched body list defaults and an independent notes theme without borrowing slide fonts/palettes. Both dialects, scope/override/blocked cases, three scalar saves and [independent source evidence](docs/testing/notes-style-import.md) cover the text slice. Full notes-page layout, native typography/editing and source style writes remain open.

[Installed-font evidence](docs/testing/real-fonts.md) now verifies eight regular-face shaping/measurement samples with actual cmap, HarfBuzz and Chromium PostScript identities, plus located coverage/style/fallback controls. Twenty PowerPoint font controls retain selected declarations and identical PDF rasters across three native save/reopen rounds; five selection differences remain explicit. [Concrete follow-up cards](docs/planning/advanced-capability-todos.md#真实字体证据之后的执行卡) separate native precedence, exact-face SVG, mixed bidi, emoji and native text/style/theme/notes editing. Full font tasks and source font writes remain open.

Another [32 paired PowerPoint controls](docs/testing/real-fonts.md#paired-native-precedence-controls) isolate short-language tags, absent/empty EA-CS, covered cross-slot references, changed supplemental maps and major collections. They expose native differences in v2's absent-slot heuristic and distinguish `ko` from `ko-KR` despite verified glyph coverage. Five mutated ZIP controls now also verify actual presentation order/theme chains. Conservative profile correction, table/notes coverage and native editing remain open; no production selection rule changed in this research batch.

The separate [declared-slot v3 font profile](docs/api/declared-slot-font-import.md) now excludes the absent-EA/CS Latin-dictionary heuristic while preserving older explicit profiles. Drawings, table choices, independent notes and empty insertion styles share located mapped/unsupported provenance, captured options and complete candidate-save validation. [Independent source evidence](docs/testing/declared-slot-font-import.md) verifies four packages with six mapped and 36 unresolved rules plus exact scalar-only preservation. Implicit native font resolution, full table/notes/native appearance and style/theme editing remain open.

[Actual SVG font controls](docs/testing/real-font-svg.md) now verify 16 native/mixed-run cases with unique aliases, actual custom PostScript faces, HarfBuzz advances and independent inline/image raster checks. [ADR0052](docs/architecture/decisions/0052-svg-text-boundaries-and-font-names.md) fixes no-wrap, hard-break/empty-line geometry, grapheme boundaries and literal text/chart font names without new public options. Full shaped wrapping, bidi/emoji, fresh-PPTX native appearance and native editing remain separate open gates.

[Fresh author-generated PowerPoint controls](docs/testing/real-font-fresh.md) now verify 16 direct-slot/map-materialized regular samples using actual PDF PostScript faces, visible body glyphs and independent ZIP/XML checks. Three native save/close/reopen rounds retain text, language, slots, run topology and exact page rasters. Five actual ZIP/PDF corruption controls are rejected. Native text/style/theme editing, table/notes and full font capability remain open.

[SC native editing evidence](docs/testing/real-font-edits.md) now covers appending Greek text, actual W6 bold, W3 synthetic italic shear and a changed Office theme, with three native save/close/reopen rounds per operation. Sixteen native artifacts retain admitted Parser projections and byte-exact no-op saves; ten actual ZIP/PDF mutations are rejected. Unused author-map rules do not return after materialization. Other scripts, mixed text, table/notes and complete native font editing remain open.

[Five more native script edit controls](docs/testing/real-font-multi.md) cover TC/Japanese/Korean/Arabic/Hebrew text additions, full-body bold/italic and Office theme changes, each followed by three native reopen rounds. Actual Cambria fallback, synthetic shear, native language changes and RTL PDF ToUnicode differences remain explicit. Sixteen Parser byte-exact no-op saves and 43 actual artifact mutations verify this slice. Original mixed-script paragraph materialization/local editing, full bidi/mark shaping, table/notes and parent font tasks remain open.

[Original paired mixed-script paragraphs](docs/testing/real-font-mixed.md) now verify six author runs materialized into three concrete families, followed by local Latin insertion/bold/italic and Office theme changes with three native reopen rounds each. Neighbor glyph sequences remain stable; 72 actual artifact/native-selection counterexamples and 17 exact Parser no-op saves verify the bounded slice. Full seven-script same-paragraph combinations, bidi/mark correctness, table/notes and parent font tasks remain open.

[The complete same-paragraph font combination](docs/testing/real-font-combined.md) now adds Latin/SC/TC/Japanese/Korean/Arabic/Hebrew and Greek in one original paragraph, with local editing/style/theme operations and three native reopen rounds each. Source/content/visual order and native proofing normalization are explicit; per-span visibility, CJK font-binding swaps, 55 actual negative controls and 17 exact Parser no-op saves verify this gate. The specific mixed-paragraph task is complete; parent appearance/table/notes dependencies, full bidi/mark conformance and source style writes remain open.

[Native table font controls](docs/testing/native-table-fonts.md) now compare major/minor references and explicit collections with empty versus filled generic slots across 36 cases and three native reopen rounds per group. Twenty declared selections match; sixteen native differences remain explicit, including short Korean tags and explicit-collection behavior. Source rules remain retained; eight exact Parser no-op saves and 79 artifact mutations verify the evidence. Native table editing, source/appearance policy and independent notes themes remain open.

[Native table Office theme evidence](docs/testing/native-table-font-themes.md) now covers four groups, local source overrides and three actual save/reopen rounds, with 216 themed PDF pages. References follow the new theme; explicit collections retain source rules and body rasters. Named baseline shifts, subset GID renumbering and ToUnicode differences remain explicit. Sixteen exact Node no-ops, four browser checks, 36 scalar candidates and 484 artifact mutations verify the slice. Table append/style/theme candidate evidence is consolidated; independent notes and parent font/round-trip gates remain open.

[Independent notes font baselines](docs/testing/native-notes-fonts.md) now cover two source-patched notes themes and 30 controls with three actual PowerPoint save/reopen rounds and Notes-print PDFs. Twenty-four declared selections match; six short-Korean/empty/bad-slot differences remain explicit. Ninety native page records, 226 artifact mutations, eight exact Node no-ops, two browser checks and 18 scalar candidates verify the bounded baseline. The subsequent bounded append/local-style evidence is linked below; theme response, blank slide-image placeholders, full page appearance and parent typography/round-trip gates remain open.

[Two-run notes source regressions](docs/testing/native-notes-font-edits.md) now cover append and localized bold/pure-italic/regular declarations across two themes, three public consumers and three scalar-save rounds (72 save/no-op paths). Font maps, language, owner scopes, exact non-character bytes and guarded style rejection are verified. Native append and Ω-only bold/pure-italic/restored-regular now have three actual save/close/reopen/Notes-print rounds per group/stage: 360 page records, 24 exact Node no-ops and 88 actual ZIP/PDF counterexamples. Original glyphs/positions and outside-Ω rasters remain stable; restored-regular full rasters match the append baseline. Six of 22 appended declared font selections match native output; sixteen differences have a separate language/slot policy task. Theme response, native consumer candidate consolidation, complete notes imagery and source style writes remain open.

[Notes-domain cardinality and isolation evidence](docs/testing/notes-master-isolation.md) now distinguishes a single notes master with per-notes fontScheme overrides from diagnosed non-conforming extra-master source. Two listed notes-master IDs still reject. Interleaved A/B/A edits agree across three public consumers (18 scalar/no-op paths); eight materialized source states and 80 repacked ZIP mutations have independent checks. This corrects the planning assumption about standard multi-notes-master authoring. Full schema validation, native theme response and parent notes/round-trip gates remain open.

[Native notes Office-theme evidence](docs/testing/native-notes-font-themes.md) now covers A/B independent notes themes and one master with distinct per-notes font overrides: 180 themed Notes pages, 90 new baseline pages and 30 first-open pages, plus 240 separate ordinary-slide PDF pages. Notes source payloads remain exact; native header kerning and slide Y shifts remain explicit. Fifteen Node plus three browser no-ops, 45 headless scalar candidates and 73 saved artifact mutations verify the slice. Shared override-part source prompts repair in the named build; no repair was performed. Native editing consolidation, full notes imagery/fields, general font equivalence and source style writes remain open.

[Notes source inventory / ADR0056](docs/api/notes-source-inventory.md) advances source-only roles, geometry/artwork and cached fields; full notes rendering, field evaluation and authoring remain open.

[Multiple authored master foundation](docs/api/multi-master-authoring.md) now connects owner themes/backgrounds through Core, Layout, SVG and native output; full artwork/template/native operation gates remain pending.

[Shared definition editing](docs/testing/shared-definition-edit.md) now supports master creation/replacement and atomic shared layout replacement without expanding the legacy structural interface. Complete property origins, artwork and native admission of this new authoring sequence remain pending.

[Theme/background origins](docs/testing/inheritance-origins.md) now records actual per-property declarations and defaults alongside ordinary IR through the shared Core resolver. Element/preset/notes origins, artwork and full reset/template gates remain open; the parent scope is unchanged.

[Text formatting origins/native fixes](docs/testing/text-format-origins.md) now cover text/shape/placeholder selections, font-clear barriers, explicit native off flags and visible layout text. A real prompt-identity failure is retained, with corrected three-round/native negative controls. Table/chart/notes origins, full live template inheritance, artwork and general typography remain open.

[Chart/direct-notes format origins](docs/testing/chart-notes-origins.md) now trace the effective chart common/title/legend/axis text and direct authored notes through shared selectors, retaining explicit defaults/null clears and scoped declaration paths. Complete Node/ESM/global output, independent XML and installed consumer checks cover this slice; table-region origins, authored notes-master/theme graphs, full template/artwork/field and parent requirements remain open.

[Table authoring format origins](docs/testing/table-format-origins.md) now cover actual regional/preset/cell merges, all supported logical cell style/paragraph/run leaves, clear barriers, implicit blanks and aliased input occurrences. Full consumer/XML checks distinguish logical border candidates from Layout shared-edge winners. Broader geometry/content origins, native reusable styles, source style writes and original template/artwork/field parent tasks remain open.

[Ordinary master artwork](docs/testing/master-artwork.md) now has owned collections, atomic replacement, shared-resource protection, separate layout/page visibility, resolved master geometry and native owner output. Six actual first-open/eighteen reopened pages verify named shape/text/image/group visibility, geometry, media and layer order, with 19 corrupt file controls. Advanced master table/chart/connector native appearance and editing, complete reset/rebinding, source import/writes and field/notes parent gates remain open.

The [presentation-format metadata correction](docs/testing/presentation-format.md) now labels fresh rounded 16:9/4:3/16:10 frames accurately and uses Custom for other ratios. Eight automated three-consumer scenes and a separate corrected 4:3 six-page source with three actual PowerPoint reopen rounds verify the metadata-only change. Historical artwork evidence is preserved; this does not close advanced artwork/template/field parent tasks.

Advanced master native checks found and [repaired cell property serialization](docs/testing/table-cell-properties.md): borders precede fill, and cell margins/anchor belong to tcPr. Isolated controls and a separate corrected source verify the actual appearance/inset correction over three reopen rounds. First-save chart styling and shared artwork editing remain open; parent task status is unchanged.

[Named shared master editing](docs/testing/master-shared-edit.md) now links public Core replacement to actual native ordinary-heading editing: two visible red pages update, hidden/blue pages remain exact, and three native reopen rounds plus a fresh Core native print agree. Six positive master-chart first-save controls have explicit XML review and eleven corrupt ZIP tests. General object/data/style editing and reset/rebinding/fields remain open; parent tasks are unchanged.

[Reviewed authoring layout rebinding](docs/testing/layout-rebind.md) now provides logical key/unique-role or explicit decisions, local preserve/adopt/reset policies and one multi-page identity transaction through managed documents. Plans recheck full authoring values, report ambiguity/staleness and retain source-write boundaries. Automated/consumer checks and a named three-page native preserve/reset/unselected control (all first-save parts, three reopen rounds, 29 ZIP/seven PDF negatives) advance AUT-07.2; complete template/index/field/cross-document/native scope remains open.


[Explicit authored placeholder indices](docs/testing/placeholder-index.md) now complete AUT-07.2.1 matching policy: key/index/role priority, review conflicts, mixed numeric allocation and actual zero/max three-round controls, with 23 ZIP/four PDF negatives and fifteen actual-file source-only no-ops. Full reset/fields/general geometry, cross-document mapping and native source projection remain in their original open tasks.

The opt-in native text profile now emits only selected eligible layout defaults. A [color/bold control](docs/testing/native-minimal-color-bold-evidence.json) matches the materialized baseline in three-page pixels, glyph traces, drawings and images through three native saves; a [font control](docs/testing/native-minimal-font-evidence.json) retains the known shared-layout metric shift. [Five further run properties](docs/testing/native-run-expanded-evidence.json) have a named three-round native control; [Seven paragraph properties](docs/testing/native-paragraph-evidence.json) have a named native three-round control; font-family metrics, script slots, opacity, frames, fields and full native rebinding remain open.
