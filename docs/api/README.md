# API Reference

This section documents the public package contracts currently implemented by PPTKit. It describes callable APIs, inputs, results, defaults, validation, and runtime boundaries.

## Packages

- [Authoring inheritance origins](inheritance-origins.md) — immutable theme/background declaration sources alongside unchanged normalized IR.
- [Text formatting origins](text-format-origins.md) — placeholder/preset/local sources, optional-font clear barriers and native false-value serialization.
- [Explicit field evaluation](field-evaluation.md) — detached page number, date and footer values using caller-supplied context.
- [Missing description inspection](accessibility-descriptions.md) — explicit authored image, chart and table description review with owner identities.
- [Text contrast observations](accessibility-contrast.md) — conservative numeric checks and indeterminate scene reporting.

- [`@pptkit/core`](core.md) — authoring state, elements, assets, themes, layouts, validation, and Canonical IR.
- [`@pptkit/layout`](layout.md) — connector and image geometry resolution over normalized documents.
- [`@pptkit/pptx-exporter`](pptx-exporter.md) — browser-neutral PPTX generation and Node.js file output.
- [`@pptkit/svg-renderer`](svg-renderer.md) — browser-oriented SVG previews from Core/Layout state.
- [`@pptkit/pdf-renderer`](pdf-renderer.md) — PDF bytes using a caller-owned Chromium browser.
- [`@pptkit/media-inspection`](media-inspection.md) — bounded, environment-neutral media structure checks.
- [Browser PNG/JPEG output](browser-raster.md) — bounded browser-only raster conversion, resource checks and limitations.
- [SVG selection frames](svg-selection.md) — detached transformed frames and rectangular hit testing.
- [SVG preview invalidation](svg-invalidation.md) — conservative per-slide SVG string reuse planning.
- [Serializable edit operations](edit-operations.md) — value-checked atomic document edits and inverse operations.
- [`@pptkit/pptx-parser`](import-session.md) — bounded inspection and source-preserving import sessions; its Node entry adds local file import/output.
- [Package overview](package-overview.md) — package responsibilities and entry points at a glance.

`@pptkit/cli` provides [inspect/compare](cli-inspection.md), [JSON authoring and preview](cli-generation.md), and scoped [preservation saves](cli-save.md). The experimental PPTX parser exposes package inspection, partial structural/content projections, original-byte no-op sessions, and explicitly bounded edit-save profiles; the linked contracts describe their limits.

- [Byte assets and export resolution](byte-assets.md) — immutable byte inputs, image profiles, resolvers and cancellation.

- [Asset lifecycle](asset-lifecycle.md) — content matching, replacement, known references and explicit cleanup.

- [Copy references](copy-references.md) — slide/group identity remapping, connector targets and self-links.

- [Native straight connectors](native-connectors.md) — experimental native attachments, explicit path fallback and application evidence.

## Reading order

1. Start with the [Quick Start](../getting-started/quick-start.md).
2. Use the [Core overview](core.md) to find the relevant authoring topic.
3. Read [Layout](layout.md) when working with resolved geometry.
4. Read the [SVG renderer](svg-renderer.md) for browser preview and QA behavior.
5. Read the [PPTX exporter](pptx-exporter.md) for runtime, asset-loading, and output behavior.

All examples use the current pre-release contract. PPTKit does not maintain documentation for superseded unpublished APIs.

- [Browser font-backed measurement](browser-measurement.md) — explicit DOM adapter, font inventory and lifecycle.

- [Alignment, distribution and stacks](composition.md) — fixed-size composition through Layout options.

- [Measured table rows](table-sizing.md) — fixed/automatic rows, merged spans and overflow policy.

See [intrinsic sizing](intrinsic-sizing.md) for explicit content/parent dimensions and bounded group fitting.

See [grid composition](grid.md) for explicit bounded tracks and spans.

See [frame constraints](constraints.md) for bounded size/position references and atomic dependency resolution.

See [table pagination](table-pagination.md) for explicit cross-slide continuation and source row maps.

- [Experimental ZIP archive inspection](archive-inspection.md) — parser foundation; semantic import remains pending.

- [Experimental XML part inspection](xml-inspection.md): bounded namespace-aware XML metadata, encoding/error policy and Node/browser usage.

- [Experimental OPC package inspection](opc-inspection.md): content types, relationship graphs, logical part names and cumulative XML budgets.
- [OPC source comparison](package-comparison.md): exact payload and explicit conservative XML profiles, graph differences, bounded reports and independent corruption evidence.
- [OPC source relationship closure](preservation-closure.md): all MCE source branches, shared/cyclic dependencies, incoming owners, footprint budgets and scoped completeness.
- [Shape-text and notes preservation saves](text-owner-save.md): explicit V2 character ownership, separate notes provenance and scoped validation/evidence.
- [Table-anchor character preservation saves](table-text-save.md): explicit V3 scope, compact/physical merge identities, unchanged grid/unknown content and validated candidates.

- [Presentation entry inspection](presentation-inspection.md): main-part resolution, Strict/Transitional entry identification and document/template/slideshow classification.

- [Markup compatibility inspection](markup-compatibility.md): explicit namespace capabilities, alternate selection, ignorable content, opaque extension islands and OPC relationship integration.

- [PPTX structural inspection](pptx-inspection.md): ordered source slides, dimensions, hidden state, master/layout/notes/theme references, unresolved inheritance, document metadata and native section inspection.

- [Content import](content-import.md): explicit text/preset-shape/group subset, property provenance and reported limitations.
- [Placeholder dimensions](placeholder-inheritance.md): bounded source matching, dimension provenance and local-value projection.
- [Text style inheritance](text-style-inheritance.md): bounded per-property text cascade, source provenance and explicit unsupported-value behavior.
- [Text color transforms and alpha](text-color-transforms.md): ordered tint/shade/HSL/alpha projection, run opacity, independent provenance budgets and recorded native differences.
- [Speaker notes import](notes-import.md): direct body rich text, notes editing, source reports and notes-page limits.
- [Table import](table-import.md): authored grids, rectangular merges, rich-text cells, direct styles and explicit source/save boundaries.

- [Import and source-preserving sessions](import-session.md): explicit Core projections, source-bound plans, original bytes, lifecycle and resource budgets; edit requires the separately selected profile and rebuild remains blocked.
- [Node.js file import and atomic output](file-save.md): stable source capture, process-local overwrite receipts, exclusive create, guarded replacement, temporary-file validation, commit warnings, concurrency and filesystem limits.
- [Core metadata edit-save](core-metadata-save.md): explicit bound core-string changes without implicit author, revision or date updates.
- [Slide visibility edit-save](slide-visibility-save.md): explicit source-bound hidden-state changes, shared-source guards and mixed edit limits.
- [Direct text-run edit-save](text-edit-save.md): explicit `text-runs-v1`, existing run characters, whole-plan guards, minimal XML/ZIP patches and independent candidate validation.
- [Picture replacement edit-save](picture-edit-save.md): explicit `replace-v1`, one source-bound image retarget, shared-media isolation and documented limitations.

- [Existing chart numeric preservation saves](chart-value-save.md): explicit opt-in cache/XLSX atomic writes, source eligibility, shared/alias rejection and candidate validation.
- [Experimental chart content import](chart-import.md): cache/workbook priority, bar/line/pie/doughnut data, approximate preview, resource limits and fresh-output boundaries.
- [Scatter charts](chart-scatter.md): independent X/Y data, numeric axes, fresh export, preview and current import limits.
- [Doughnut charts](chart-doughnut.md): multi-ring authoring, geometry, bounded import, source edits and native evidence.

- [Slide Fade transitions](slide-transitions.md): immutable page timing, exact milliseconds, native output/import, static preview and preservation boundaries.

- [SmartArt component inventory](smartart-inventory.md): effective source objects, four component associations, explicit caches and direct shared resources; no preview or node editing claim.

- [Cached SmartArt preview](diagram-preview.md): explicit content-mode cached shapes, visible unavailable placeholders, unknown freshness and source-save boundaries.

- [SmartArt source data graphs](diagram-data.md): direct points/connections, typed references, stored text and cumulative inspection budgets; no semantic node editing.

- [Explicit SmartArt mappings](diagram-mapping.md): presentation evidence and cache references normalized by actual data/cache pairs, without inferred text ownership.
- [Direct SmartArt text observations](diagram-text-comparison.md): opt-in exact paragraph comparisons with explicit associations and unknown freshness.

- [Versioned import capabilities](import-capabilities.md): exact subjects, complete operation rows, explicit unknown/conflict states and task/artifact references.

- [Finite timeline arithmetic](finite-timeline.md): detached integer-time schedules, algebraic repetition, exact boundaries, immutable random-access samples and transported-plan validation.

- [Foreground event timeline replay](event-timeline.md): explicit logs/visits/current inventories, deterministic gate consumption and frozen observations.

- [Slide-bound event targets](slide-event-timeline.md): owned authored programs, IR/Layout transport, copy/delete/replace guards and explicit static-output omission.


- [Native timing source inventory](native-timing-inventory.md): bounded raw/effective timing observations and conservative source candidates without authored mapping or playback.


- [Stable paragraph target saves](paragraph-target-save.md): explicit paragraph-index preservation for scalar text edits, with raw owner proofs and unchanged default guards.


- [Atomic table editing](table-editing.md): pure row/column insertion/deletion and merge/split transactions, explicit content/frame policies, complete anchor mappings and source-write boundaries.

- [Direct table borders](table-borders.md): sparse overrides, bounded shared edge resolution, merged-native perimeter rules and independent partial-native evidence.

- [Table region style authoring](table-style.md): deterministic whole/band/edge/corner/local cascade, shared effective cells and explicit materialized-native output.

- [Source table-style projection](table-style-import.md): inline/reference/default selection, independent theme components, solid/noFill matrix references, reference text colors and explicit uniform script-font collections with bounded provenance and native/source-write limits.
- [Explicit script font families](script-fonts.md): Core family slots, shared Unicode selection, concrete measurement fragments, SVG/PPTX output and source import boundaries.
- [Source script-font profile](script-font-import.md): opt-in independent font fields, scoped provenance, notes boundaries and captured-profile candidate validation.

- [Supplemental source fonts](supplemental-font-import.md): explicit v2 collection transport after slot inheritance, per-script source provenance, captured scalar-save validation and native limits.

- [Notes-master text inheritance](notes-style-import.md): explicit independent notes theme, type-based placeholder/list defaults, bounded provenance and captured scalar-save validation.

- [Declared-slot supplemental import](declared-slot-font-import.md): explicit v3 excludes missing-bucket Latin heuristics, reports unresolved source rules and retains candidate/save guards.

- [Notes page source inventory](notes-source-inventory.md): bounded read-only roles, raw geometry, artwork and cached fields; no evaluation or authored mapping.

- [Master artwork and visibility](master-artwork.md): shared collections, owner resources and separate layout/page display switches.
- [Multiple authored masters](multi-master-authoring.md): explicit layout ownership, theme/background domains and native graph output.

[Chart/notes formatting verification](../testing/chart-notes-origins.md) extends the [text origin contract](text-format-origins.md) to concrete chart selections and direct authored notes.

[Table cell formatting origins](table-format-origins.md) describe actual regional/preset/local declarations before shared physical border resolution.
