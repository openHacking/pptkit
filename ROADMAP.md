# PPTKit Roadmap

PPTKit is building a format-independent presentation document engine with editable PPTX output. This roadmap describes capability milestones, not a promise of complete Office parity or a single release containing every feature.

The [Implementation Backlog](docs/planning/implementation-backlog.md) contains task IDs, dependencies, ownership, and acceptance criteria. This file owns milestone order; the backlog owns task status and delivery details. Update both when scope or milestone order changes.

## Status legend

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
- **Layout:** detached resolution for connector anchors/bounds, image contain/cover, nested groups, and chart geometry. Core also provides deterministic text-height estimation.
- **PPTX export:** browser-neutral bytes, Node output, native themes/layouts/tables/notes/media relationships, validation failures, and recoverable asset warnings.
- **Browser preview:** per-slide hybrid SVG, current element families, asset/fidelity warnings, workbench integration, and Chromium visual fixtures.
- **Developer delivery:** package manifests/exports, a local npm release script with tests, API/getting-started documentation, typed examples, browser bundle checks, and CI. CLI commands remain a bootstrap shell.

These foundations do not establish a formal cross-application compatibility matrix, complete layout fidelity, or blanket round-trip guarantees.

## M0 — Reliability and release baseline — Next

Backlog: **REL-01–REL-06**.

Audit public contracts and capability status; establish PowerPoint/LibreOffice open-save fixtures, performance/output-size baselines, release/compatibility policy, consumer-package verification, and a traceable regression corpus. Extend existing documentation and automation rather than rebuilding them.

## M1 — Measured layout and composition — Next

Backlog: **LAY-01–LAY-08**.

Add pluggable text/font measurement, overflow and boundary diagnostics, alignment/stack/grid composition, intrinsic sizing, table measurement, explicit pagination/continuation, and preview/export layout consistency. Preserve deterministic defaults and isolate environment-backed measurement outside Core.

## M2 — Rich authoring and templates — Planned

Backlog: **AUT-01–AUT-09**.

Expand text and international typography, geometry/paths, paints/effects, image assets, tables, chart families, masters/layout inheritance, and slide reuse. Features ship as complete vertical slices, each with explicit preview and export behavior.

## M3 — PPTX import and scoped round-trip — Planned

Backlog: **IMP-01–IMP-10**.

Introduce a parser, bounded package inspection, supported semantic import, master/theme inheritance resolution, and preservation-aware saves. Classify every encountered feature as editable, preserved, fallback, or unsupported. Define separate guarantees for no-op saves, supported edits, and destructive operations; do not promise arbitrary PPTX fidelity.

Import begins after the contract audit and fixture corpus. Initial preservation work can proceed before broader M2 authoring capabilities; unsupported features remain classified until their semantic slices exist.

## M4 — SmartArt, media, motion, and review — Planned / Research

Backlog: **ADV-01–ADV-12**.

- SmartArt: preserve existing diagrams first, then support an explicit editable subset, then author selected diagram families.
- Media: audio/video assets, poster frames, playback semantics, and portable fallback behavior.
- Motion: slide transitions, element timelines, triggers, sequencing, then interactive preview playback.
- Review: comments, author/thread identity, accessibility coverage, and incremental change events for consumers. Real-time collaboration infrastructure belongs to applications.
- Advanced Office objects: inspect/preserve OLE and controls without executing them; research embedded fonts and advanced presentation settings.

Each feature declares native editability, preservation scope, preview limitations, and cross-application behavior. Unsupported advanced features remain visible in capability reports.

## M5 — Formats, tooling, and ecosystem — Planned

Backlog: **ECO-01–ECO-09**.

Add practical CLI workflows, directional SVG parsing, portable/raster/PDF output adapters, preview/editor integration contracts, plugin capability negotiation, engine examples, interoperability reporting, and developer inspection tools. Guided authoring products stay in the separate PPTKit Presentation repository.

## M6 — Stable v1.0 contracts — Planned

Backlog: **STB-01–STB-04**.

Stabilize deliberately selected public APIs and IR rules; publish tested runtime/application support, regression budgets, migration guidance, and maintainer procedures. v1.0 requires an explicit supported capability set, not completion of every advanced Office feature.

## Capability delivery rule

A task is complete only when its public semantics, validation/normalization, required Layout behavior, PPTX export or documented fallback, tests, examples, API reference, and architecture documentation are complete where applicable. Documentation-only and research tasks use their own stated acceptance criteria.

No feature is marked implemented solely because its type exists, an XML part is emitted, a preview looks correct, or an unknown part is copied. Compatibility claims require named fixtures and recorded application results.
