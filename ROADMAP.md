# PPTKit Roadmap

PPTKit is building a format-independent presentation document engine with editable PPTX output. This roadmap separates implemented foundations from release gates and longer-term capability tracks; it is not a promise that every later item ships in one release.

## Status legend

- **Implemented** — present in Core, Layout, Exporter, tests, and current documentation.
- **Next** — required for the first public preview or the next fidelity milestone.
- **Later** — intended direction after the relevant foundations are stable.
- **Deferred** — explicitly outside the current roadmap sequence.

## Implemented foundation

- **Core contract:** method-managed authoring state, readonly collections, document-wide IDs, complete diagnostics, and Canonical Presentation IR v1.
- **Editable content:** rich text, theme/direct colors, paints, strokes, actions, image fitting/crop, transforms, and accessibility metadata.
- **Structured elements:** preset shapes, anchored connectors, nested groups, editable tables with spans, notes, hidden slides, sections, tags, and custom data.
- **Templates:** presentation themes, reusable slide layouts, typed placeholders, placeholder binding, and explicit inheritance.
- **Layout:** detached resolution for connector anchors, connector bounds, image contain/cover, and nested group traversal.
- **PPTX export:** browser-neutral bytes, Node file output, native themes/layouts/tables/notes/media relationships, validation failures, and recoverable asset warnings.
- **Browser preview:** deterministic per-slide hybrid SVG from Layout state, all current element families, asset/fidelity warnings, workbench integration, and Chromium visual fixtures.

## Public preview gate — Next

The first public preview requires more than package publication:

- complete public API and architecture review across Core, Layout, Exporter, and CLI
- user-facing API reference, tutorials, troubleshooting, and runnable examples
- package manifests, exports, release automation, changelog policy, and preview versioning
- deterministic package fixtures, key XML assertions, and rendered visual regression coverage
- PowerPoint and LibreOffice open/save regression matrix across supported platforms
- explicit compatibility policy beginning with the first published preview
- performance and output-size baselines for representative decks

## Layout fidelity — Next

- pluggable text measurement and font metrics
- overflow and clipping diagnostics with actionable element paths
- alignment, gaps, stacking, grids, and constraint-based composition
- intrinsic sizing for text, images, groups, and tables
- table row measurement, overflow behavior, and pagination
- slide boundary checks and intentional cross-slide continuation

Layout remains deterministic and independently testable; exporters do not become an implicit layout engine.

## Import and round-trip — Later

- introduce `@pptkit/pptx-parser` and a format-aware package model
- classify parsed structures as editable, preserved, fallback, or unsupported
- preserve unknown/partially understood OOXML without leaking it into common authoring types
- define degradation diagnostics and round-trip fixtures
- establish scoped guarantees for open → inspect/edit → save workflows
- keep parser and exporter implementations independent rather than assuming symmetry

## Format and ecosystem — Later

- broader preset geometry plus custom paths
- SVG parsing and rendering boundaries
- higher-fidelity browser preview text measurement and cross-browser visual baselines
- higher-level authoring/layout recipes without expanding Core into a catch-all API
- practical CLI generation and inspection workflows
- plugin extension points with explicit ownership and capability negotiation
- example gallery, templates, framework integrations, and benchmark suites

## v1.0 readiness — Later

- stable and intentionally small public APIs
- documented compatibility and IR-versioning rules
- stable Core → Layout → Exporter dependency contracts
- tested failure, fallback, and warning semantics
- supported-runtime and cross-application compatibility matrix
- performance budgets and regression baselines
- complete user, API, architecture, migration, and maintainer documentation

## Deferred capabilities

These depend on stable identity, rich text, themes, groups, assets, layout, and preservation contracts. Each should ship as an independent vertical slice rather than a large flat options expansion.

- native charts and SmartArt
- audio and video
- transitions and animation timelines
- comments and collaborative review metadata
- OLE, controls, and embedded fonts
- presentation show, print, and advanced delivery settings

## Capability delivery rule

A roadmap item is marked implemented only when its public semantics, Core validation/normalization, required Layout behavior, PPTX export or documented fallback, tests, examples, API reference, and architecture documentation are complete.
