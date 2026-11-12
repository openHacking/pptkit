# Core Authoring Model

`@pptkit/core` separates ergonomic, method-managed authoring state from immutable-in-practice, fully explicit Canonical Presentation IR. This document explains why the separation exists and which invariants Core owns.

## Why Core exists

Applications need a format-independent way to describe presentations without constructing OOXML parts, relationship IDs, ZIP paths, or renderer-specific objects. Core is that semantic boundary.

Core owns:

- presentation, slide, element, asset, layout, and placeholder contracts
- document-level stable identity and object ownership
- method-managed authoring operations
- native chart elements (bar, line, pie, doughnut) with categories and series; Cartesian charts also expose axes
- format-independent themes, paints, rich text, actions, and accessibility
- document-level named text style presets for reusable frame, paragraph, and run defaults
- deterministic intrinsic height estimation for text boxes with a fixed width
- complete validation and normalization into Canonical IR v2

Core does not own:

- filesystem or network access
- filesystem/network-backed binary loading and image decoding
- browser/native font measurement, automatic placement, overflow, or pagination
- PPTX parts, relationships, XML, ZIP encoding, or file output
- parser preservation fragments for unknown OOXML

Those responsibilities belong to Layout, exporters, runtime adapters, and future parsers.

## Two representations with different jobs

### Authoring state

The authoring model favors ordinary application code. IDs and most defaults may be omitted, text may be a string, theme objects may be partial, and an element may derive its box from a placeholder.

State is changed through document and slide methods. Collection getters return readonly frozen snapshots. This keeps ownership checks and ID accounting inside Core instead of allowing arbitrary array mutations to create duplicate IDs, stale references, or detached children.

### Canonical Presentation IR

Canonical IR favors predictable downstream processing. Every ID, box, transform, style, crop, background, accessibility value, text paragraph, and inheritance result is explicit. It contains no authoring class instances and shares no mutable references with authoring state.

The exact contract is documented in [Canonical Presentation IR v2](canonical-ir-v2.md).

## Data flow

```text
application input
       │
       ▼
createPresentation()
       │
       ▼
method-managed authoring document
       │
       ├── validatePresentation() ──► complete diagnostics
       │
       ▼
normalizePresentation()
       │
       ▼
Canonical Presentation IR v2
       │
       ▼
@pptkit/layout
       │
       ├──────────────► preview/render paths (planned)
       ▼
@pptkit/pptx-exporter ──► editable .pptx
```

The exporter convenience API accepts an authoring document, but internally it normalizes exactly once and passes normalized data through Layout before serialization.

## Method-managed state

Presentations expose add, insert, move, remove, and duplicate operations for slides. Slides expose the same operations for elements.

This model provides four guarantees:

1. Duplicate explicit IDs are rejected before state changes.
2. Insert and move indices have one checked interpretation.
3. Removing an object releases every ID owned by its subtree.
4. Duplicating an object deep-copies content, generates fresh nested identities and remaps modeled internal references.

The returned slide and element objects are immutable views of accepted input. Metadata, theme, notes, layout definitions, assets, tags, and custom data are cloned and frozen when stored.

## Identity ownership

The document owns all IDs, including elements inside slides, layouts, and nested groups. Element IDs are globally unique within the document so actions and connectors can use stable references without relying on array positions.

Slide duplication assigns a new slide ID and regenerates every copied element ID. Group duplication regenerates the group and all descendants. A private two-pass copy prepares all new IDs before rewriting internal connector targets. Slide copies also remap self-slide actions in elements, rich text, shape/table text and notes. External targets, shared assets/layouts and custom JSON remain unchanged. See [copy reference rules](../api/copy-references.md). Removal releases the same subtree atomically. Assets and layouts also have stable document-level IDs, although their lifecycle APIs differ from slides and elements.

ID generation is an authoring convenience, not an IR fallback. By the time normalized data exists, every required identity is present.

## Ordering and coordinates

Slide and element array order is semantic:

- slide order is presentation order
- element order is drawing order, from back to front
- no parallel `zIndex` exists

All public geometry uses points. Group children use the group's local `coordinateSize`; Layout and exporters apply nested transforms. Connector endpoints may use points or stable element anchors. Core validates reference scope, while Layout resolves anchors to final points.

## Validation and failure model

Local operation errors fail immediately: duplicate IDs, unknown IDs, conflicting asset registrations, and invalid collection indices throw from the mutating method.

Cross-document semantic errors are exhaustive. `validatePresentation()` returns all diagnostics it can find, including invalid geometry and style ranges, missing assets/layouts/placeholders/slides, invalid connector references, malformed tables, and invalid theme values.

`normalizePresentation()` runs validation first. When any error exists, it throws one `PresentationValidationError` whose `diagnostics` collection contains all errors. It never emits a partially normalized document.

## Defaults and inheritance

Authoring input can remain concise, but ambiguity ends at normalization. Core materializes:

- metadata, size, theme roles, and fonts
- layout selection and effective backgrounds
- base element name, box, transform, opacity, hidden state, and accessibility
- text frame, paragraph, run, bullet, numbering, font, color, and language values
- text bodies embedded in shapes while preserving one shape identity
- optional explicit East Asian/complex-script font families, inherited independently with authoring null resets; shared deterministic selection does not change author run identities
- paints, strokes, image fit/crop, connector style, and table cell style

Text and visual precedence is fixed: local run/element values, paragraph/frame values, placeholder/layout values, presentation theme, then Core defaults. Normalized slides record background origin explicitly.

Layout and exporters may transform explicit semantics into output-specific forms, but they do not invent Core business defaults.

## Asset lifecycle boundary

Assets belong in Core because elements need stable references before layout or export. Core owns IDs, sources, MIME hints, source dimensions, accessibility metadata, and caller-controlled deduplication keys.

Path/URL sources remain descriptive; byte sources are immutable Base64 snapshots captured by pure registration conversion. The browser exporter fetches URLs; the Node adapter additionally reads paths. Loading failures are exporter warnings rather than Core validation results because they depend on environment and time.

Core also owns explicit stable-ID asset replacement, known image-reference inventory, guarded removal and opt-in byte content matching. Reference scans include nested groups, hidden content and every reusable layout definition. Caller-provided keep-lists protect IDs referenced outside this known graph; they do not import OOXML preservation state into Core. Source bytes and format relationships remain session-owned. See [asset lifecycle](../api/asset-lifecycle.md).

## Format boundaries and preservation

Core expresses common template and document semantics such as layouts, placeholders, notes, hidden slides, sections, tags, and custom JSON data. It never exposes OOXML part names or relationships.

A future PPTX parser will need a preservation model for unknown or non-editable package structures. That model belongs at the parser/package boundary and must not expand the authoring API into a mirror of OOXML. Parser preservation is architectural direction, not a current implemented feature.

## Adding a capability

A capability is available only when its vertical slice is complete:

1. Core authoring and normalized contracts express it without format leakage.
2. Validation defines invalid states and references.
3. Normalization materializes all defaults and detaches state.
4. Layout implements required geometry or explicitly documents that none is required.
5. Exporters either serialize it or report a documented unsupported outcome.
6. Focused tests, cross-package tests, examples, API reference, and architecture docs agree.

This rule prevents “modeled but not exported” features and exporter-only defaults. The chart element (bar, line, pie, doughnut) is a vertical slice: Core authoring and normalization, SVG rendering, and native OOXML chart export shipped together.

## Internal organization

Core is organized by responsibility rather than element type:

- `types` defines public semantic contracts
- `document` owns state and identity
- `validation` reports document problems without side effects
- `normalization` performs pure authoring-to-IR transformations
- `factories` exposes construction entry points
- `utils` contains state-independent helpers

The package root only exports the intentional public surface. Core must remain independent of Layout and every exporter.

### Text style presets

`PresentationInit.textStylePresets` is an immutable, document-level registry of
named partial text styles. Text elements, shape text bodies, placeholders, and
table cells may reference a preset by name. Presets do not inherit from one
another. During normalization, explicit local values override the preset,
which overrides placeholder values and then Core defaults. Preset names are an
authoring convenience and never appear in Canonical IR; the normalized result
contains fully materialized text styles.

## Optional slide-bound foreground timing

Slides retain a frozen authored [event stage tree](../api/slide-event-timeline.md), not caller inventory/log/clock state. Setter/init binds actual owned root/descendant IDs. Atomic replaceElement preserves root identity and validates timed subtree targets before registry commit; referenced deletion blocks. Whole-slide copies use the allocator map to remap page/object/trigger identities, while object-only copies keep the page-scoped program. Shared layout objects/subtargets and native effects are not admitted.


[ADR0040](decisions/0040-atomic-table-structure-edits.md) keeps table topology edits in a pure Core transformation. The complete candidate and original-anchor mapping are detached; `replaceElement` remains the sole document ownership commit. No native/source cell identity is introduced.

[ADR0042](decisions/0042-table-region-style-materialization.md) adds table-local region rules resolved by Core into effective cells. A bounded physical-grid proof materializes implicit blanks; edits recompute authoring regions, while Layout pagination retains resolved source-row formatting. The optional IR origin marker is not an unresolved rule registry or native style identity.

[ADR0057](decisions/0057-master-theme-ownership.md) adds optional immutable initializer masters and explicit layout ownership. Core resolves each bound theme/background once; Layout/Renderer/Exporter consume the same domain. Complete artwork, mutation/rebinding and source master import remain open.


[ADR0058](decisions/0058-reviewed-layout-rebinding.md) adds reviewed value plans and a factory document transaction for selected-page layout rebinding. Read-only `PresentationData`/`PresentationSlideData` keep pure normalization/validation independent of mutation methods. Ambiguous bindings remain review findings, current data is rechecked before apply, and one ownership transaction validates all selected slides. Shared artwork stays in target owners; complete reset/native/source gates remain open.
