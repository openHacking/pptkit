# Package Boundaries

PPTKit packages are separated by cohesive responsibility, not by element type. The current dependency direction follows the Authoring Model → Canonical IR → Layout Result → format package flow described in [Document Models](document-models.md).

## Implemented packages

### `@pptkit/core`

Owns format-independent authoring and shared document contracts:

- presentation, slide, element, asset, theme, layout, and placeholder types
- method-managed state and document-wide identity
- complete validation diagnostics
- Canonical Presentation IR v1 normalization

Core contains no layout resolution, file I/O, network access, binary asset loading, XML, ZIP, or PPTX relationship behavior.

### `@pptkit/layout`

Owns deterministic geometric transformations over normalized presentation data:

- authoring and normalized entry points
- detached layout result contracts
- connector anchor and bound resolution
- image contain/cover resolution
- recursive group processing

Future constraints, measurement, overflow, and pagination belong here when implemented. Layout does not read/write files or serialize OOXML.

### `@pptkit/pptx-exporter`

Owns editable PPTX output:

- one-time Core normalization and Layout orchestration
- runtime-specific asset loading
- OOXML parts, relationships, content types, and metadata
- ZIP package generation
- browser-neutral bytes and Node file output
- recoverable export warnings

The exporter translates explicit Core/Layout semantics and does not own authoring defaults or public presentation types.

### `@pptkit/svg-renderer`

Owns browser-oriented preview transformation over detached Layout state:

- per-slide standalone SVG serialization
- SVG geometry, connector, image, crop, and group rendering
- native SVG simple text with deterministic wrapping, plus XHTML `foreignObject`
  fallback for mixed-run rich text and tables
- runtime-neutral asset URL resolution hooks and fidelity warnings

The renderer has no filesystem, framework, editor, PPTX, OOXML, ZIP, or parser concerns.
It depends only on public Core and Layout exports and never becomes authoritative state.

### `@pptkit/presentation-workflow`

Owns browser-neutral guided-deck contracts and deterministic presentation recipes:

- versioned deck sessions, source evidence, asset references, and build reports
- curated theme/slide-role authoring used by browser and Node skill runtimes
- byte-oriented source parsing orchestration and PPTX package inspection
- deck and layout quality checks shared by runtime adapters

It depends on public Core contracts and contains no filesystem, process, browser UI, or
network behavior. Browser `File`/IndexedDB and Node path/file adapters remain outside it.

### `@pptkit/cli`

Owns thin command-line orchestration. Its current public surface is minimal; future generation and inspection commands must call package APIs rather than reimplement document, layout, parser, or exporter logic.

## Planned packages

### `@pptkit/pptx-parser`

Will own PPTX package inspection, supported semantic conversion, diagnostics, and preservation/degradation of unknown structures. It must not depend on exporter internals as if parse and export were exact inverses.

### `@pptkit/svg-parser`

Will own directional SVG parsing without duplicating the renderer or turning Core into an SVG DOM.

## Dependency rules

- Dependencies flow from format-independent contracts toward specialized transformations and side effects.
- Packages import other packages through public exports, never `dist` or private source paths.
- Public entry points orchestrate; internal modules own detailed transformations and side effects.
- Core authoring state remains inside its document/registry layers; validation and normalization remain pure.
- Layout and normalized output never share mutable state with their inputs.
- CLI, UI, editor, and framework concerns stay outside Core.
- A feature is not available until its required Core, Layout, exporter/fallback, tests, examples, and docs are complete.

## Anti-patterns

- packages split by text/image/shape/table object category
- raw OOXML or relationship IDs in public authoring types
- exporter-only fonts, colors, borders, or layout defaults
- parser preservation fragments exposed as ordinary editable elements
- browser editor state built directly on raw PPTX package structures
- private cross-package imports or circular dependency shortcuts

These rules keep changes reviewable and make future formats possible without rewriting the authoring foundation.
