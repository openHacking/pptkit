# Core Authoring Model

This document explains how `@pptkit/core` models presentation state for developers and how that state is prepared for downstream packages.

## Why `@pptkit/core` Exists

`@pptkit/core` is the developer-facing entry point for PPTKit.

It owns:

- presentation construction
- typed slide elements
- asset registration and references
- normalization into a shared, validation-friendly structure

It does not own:

- file export
- package-level PPTX structures
- binary asset loading
- final layout fidelity decisions

## Two Layers Inside Core

`@pptkit/core` intentionally exposes two related layers:

- `Authoring Model`
- `Normalized Document`

The authoring model stays ergonomic and mutable enough for ordinary application code.

The normalized document is the contract for downstream systems. It is copied, validated, and stripped of authoring-time ambiguity before layout or export consume it.

## Data Flow

```text
application input
      |
      v
createPresentation()
      |
      v
authoring document
      |
      v
normalizePresentation()
      |
      v
normalized presentation
      |
   +--+--+
   |     |
   v     v
layout  export
```

This keeps `@pptkit/layout` and `@pptkit/pptx-exporter` from depending on mutable authoring objects or ad hoc conventions.

## Authoring Model Responsibilities

The authoring model should remain easy to construct in application code:

- `createPresentation()` creates a document shell with defaults
- `addSlide()` appends typed slide content
- `registerAsset()` registers reusable assets for element references
- `getAsset()` allows lookups by id

The public element set is intentionally small in the first implementation:

- text
- image
- shape

That is enough to make the package useful without freezing a much broader API too early.

## Normalization Responsibilities

`normalizePresentation()` is the boundary that turns authoring state into shared document state.

Normalization currently does four things:

- applies defaults such as presentation size
- validates identifiers, geometry, and asset references
- deep-copies the document into downstream-safe structures
- produces a stable shape for layout and export packages

This separation is important because layout and export need predictability more than authoring convenience.

## Asset Lifecycle Boundary

Assets are part of the core model because slides need stable references before export starts.

Core owns:

- asset identity
- source metadata
- optional dimensions and alt text
- deduplication keys for repeat registration

Core does not own:

- reading local files
- downloading remote resources
- storing binary payloads
- hashing bytes
- packaging media into PPTX relationships and parts

Those concerns are downstream responsibilities because they are tied to environment and output format choices.

## Extension Direction

The current model is designed to grow without changing package boundaries:

- theme and style inheritance can expand from the existing style shells
- richer text can evolve from `text` elements into runs or spans
- additional element kinds can extend the union
- future binary asset support can be added as a new asset source mode if required

The key constraint is that new authoring features should still normalize into a stable downstream contract instead of leaking export details back into the public API.

## Internal Architecture Rules

Core is organized by responsibility rather than by element type:

- `types` defines public contracts only.
- `document` owns mutable authoring state and asset registration.
- `validation` and `normalization` are side-effect-free flows.
- `utils` provides state-independent helpers, while `factories` exposes stable construction entry points.
- `src/index.ts` only aggregates the public API; it contains no domain logic.

Normalization must detach every output reference from authoring state. Core must not add layout resolution, file IO, network downloading, or PPTX-specific concepts. Each new element or asset capability must include authoring types, normalization mapping, validation, smoke or unit coverage, and the relevant API or architecture documentation.
