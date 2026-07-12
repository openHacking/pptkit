# Canonical IR v0.1 Draft

This document defines the minimum canonical presentation shape needed to guide the first implementation steps.

It is intentionally narrower than the long-term document model.

## Status

Draft

## Goals

- Separate authoring convenience from normalized internal state.
- Give layout and export work a shared handoff contract.
- Avoid dragging OOXML-specific details into `@pptkit/core`.

## Minimum v0.1 Shape

At the bootstrap stage, the canonical IR should support only:

- presentation metadata with optional title
- an ordered slide list
- per-slide ordered element lists
- stable slide identifiers
- element `type` as an open string field until the first typed unions are ready

Representative shape:

```ts
interface CanonicalPresentation {
  title?: string;
  slides: CanonicalSlide[];
}

interface CanonicalSlide {
  id: string;
  elements: CanonicalElement[];
}

interface CanonicalElement {
  type: string;
}
```

## Layer Boundary

### Authoring Model

The authoring model may expose helpers such as `createPresentation()` and `addSlide()`.

It may keep ergonomic defaults and mutable construction APIs.

### Canonical IR

The canonical IR should be:

- explicit
- serializable
- validation-friendly
- independent from OOXML packaging

### PPTX Package Model

The PPTX package model should remain downstream of normalization and layout.

It must not become the source of truth for early authoring types.

## What v0.1 Explicitly Does Not Model

- exact geometry
- themes and slide masters
- text runs and paragraph richness
- image embedding metadata
- speaker notes
- animations and transitions
- parser preservation state
- OOXML relationship graphs

## Normalization Expectations

When normalization is implemented, it should at minimum:

- ensure every slide has an identifier
- preserve slide ordering
- ensure missing element lists normalize to empty arrays
- reject or warn on obviously invalid top-level shapes

## Follow-Up Questions

The next round of design work should answer:

- when `CanonicalElement` becomes a discriminated union
- whether geometry belongs directly in the canonical IR or in a post-layout structure
- how validation diagnostics are represented

