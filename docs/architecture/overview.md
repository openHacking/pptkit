# Architecture Overview

PPTKit is being designed as a multi-package presentation toolkit with explicit boundaries between authoring, normalization, layout, parsing, preview, and file export.

## High-Level Goal

Provide a clean developer-facing model for generating editable presentation output without forcing users to work directly with low-level PPT structures.

## Three-Layer Model

The architecture centers on three related but distinct layers:

- `Authoring Model`
- `Canonical Presentation IR`
- `PPTX Package Model`

These layers exist to keep user-facing APIs, normalized document behavior, and file-format concerns from collapsing into one abstraction.

## System View

At a high level, the system is expected to look like this:

```text
Structured Content
        |
        v
 Authoring Model
        |
        v
 normalize
        |
        v
Canonical Presentation IR
        |
   +-----+------+
   |            |
   v            v
preview path   export path
   |            |
   v            v
SVG output   PPTX Package Model
                |
                v
              .pptx
```

The reverse direction starts from package-level data:

```text
.pptx
  |
  v
PPTX Package Model
  |
  v
parse path
  |
  v
Canonical Presentation IR
```

## Package Mapping

The initial package vocabulary maps onto this flow:

- `@pptkit/core` owns the authoring-facing model and shared document structures.
- `@pptkit/layout` resolves layout against normalized document state.
- `@pptkit/pptx-exporter` owns translation from normalized structures into package output.
- `@pptkit/pptx-parser` owns translation from package input back into normalized structures.
- `@pptkit/svg-parser` and `@pptkit/svg-renderer` support directional SVG workflows without becoming a second PPTX exporter.
- `@pptkit/cli` provides local workflow entry points on top of these packages.

## Forward and Reverse Flows

### Authoring to Export

```text
Authoring API
   |
   v
Authoring Model
   |
   v
Normalization
   |
   v
Canonical Presentation IR
   |
   v
Layout Resolution
   |
   v
PPTX Export
   |
   v
PPTX Package Model
   |
   v
.pptx
```

### PPTX Parse to Normalized Structures

```text
.pptx
   |
   v
PPTX Package Model
   |
   v
Package Inspection
   |
   v
Normalization
   |
   v
Canonical Presentation IR
```

## Why This Framing Matters

This separation keeps the system from inheriting the main limitations of monolithic PPT generators:

- command-style APIs become the only authoring surface
- layout gets mixed into file writing logic
- parser and exporter assumptions leak into public models
- preview and editing flows become harder to add later

It also gives PPTKit room to support both editable output and high-fidelity preview workflows without pretending they are always the same problem.

## Architecture Priorities

- Clear package ownership
- Stable normalization boundary
- Strong separation between layout and file-format concerns
- Explicit preview, export, and parse paths
- Honest handling of fidelity tradeoffs

## Current Maturity

The implementation is not in place yet. These docs exist to keep early package and API decisions aligned before code hardens around the wrong abstractions.
