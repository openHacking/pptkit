# Preview and Editing

Browser preview and future editing flows should operate on normalized document state, not on raw PPTX package state.

## Core Flow

The conceptual split should be:

```text
Canonical Presentation IR
        |
   +----+----+
   |         |
   v         v
SVG renderer PPTX exporter
   |         |
   v         v
preview    delivery
```

In other words:

- `IR -> SVG renderer -> preview`
- `IR -> PPTX exporter -> delivery`

The implemented renderer resolves the IR through `@pptkit/layout` and emits standalone
per-slide SVG strings. SVG geometry handles shapes, connectors, images, groups, and
simple text with explicit PowerPoint-oriented lines and baselines. Browser XHTML
`foreignObject` remains the fallback for mixed-run rich text and tables. Consumers may display these
strings in a gallery, but must keep authoring/normalized state as the source of truth.

Application-level review sessions, persistence, transfer protocols, and download behavior
remain consumer concerns. The guided implementation maintained by PPTKit lives in the
external [PPTKit Presentation](https://github.com/openHacking/pptkit-presentation)
repository and consumes only the public engine APIs.

## Fidelity boundary

The preview is designed for modern-browser QA: missing images, clipping, overlap,
drawing order, transforms, table structure, and broad visual composition. Browser font
metrics differ from PowerPoint, and hybrid SVG is not a promise of portable SVG or
pixel-identical Office rendering. Renderer warnings are evidence to review, not edits to
the document.

## Why Not Edit Raw PPTX State

Raw package structures are too format-specific to serve as the main editing surface.

Using raw PPTX package state directly for preview or editing would make it harder to support:

- undo and redo
- structured editing
- AI-assisted modifications
- collaboration
- future non-PPTX outputs

It would also leak package-specific details into workflows that should operate on document semantics instead.

## Benefits of Editing IR

Using normalized document state for preview and editing gives the project a better foundation for:

- shared rendering logic
- deterministic updates
- future viewers
- diagnostics and validation
- cleaner separation between authored state and derived output

## Authored vs Derived State

One useful architecture lesson is the difference between authored state and derived state.

For PPTKit, the `Canonical Presentation IR` should be treated as authored or authoritative document state, while preview artifacts and exported packages should be treated as derived outputs.

That distinction helps prevent preview logic, exporter quirks, or package details from becoming the source of truth for the whole system.
