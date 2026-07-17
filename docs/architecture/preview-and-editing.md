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

The browser-first presentation workflow sends session JSON and binary assets through one
resumable, integrity-checked chunk protocol and stores the completed `DeckSessionV2` and
asset Blobs in IndexedDB. The session owns the semantic deck, extracted evidence, and
asset metadata; object URLs, SVG strings, PPTX bytes, and build reports are derived state.
A deployed static HTTPS application can therefore restore review state without uploading
presentation data or embedding binary data in the session JSON.

The preview publishes protocol compatibility, native-page API checks, and resumable
transfer progress as JSON in a hidden read-only DOM bridge. Keeping observation in the DOM
and mutation in the existing transfer controls lets browser hosts with isolated evaluation
contexts inspect the workflow without assuming access to the page's `window` object.
The mutation controls are progressively disclosed from a compact agent-connection trigger:
the payload form occupies no review-layout space after the session and declared assets are
available, while the trigger remains reachable for revisions and visible failure recovery.
This presentation-layer behavior does not change the transfer protocol or make the preview
state authoritative.

PPTX generation is deferred until the user requests a download. Previewing does not
allocate package bytes, and a failed package inspection withholds the PPTX while still
returning an actionable report.

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
