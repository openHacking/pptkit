# Package Boundaries

PPTKit packages are separated by cohesive responsibility, not by element type. The current dependency direction follows the Authoring Model → Canonical IR → Layout Result → format package flow described in [Document Models](document-models.md).

## Implemented packages

### `@pptkit/core`

Owns format-independent authoring and shared document contracts:

- presentation, slide, element, asset, theme, layout, and placeholder types
- method-managed state and document-wide identity
- complete validation diagnostics
- Canonical Presentation IR v2 normalization

Core contains no layout resolution, file I/O, network access, binary asset loading, XML, ZIP, or PPTX relationship behavior.

### `@pptkit/layout`

Owns deterministic geometric transformations over normalized presentation data:

- authoring and normalized entry points
- detached layout result contracts
- connector anchor and bound resolution
- image contain/cover resolution
- recursive group processing

Opt-in measurement analysis and overflow diagnostics live here. Future constraints, font-backed adapters, and pagination also belong here. Layout does not read/write files or serialize OOXML.

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
It depends only on public Core and Layout exports and never becomes authoritative state. The separate opt-in [`/media` browser adapter](../api/browser-media.md) owns explicitly resolved/mounted HTML media, local Blob URLs, playback events and disposal, and can explicitly bind a slide-local Layout media identity without positioning or navigation; the static SVG entry does not import it or initiate playback.

### `@pptkit/pdf-renderer`

Owns PDF byte output over public Layout and SVG renderer contracts. It accepts a
caller-owned Chromium browser, isolates each print in a fresh network-denied
context, and reports rendering limitations. It does not launch browsers, write
files or introduce a reverse dependency into SVG, Layout or Core.

### `@pptkit/media-inspection`

Owns environment-neutral, bounded media structure checks shared by output
adapters and the PPTX parser. It has no workspace, filesystem, network or format-writer dependency.
The parser's opt-in PCM WAVE source inventory, the existing PPTX JPEG, PNG/APNG and GIF inspections and PDF's opt-in JPEG and
static-image structure policies consume its public entry. Pixel decoding,
general animation semantics and complete image validation remain outside its
contract.

The public ISO BMFF inventory reads bounded top-level box framing and file-type brands only; a separate bounded movie-clock inspection is shared by Parser and Exporter and does not validate tracks; it does not infer a codec or playback. Parser's optional [video source inventory and projection](../api/video-source-inventory.md) retain direct, companion and poster relationship identities without fetching links; an explicit projection creates separate Core video/poster assets, SVG remains static, and an explicit experimental PPTX profile packages bounded MP4 sources. See [the framing contract](../api/iso-bmff-inspection.md).

Core owns format-independent audio and poster asset identities, immutable byte capture and reference/copy semantics. Layout carries the audio object without resolving playback. SVG renders only a static poster or placeholder; the PPTX writer defaults to rejection and offers an explicit bounded PCM WAVE profile with native media relationships and timing. See [audio authoring](../api/audio-inspection.md#authoring-and-static-preview).

### `@pptkit/cli`

Owns thin command-line orchestration, local regular-file capture and bounded/versioned report transport. Inspect/compare and create-only preservation save call public Parser APIs; JSON generation and preview use public Core, Exporter, Parser and Renderer APIs. No parsing or package-write implementation is duplicated. Broader authoring and edit commands must likewise use public package APIs. See [ADR0032](decisions/0032-read-only-cli-inspection.md) and the [save contract](../api/cli-save.md).

Guided presentation workflow, skill, preview UI, themes, and authoring recipes are
application-level consumer concerns. They live in the external
[PPTKit Presentation](https://github.com/openHacking/pptkit-presentation) repository
and depend only on the engine's public package APIs.

### `@pptkit/pptx-parser` (experimental)

The experimental package now owns bounded ZIP32 archive inventory, CRC/size checks and per-part namespace-aware XML inspection through public Node/browser entries. A bounded OPC graph inspector now owns logical part names, content types, relationship source/target checks and cumulative XML budgets. MCE projection and presentation entry/structure inspection add effective source ordering, dimensions and typed structural references. Structural Core projection and original-byte no-op sessions are implemented; partial content semantics and opt-in [existing text-run preservation profiles](../api/text-edit-save.md) are available. A separate `/node` adapter owns stable local-file capture and publication while the root/browser entries remain byte-only. Complete inheritance, full presentation extension policy and broader preservation reconciliation remain pending. Parser owns source-to-Core bindings, normalized change analysis, raw source guards and lexical XML/ZIP candidate validation; Core receives no source state. It uses fflate, a private pinned saxes adapter and the public Core entry, with no Exporter dependency yet. Browser bundling includes the XML dependency and notices; no XML library or Node types leak into public contracts. Preservation ownership follows [ADR 0001](decisions/0001-import-preservation-save.md), [ADR 0008](decisions/0008-direct-text-preservation-save.md) and [ADR 0011](decisions/0011-node-file-publication.md); parser code must not depend on exporter internals as if parse and export were exact inverses.

SmartArt source discovery and component relationship/root checks belong to Parser's structural inspection. [ADR0016](decisions/0016-smartart-component-inventory.md) keeps raw XML identities and associations outside Core/Layout; the optional public records provide inventory without a diagram authoring model, cached preview or native editing permission.

[ADR0020](decisions/0020-smartart-source-data-graph.md) extends Parser's format-owned inspection to direct data-model points/connections, typed references and stored text. Its presentation-specific options own cumulative graph budgets and are preserved through file capture/candidate validation. Core/Layout receive no source IDs, parser trees or node-edit authority; raw cache model references remain distinct from semantic visual mapping.

Parser's [explicit SmartArt mapping](decisions/0021-smartart-explicit-model-mapping.md) retains per-data-part presentation evidence and per-actual-pair cache references, reusing bounded identity indices and source scans. It does not add authoring IDs or layout policy to Core, create new packages, or permit diagram source writes.

## Planned packages

Minimum native SmartArt generation remains in standalone research scripts under
[ADR0023](decisions/0023-native-smartart-generation-research.md). The fresh-package
builder consumes public APIs and explicit offline component bytes; no production
package imports the fixture/presentation writers or borrowed layout resources.
Producer rendering evidence does not introduce a Core authoring model, source
writer, layout interpreter or native-edit admission.

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


## Injected analysis boundary

Layout accepts optional synchronous text/font providers and diagnostic options through its public entry points. Frozen detached requests cross this boundary; provider results produce detached measurements and warnings, without changing authored geometry or Core's deterministic height estimate. Layout owns coordinate transforms, margins/alignment and overflow comparisons; the provider owns font metrics and text shaping/wrapping. Core receives no font, filesystem, network or DOM dependency. The separate `@pptkit/layout/dom` adapter owns explicit browser measurement side effects; resolved line composition remains future work.

Exporter and Renderer accept Layout options on document-facing entry points and forward diagnostic identities/paths as warnings. Rendering an existing Layout result reuses its analysis. Default calls retain previous output; analysis does not claim native application layout parity. See the [Layout API](../api/layout.md#opt-in-measurement-and-diagnostics).

The [DOM measurement adapter](../api/browser-measurement.md) is an explicit runtime entry, not imported by the pure Layout entry. It consumes caller-owned loaded fonts and a document, creates a disposable isolated host, and returns synchronous providers. It has no implicit filesystem/network/font loading, and introduces no new package or dependency.

Fixed-size alignment/distribution/stack rules belong to Layout. They operate before connector/image/chart resolution and are passed through public export/preview options; Core remains unaware of composition algorithms. Explicit intrinsic sizing also belongs to Layout and runs between table measurement and composition; see [its contract](../api/intrinsic-sizing.md). Bounded grid placement also belongs to Layout; its pure track solver consumes resolved frames. Explicit bounded affine frame constraints also belong to Layout and resolve only within each composition rule; cyclic dependencies are rejected before any mutation.

Explicit table pagination is a Layout-owned structural transformation. It inserts detached continuation slides, preserves source IDs on the first fragment and exposes a source-row map through LayoutResult. Format writers consume resolved order/geometry. Asset registries and source authoring state are not duplicated or mutated; generic cross-document copy remains a separate Core capability.

Parser now depends on the public Core entry to construct the explicit structural projection and compare normalized snapshots. It owns original-source bytes, session/plan lifecycle and no-op save preparation. Its Node-only file adapter adds no reverse cross-package edge: it coordinates the public session lifecycle and local filesystem operations. Exporter has no reverse dependency; Parser → Exporter remains a planned edge for semantic rebuild. See [session ownership](../api/import-session.md) and the [Node file transaction](../api/file-save.md).


## Byte asset ownership

Core converts raw Uint8Array registration input into immutable, serializable Base64 using pure local transformation. It owns no image codec, file handle or network client. Exporter owns byte decoding, its PNG/JPEG/GIF header profile, resolver requests, stream/file budgets and cancellation. The separate media-inspection package owns the shared bounded JPEG structure profile used by exporter and PDF output. Renderer serializes supported bytes as data URLs and leaves image decoding to the browser. Layout continues to consume declared dimensions. See [byte asset contracts](../api/byte-assets.md).


## Native connector boundary

Layout retains authored connector target identities/anchors alongside resolved coordinates. These are format-independent provenance, not package IDs. Exporter translates the experimental native-straight profile using part-local preallocated drawing IDs and known preset site tables. Hidden, cross-part and unsupported sites fall back with diagnostics; cross-group targets outside the immediate coordinate scope fail validation, and default path serialization remains unchanged for valid geometry. No XML or new dependency enters Core/Layout. See [native profile](../api/native-connectors.md).

Parser resolves imported connector drawing IDs only after building the full slide/group forest, then maps supported same-scope targets to Core element IDs and anchors. Parser also maps direct supported click actions while keeping relationship and XML details in its own package. Unmapped endpoint references remain explicit findings with free cached geometry. Content-projection consumers request native attachment serialization explicitly; fresh path-mode exports serialize resolved connector geometry without attachment semantics.

## Finite time arithmetic

Core owns the pure [finite timeline utility](../api/finite-timeline.md): typed abstract spans/groups, immutable versioned schedules, bounded exact arithmetic and random-access samples. It introduces no authoring/Canonical IR/Layout fields and no clocks, listener/playback state, drawing targets or native effects. Layout remains geometric; format adapters retain all source timing IDs/XML. Event/target/effect integration must separately satisfy the roadmap/native mapping gates. See [ADR0035](decisions/0035-public-finite-timeline-arithmetic.md).

Core also owns [detached foreground event replay](../api/event-timeline.md) over caller page/visit/element identities and complete finite schedules. It validates current inventories/logs, returns frozen observations and holds no listener/clock/global registry. Format timing IDs, authored target transport/copy/delete reconciliation and native effects remain outside this numerical utility under ADV-06.1.2.2/.3.

[ADR0037](decisions/0037-slide-event-target-transport.md) adds optional owned slide timing source in Core and compiled semantic IR/Layout metadata. Core owns reference-safe set/replace/delete/copy; Layout validates actual IDs and refuses timed page splits without a reconciliation policy. Exporter owns default typed refusal and explicit normalized-candidate omission; static Renderer warns. Parser never fabricates authored targets from unknown native source timing. Dependencies remain Core → geometric/format consumers, with no raw format IDs in the shared model.

The Parser owns the explicit [notes source inventory](../api/notes-source-inventory.md): source roles, raw geometry, nonvisual attributes and stored field caches. It adds no Core/IR, Layout/Renderer or filesystem dependency, and grants no preservation-write permission.
