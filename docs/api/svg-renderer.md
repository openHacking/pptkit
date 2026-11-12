# `@pptkit/svg-renderer`

`slideActionMode: "document-links"` opts into fragment links between selected
slides when their returned SVGs are embedded inline in the same document. The
default `"metadata"` retains metadata-only slide actions and warnings. The opt-in
uses unique IDs based on original slide positions; use distinct `idNamespace`
values for separate render results in one document. An action whose target is not
selected stays metadata-only and emits `slide-action-not-linked`. SVGs loaded as
separate images or documents do not share these targets, and raster outputs cannot
retain links. An object action containing a run action can produce nested links,
whose behavior is viewer-dependent; a shape with no painted hit area may not be
clickable in the browser even if a PDF printer emits a link annotation.
Incremental rendering tracks the complete target table and invalidates
cached slides when the table changes. Chromium printing and MuPDF inspection for
one text and one shape link are described in [PDF research](../testing/pdf-print-observation.md);
this option does not itself generate PDF. Click navigation for one text and one
painted shape sample is exercised in Chromium, Firefox and Playwright WebKit via
`pnpm exec playwright test --config packages/svg-renderer/playwright.document-links.config.mjs`.

`@pptkit/svg-renderer` creates browser-oriented SVG previews from a PPTKit
`PresentationDocument` or an already resolved `LayoutResult`. It returns one standalone
SVG string per slide and never reads files, writes output, parses PPTX, or mutates its
input.

## Render an authoring document

```ts
declare function renderPresentationToSvg(
  document: PresentationDocument,
  options?: SvgPresentationRenderOptions,
): Promise<SvgRenderResult>;
```

The convenience entry normalizes the authoring document once, calls
`resolveNormalizedLayout()`, and renders the detached layout result.

## Render a layout result

```ts
declare function renderLayoutToSvg(
  layout: LayoutResult,
  options?: SvgRenderOptions,
): Promise<SvgRenderResult>;
```

Use this entry when a pipeline already owns normalized and resolved state. Both entry
points produce the same output for the same presentation.

## Results and warnings

`slideIndices?: readonly number[]` selects zero-based slide positions for rendering in the supplied order. Positions must be distinct, safe integers in range; an empty array produces no rendered slides. Omission renders every slide as before. Returned slide `index` and SVG IDs retain the original layout positions, and the input selection is captured before asynchronous asset resolution. The renderer resolves assets and emits rendering warnings only for selected slides. Layout analysis diagnostics with a `slideId` are restricted to selected IDs; diagnostics without one remain global. The authoring entry still normalizes and resolves the complete presentation before rendering, so selection is not a bound on layout work or memory.

`maxSvgBytes?: number` optionally caps the sum of UTF-8 bytes in the returned SVG strings. It accepts a nonnegative safe integer; zero allows only an empty selection. The renderer charges each completed slide before retaining it and throws exported `SvgRenderLimitError` (`code: "svg-limit-exceeded"`, `limit`, zero-based `slideIndex`) on overflow, with no partial result. Unpaired surrogates are charged as UTF-8 replacement characters. Later slides and their asset resolvers do not run after a failure. Omission leaves the prior unlimited behavior. The current slide's fragments and SVG are constructed before the check; layout, warnings, serialized envelopes and transient buffers are outside this limit, so it is not a peak-memory or CPU bound.

XML 1.0 cannot represent certain controls, unpaired surrogates or U+FFFE/U+FFFF. If they occur in rendered slide output, the SVG replaces them with U+FFFD and reports `svg-xml-character-replacement` with the slide ID and path. Displayed chart labels use the more specific `chart-text-xml-replacement` warning with chart identity. The original document is not changed; warning locations do not identify every individual character. Byte limits apply to the returned replacement text. Escaped carriage returns in attributes and text are serialized as character references, though ordinary text layout may interpret them as line breaks.

`signal?: AbortSignal` requests cancellation. A pre-aborted signal rejects before authoring normalization; rendering checks between slides and elements, and stops waiting promptly for a pending asset resolver. The resolver receives an optional second `{ signal }` argument and remains responsible for cancelling its own underlying work. Abort rejects with exported `SvgRenderAbortError` (`code: "operation-aborted"`) and returns no partial result; ordinary resolver failures still produce warnings. Listeners are removed after completion or abort, and a late resolver rejection is consumed. Normalization, layout, and synchronous work inside a single element are not preemptible; cancellation is observed after they yield. Supplying no signal retains the existing behavior.

```ts
interface SvgRenderResult {
  width: number;
  height: number;
  slides: SvgRenderedSlide[];
  warnings: SvgRenderWarning[];
  status: "rendered" | "rendered-with-warnings";
}

interface SvgRenderedSlide {
  slideId: string;
  index: number;
  hidden: boolean;
  svg: string;
}

interface SvgRenderWarning {
  code: string;
  message: string;
  slideId?: string;
  elementId?: string;
  assetId?: string;
  layoutId?: string;
  path?: string;
  fontDependencies?: { family: string; bold: boolean; italic: boolean }[];
}
```

Hidden slides remain in the result and are marked with `hidden: true`. Hidden elements
are omitted. Layout elements render before slide-local elements, preserving their
drawing order.

`backgroundMode: "transparent"` omits the resolved slide background rectangle;
inherited artwork and slide objects remain. The default `"slide"` retains existing
behavior. The mode participates in incremental cache invalidation. For PNG and JPEG
conversion, see the separate [browser raster adapter](browser-raster.md).

`portabilityDiagnostics: "warn"` reports selected dependencies on each rendered object:
`svg-foreign-object-dependency` for XHTML fallback, `svg-system-font-dependency` when
output references fonts supplied by the viewing environment, and
`svg-external-image-dependency` for emitted HTTP(S), relative, or blob image URLs.
`svg-inline-svg-image-unverified` identifies nested SVG image data that has not been inspected. Warnings include
slide and element IDs, and image warnings also include the asset ID; repeated uses are
reported separately. Hidden objects and unselected slides are excluded. The option
defaults to `"off"`, does not change SVG, and participates in incremental cache
invalidation. Resolver output determines the image warning, rather than the original
asset source. A data URL is treated as inline but its payload is not inspected for
external dependencies or validated. Font warnings list distinct emitted family,
bold and italic combinations in first-use order, including chart labels and selected
script-specific families. Names are recorded before final XML character replacement.
This is a list of declared requests, not resolved font
files: a viewer can substitute another face, fail to load one, or apply synthetic
styles. The list does not imply that embedded data URLs and other SVG features are portable.
It is a dependency aid, not a standalone-viewer compatibility certification.

`mixedTextMode: "native-no-wrap"` opts eligible mixed-run text boxes and shape text into SVG
`text`/`tspan` output. Eligibility requires no wrapping, no tabs or run actions,
no RTL paragraph or run and no direction-sensitive Unicode. Vertical text retains
its existing path. Other mixed-run content keeps XHTML. The default, `"browser"`,
preserves the existing output. Each newly admitted object reports
`svg-native-mixed-text-approximation`: font availability, baselines, shaping and
spacing remain viewer-dependent; fonts are embedded only when explicitly supplied.
The mode participates
in incremental cache invalidation. It does not convert table text by itself and does not certify
standalone SVG compatibility. [Named MuPDF comparisons](../testing/native-mixed-svg.md)
confirm visible and extractable text for one text-box and one rotated-shape example,
not general fidelity.

`mixedTextMode: "native"` additionally admits eligible wrapped mixed-run text and
ordinary horizontal shape text. It uses the same approximate grapheme-aware line
layout as simple native text; its line breaks and baselines are not measured against
the viewer's font. This additional path does not admit RTL, direction-sensitive
Unicode, tabs or run actions; existing native cases and vertical-text routing are
unchanged. Separate SVG line elements can introduce
line breaks or spaces in extracted/selected text even when none occur in the source.
The mode retains text instead of converting it to outlines, so viewer font
availability and shaping still apply. A [MuPDF specimen](../testing/native-wrapped-svg.md)
documents both visible output and changed extracted whitespace; it is not a
general text-fidelity guarantee.

`tableMode: "native"` opts tables into native SVG cells, text, clipping and shared
borders. It resolves merged-cell geometry using Layout's bounded border-grid
resolver and emits row/cell/header roles and span indices. Cell accessible names
come from full authored runs and explicit list markers, including text clipped
visually, so heuristic line wrapping does not insert spaces into spoken names.
Every converted table
reports `svg-native-table-approximation`: text wrapping uses the renderer's
heuristic, fonts remain external, and shared-border resolution can differ from
XHTML and PowerPoint. The table frame clips the grid, and each cell clips its text.
Quarter-turn cell text is rotated as a horizontal layout with physical margins
mapped to its rotated frame; it reports `table-text-direction-approximation` once
per table. Other vertical text, run actions, tabs, RTL or direction-sensitive Unicode, invalid grids
and border resolver limit failures cause the whole table to retain XHTML with
`svg-native-table-fallback`; no partial table conversion occurs. Default `"browser"`
preserves existing output. This mode participates in incremental cache invalidation.
SVG role markup does not guarantee a particular accessibility tree in all viewers.
The [MuPDF samples](../testing/native-table-svg.md) cover one merged table and
two quarter-turn cells, not
general typography, accessibility or native Office parity.

## Explicit embedded fonts

`embeddedFonts` accepts caller-supplied font programs; no font is loaded from the
filesystem or network by the renderer. Each face specifies `family`, `weight`
(`400` or `700`), `italic`, `format` (`woff2`, `woff`, `truetype`, `opentype`)
and canonical `base64`. The matching family and style must also be requested by
the document. The renderer copies these values before awaiting asset resolution,
and emits them as `@font-face` data URLs in every selected slide. Omission or an
empty array preserves the previous output. Each slide with supplied faces emits
`svg-embedded-font-unverified`.

At most 32 faces and 8 MiB of decoded font data are accepted per render. Names
must be nonempty and XML-safe without controls; family/style duplicates are
rejected case-insensitively. The format signature and canonical Base64 are checked,
but the complete font structure, internal family, glyph coverage, decoding and
viewer support are not verified. Invalid input rejects the whole render. The SVG
byte budget includes each repeated embedded font, but applies after constructing
each slide; it is not a peak-memory limit. Incremental caches include the full
captured font input in their shared snapshot, so changing any face invalidates all
slides and large programs increase cache memory. SVG embedded inline in a larger
HTML document can register its `@font-face` globally: use distinct family aliases
when different documents could collide. CSP and isolated-image support vary by
viewer. The optional dependency report remains a list of font requests, not proof
that an embedded font was selected.

One Chromium test uses a supplied WOFF2 face under an alias containing quotes,
checks a custom font was actually selected for inline SVG, and compares the inline
SVG and isolated SVG-image pixels. An unembedded image produces different pixels;
external requests are blocked. This does not establish support in non-browser SVG
viewers, font coverage, all styles or visual equivalence with PowerPoint. A separate
[MuPDF observation](../testing/embedded-font-mupdf.md) used a fallback face for the
same supplied WOFF2, demonstrating that embedding alone does not guarantee selection.

## Asset resolution

URL assets are used directly. Path assets produce a visible placeholder and an
`asset-path-unsupported` warning unless a resolver supplies a browser-safe URL:

```ts
type SvgAssetResolver = (
  asset: NormalizedAsset,
) => string | undefined | Promise<string | undefined>;

interface SvgRenderOptions {
  idNamespace?: string;
  resolveAsset?: SvgAssetResolver;
}
```

Resolvers may return HTTP(S), data, or blob URLs. Resolver failures become warnings and
do not prevent other slides from rendering. Applications own blob URL lifetime.

## Supported preview semantics

- SVG geometry for the eight Core shapes, connectors, arrows, image transforms, crop,
  opacity, nested groups, backgrounds, fills, and strokes.
- Bar, line, and pie charts consume Layout's shared plot box, scale, ticks,
  category label/data positions, category tick positions, legend item boxes, and
  series geometry. Their fixed layer order is chart
  background, plot background, gridlines, clipped series, axes/labels, title,
  and legend. Line and marker visibility are independent for data and legend;
  a line series without markers keeps its line-shaped legend glyph. Pie legends
  use the resolved per-point colors.
- Native SVG `<text>/<tspan>` output for simple text, with deterministic line breaks
  and baselines calibrated toward DrawingML output.
- XHTML `foreignObject` fallback for most mixed-run rich text, shape text, and tables with
  row/column spans.
- Deterministic internal definition IDs, accessibility labels, URL actions, and
  metadata for cross-slide actions. Definitions are unique across the slides of one
  render result, including inherited artwork; they depend on traversal order and
  are not a public identity API. Independently rendered copies with the same slide
  identity and position can reuse IDs. To embed them in one DOM, provide a distinct,
  nonempty `idNamespace` to each render call; it namespaces root, title and internal
  definitions without changing slide or element identities. The caller is responsible
  for choosing distinct namespaces. An omitted namespace preserves default output;
  an empty or non-string value throws `TypeError`. The option is deterministic and
  does not prevent collisions with arbitrary IDs supplied by the surrounding page.
- Explicit warnings for unsupported path assets, resolver failures, degraded image
  crop, `autoFit: resize`, unsafe URLs, and standalone cross-slide actions.

## Fidelity and security boundary

The output targets current Chromium, Firefox, and Safari. Native text is the portable
default for simple content. Remaining `foreignObject` content is not guaranteed in
non-browser SVG viewers, PowerPoint SVG import, or arbitrary rasterizers.
Browser and PowerPoint font engines differ, so previews help find clipping, overlap,
missing assets, structural mistakes, and obvious visual regressions; they are not
pixel-identical PowerPoint evidence.

Core point values are emitted as CSS pixel values inside XHTML `foreignObject` content.
This keeps text, margins, and tables in the same SVG user-unit coordinate system as the
surrounding point-based geometry; using CSS `pt` there would apply the browser's 4/3
point-to-pixel conversion a second time and make text overflow its authored bounds.

Simple text does not delegate wrapping or vertical placement to CSS. The renderer emits
explicit native SVG lines, applies normalized shrink scaling, and positions alphabetic
baselines from the authored top inset. Mixed-run and shape-text `foreignObject` nodes
keep vertical overflow visible to avoid silently cutting off glyphs; table cells remain
clipped to their explicit bounds.

`frame.wrap: false` disables native soft wrapping while preserving explicit CRLF/CR/LF breaks. Repeated and trailing breaks retain empty-line vertical space; source leading/interior/trailing whitespace is preserved. Whitespace used as a soft wrap opportunity is omitted at the new line. Joined text is segmented into graphemes before selected-font fragments are placed on lines, so a forced break does not split combining/ZWJ/flag clusters. An oversized cluster may overflow its frame. Width estimates remain heuristic, rather than shaped Layout line fragments or a full Unicode line-breaking/bidi implementation.

For native no-wrap text, a paragraph with one explicit stop, no list or indent, left paragraph alignment, ASCII alphanumeric or left-to-right Han, Hiragana, Katakana and Hangul words separated by single spaces on each side of at most one tab per hard line uses that stop as an SVG text anchor. Decimal suffixes remain restricted to a single number with a period. Single runs and mixed runs without run actions are admitted, including mixed font sizes; run actions remain on the XHTML path. Mixed-size or admitted East Asian tab lines use the largest run size as an approximate shared baseline. In 12 first-open Chinese-prefix shape samples with Hiragino Sans GB, [PowerPoint versus Chromium](../testing/tab-svg-cjk-shape-powerpoint-evidence.json) shows up to 0.331 pt horizontal and 0.247 pt baseline difference across 32 suffix characters; Japanese, Korean and other fonts have no corresponding native evidence. In 12 first-open Arial shape samples, [PowerPoint versus Chromium](../testing/tab-svg-shape-mixed-size-powerpoint-evidence.json) shows up to 0.596 pt horizontal and 0.838 pt baseline difference across 44 suffix characters; this is not full visual parity and other fonts remain unverified. Shape text uses this path when each tabbed paragraph meets those restrictions and each other paragraph has a single run without an action; its geometry and transform remain shared with the shape. Shapes with no tabs retain the existing path. Left, center and right align the second chunk; decimal aligns the period in a numeric suffix. The source tab remains in a zero-size hidden text span. This does not resolve collisions with the prefix, multiple tabs, wrapping or tables. All tabbed text still emits `text-tab-layout-unsupported` because these and font-dependent cases can differ from the exported presentation. [PowerPoint and Chromium controls](../testing/tab-stops-libreoffice.md) compare 12 named no-wrap text-box samples, 12 shape-text samples and one fitting multiword sample; they are not general preview equivalence or native shape-transform evidence.

Native grapheme segmentation uses the runtime's `Intl.Segmenter`; runtimes without it use the escaped XHTML path. Empty XHTML paragraphs and trailing empty lines use an `aria-hidden` display placeholder in the last effective style, outside authored runs. Font-family strings are literal names, serialized as quoted CSS family names for native text/chart labels and XHTML; numeric words, commas and quotes do not become font lists or invalid CSS tokens.

`foreignObject` is not a universal requirement or a fidelity guarantee. It is a bounded
fallback until Layout can expose resolved mixed-run lines. Keeping the native and
fallback paths explicit avoids pretending that browser CSS and PowerPoint share font
metrics.

All authored text and attributes are escaped and unsafe URL protocols are omitted. The
renderer does not accept arbitrary SVG or HTML fragments. Applications should still
apply their normal Content Security Policy when inserting generated SVG into a page.

See the checked [browser example](../examples/svg-preview.ts) and
[Preview and Editing](../architecture/preview-and-editing.md).


## Layout analysis options

`SvgPresentationRenderOptions` extends `SvgRenderOptions` with `layout?: LayoutOptions`, forwarded by `renderPresentationToSvg`. `renderLayoutToSvg` consumes any existing `layout.analysis` without running measurement again. Both entry points forward analysis diagnostics once as warnings with object identity and path. Analysis alone leaves SVG geometry and wrapping unchanged. See [Layout analysis](layout.md#opt-in-measurement-and-diagnostics) for provider requirements and measurement limits.

Tables successfully resolved with [table sizing](table-sizing.md) render positioned XHTML cells from `cellLayout`, with explicit clipping and table/row/cell semantics. This enforces resolved row geometry instead of relying on HTML table minimum heights. Other tables retain the existing HTML table path.

## Byte-backed resources

See [byte assets and explicit export resolution](byte-assets.md) for immutable registration snapshots, supported image signatures, self-contained preview, export resolvers, resource budgets and cancellation.

For directly embedded JPEG byte assets, the renderer examines at most 65,535 decoded header bytes and reports `jpeg-preview-precision-unverified` if it finds a frame declaring a sample precision other than 8 bits. If a segment or marker crosses that prefix while more bytes exist, it reports `jpeg-preview-inspection-incomplete`. It still emits the original JPEG; this is not image validation, decoding or a promise of viewer support. Malformed or truncated headers, URL/path assets and resolver output are not certified by the absence of a warning. A [named 12-bit comparison](../testing/jpeg-12-native.md) found Chromium returning a successful image decode with transparent pixels where PowerPoint displayed the image; an explicit resolver can provide a caller-owned compatible conversion.

Page [Fade transition](slide-transitions.md) data does not alter static SVG. Complete final slide content remains visible; playback/navigation is not supplied by this renderer.

Bound foreground event plans remain detached metadata in Layout. Static rendering keeps authored geometry and emits `event-timeline-static` per timed slide; it never evaluates a clock or embeds native effects. See [the slide-bound contract](slide-event-timeline.md).

See [direct table borders](table-borders.md) for shared SVG physical segments over exact HTML cells, alpha handling and JSON Layout metadata validation.

[Table region styles](table-style.md) resolve in Core; SVG consumes canonical cell paints/text and the same shared physical borders as PPTX without its own regional cascade.

[Exact-face SVG evidence](../testing/real-font-svg.md) verifies 16 native/mixed-run cases in Chromium with actual custom PostScript faces, independent HarfBuzz advances, visible PNG ink and isolated SVG-image raster controls. That experiment manages its own font embedding; the optional renderer input above is separately tested. These observations do not extend native Office or cross-browser typography guarantees. See [ADR0052](../architecture/decisions/0052-svg-text-boundaries-and-font-names.md), [Intl segmentation](https://tc39.es/ecma402/#sec-intl.segmenter) and [CSS literal family-name syntax](https://www.w3.org/TR/css-fonts-4/#family-name-syntax).

Run styles with `rightToLeft: true` request DrawingML selective run-level direction and complex-script font selection. SVG preview reports `text-run-rtl-approximation`; the built-in browser text measurer rejects these runs because browser layout cannot reproduce the selective override. Explicit `false` does not trigger this limitation.
