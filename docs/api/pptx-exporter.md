# `@pptkit/pptx-exporter`

`@pptkit/pptx-exporter` validates and normalizes a Core authoring document, resolves layout once, loads image assets, serializes editable OOXML parts, and returns a deterministic PPTX package. The default entry is browser-neutral; the `/node` entry additionally reads local files and writes output.

## Browser-neutral `generatePptx(document)`

```ts
declare function generatePptx(
  document: PresentationDocument,
  options?: GeneratePptxOptions,
): Promise<GeneratePptxResult>;
```

```ts
import { generatePptx } from "@pptkit/pptx-exporter";

const result = await generatePptx(presentation);

const blob = new Blob([result.bytes], {
  type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
});

console.log(result.status, result.slideCount, result.byteLength);
```

Escaped XML text and attribute values, including document metadata, reject XML 1.0 forbidden characters and isolated UTF-16 surrogates with `RangeError`; no PPTX bytes are returned. Valid supplementary Unicode characters are supported.

The default entry loads `source.type: "url"` assets using `fetch`. HTTP(S), data, and blob URLs are accepted when supported by the runtime. Local path assets are omitted with a warning. Use [`externalImageValidation: "signature"`](byte-assets.md#optional-validation-for-external-images) to validate default URL and Node path loads before packaging.

## Node.js entry

The Node entry exports both `generatePptx()` and `writePptx()` and supports byte, URL and path assets.

```ts
import { writePptx } from "@pptkit/pptx-exporter/node";

const result = await writePptx(presentation, {
  output: "./output/deck.pptx",
});

console.log(result.status, result.output, result.warnings);
```

### `writePptx(document, options)`

```ts
declare function writePptx(
  document: PresentationDocument,
  options: WritePptxOptions,
): Promise<WritePptxResult>;
```

`output` is required. Parent directories are created as needed, and the returned `output` path identifies the written file.

## Result types

```ts
interface GeneratePptxResult {
  bytes: Uint8Array;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "generated" | "generated-with-warnings";
}

interface WritePptxResult {
  output: string;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "written" | "written-with-warnings";
}

interface ExportWarning {
  code: string;
  message: string;
  slideId?: string;
  elementIndex?: number;
  assetId?: string;
  elementId?: string;
  layoutId?: string;
  path?: string;
}
```

`byteLength` is the generated package size. A non-empty warning collection changes the status but does not discard otherwise valid presentation content.

## Failure and warning behavior

| Condition | Outcome |
| --- | --- |
| Core validation errors | Rejects with `PresentationValidationError` containing all error diagnostics. |
| URL request failure | Omits the affected image and returns an asset warning. |
| Missing local file in Node | Omits the affected image and returns an asset warning. |
| Path asset through the default entry | Omits the image and explains that `/node` is required. |
| Output directory does not exist | Node entry creates it. |
| Output write fails | `writePptx()` rejects with the filesystem error. |

Applications should display or log warnings even when output was generated successfully.

Warning batches retain their diagnostic order and are assembled without passing the entire batch as function arguments. This removes an engine argument-count failure for large batches; generating the package and retaining all warnings still consumes memory and has no general document-wide memory guarantee.

## Exported presentation features

- Editable rich text paragraphs and runs, direct/theme colors, theme fonts, bullets, numbering, hyperlinks, and slide actions.
- Images with stretch, contain, cover, crop, opacity, accessibility, and media relationships.
- Shapes, connectors, nested groups, and native editable tables with cell spans.
- Native editable charts (`c:chartSpace`) for vertical/horizontal bar, line, and
  pie, doughnut, [area](chart-area.md), and [scatter](chart-scatter.md). The exporter writes explicit axes, scales, gridlines, ticks, fonts,
  markers, gaps, manual plot layout, native legend position, and per-point pie
  colors. Doughnut geometry and bounds are described in [doughnut charts](chart-doughnut.md); a 1–9% hole emits `doughnut-small-hole-office-extension` with its owning element and path because older Transitional descriptions specify a 10% minimum. Right/bottom legend flow is left to the corresponding native position
  instead of a manual legend box that Office may reflow into extra rows. Each chart
  owns a minimal embedded XLSX with category/series formulas and caches, so
  Native interactive Edit Data availability has not been established by the embedded workbook alone. In one PowerPoint environment, the command was disabled for both a PPTKit scatter sample and an independent control; see the [recorded observation](../testing/chart-scatter-edit-data-availability.json). Fresh export writes explicit `null` values as blank workbook cells and omitted numeric cache points, with `c:dispBlanksAs` set to the selected `blankDisplay` (`gap` or `zero`, plus `span` for lines and scatter). Without an explicit mode, blank charts default to `gap` and dense charts retain `zero`; stacked area charts default to `zero` even with blanks. Sparse array holes and non-finite values are rejected rather than converted to zero.
- Reusable slide layouts, typed placeholders, backgrounds, theme parts, hidden slides, and speaker notes.
- Presentation metadata plus PPTKit-owned slide metadata preservation for section, tags, and custom data.

Core supplies explicit normalized defaults. The exporter translates those semantics into OOXML but does not choose business defaults.

## Runtime boundary example

```ts
const remote = presentation.registerAsset({
  kind: "image",
  source: { type: "url", value: "https://example.com/image.png" },
  mimeType: "image/png",
});

// Works in the default and Node entries when the URL is reachable.
slide.addElement({
  type: "image",
  assetId: remote.id,
  box: { x: 80, y: 120, width: 320, height: 180 },
});
```

For asset registration and dimensions, see [Core assets](core/assets.md). For geometric resolution before export, see [`@pptkit/layout`](layout.md).


## Layout analysis options

`GeneratePptxOptions` accepts `layout?: LayoutOptions`. `WritePptxOptions` extends it with required `output: string`. Both default/browser and Node generators forward these options to Layout. See [measurement and diagnostics](layout.md#opt-in-measurement-and-diagnostics).

`firstSlideNumber?: number` overrides the authored starting slide number in native PPTX output (integer 0–9999). When omitted, the authored value is used; when both are absent, the attribute remains absent. This setting is independent of Layout's field-cache refresh and application-specific numbering behavior. See [field evaluation](field-evaluation.md).

An explicitly authored Core section collection exports as the native section extension when its ordered memberships cover every slide exactly once in presentation order; invalid or interleaved membership rejects export. Empty sections and repeated or absent names are supported. A legacy slide `section` label remains independent PPTKit slide metadata. One [PowerPoint save/reopen observation](../testing/authored-native-sections.md) covers named and empty sections; it does not establish source-preserving section editing or general native interoperability.

Analysis diagnostics become export warnings with their code, message, path and element/slide/layout identity. Bounds and severity remain available through direct Layout analysis. Provider failures are recoverable warnings; invalid option configuration rejects before package generation. Analysis alone does not change PPTX bytes, native text formatting, or frame dimensions.

## Byte-backed resources

See [byte assets and explicit export resolution](byte-assets.md) for immutable registration snapshots, supported image signatures, self-contained preview, export resolvers, resource budgets and cancellation.

## Connector representation

`connectorMode: "native"` opts into the [experimental native straight profile](native-connectors.md), including cardinal attachment bindings and located fallback diagnostics. The default `"path"` mode retains existing custom-path output. Native bent/curved and transformed-target profiles remain pending.

Internal slide actions serialize as explicit slide-jump commands. Relationships emitted from slide parts use a same-directory `slideN.xml` target; notes-part relationships use `../slides/slideN.xml`. External URL actions remain external hyperlink relationships.

Notes export retains normalized paragraphs that contain text, formatting or actions, including all-empty paragraphs with authored style/action data. It omits the plain default empty paragraph synthesized by normalization. Importing only [speaker-notes text](notes-import.md) does not preserve source notes-page artwork, master layout or inherited style provenance in a fresh export.

Fresh output now uses a fixed 540×720pt notes page with a native slide-image placeholder in the upper area and a separate lower body. The image preserves the actual slide aspect ratio; each local notes page follows its own slide backlink and shares one notes master. This repairs blank native thumbnails without raster assets or new authoring options. [Three-size native evidence](../testing/authored-notes-pages.md) covers copied, absent and styled-empty notes across three reopen rounds. The initial styled-empty native loss is corrected by the shared [empty-paragraph insertion properties](../testing/notes-insertion-styles.md): the paragraph mark retains the final empty run style, without copying its action. Native zero-character run identities remain distinct from insertion styling. Custom/imported notes artwork, fields, notes preview and overflow pagination remain outside this default layout.

## Page transitions

Experimental [Fade transitions](slide-transitions.md) serialize native page timing with exact p14 milliseconds and a coarse compatibility fallback. Static preview and native playback have separate acceptance boundaries.

## Authored event timeline admission

Slides with the [bound foreground timing extension](slide-event-timeline.md) reject by default with `PptxExportError` and frozen per-stage diagnostics before asset/file work. Native timing mapping is not implemented. Explicit `eventTimelineMode: "omit"` authorizes static output and emits `event-timeline-omitted` for every stage, without mutating the document; this omission occurs before requested pagination. Earlier untimed calls retain their output.

See [direct table borders](table-borders.md) for merged logical/perimeter border encoding, `table-border-segmented-native` and `table-transform-grouped`. Rotated/flipped native tables use groups; identity/action stay on the root, opacity applies once, and native application parity remains partial.

[Regional table authoring](table-style.md) emits `table-style-materialized`: effective direct native cells/theme references are retained, reusable native table-style definitions are not. Independent saved fill/text observations do not certify region reevaluation in native editors.

Script font maps use [warned materialization](script-fonts.md#script-font-map-extension): current concrete text fonts populate all three native slots, runs may split and reusable script/language rules are not retained. `script-font-map-materialized` identifies slide/element, notes or layout-placeholder scope.

## Presentation format metadata

Fresh output describes the serialized slide frame in `docProps/app.xml` with `PresentationFormat`: exact 16:9, 4:3 and 16:10 ratios use `On-screen Show (16:9)`, `On-screen Show (4:3)` and `On-screen Show (16:10)`; other ratios use `Custom`. Classification uses the same rounded integer EMU dimensions as `p:sldSz`, with exact integer arithmetic. Portrait, square and paper sizes remain custom. The label neither resizes the frame nor selects a native `p:sldSz@type` preset. Default 16:9 output remains byte-compatible.

This is descriptive extended-property metadata ([Microsoft property reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.extendedproperties.properties?view=openxml-3.0.1)). Fresh PPTX generation separately requires each rounded slide dimension to be within 914400–51206400 EMU, the [PresentationML coordinate bounds](https://iso.ieclist.ink/iec/ISO%20IEC%2029500-1-2012%20PDF.pdf); values outside this range throw `RangeError` before asset loading instead of producing an invalid package. Fractional point inputs that round to a boundary remain valid. This is a change for callers that previously exported out-of-range sizes; Core and SVG size rules are unchanged. The notes thumbnail uses the same rounded slide dimensions. Source-only Parser no-op retains the original package, including any original label. See [format verification](../testing/presentation-format.md) for automated and actual native scope.

## Native table cell properties

Fresh table output places normalized cell margins and vertical alignment on `a:tcPr` and emits cell borders before fill, as required by the native cell property model. This repairs PowerPoint ignoring body-frame cell insets and discarding out-of-order borders; existing affected table output intentionally changes. Text/shape frames and source-only no-op are unaffected. See [actual failure, isolated controls and corrected reopen evidence](../testing/table-cell-properties.md). Regional styles remain materialized direct formatting.

Nonempty table paragraphs also retain the final normalized run's font slots on the paragraph mark. PowerPoint includes that mark in line metrics; leaving its font implicit can move explicitly formatted text after changing the owning master theme. This adds no paragraph-mark size, color or action. All-empty paragraph insertion styling keeps its existing full-style behavior. See the [isolated controls and bounded rebind evidence](../testing/layout-rebind.md); fresh table bytes intentionally change, and older native files remain historical evidence rather than current producer byte matches.

## Native placeholder text inheritance

`nativePlaceholderText` defaults to `"materialize"`, preserving effective slide text formatting. `"inherit"` opts eligible reset Western font family, color, bold, italic, underline, strike, font size, language and explicit run RTL defaults, plus paragraph alignment, indent, hanging indent, line spacing, space before/after and bullet, and frame wrap, vertical alignment, writing direction and margins into native layout-default inheritance; unsupported cases stay materialized and produce `native-text-inheritance-materialized`. Only the selected eligible defaults are written; adding a font default can alter line metrics for other shapes on the shared layout, so the opt-in emits `native-placeholder-metrics-may-change` for layouts receiving run or paragraph defaults. Frame-only inheritance uses existing layout body properties and does not add those defaults. See [layout rebinding](layout-rebind.md) for exact conditions and current limits.

Distinct background or shape gradient stops that round to the same OOXML position generate `background-gradient-stops-collapsed` or `shape-gradient-stops-collapsed`, with the authored owner and stop path. Intentionally equal positions do not trigger this warning.
