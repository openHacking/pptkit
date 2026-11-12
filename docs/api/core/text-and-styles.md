# Text and Styles

PPTKit separates text-container behavior, paragraph layout, and character styling so that each value has one owner.

## Structured text

```ts
interface TextParagraphInput {
  runs: TextRunInput[];
  style?: TextParagraphStyleInput;
}

interface TextRunInput {
  text: string;
  style?: TextRunStyleInput;
  action?: ElementAction;
}
```

String content is normalized into paragraphs and runs. Use structured content for mixed emphasis, lists, language metadata, or run-level links.

## Document text style presets

Documents may define immutable named partial styles once at construction time:

```ts
interface TextStylePresetInput {
  frame?: TextFrameStyleInput;
  paragraph?: TextParagraphStyleInput;
  run?: TextRunStyleInput;
}

type TextStylePresetMap = Readonly<Record<string, TextStylePresetInput>>;
```

```ts
const presentation = createPresentation({
  textStylePresets: {
    title: {
      frame: { margin: 0, verticalAlign: "middle" },
      paragraph: { align: "center" },
      run: { fontSize: 32, bold: true },
    },
  },
});

slide.addElement({
  type: "text",
  content: "Preset-backed title",
  textStylePreset: "title",
  box: { x: 48, y: 48, width: 600, height: 48 },
});
```

Text elements, shape text bodies, placeholders, and table cells can reference a
preset. The fixed precedence is: explicit local value, referenced preset,
placeholder style, theme, then Core default. Presets cannot inherit from one
another. Unknown names and invalid preset values are reported by validation.

| Preset field | Applies to |
| --- | --- |
| `frame` | Text elements and Shape text bodies; for table cells, `margin` and `verticalAlign` map to the cell. |
| `paragraph` | Text elements, Shape text bodies, placeholders, and table cells. |
| `run` | Text elements, Shape text bodies, placeholders, and table cells. |

Preset keys are caller-defined, non-empty strings, such as `title`, `body`, or
`cardLabel`; they are not a fixed enum. An unknown reference produces the
`missing-text-style-preset` validation diagnostic.

When a text element has a fixed `x`, `y`, and `width` but no `height`, Core calculates a deterministic estimated height during normalization. The estimate includes explicit paragraphs and line breaks, approximate wrapping, the largest run font size in each paragraph, line spacing, paragraph spacing, bullet indentation, frame margins, and a small font-size-based trailing safety buffer for renderer-specific ascent/descent differences. This keeps the Canonical IR and exported text box bounds complete while allowing authoring inputs to omit a fragile single-line height.

```ts
slide.addElement({
  type: "text",
  box: { x: 48, y: 48, width: 600 },
  frame: {
    margin: { left: 12, right: 12 },
    verticalAlign: "middle",
    wrap: true,
    autoFit: { mode: "shrink", fontScale: 0.8 },
  },
  content: [
    {
      style: {
        align: "left",
        spaceAfter: 8,
        bullet: { type: "bullet", character: "•" },
      },
      runs: [
        { text: "Reliable ", style: { fontSize: 24 } },
        {
          text: "editable output",
          style: { fontSize: 24, bold: true, color: { theme: "accent1" } },
          action: { type: "url", url: "https://example.com/details" },
        },
      ],
    },
  ],
});
```

## Frame style

| Field | Type | Core default |
| --- | --- | --- |
| `margin` | `number \| Partial<Insets>` | top/bottom `3.5`, left/right `7` pt |
| `verticalAlign` | `"top" \| "middle" \| "bottom"` | `"top"` |
| `wrap` | `boolean` | `true` |
| `autoFit` | `{ mode: "none" \| "resize" } \| { mode: "shrink"; fontScale?: number }` | `{ mode: "none" }` |

A numeric margin applies to all sides. `fontScale` is in `0..1` and defaults to `1` when shrink mode is selected. Layout measurement is not yet implemented; these values are semantic authoring/export instructions.

## Paragraph style

| Field | Type | Core default |
| --- | --- | --- |
| `align` | `left`, `center`, `right`, `justify` | `left` |
| `direction` | `ltr`, `rtl` | absent; inherit/default |
| `indent` | non-negative points | `0`; `27` when enabling a list |
| `hanging` | non-negative points | `0`; `27` when enabling a list |
| `lineSpacing` | positive multiplier | `1` |
| `spaceBefore` / `spaceAfter` | non-negative points | `0` |
| `bullet` | none, bullet, or number | none |
| `tabStops` | ordered `{ position, align }` entries | absent |

`direction` sets the paragraph's base direction independently of alignment and
run language. It participates in paragraph-style inheritance; omitting it
retains the existing default behavior. Fresh PPTX writes `a:pPr/@rtl` when the
value is explicit. Import projects supported direction declarations through
the existing inheritance chain; a false presentation-level default is treated
as the existing LTR baseline, while higher-level `rtl="0"` remains an explicit
override. SVG uses browser paragraph layout for RTL direction. Native and
browser bidi details may differ, and source-preserving style edits remain outside
the text-run save profiles. See the [direction evidence](../../testing/bidi-first-open-observation.md#explicit-paragraph-direction-source-boundary).

`tabStops` sets paragraph tab stops in points relative to its left margin; alignment is `left`, `center`, `right`, or `decimal`. Positions must be nonnegative and strictly increasing after conversion to DrawingML EMUs, within the signed 32-bit coordinate range. An authored array replaces the inherited array; `[]` explicitly clears it. Fresh PPTX writes the stops as `a:tabLst`; content import maps explicit, ordered lists of at most 256 stops, including empty lists, through the existing paragraph cascade and source report. Missing position or alignment, unsupported values, negative or nonincreasing positions and longer lists block inherited stops and are reported as unsupported; malformed numeric coordinates reject the import. This projection does not add a source-preserving tab-stop edit profile. [Browser measurement](../browser-measurement.md) and [SVG preview](../svg-renderer.md#fidelity-and-security-boundary) support different restricted explicit-stop profiles; SVG retains `text-tab-layout-unsupported` for tabbed text. In [one LibreOfficeDev PDF observation](../../testing/tab-stops-libreoffice.md), explicit stops did not change the measured positions. A separate PowerPoint first-open and print observation confirms the four alignments for its 12 named no-wrap samples; one appended character was saved and reopened, but general PowerPoint editing, default stops and broader visual alignment remain unverified.

Fresh export reports `tab-stops-libreoffice-zero-indent` when a paragraph containing tabs emits explicit stops and zero first-line indent: the cited LibreOfficeDev control ignored the stops in that combination. The explicit zero remains in the file to preserve formatting against inherited paragraph styles.

<!-- doc-test: docs/examples/tab-stops.ts -->
```ts
import { createPresentation } from "@pptkit/core";

export const deck = createPresentation();
deck.addSlide().addElement({
  type: "text", box: { x: 40, y: 40, width: 300, height: 120 },
  content: [{
    style: { tabStops: [{ position: 72, align: "left" }, { position: 180, align: "decimal" }] },
    runs: [{ text: "Item\tPrice\t12.50" }],
  }],
});
```

```ts
type TextBulletInput =
  | { type: "none" }
  | { type: "bullet"; character?: string }
  | {
      type: "number";
      style?: "arabicPeriod" | "arabicParen" |
        "alphaLowerPeriod" | "alphaUpperPeriod";
      startAt?: number | "continue";
    };
```

The default bullet character is `•`; numbered lists default to `arabicPeriod` starting at `1`.
When a paragraph enables bullets or numbering without specifying `indent` or
`hanging`, Core uses PowerPoint's first-level list defaults: both are `27` pt.
Explicit indentation values are preserved.

`startAt` can be a positive safe integer or `"continue"` on an individual paragraph. Continuation requires the immediately preceding paragraph in the same content array to explicitly declare a numbered bullet with the same effective numbering style. Empty numbered paragraphs count; an unnumbered paragraph breaks the sequence. A numeric start restarts it. Omitting `startAt` retains the existing default of 1. Presets and placeholder defaults cannot declare `"continue"`. Core recalculates numbers after edits and rejects an orphan, style mismatch, or unsafe increment; normalization leaves the authoring input unchanged and contains numeric starts only.

Fresh PPTX export requires every emitted value to be 1–32767 and rejects an out-of-range value; this changes earlier behavior that could generate an invalid OOXML value. Export materializes continuation as explicit starts and reports `numbered-list-continuation-materialized`: the authoring dependency is not preserved for subsequent native editing or fresh import. `formatListMarker` formats a normalized marker without tracking continuation or changing the document. SVG preview and browser measurement use the same base-26 alphabetic sequence (`z`, `aa`, `ab`, …, `ba`).

<!-- doc-test: docs/examples/list-continuation.ts -->
```ts
import { createPresentation } from "@pptkit/core";

export const deck = createPresentation();
deck.addSlide().addElement({
  type: "text", box: { x: 40, y: 40, width: 300, height: 120 },
  content: [
    { style: { bullet: { type: "number", startAt: 4 } }, runs: [{ text: "First" }] },
    { style: { bullet: { type: "number", startAt: "continue" } }, runs: [{ text: "Next" }] },
  ],
});
```

<!-- doc-test: docs/examples/list-markers.ts -->
```ts
import { formatListMarker } from "@pptkit/core";

export const marker = formatListMarker({
  type: "number", style: "alphaLowerPeriod", startAt: 53,
});
```

For alphabetic starts above 27, SVG emits `alphabetic-numbering-application-dependent`: [Microsoft's OOXML description](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.autonumberedbullet?view=openxml-3.0.1) gives `aaa` for 53, while [one LibreOfficeDev headless PDF observation](../../testing/alphabetic-numbering-libreoffice-evidence.json) produced `ba`. PowerPoint behavior and editing are untested. Nested list levels are not implemented.

## Run style

| Field | Type | Core default |
| --- | --- | --- |
| `fontFamily` | family string or theme font role | theme body font |
| `rightToLeft` | `boolean` | absent; inherit/default |
| `eastAsianFontFamily` / `complexScriptFontFamily` | family string, heading/body role, or authoring null reset | inherited slot, otherwise `fontFamily` |
| `fontSize` | positive points | `18` |
| `bold` / `italic` | `boolean` | `false` |
| `underline` / `strike` | `boolean` | `false` |
| `color` | `ColorValue` | theme `text1` |
| `colorOpacity` | finite number in `0..1` | `1` (opaque) |
| `language` | language tag string | `en-US` |

`rightToLeft` maps the DrawingML run-level `a:rtl` declaration, independently of paragraph direction. Explicit `false` overrides inherited `true`. When true, the shared font selection uses the complex-script family, falling back to the default family; explicit per-script mappings take precedence. Fresh PPTX and import preserve the declaration, while source-preserving style edits remain blocked. SVG emits `text-run-rtl-approximation` because browsers do not implement DrawingML’s selective weak/neutral character override; the built-in browser measurer rejects such runs. This selection is a deterministic policy, not a promise of native font parity: a named PowerPoint PDF observation used the Latin font for RTL-declared digits. Strong left-to-right characters with this property have no generally established native behavior. See the [native observations](../../testing/bidi-first-open-observation.md#run-level-direction-is-a-separate-source-feature).

Run color opacity inherits through presets/placeholders independently from element opacity. An explicit `1` resets inherited translucency; `0` retains invisible text with unchanged layout. Normalized IR omits opaque `colorOpacity`, and existing normalized runs without the field remain opaque. Export combines run and element alpha; native SVG and mixed-run HTML apply both. See [text colors and alpha](../text-color-transforms.md) for import transforms, precision and native limitations.

[Explicit script fonts](../script-fonts.md) preserve independent families through Core, Layout measurement, SVG and native fresh output. Undefined inherits each slot; null clears only that inherited slot, and normalized absent uses the default family. `resolveTextFontSegments` exposes the shared deterministic Unicode selection policy and UTF-16 source offsets. Multi-family semantic source import, script-language fallback and complete shaping/bidi remain separate gates.

## Colors and paints

```ts
type ColorValue = string | { theme: ThemeColorRole };

type PaintInput =
  | { type: "none" }
  | { type: "solid"; color: ColorValue; opacity?: number };
```

Direct colors are six-digit RGB strings with an optional leading `#`. Theme colors use semantic roles: `background1`, `background2`, `text1`, `text2`, `accent1` through `accent6`, `hyperlink`, and `followedHyperlink`. Paint opacity is `0..1` and defaults to `1`.

## Strokes

```ts
interface StrokeStyleInput {
  paint?: PaintInput;
  width?: number;
  dash?: "solid" | "dash" | "dot" | "dashDot";
  beginArrow?: ArrowType;
  endArrow?: ArrowType;
}
```

Stroke width is a non-negative number of points. Arrow types are `none`, `arrow`, `triangle`, `stealth`, `diamond`, and `oval`. The Core default stroke has no paint, width `1`, a solid dash pattern, and no arrowheads.

## Transforms and opacity

```ts
interface TransformInput {
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}
```

Rotation uses degrees. Flips apply within the element box. Element opacity multiplies visual paint opacity during export and must be between `0` and `1`.

## Style precedence

Normalization resolves text using this fixed order, from most specific to least specific:

1. Run, paragraph, or frame-local values.
2. The referenced document text style preset.
3. Bound placeholder defaults from the selected layout.
4. Presentation theme values.
5. Core defaults.

Normalized output contains explicit values. Layout and exporters do not invent business defaults.

## Validation

Invalid colors, opacity, font sizes, line spacing, margins, indents, numbering starts, shrink scales, and stroke widths produce diagnostics. See [Validation and IR](validation-and-ir.md) for error handling.

See [direct table borders](../table-borders.md) for sparse `TableCellStyleInput.borders` and inheritance from uniform `stroke`.

[Table region styles](../table-style.md) provide table-local cell/text defaults, alternating bands, edge/corner regions and local override priority; resolved theme roles remain semantic.

Script-specific overrides now also accept `scriptFontFamilies`, a bounded map with per-key and whole-map resets. See the [map selection/output contract](../script-fonts.md#script-font-map-extension) and [typed example](../../examples/script-font-map.ts); fresh native output materializes current fonts and warns about reusable-rule/run-topology loss.

A text frame's optional `textDirection` accepts `"horizontal"`, `"rotate90"`, `"rotate270"`, `"eastAsianVertical"`, `"mongolianVertical"`, `"stackedVertical"`, or `"stackedVerticalRtl"`. It changes the writing direction inside its frame, independently of the shape's geometry and transform. An explicit horizontal value can override an inherited rotation. PPTX writes the direction on the text body. SVG rotates a horizontal layout inside the physical frame, remapping the insets; preview and bounds diagnostics report that wrapping and font metrics can differ from native PowerPoint. Intrinsic sizing of quarter-turned rectangular frames uses physical height as the inline axis. When both axes are content-sized, the solver can wrap at the height limit before sizing width. If measured text cannot fit, the intrinsic stage rolls back. Vertical parallelogram text supports content sizing on one or both physical axes. A single content rule keeps the other axis fixed. With wrapping enabled, the solver tries a bounded set of inline lengths and commits only if the text fits when remeasured in the resulting inset. If it cannot verify a fit, it retains the original geometry. Table cells continue to use their own cell style direction. See the [native mode matrix](../../testing/vertical-writing-matrix.md).

East Asian vertical writing maps to DrawingML `eaVert` for text frames and cell styles. The SVG preview uses CSS vertical right-to-left column flow with mixed glyph orientation. Layout can measure fixed frames with the built-in browser adapter or a custom vertical provider. Browser vertical measurement is an estimate and does not cover tabs, list markers or RTL overrides. Intrinsic sizing supports rectangular East Asian frames with a vertical measurer. A height cap may trigger wrapping before width is solved. East Asian automatic table row sizing uses a vertical provider and rechecks the final cell bounds. The same parallelogram contract applies.

Mongolian and stacked directions map to DrawingML `mongolianVert`, `wordArtVert`, and `wordArtVertRtl` respectively in both text frames and cell styles. SVG approximates them using left-to-right columns with mixed glyph orientation (Mongolian), left-to-right columns with upright glyphs (stacked), or right-to-left columns with upright glyphs (stacked RTL). Fixed-frame and cell bounds measurement uses the browser adapter or a custom vertical provider. These modes also support rectangular intrinsic sizing and table auto rows with a vertical measurer. The same parallelogram contract applies.
