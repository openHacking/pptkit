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
| `indent` | non-negative points | `0`; `27` when enabling a list |
| `hanging` | non-negative points | `0`; `27` when enabling a list |
| `lineSpacing` | positive multiplier | `1` |
| `spaceBefore` / `spaceAfter` | non-negative points | `0` |
| `bullet` | none, bullet, or number | none |

```ts
type TextBulletInput =
  | { type: "none" }
  | { type: "bullet"; character?: string }
  | {
      type: "number";
      style?: "arabicPeriod" | "arabicParen" |
        "alphaLowerPeriod" | "alphaUpperPeriod";
      startAt?: number;
    };
```

The default bullet character is `•`; numbered lists default to `arabicPeriod` starting at `1`.
When a paragraph enables bullets or numbering without specifying `indent` or
`hanging`, Core uses PowerPoint's first-level list defaults: both are `27` pt.
Explicit indentation values are preserved.

## Run style

| Field | Type | Core default |
| --- | --- | --- |
| `fontFamily` | family string or theme font role | theme body font |
| `fontSize` | positive points | `18` |
| `bold` / `italic` | `boolean` | `false` |
| `underline` / `strike` | `boolean` | `false` |
| `color` | `ColorValue` | theme `text1` |
| `language` | language tag string | `en-US` |

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
