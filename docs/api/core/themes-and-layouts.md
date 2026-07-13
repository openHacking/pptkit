# Themes and Slide Layouts

Themes provide semantic colors and fonts. Slide layouts provide reusable backgrounds, static elements, and typed placeholders. Core models these concepts without exposing PPTX master or relationship details.

## Presentation theme

```ts
interface PresentationThemeInput {
  name?: string;
  colors?: Partial<Record<ThemeColorRole, string>>;
  fonts?: {
    heading?: string;
    body?: string;
  };
}
```

Any omitted role inherits the Core theme. Theme colors must be six-digit RGB strings. The normalized default theme is named `PPTKit`, uses Aptos/Aptos Display, a white primary background, black primary text, and a standard six-accent palette.

Use semantic references in elements when colors should follow the presentation theme:

```ts
const presentation = createPresentation({
  theme: {
    name: "Product",
    colors: {
      accent1: "2457D6",
      accent2: "F97316",
    },
    fonts: {
      heading: "Aptos Display",
      body: "Aptos",
    },
  },
});
```

## `defineSlideLayout(input)`

```ts
defineSlideLayout(input: SlideLayoutInput): SlideLayoutDefinition;

interface SlideLayoutInput {
  id: string;
  name: string;
  background?: PaintInput;
  elements?: PresentationElementInput[];
  placeholders?: PlaceholderDefinitionInput[];
}
```

The method creates and returns an immutable layout definition. Layout IDs are caller-supplied and stable. Duplicate layout IDs and duplicate nested element IDs throw immediately. Cross-reference and placeholder errors are collected during validation.

Static `elements` draw as part of every slide using the layout. A slide references the layout through `layoutId`; no layout object is copied into authoring slide state.

## Placeholders

```ts
interface PlaceholderDefinitionInput {
  key: string;
  kind: "title" | "subtitle" | "body" | "image" |
    "table" | "footer" | "slideNumber";
  box: Box;
  textStyle?: {
    frame?: TextFrameStyleInput;
    paragraph?: TextParagraphStyleInput;
    run?: TextRunStyleInput;
  };
}
```

Placeholder keys must be unique inside a layout. An element binds by setting `placeholderKey`. When the element omits `box`, normalization uses the placeholder box. Text placeholders can also supply frame, paragraph, and run defaults.

```ts
presentation.defineSlideLayout({
  id: "title-and-body",
  name: "Title and Body",
  background: { type: "solid", color: { theme: "background1" } },
  elements: [
    {
      type: "shape",
      shape: "rect",
      box: { x: 0, y: 0, width: 12, height: 540 },
      style: { fill: { type: "solid", color: { theme: "accent1" } } },
    },
  ],
  placeholders: [
    {
      key: "title",
      kind: "title",
      box: { x: 48, y: 40, width: 760, height: 72 },
      textStyle: { run: { fontFamily: { theme: "heading" }, fontSize: 36, bold: true } },
    },
    {
      key: "body",
      kind: "body",
      box: { x: 48, y: 132, width: 760, height: 320 },
    },
  ],
});

const slide = presentation.addSlide({ layoutId: "title-and-body" });
slide.addElement({ type: "text", content: "Quarterly update", placeholderKey: "title" });
slide.addElement({ type: "text", content: "Highlights and decisions", placeholderKey: "body" });
```

## Background inheritance

The effective slide background is selected in this order:

1. Slide-local `background`.
2. Selected layout background.
3. Theme/Core default background.

Normalized slides record both the effective paint and `backgroundSource: "slide" | "layout" | "theme"`.

## Text inheritance

For a bound text element, placeholder text defaults sit between element/paragraph/run values and the presentation theme/Core defaults. Normalization materializes the result; exporters receive no unresolved inheritance.

## Failure behavior

- `defineSlideLayout()` throws immediately for duplicate layout IDs or element IDs it cannot own safely.
- Validation reports missing slide layouts, duplicate placeholder keys, invalid placeholder boxes, invalid placeholder text styles, and element bindings to missing placeholders.
- A non-existent `layoutId` or `placeholderKey` prevents normalization.
