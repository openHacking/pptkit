# Canonical Presentation IR v1

Canonical Presentation IR is the stable, format-independent contract between Core and downstream layout, preview, parser, and exporter work. The implemented version is identified by `irVersion: 1`.

## Top-level shape

```ts
interface NormalizedPresentation {
  irVersion: 1;
  id: string;
  metadata: NormalizedPresentationMetadata;
  size: PresentationSize;
  theme: NormalizedPresentationTheme;
  slides: NormalizedSlide[];
  assets: NormalizedAsset[];
  layouts: NormalizedSlideLayout[];
}
```

The IR contains data only. It has no authoring methods, filesystem handles, network clients, OOXML fragments, relationship IDs, or package paths.

## Metadata, theme, and assets

```ts
interface NormalizedPresentationMetadata {
  title: string;
  author: string;
  company?: string;
  subject?: string;
  description?: string;
  language: string;
  keywords: string[];
  revision: number;
}

interface NormalizedPresentationTheme {
  name: string;
  colors: Record<ThemeColorRole, string>;
  fonts: { heading: string; body: string };
}

interface NormalizedAsset {
  id: string;
  kind: "image";
  source: { type: "path" | "url"; value: string };
  mimeType?: string;
  width?: number;
  height?: number;
  accessibility: { description?: string; decorative: boolean };
  dedupeKey?: string;
}
```

Theme colors remain semantic inputs for elements but every theme role maps to a concrete RGB value in the normalized presentation.

## Slide and layout schema

```ts
interface NormalizedSlideLayout {
  id: string;
  name: string;
  background: NormalizedPaint;
  elements: NormalizedElement[];
  placeholders: NormalizedPlaceholderDefinition[];
}

interface NormalizedSlide {
  id: string;
  layoutId: string;
  background: NormalizedPaint;
  backgroundSource: "slide" | "layout" | "theme";
  elements: NormalizedElement[];
  notes: NormalizedTextParagraph[];
  hidden: boolean;
  section?: string;
  tags: string[];
  customData: Record<string, JsonValue>;
}

interface NormalizedPlaceholderDefinition {
  key: string;
  kind: "title" | "subtitle" | "body" | "image" |
    "table" | "footer" | "slideNumber";
  box: Box;
  textStyle: {
    frame: NormalizedTextFrameStyle;
    paragraph: NormalizedTextParagraphStyle;
    run: NormalizedTextRunStyle;
  };
}
```

## Common element schema

```ts
interface NormalizedElementBase {
  id: string;
  name: string;
  box: Box;
  transform: { rotation: number; flipH: boolean; flipV: boolean };
  opacity: number;
  hidden: boolean;
  accessibility: { description?: string; decorative: boolean };
  action?:
    | { type: "url"; url: string; tooltip?: string }
    | { type: "slide"; slideId: string; tooltip?: string };
  placeholderKey?: string;
}
```

## Type-specific element schema

```ts
interface NormalizedTextElement extends NormalizedElementBase {
  type: "text";
  content: NormalizedTextParagraph[];
  plainText: string;
  frame: NormalizedTextFrameStyle;
}

interface NormalizedImageElement extends NormalizedElementBase {
  type: "image";
  assetId: string;
  fit: "stretch" | "contain" | "cover" | "crop";
  crop: { left: number; top: number; right: number; bottom: number };
}

interface NormalizedShapeElement extends NormalizedElementBase {
  type: "shape";
  shape: "rect" | "roundRect" | "ellipse" | "triangle" |
    "diamond" | "arrow" | "chevron";
  style: NormalizedShapeStyle;
  text?: {
    content: NormalizedTextParagraph[];
    plainText: string;
    frame: NormalizedTextFrameStyle;
  };
}

interface NormalizedConnectorElement extends NormalizedElementBase {
  type: "connector";
  start: Point | { elementId: string; anchor?: ConnectorAnchor };
  end: Point | { elementId: string; anchor?: ConnectorAnchor };
  route: Point[];
  style: NormalizedStrokeStyle;
}

interface NormalizedGroupElement extends NormalizedElementBase {
  type: "group";
  coordinateSize: Size;
  children: NormalizedElement[];
}

interface NormalizedTableElement extends NormalizedElementBase {
  type: "table";
  columns: number[];
  rows: Array<{
    height: number;
    cells: Array<{
      content: NormalizedTextParagraph[];
      rowSpan: number;
      colSpan: number;
      style: NormalizedTableCellStyle;
    }>;
  }>;
}
```

`NormalizedElement` is the union of these six interfaces. A shape may optionally
carry an editable text body; this remains one element with one identity rather
than becoming a shape plus an independent text element.

## Normalized style schema

```ts
type NormalizedPaint =
  | { type: "none" }
  | { type: "solid"; color: ColorValue; opacity: number };

interface NormalizedStrokeStyle {
  paint: NormalizedPaint;
  width: number;
  dash: "solid" | "dash" | "dot" | "dashDot";
  beginArrow: ArrowType;
  endArrow: ArrowType;
}

interface NormalizedTextFrameStyle {
  margin: { top: number; right: number; bottom: number; left: number };
  verticalAlign: "top" | "middle" | "bottom";
  wrap: boolean;
  autoFit:
    | { mode: "none" }
    | { mode: "shrink"; fontScale: number }
    | { mode: "resize" };
}

interface NormalizedTextParagraph {
  runs: NormalizedTextRun[];
  style: {
    align: "left" | "center" | "right" | "justify";
    indent: number;
    hanging: number;
    lineSpacing: number;
    spaceBefore: number;
    spaceAfter: number;
    bullet: NormalizedTextBullet;
  };
}

interface NormalizedTextRun {
  text: string;
  style: {
    fontFamily: string | { theme: "heading" | "body" };
    fontSize: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strike: boolean;
    color: ColorValue;
    language: string;
  };
  action?: ElementAction;
}
```

Shape styles contain explicit `fill` and `stroke`. Table cell styles contain explicit fill, stroke, four margins, and vertical alignment.

## Invariants

- Every slide, element, nested group child, asset, and layout has a stable ID.
- Element IDs are globally unique within the document.
- All geometry uses points; crop values use normalized `0..1` fractions.
- Array order is semantic presentation/drawing order; no `zIndex` exists.
- Every element has an explicit name, box, transform, opacity, hidden state, and accessibility object.
- Every paint, stroke, text style, crop, fit, table span/style, background, metadata field, theme role, and placeholder text style is materialized.
- Text is always paragraphs containing runs; text elements additionally contain `plainText`.
- Slides reference layouts by stable ID and record the source of their effective background.
- References have passed Core validation before ordinary normalization returns.
- Output shares no mutable references with authoring state.

## Slide and layout semantics

Normalized slides contain layout identity, effective background and source, slide-local elements, structured notes, hidden state, optional section, tags, and JSON-compatible custom data.

Normalized layouts contain a stable ID/name, effective background, static elements, and typed placeholders. Placeholder definitions include explicit boxes and fully normalized frame, paragraph, and run defaults.

Layout elements are not copied into slide-local element arrays. Downstream consumers combine the referenced layout and slide while preserving their distinct semantics.

## Elements

IR v1 supports text, image, shape, connector, group, and table elements.

- Text contains structured paragraphs/runs and a normalized frame.
- Images contain stable asset references, explicit fit mode, and four crop edges.
- Shapes contain a preset semantic shape and explicit fill/stroke.
- Connectors retain point or element-anchor intent until Layout resolves it.
- Groups retain a local coordinate size and recursively normalized children.
- Tables contain explicit column widths, row heights, cell spans, rich text, and cell styles.

## Default ownership

Core normalization is the sole owner of business defaults and semantic inheritance. Layout may resolve geometry, and exporters may map values to format-specific constructs, but neither stage chooses missing fonts, colors, borders, or text behavior.

## Versioning

Incompatible changes require a new `irVersion`. The repository is pre-release and does not maintain a parallel compatibility layer for superseded unpublished IR shapes. Documentation, tests, Layout, and exporters move together when the version changes.

See [Core Authoring Model](core-authoring-model.md) for the design rationale and [Core validation and IR API](../api/core/validation-and-ir.md) for callable behavior.
