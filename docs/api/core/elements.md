# Elements

All elements share one base contract and belong to exactly one presentation document. Authoring inputs may omit IDs; stored elements and Canonical IR entries always have IDs.

## Shared fields

```ts
interface ElementBaseInput {
  id?: string;
  name?: string;
  box?: Box;
  transform?: TransformInput;
  opacity?: number;
  hidden?: boolean;
  accessibility?: ElementAccessibilityInput;
  action?: ElementAction;
  placeholderKey?: string;
}
```

| Field | Default | Description |
| --- | --- | --- |
| `id` | generated | Globally stable document element ID. |
| `name` | generated from ID | Human-readable object name. |
| `box` | placeholder box or zero box | Position and size in points. Required in practice unless a valid placeholder supplies it; connectors derive their box during layout. |
| `transform` | no rotation or flips | Rotation in degrees plus horizontal and vertical flipping. |
| `opacity` | `1` | Whole-element opacity in the inclusive range `0..1`. |
| `hidden` | `false` | Excludes the element from visible output without deleting it. |
| `accessibility` | non-decorative, no description | Description and decorative status. |
| `action` | omitted | URL or stable slide-ID navigation. |
| `placeholderKey` | omitted | Binds the element to a placeholder on the slide's layout. |

```ts
type ElementAction =
  | { type: "url"; url: string; tooltip?: string }
  | { type: "slide"; slideId: string; tooltip?: string };
```

Slide actions use `slideId`, never a page number that can become stale after reordering.

## Text

```ts
interface TextElementInput extends Omit<ElementBaseInput, "box"> {
  type: "text";
  content: string | TextParagraphInput[];
  box?: Omit<Box, "height"> & { height?: number };
  frame?: TextFrameStyleInput;
}
```

`content` is required. A string is an authoring convenience; normalized output always contains paragraphs and runs. See [Text and styles](text-and-styles.md) for rich text and inheritance.

Text boxes may omit `height` when `x`, `y`, and `width` are provided. Core estimates the intrinsic height from normalized paragraphs, wrapping, font sizes, line spacing, paragraph spacing, indentation, and text-frame margins. An explicit `height` remains authoritative. Width is never inferred because it defines the wrapping boundary.

```ts
slide.addElement({
  type: "text",
  content: "A simple editable text box",
  box: { x: 48, y: 48, width: 420, height: 48 },
  action: { type: "url", url: "https://example.com" },
});
```

## Image

```ts
interface ImageElementInput extends ElementBaseInput {
  type: "image";
  assetId: string;
  fit?: "stretch" | "contain" | "cover" | "crop";
  crop?: Partial<{ left: number; top: number; right: number; bottom: number }>;
}
```

| Field | Default | Notes |
| --- | --- | --- |
| `assetId` | required | Must reference a registered image asset. |
| `fit` | `stretch` | `contain` and `cover` require source dimensions for geometric resolution. |
| `crop` | all edges `0` | Edge fractions use normalized `0..1` coordinates. Used directly with `crop`; layout derives values for `cover`. |

```ts
slide.addElement({
  type: "image",
  assetId: hero.id,
  box: { x: 480, y: 80, width: 400, height: 240 },
  fit: "cover",
  accessibility: { description: "Product dashboard overview" },
});
```

## Shape

```ts
interface ShapeElementInput extends ElementBaseInput {
  type: "shape";
  shape: "rect" | "roundRect" | "ellipse" | "triangle" |
    "diamond" | "arrow" | "chevron";
  style?: ShapeStyleInput;
}
```

Shapes default to no fill and no visible stroke. Supply explicit paint when the shape should be visible.

```ts
slide.addElement({
  type: "shape",
  shape: "roundRect",
  box: { x: 48, y: 140, width: 240, height: 100 },
  style: {
    fill: { type: "solid", color: { theme: "accent1" }, opacity: 0.15 },
    stroke: {
      paint: { type: "solid", color: { theme: "accent1" } },
      width: 2,
    },
  },
});
```

## Connector

```ts
type ConnectorEndpointInput =
  | Point
  | { elementId: string; anchor?: "top" | "right" | "bottom" | "left" | "center" };

interface ConnectorElementInput extends ElementBaseInput {
  type: "connector";
  start: ConnectorEndpointInput;
  end: ConnectorEndpointInput;
  route?: Point[];
  style?: StrokeStyleInput;
}
```

Connectors can use absolute points or stable element references. `route` contains intermediate polyline points. Layout resolves references to points and derives the final connector box. Referenced elements must be in the same slide/layout scope.

```ts
slide.addElement({
  type: "connector",
  start: { elementId: "source", anchor: "right" },
  end: { elementId: "target", anchor: "left" },
  style: {
    paint: { type: "solid", color: "404040" },
    width: 2,
    endArrow: "triangle",
  },
});
```

## Group

```ts
interface GroupElementInput extends ElementBaseInput {
  type: "group";
  coordinateSize: Size;
  children: PresentationElementInput[];
}
```

`box` positions and scales the group in its parent coordinate system. `coordinateSize` defines the local coordinate system used by children. Both dimensions must be positive. Groups may be nested; every descendant receives a globally unique ID.

```ts
slide.addElement({
  type: "group",
  box: { x: 500, y: 80, width: 300, height: 180 },
  coordinateSize: { width: 600, height: 360 },
  children: [
    {
      type: "shape",
      shape: "diamond",
      box: { x: 40, y: 40, width: 160, height: 120 },
      style: { fill: { type: "solid", color: { theme: "accent2" } } },
    },
  ],
});
```

## Table

```ts
interface TableElementInput extends ElementBaseInput {
  type: "table";
  columns: number[];
  rows: Array<{
    height?: number;
    cells: Array<{
      content: TextContentInput;
      rowSpan?: number;
      colSpan?: number;
      style?: TableCellStyleInput;
    }>;
  }>;
}
```

Column widths and optional row heights are points. Spans are positive integers and a row may not span more logical columns than the table declares. Core validates structure; automatic measurement, overflow, and pagination belong to Layout.

```ts
slide.addElement({
  type: "table",
  box: { x: 48, y: 280, width: 600, height: 160 },
  columns: [240, 180, 180],
  rows: [
    {
      cells: [{
        content: "Quarterly results",
        colSpan: 3,
        style: { fill: { type: "solid", color: { theme: "accent1" } } },
      }],
    },
    { cells: [{ content: "Q1" }, { content: "Q2" }, { content: "Q3" }] },
  ],
});
```

## Validation summary

Element validation covers geometry, opacity, paint and stroke ranges, text structure, image assets/crops, connector references, group coordinate sizes, table spans, placeholder binding, and action targets. Validation returns structured diagnostics rather than stopping at the first cross-document error.
