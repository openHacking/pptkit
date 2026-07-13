# `@pptkit/core`

`@pptkit/core` is the authoring-facing package for PPTKit. It owns presentation construction, typed slide elements, asset registration, and the normalized document boundary that downstream packages consume.

## Install

```bash
pnpm add @pptkit/core
```

## Quick Example

```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";

const presentation = createPresentation({
  title: "Hello PPTKit",
});

const heroImage = presentation.registerAsset({
  kind: "image",
  source: {
    type: "path",
    value: "./assets/hero.png",
  },
  mimeType: "image/png",
  altText: "Quarterly launch hero graphic",
  dedupeKey: "launch-hero",
});

presentation.addSlide({
  elements: [
    {
      type: "text",
      text: "Hello, PPTKit",
      box: { x: 48, y: 48, width: 400, height: 32 },
      style: { fontSize: 24, fontWeight: "bold" },
    },
    {
      type: "image",
      assetId: heroImage.id,
      box: { x: 48, y: 112, width: 320, height: 180 },
      altText: "Launch hero preview",
    },
  ],
});

const normalized = normalizePresentation(presentation);

console.log(normalized.slides.length);
```

## Public API

### `createPresentation(init?)`

Creates an in-memory presentation document.

```ts
declare function createPresentation(init?: PresentationInit): PresentationDocument;
```

`PresentationInit`:

```ts
interface PresentationInit {
  id?: string;
  title?: string;
  size?: Partial<PresentationSize>;
}
```

Default size:

```ts
{
  width: 960,
  height: 540,
  unit: "pt"
}
```

### `presentation.addSlide(input?)`

Appends a slide to the document and returns the created slide.

```ts
interface PresentationSlideInput {
  id?: string;
  elements?: PresentationElementInput[];
}
```

If `id` is omitted, PPTKit generates a stable slide id such as `slide-1`.

### `presentation.registerAsset(input)`

Registers an asset for later element references.

```ts
interface PresentationAssetInput {
  id?: string;
  kind: "image";
  source: { type: "path" | "url"; value: string };
  mimeType?: string;
  width?: number;
  height?: number;
  altText?: string;
  dedupeKey?: string;
}
```

Current behavior:

- only image assets are supported
- repeat registrations with the same `id` or `dedupeKey + source` return the existing asset when metadata matches
- conflicting duplicate registrations throw an error

### `presentation.getAsset(assetId)`

Returns a previously registered asset when present.

### `normalizePresentation(document)`

Converts a mutable authoring document into a normalized structure for layout and export.

```ts
declare function normalizePresentation(
  document: PresentationDocument,
): NormalizedPresentation;
```

Normalization guarantees:

- generated defaults are materialized
- normalized arrays and objects do not share mutable references with authoring state
- image elements must reference an existing asset
- duplicate slide ids and asset ids are rejected
- negative or non-finite geometry values are rejected

## Types

### Geometry

```ts
type LengthUnit = "pt";

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface PresentationSize {
  width: number;
  height: number;
  unit: LengthUnit;
}
```

### Elements

```ts
type PresentationElementInput =
  | TextElementInput
  | ImageElementInput
  | ShapeElementInput;

interface TextElementInput {
  type: "text";
  text: string;
  box: Box;
  style?: TextStyle;
}

interface ImageElementInput {
  type: "image";
  assetId: string;
  box: Box;
  altText?: string;
}

interface ShapeElementInput {
  type: "shape";
  shape: "rect" | "ellipse" | "line";
  box: Box;
  style?: ShapeStyle;
}
```

Minimal style shells:

```ts
interface TextStyle {
  fontSize?: number;
  fontFamily?: string;
  fontWeight?: "normal" | "bold";
  color?: string;
  align?: "left" | "center" | "right";
}

interface ShapeStyle {
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
}
```

### Normalized Output

`NormalizedPresentation` keeps the same document, slide, asset, and element concepts, but every entry is copied into a validation-friendly structure with defaults applied.

Downstream packages should consume normalized data rather than authoring state directly.

## Validation Rules

`@pptkit/core` throws explicit errors for:

- duplicate slide ids
- duplicate asset ids with conflicting metadata
- image elements that reference missing assets
- negative or non-finite box dimensions
- invalid presentation or asset dimensions

## Asset Lifecycle Boundary

`@pptkit/core` owns:

- asset ids
- source metadata
- optional descriptive metadata such as MIME type, dimensions, and alt text
- deduplication keys

`@pptkit/core` does not currently own:

- file reads
- binary payload storage
- hashing bytes for deduplication
- packaging media into `.pptx`

Those responsibilities belong in downstream packages such as `@pptkit/pptx-exporter`.

## TODO

- richer text styling and text runs
- theme and token-aware style inheritance
- grouped elements, tables, speaker notes, and animations
- binary asset payload support if a future workflow requires in-memory authoring
