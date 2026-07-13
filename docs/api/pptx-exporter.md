# `@pptkit/pptx-exporter`

`@pptkit/pptx-exporter` validates and normalizes a Core authoring document, resolves layout once, loads image assets, serializes editable OOXML parts, and returns a deterministic PPTX package. The default entry is browser-neutral; the `/node` entry additionally reads local files and writes output.

## Browser-neutral `generatePptx(document)`

```ts
declare function generatePptx(
  document: PresentationDocument,
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

The default entry loads `source.type: "url"` assets using `fetch`. HTTP(S), data, and blob URLs are accepted when supported by the runtime. Local path assets are omitted with a warning.

## Node.js entry

The Node entry exports both `generatePptx()` and `writePptx()` and supports URL and path assets.

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
  options: { output: string },
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

## Exported presentation features

- Editable rich text paragraphs and runs, direct/theme colors, theme fonts, bullets, numbering, hyperlinks, and slide actions.
- Images with stretch, contain, cover, crop, opacity, accessibility, and media relationships.
- Shapes, connectors, nested groups, and native editable tables with cell spans.
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
