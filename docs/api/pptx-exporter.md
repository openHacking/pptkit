# `@pptkit/pptx-exporter`

`@pptkit/pptx-exporter` converts an `@pptkit/core` presentation into a minimal editable PowerPoint package. Byte generation is runtime-neutral; filesystem delivery is a Node.js adapter.

## Runtime-neutral API

```ts
interface GeneratePptxResult {
  bytes: Uint8Array;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "generated" | "generated-with-warnings";
}

declare function generatePptx(
  document: PresentationDocument,
): Promise<GeneratePptxResult>;
```

Import `generatePptx` from `@pptkit/pptx-exporter` in either a browser or Node.js. The caller owns delivery: for example, a browser application may wrap `bytes` in a Blob.

## Node.js API

```ts
interface WritePptxOptions {
  output: string;
}

interface WritePptxResult {
  output: string;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "written" | "written-with-warnings";
}

declare function writePptx(
  document: PresentationDocument,
  options: WritePptxOptions,
): Promise<WritePptxResult>;
```

Import `writePptx` from `@pptkit/pptx-exporter/node`. The output directory is created automatically and an existing file is overwritten. The Node.js subpath also exports `generatePptx` with local-path asset support.

## Supported content and assets

| Core element | PPTX output |
| --- | --- |
| `text` | Editable text box with font family, size, weight, color, alignment, paragraph line spacing, line breaks, and optional shrink-to-fit |
| `shape: rect` | Editable rectangle |
| `shape: ellipse` | Editable ellipse |
| `shape: line` | Editable line |
| `image` | Editable picture with alt text and packaged media |

The default entry supports URL sources through the runtime `fetch`, including HTTP(S), data, and blob URLs. Browser HTTP(S) requests obey CORS. A `path` source used through the default entry produces an `asset-read-failed` warning; the Node.js entry supports both `path` and `url`.

Core normalization errors fail the operation. Recoverable asset failures omit only the affected image and are reported through `asset-read-failed` and `image-omitted` warnings. Unsupported elements use `unsupported-element`.

Solid slide backgrounds are exported as native PowerPoint slide backgrounds rather than full-slide shapes. Fonts are referenced by family name and are not embedded; PowerPoint may substitute a font that is unavailable on the viewing machine.

Themes, rich text runs, grouped elements, tables, animations, speaker notes, and PPTX parsing remain outside the current implementation.
