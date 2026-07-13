# `@pptkit/pptx-exporter`

`@pptkit/pptx-exporter` converts a normalized `@pptkit/core` presentation into a minimal editable PowerPoint package.

## API

```ts
interface ExportPptxOptions {
  output: string;
}

interface ExportWarning {
  code: string;
  message: string;
  slideId?: string;
  elementIndex?: number;
  assetId?: string;
}

interface ExportResult {
  output: string;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "written" | "written-with-warnings";
}

declare function exportPptx(
  document: PresentationDocument,
  options: ExportPptxOptions,
): Promise<ExportResult>;
```

The output directory is created automatically and an existing output file is overwritten. Core normalization errors still fail the operation. Recoverable image read errors produce warnings and omit only the affected image.

## Supported content

| Core element | PPTX output |
| --- | --- |
| `text` | Editable text box with basic font, weight, color, and alignment |
| `shape: rect` | Editable rectangle |
| `shape: ellipse` | Editable ellipse |
| `shape: line` | Editable line |
| `image` | Editable picture with alt text and packaged media |

Image assets support `path` and `url` sources. URL sources use the runtime `fetch` implementation and must return a successful response.

## Diagnostics

Warnings currently include `asset-read-failed`, `image-omitted`, and `unsupported-element`. The exporter never silently treats a core validation error as a warning.

The package model is private to the exporter; callers work with core authoring types and the structured result above.

## Current limitations

Themes, rich text runs, grouped elements, tables, animations, speaker notes, and PPTX parsing are outside this first export implementation.
