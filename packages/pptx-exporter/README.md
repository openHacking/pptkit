# `@pptkit/pptx-exporter`

Editable PowerPoint export for PPTKit.

The exporter consumes resolved IR and emits a deterministic `.pptx` package with:

- rich text paragraphs, runs, bullets, theme fonts, and links
- image embedding, opacity, crop, contain, and cover behavior
- editable shapes, connectors, nested groups, and native tables
- native text bodies embedded in editable shapes
- presentation themes, reusable slide layouts, and placeholders
- slide backgrounds, hidden slides, speaker notes, and PPTKit metadata extensions
- recoverable structured warnings for asset-loading failures

The default entry loads URL assets. `@pptkit/pptx-exporter/node` additionally loads local paths and exposes `writePptx()`.

The package owns asset I/O, OOXML parts and relationships, ZIP generation, and output diagnostics. It does not own authoring or layout semantics.

See the [PPTX exporter API reference](../../docs/api/pptx-exporter.md) and [Rendering Pipeline](../../docs/architecture/rendering-pipeline.md).
