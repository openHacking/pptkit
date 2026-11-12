# `@pptkit/core`

Authoring model and Canonical Presentation IR for PPTKit.

## Capabilities

- method-managed presentations, slides, elements, assets, and slide layouts
- rich text paragraphs and runs
- native text-bearing shapes and document-level text style presets
- theme colors and fonts
- text, image, shape, connector, group, table, and chart (bar/line/area/pie/doughnut/scatter) elements
- placeholders, speaker notes, hidden slides, sections, tags, and custom data
- structured validation and fully materialized IR v2 normalization

## Example

```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";

const presentation = createPresentation({
  metadata: { title: "Quarterly Update", author: "Example Team" },
  theme: { colors: { accent1: "2457D6" } },
  textStylePresets: {
    title: { frame: { margin: 0, verticalAlign: "middle" }, paragraph: { align: "center" }, run: { fontSize: 32, bold: true } },
  },
});

const slide = presentation.addSlide();
slide.addElement({
  type: "text",
  content: [{
    runs: [
      { text: "Quarterly ", style: { fontSize: 32 } },
      { text: "Update", style: { fontSize: 32, bold: true, color: { theme: "accent1" } } },
    ],
  }],
  box: { x: 48, y: 48, width: 500, height: 60 },
});

const normalized = normalizePresentation(presentation);
console.log(normalized.irVersion); // 1

slide.addElement({
  type: "shape",
  shape: "roundRect",
  box: { x: 48, y: 120, width: 500, height: 80 },
  text: { content: "Editable card", textStylePreset: "title" },
});
```

Collections are exposed as readonly snapshots. Use document and slide methods to add, insert, move, remove, or duplicate content.

See [the Core API reference](../../docs/api/core.md) and [authoring architecture](../../docs/architecture/core-authoring-model.md).

Experimental [finite timeline arithmetic](../../docs/api/finite-timeline.md) compiles abstract spans into immutable versioned intervals and samples them at absolute integer milliseconds. It has no drawing/effect targets, clock, Core/IR mutations or native animation serialization. Event/effect/target contracts and native playback remain separate tasks.

Experimental [foreground event replay](../../docs/api/event-timeline.md) compiles explicit target inventories and finite gate stages, evaluates ordered visit-scoped clicks and returns immutable observations. Current inventories detect stale targets; clock/controller, authoring/IR transport, native effects and playback remain separate.

Slides can explicitly attach [bound foreground timing](../../docs/api/slide-event-timeline.md). Owned targets follow whole-slide copy and survive IR/Layout transport; referenced deletion rejects and `replaceElement` preserves logical identity. Native output remains unavailable and requires explicit static omission rather than silent loss.


Use `editTable(table, operations, { frame: "fit-grid" | "preserve" })` for atomic compact-grid structural candidates and source-anchor mappings, then commit with `replaceElement`. See the [contract](../../docs/api/table-editing.md).

Direct cell styles now accept sparse four-sided `borders`; see the [contract](../../docs/api/table-borders.md). Full table-style inheritance remains open.

Table-local [region styles](../../docs/api/table-style.md) add whole/band/first-last/corner/text defaults with bounded Core materialization and explicit local overrides.
