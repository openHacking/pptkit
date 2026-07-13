# Presentations and Slides

Presentation and slide state is changed through methods. The exposed `slides`, `elements`, `assets`, and `layouts` collections are frozen snapshots, so array mutation cannot bypass identity and ownership checks.

## `createPresentation(init?)`

```ts
declare function createPresentation(
  init?: PresentationInit,
): PresentationDocument;
```

Creates one authoring document. Inputs are cloned and frozen when accepted; later changes to the caller's objects do not mutate the document.

### `PresentationInit`

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `id` | `string` | generated document ID | Stable identity for the presentation. |
| `metadata` | `PresentationMetadataInput` | empty authoring metadata | Title, authoring and discovery metadata. |
| `size` | `Partial<PresentationSize>` | `960 × 540 pt` | Presentation dimensions. Only points are supported. |
| `theme` | `PresentationThemeInput` | Core theme | Partial theme overrides. |

`PresentationInit` intentionally contains no legacy top-level metadata aliases.

### Metadata fields

| Field | Type | Normalized default |
| --- | --- | --- |
| `title` | `string` | `"PPTKit Presentation"` |
| `author` | `string` | `"PPTKit"` |
| `company` | `string` | omitted |
| `subject` | `string` | omitted |
| `description` | `string` | omitted |
| `language` | `string` | `"en-US"` |
| `keywords` | `string[]` | `[]` |
| `revision` | `number` | `1` |

## Presentation collections

```ts
interface PresentationDocument {
  readonly id: string;
  readonly metadata: Readonly<PresentationMetadataInput>;
  readonly size: Readonly<PresentationSize>;
  readonly theme: Readonly<PresentationThemeInput>;
  readonly slides: readonly PresentationSlide[];
  readonly assets: readonly PresentationAsset[];
  readonly layouts: readonly SlideLayoutDefinition[];
}
```

Every collection access returns a readonly snapshot. An earlier snapshot does not change when later operations add or remove entries. Element array order is the only drawing-order contract: later elements draw above earlier elements. Core has no `zIndex`.

## Slide operations

```ts
addSlide(input?: PresentationSlideInput): PresentationSlide;
insertSlide(index: number, input?: PresentationSlideInput): PresentationSlide;
moveSlide(slideId: string, toIndex: number): void;
removeSlide(slideId: string): PresentationSlide;
duplicateSlide(slideId: string, toIndex?: number): PresentationSlide;
```

| Method | Behavior | Failure behavior |
| --- | --- | --- |
| `addSlide` | Appends and returns a slide. | Throws on a duplicate explicit slide or nested element ID. |
| `insertSlide` | Inserts before the entry currently at `index`; `length` appends. | `RangeError` when `index` is outside `0..length`. |
| `moveSlide` | Moves an existing slide to an index in the current collection. | Throws for an unknown ID; `RangeError` outside `0..length - 1`. |
| `removeSlide` | Removes and returns the slide, releasing all nested element IDs. | Throws for an unknown ID. |
| `duplicateSlide` | Deep-copies the slide and assigns fresh IDs to all copied elements, including group descendants. | Throws for an unknown ID or invalid explicit destination index. |

When `duplicateSlide()` has no destination, the copy is inserted immediately after the source.

### `PresentationSlideInput`

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `id` | `string` | generated | Stable slide identity. |
| `layoutId` | `string` | synthesized blank layout | Reference to a defined layout. |
| `background` | `PaintInput` | inherited | Slide-local background override. |
| `elements` | `PresentationElementInput[]` | `[]` | Convenience input; elements are still inserted through document ownership checks. |
| `notes` | `TextContentInput` | empty | Speaker notes. |
| `hidden` | `boolean` | `false` | Whether the slide is hidden during a show. |
| `section` | `string` | omitted | Format-independent section label. |
| `tags` | `string[]` | `[]` | Application tags. |
| `customData` | `Record<string, JsonValue>` | `{}` | JSON-compatible application metadata. |

## Element operations

```ts
interface PresentationSlide {
  addElement(input: PresentationElementInput): PresentationElement;
  insertElement(index: number, input: PresentationElementInput): PresentationElement;
  moveElement(elementId: string, toIndex: number): void;
  removeElement(elementId: string): PresentationElement;
  duplicateElement(elementId: string, toIndex?: number): PresentationElement;
}
```

These methods follow the same append, insert, move, remove, duplicate, unknown-ID, and index rules as slide operations. IDs are globally owned by the document, not merely unique inside one slide. Duplicating a group regenerates the group ID and every descendant ID.

## Ordering example

```ts
const presentation = createPresentation();
const first = presentation.addSlide({ id: "first" });
presentation.addSlide({ id: "second" });
presentation.moveSlide("second", 0);

const back = first.addElement({
  id: "background-card",
  type: "shape",
  shape: "rect",
  box: { x: 40, y: 40, width: 300, height: 160 },
  style: { fill: { type: "solid", color: "DDEBFF" } },
});

first.addElement({
  id: "label",
  type: "text",
  content: "Drawn above the card",
  box: { x: 64, y: 80, width: 240, height: 40 },
});

first.moveElement(back.id, 1); // The card now draws above the label.
const copy = presentation.duplicateSlide("first");
presentation.removeSlide(copy.id);
```

## Immediate errors versus diagnostics

Operations reject states they can identify locally, such as duplicate IDs, unknown IDs, and invalid insertion positions. Cross-document issues such as missing assets, missing layouts, invalid action targets, invalid geometry, and style ranges are reported by [`validatePresentation()`](validation-and-ir.md).
