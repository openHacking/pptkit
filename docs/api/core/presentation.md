# Presentations and Slides

Presentation and slide state is changed through methods. The exposed `slides`, `elements`, `assets`, and `layouts` collections are frozen snapshots, so array mutation cannot bypass identity and ownership checks.

## `createPresentation(init?)`

```ts
declare function createPresentation(
  init?: PresentationInit,
): ManagedPresentationDocument;
```

Creates one authoring document. Inputs are cloned and frozen when accepted; later changes to the caller's objects do not mutate the document.

`ManagedPresentationDocument` extends the existing structural `PresentationDocument` with shared master/layout editing; custom base-interface implementations need no new methods. See [master ownership and editing](../multi-master-authoring.md#editing-shared-definitions) for complete replacement, validation and rollback semantics.

### `PresentationInit`

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `id` | `string` | generated document ID | Stable identity for the presentation. |
| `firstSlideNumber` | `number` | absent | Nonnegative safe integer for native starting slide numbering; `setFirstSlideNumber(value?)` replaces or clears it. PPTX output enforces its own 0–9999 range. |
| `metadata` | `PresentationMetadataInput` | empty authoring metadata | Title, authoring and discovery metadata. |
| `size` | `Partial<PresentationSize>` | `960 × 540 pt` | Presentation dimensions. Only points are supported. |
| `theme` | `PresentationThemeInput` | Core theme | Partial theme overrides. |
| `masters` | `readonly SlideMasterInput[]` | `[]` | Shared independent theme/background owners selected by layout `masterId`. |
| `textStylePresets` | `Readonly<Record<string, TextStylePresetInput>>` | `{}` | Immutable named frame, paragraph, and run defaults. |

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
  readonly textStylePresets: Readonly<Record<string, TextStylePresetInput>>;
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

Factory-managed `addSlide` and `insertSlide` stage the slide, nested element identities and timeline before insertion. A failed insertion does not consume generated identities or install part of the slide. If a caller-provided getter reenters and changes slide ordering or element ownership during staging, the insertion rejects; that separate reentrant change remains in place.

When `duplicateSlide()` has no destination, the copy is inserted immediately after the source. Its generated slide and element identities are staged together; a failed copy leaves both sequences available for retry. Internal connector targets and self-slide actions are remapped to copied identities; other page links, assets and layout bindings retain their targets. See [copy reference semantics](../copy-references.md).

### `PresentationSlideInput`

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `id` | `string` | generated | Stable slide identity. |
| `layoutId` | `string` | synthesized blank layout | Reference to a defined layout. |
| `background` | `BackgroundPaintInput` | inherited | Slide-local background override. |
| `elements` | `PresentationElementInput[]` | `[]` | Convenience input; elements are still inserted through document ownership checks. |
| `notes` | `TextContentInput` | empty | Speaker notes. |
| `hidden` | `boolean` | `false` | Whether the slide is hidden during a show. |
| `section` | `string` | omitted | Format-independent section label. Native import may project a source member name here; this does not represent a native section identity or empty-section collection. |
| `tags` | `string[]` | `[]` | Application tags. |
| `customData` | `Record<string, JsonValue>` | `{}` | JSON-compatible application metadata. |

`slide.setNotes(content?)` replaces speaker notes with an owned frozen snapshot; omit content to clear them. Existing returned notes and normalized snapshots remain independent. Copies include the current notes and remap self-slide actions. Custom `PresentationSlide` implementations must implement this new method. See [notes import/editing](../notes-import.md) for source projection and save limits.

Backgrounds also accept `{ type: "linearGradient", angle, scaled?, stops }` on masters, layouts, and slides. `angle` is in degrees from the positive x-axis (clockwise on the slide), in `[0, 360)`; `scaled` defaults to `false` and uses normalized fill-box coordinates when true. Provide 2–256 stops in nondecreasing order, each with `position` in `[0, 1]`, a color, and optional opacity in `[0, 1]` (default 1). Equal stop positions are allowed. PPTX output quantizes positions to 1/100000 and angles to 1/60000 degree. Linear gradients are also accepted in ordinary shape fills; line, table and chart gradients remain unsupported. Authored shape gradients rotate and flip with their shape; the opt-in `artwork-v8` importer reconstructs a bounded shared master/layout subset and `artwork-v9` adds a bounded slide-local subset. SVG emits `shape-gradient-approximation` because PowerPoint interpolation and geometry may differ. SVG emits `background-gradient-approximation` for each gradient slide because PowerPoint color and opacity interpolation can differ. A scoped native sample informed scaled geometry, but general PowerPoint visual parity is not established.

Ordinary shape fills also accept `{ type: "pattern", pattern, foreground, background, foregroundOpacity?, backgroundOpacity? }`, with `pattern` equal to `"horizontal"`, `"vertical"`, `"cross"`, `"downwardDiagonal"`, `"upwardDiagonal"`, or `"diagonalCross"`, or one of `"percent5"`, `"percent10"`, `"percent20"`, `"percent25"`, `"percent30"`, `"percent40"`, `"percent50"`, `"percent60"`, `"percent70"`, `"percent75"`, `"percent80"`, and `"percent90"`. Percentage previews use a circular tile with nominal foreground coverage and do not reproduce the Office preset texture. Colors accept RGB or theme roles; opacities are in `[0, 1]` and default to 1. PPTX exports an editable pattern fill and composes object/group opacity. SVG uses an 8-point approximate tile and reports `shape-pattern-approximation`; spacing, compositing and general appearance have not been validated against PowerPoint. One opaque black-and-white sample containing all twelve percentage presets has a [PowerPoint save/reopen observation](../../testing/percent-pattern-powerpoint-evidence.json), checking the saved preset values and retained UI state. The [saved file](../../../tests/fixtures/percent-pattern-powerpoint-resave.pptx) also passes three programmatic import/rebuild rounds and a final reimport; this does not establish visual parity or interactive editing. Pattern fills on backgrounds, lines, tables and charts remain unsupported; opt-in `artwork-v10` imports these explicit shape patterns as a partial content projection.

Shape styles may also specify `shadow: { color, opacity?, blur, distance, direction }` for an outer shadow. Blur and distance are points in `[0, 1000]`; direction is clockwise degrees in `[0, 360)`, and opacity defaults to 1. The shadow follows shape rotation and uses the shape's theme scope. PPTX exports a DrawingML outer shadow, with authored effect opacity separate from shape paint opacity. When the shape, ancestor group, fill or stroke is translucent, export reports `shape-shadow-transparency-application-difference`: target applications may compose transparency differently. SVG applies a bounded approximate filter to shape geometry and stroke, not shape text, and reports `shape-shadow-approximation`. An unblurred half-opacity shape rendered a darker shadow in one LibreOfficeDev PDF comparison than SVG; opacity composition and native PowerPoint appearance remain unverified. Layout offers an opt-in approximate shadow-boundary diagnostic. Opt-in `artwork-v11` imports a bounded explicit outer-shadow subset; preservation editing of shadow effects remains unsupported.

Factory-owned `ManagedPresentationSlide` instances provide `setBackground(value?)`, which replaces the local background with an owned frozen snapshot. Omit the value to restore layout/master inheritance. If snapshotting fails, the previous value remains intact; validation still occurs when the document is validated or normalized.

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

These methods follow the same append, insert, move, remove, duplicate, unknown-ID, and index rules as slide operations. IDs are globally owned by the document, not merely unique inside one slide. Duplicating a group regenerates the group ID and every descendant ID, remapping connector targets inside that copied subtree. References outside the subtree remain unchanged.

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

Factory-owned slides also provide `reorderElements(ids)`: pass every current top-level element ID exactly once to replace the drawing order atomically. It retains the existing objects and their IDs, nested groups, references and timing program. Invalid, duplicate, missing or nested IDs reject before the operation changes the order; if caller code mutates the slide while the ID array is being read, the operation rejects without undoing that caller mutation. An empty slide accepts an empty order. This changes visual stacking and emitted object order, not an independent accessibility reading-order setting. Reacquire the current IDs before each operation; IDs from another slide or stale snapshots are not accepted merely because they once existed.

`reorderGroupElements(groupId, ids)` performs the same complete permutation for the direct children of an owned group, including a nested group. It preserves IDs and references but replaces the frozen top-level group snapshot containing the target, so callers must reacquire object references afterward; the old snapshot remains unchanged. The operation validates the current timing program before committing and rejects a non-group ID, unknown group, descendants that are not direct children, duplicates and incomplete lists. A reentrant caller mutation is not rolled back. Only slide-owned groups are editable through this method; layout and master definitions have separate ownership. Group order affects stacking and does not establish an independent reading order.

A [named LibreOfficeDev headless experiment](../../testing/reorder-libreoffice.md) observes top-level and nested orders through three saves, while recording pixel changes and removal of group description attributes. It does not establish interactive or PowerPoint behavior.

## Immediate errors versus diagnostics

Operations reject states they can identify locally, such as duplicate IDs, unknown IDs, and invalid insertion positions. Cross-document issues such as missing assets, missing layouts, invalid action targets, invalid geometry, and style ranges are reported by [`validatePresentation()`](validation-and-ir.md).

## Page transitions

`addSlide({ transition })` and `slide.setTransition()` snapshot or clear the experimental [Fade transition](../slide-transitions.md). Copies retain independent transition values. The optional IR field contains timing data; Core does not play slides.

## Bound foreground timing

`setEventTimeline()` retains validated authored stages; `getEventTimelinePlan()` binds them to current owned IDs. Whole-slide copies remap target/trigger identities, and referenced object/group removal rejects. `replaceElement()` atomically updates a top-level object with its root identity retained. See [the full profile](../slide-event-timeline.md) for budgets, subtree guards, object-only copy behavior and native-output refusal.


Table edits compose through [the pure table transaction](../table-editing.md) and `replaceElement`; the grid transformation adds no registry mutation or source-write permission.

Tables can use table-local [region style rules](../table-style.md), resolved into canonical cells; these introduce no native-style GUID registry or new theme mutation API.


Placeholder definitions may declare a layout-local `index` (unsigned 32-bit integer, including zero). Explicit indices are unique within the layout; omitted indices remain omitted in authoring/IR. See [index identity, matching and native allocation](../layout-rebind.md) and [verification limits](../../testing/placeholder-index.md). Native serialization numbers generated for unindexed definitions do not become Core identities.
Validation also requires a nonempty string key and a supported `PlaceholderKind`. Malformed JavaScript values receive located `invalid-placeholder-key`/`invalid-placeholder-kind` diagnostics before normalization, matching or export; unsupported kinds are not coerced into body prompts. Definition replacement remains atomic on failure.

Managed slides expose `setHidden(boolean)` to change slideshow visibility; `hidden` remains read-only, and invalid values throw before mutation. Legacy custom slide implementations may omit this optional capability. [Source-preserving save](../slide-visibility-save.md) requires an explicit Parser profile.

Managed documents expose `setMetadata(metadata)` to replace the full metadata value with a validated, detached snapshot. Spread the existing metadata when updating one field; invalid input leaves the old value unchanged. Legacy custom documents may omit this optional method. [Source-preserving metadata saves](../core-metadata-save.md) require an explicit Parser profile.
