# Notes page source inventory

`inspectPptxNotes(bytes, options?)` is a standalone read-only API from `@pptkit/pptx-parser`, available in Node, browser ESM and the `PPTKit` browser global. Its schema is `pptkit-notes-source-inventory-v1`. Existing `inspectPptx`, Core import and save profiles retain their contracts.

```ts
import { inspectPptxNotes } from "@pptkit/pptx-parser";
const inventory = await inspectPptxNotes(bytes, {
  notesLimits: { maxRecords: 100_000 },
});
for (const page of inventory.pages) {
  console.log(page.slidePart, page.notesPart, page.body.status);
}
```

See the [typed example](../examples/notes-source-inventory.ts) and [ADR0056](../architecture/decisions/0056-notes-source-inventory.md).

For local files, `pptkit inspect deck.pptx --notes-inventory --json` adds this inventory to the [CLI inspection envelope](cli-inspection.md), including the captured input hash. Omitting `--json` displays a compact summary. Both retain the same source-only boundaries.

## Observations and interpretation

Pages follow presentation relationship order, including slides without notes. `notesSize` uses the existing structural inspection dimensions in EMU. Each page identifies its actual notes/master/master-theme/local-theme-override relationships; a missing edge never borrows another slide's theme or the presentation's listed master. The existing OPC/structure validator checks relationship types, part roots/content types and the notes-to-slide backlink. Referenced unlisted masters remain diagnosed source observations, not standard multi-master authoring support.

Distinct listed or referenced notes/master parts are scanned once. Records retain part and one-based **raw XML element preorder**, expanded names and literal decoded attributes. Prefixes and lexical number strings are retained; entity spelling, encoding, comments and the original XML bytes are not reconstructed. Source locations are identities within this input, not stable IDs across native resaves.

Drawings distinguish the group root, body, slide image, header, footer, date, slide number, unclassified placeholders and artwork. These are source classifications, not renderer objects. Effective drawings use the selected MCE projection for ownership and nonvisual placeholder structure, while their records point to actual raw nodes. Inactive branches remain `raw-only`; ambiguous nonvisual children there are left unselected instead of rejecting the selected view; unknown drawing roots and extension islands retain their names and source attributes without interpreting their descendants as native content. ProcessContent-promoted children can be effective under a raw-only wrapper. Raw-only drawings never become direct body candidates.

`body.status` distinguishes absent, unique direct shape body, multiple direct bodies and missing/ambiguous local text bodies. Grouped bodies and master prompt text never become page body text. Placeholder candidates match **explicit type** within the actual master only: idx and name never select a candidate. `matched-direct-type` means one direct source candidate, not resolved geometry or a native rendering guarantee. Missing type/master/candidate, duplicates and groups remain explicit. No schema default placeholder type is guessed.

Drawing identity includes raw id/name and the nonvisual attributes, including any stored hidden/description values. Duplicate effective drawing IDs use numeric uint32 identity, including admitted signed/zero-padded lexical aliases; malformed IDs remain uninterpreted findings. This lookup profile does not validate the entire OOXML schema.

Transforms and preset/custom geometry retain their own attributes and direct components. Offsets, extents, rotation and flips remain strings. Custom paths, crop, inherited transforms, visibility and clipping are **unresolved source-only**. Root visibility flags and header/footer flags are observations, not an evaluated display policy. The API does not flatten artwork into ordinary slide Layout or Renderer output.

Text records contain direct decoded text/CDATA. Text and field records include the enclosing DrawingML paragraph raw element index when present, grouping records without asserting editability or evaluating a field. Field records retain id/type and stored cache segments selected through MCE (raw source identities retained), distinguishing an empty text segment from no segment. Effective field caches use direct text children after MCE projection, retaining raw source references through wrappers; inactive cache alternatives are not selected values. Unknown field types remain observable. This inventory is not a full XML tree; use XML inspection for unlisted descendants. Dates and slide/page numbers are never evaluated; no clock, locale or native application is consulted. Extension children are named but not parsed into fields or artwork. A field cache is not evidence of the currently displayed value.

## Resources and failure

Existing ZIP, XML, OPC and MCE limits/cancellation apply. Additional report-wide `notesLimits` have these defaults:

| Limit | Default | Accounting |
| --- | ---: | --- |
| maxWorkUnits | 1,000,000 | Inventory traversals and candidate work |
| maxRecords | 100,000 | Report objects, attributes, locations and candidates |
| maxFindings | 10,000 | Notes and inherited structural findings |
| maxTextCharacters | 1,000,000 | Captured strings in UTF-16 code units |

`workUnits`, `recordCount` and `textCharacters` are inventory accounting, not total infrastructure allocation or serialized byte size. ZIP/XML/MCE/source indexing have their own budgets. Limits must be positive safe integers; unknown keys and accessors reject without invoking getters. Bytes and option values are captured before asynchronous work. Exhaustion throws `PackageInspectionError` with `package-limit-exceeded`; cancellation uses `operation-aborted`. There is no partial-success result. Reports are deeply frozen detached plain data and support JSON serialization.

`geometryPolicy` is always `unresolved-source-only`; `fieldEvaluation` and `authoredMapping` are always `not-performed`. The report grants no source write authority and introduces no Core notes-page authoring model. The existing [direct notes import](notes-import.md), [notes text style profile](notes-style-import.md) and [scalar preservation saves](text-owner-save.md) remain separate.

[Focused tests](../../packages/pptx-parser/test/notes-inventory.test.mjs), [independent DOM/browser parity](../../scripts/notes-inventory-parity.test.mjs) and [captured native-file inspection](../testing/notes-source-inventory.md) cover source identities, isolation, failures and exact no-op preservation. Native field evaluation, complete notes rendering and broader authoring/editing remain pending.
