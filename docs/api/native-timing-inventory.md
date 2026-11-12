# Native timing source inventory

`inspectPptxTiming(bytes, options?)` is an explicit, read-only API from `@pptkit/pptx-parser` (Node, browser ESM and browser global builds). It returns `PptxTimingInspection` with schema `pptkit-native-timing-inventory-v1`. It does not change `inspectPptx`, content import or edit-save authority.

```ts
import { inspectPptxTiming } from "@pptkit/pptx-parser";
const report = await inspectPptxTiming(bytes, {
  signal: controller.signal,
  timingLimits: { maxRecords: 100_000 },
});
for (const slide of report.slides) {
  for (const node of slide.nodes) {
    console.log(node.source, node.role, node.drawingReference?.status);
  }
}
```

See the [typed example](../examples/native-timing-inventory.ts) and [ADR0038](../architecture/decisions/0038-native-timing-source-inventory.md).

## Source observations

Slides follow presentation relationship order, not filenames. Both Strict and Transitional dialects are supported. Each timing element retains its expanded namespace/local name, decoded attributes and direct text/CDATA, outer timing root index, raw parent index and nearest raw `cTn` ancestor. Source locations identify part, one-based raw element preorder and expanded-name sibling path. These are observations of this package, not stable identities across producer resaves. Entity spelling, comments, encoding and original XML bytes are not reconstructed by this report.

`effective` means membership in the selected MCE XML view under a direct slide timing container; it does **not** certify native activation or execution. Inactive branches, opaque extension islands and other timing containers remain `raw-only`. Unknown nodes are retained and reported. Native nodes within foreign containers do not gain execution semantics from their view label.

Drawing inventory retains `cNvPr` identity/name/owner observations under the slide drawing tree. Drawing IDs and timing IDs have separate, part-local candidate indexes. `spid` and `tn@val` references retain their original strings. The conservative lookup profile accepts canonical decimal uint32 strings for `located`; this is not general OOXML schema validation. Noncanonical strings, including numeric aliases, remain source-only. Alias candidates are shown, never silently selected. Multiple candidates are ambiguous; missing references are reported. Raw-only references have no effective candidates.

`located` means one supported source identity candidate, not a bound Core object, editable effect or executable target. `spTgt` children identify subobjects; build references and extra target attributes remain uninterpreted. Roles are inventory classifications, not a semantic validator. Durations, delay, repeat, autoReverse, conditions and behavior properties remain decoded strings without unit conversion or scheduling.

## Resource and failure contract

Existing ZIP/XML/OPC limits and cancellation apply, plus shared report-wide `timingLimits`:

| Limit | Default | Accounting |
| --- | ---: | --- |
| `maxWorkUnits` | 1,000,000 | Inventory traversal/enqueue, identity and reference operations |
| `maxRecords` | 100,000 | Published nodes, attributes, locations, candidates, findings and slide records |
| `maxFindings` | 10,000 | Findings across all slides |
| `maxTextCharacters` | 1,000,000 | Captured strings and constructed expanded paths, in UTF-16 code units |

Metrics `workUnits`, `recordCount`, `textCharacters` describe inventory accounting. Shared location references are not counted anew for every JSON serialization occurrence; these metrics are not serialized byte size or total infrastructure allocation. XML parsing/MCE/source indexing remain governed by their own budgets. Paths are charged before construction.

Overrides must be positive safe integers; unknown/accessor limit properties reject. Bytes and option values are captured before asynchronous work. Exhaustion throws `PackageInspectionError` with `package-limit-exceeded`; cancellation uses `operation-aborted`. No partial/truncated success is returned. Independent retries do not share mutable state. The result is deeply frozen plain data and can be JSON serialized; serialization does not itself verify provenance.

## Admission and evidence

`nativePlayback` and `authoredMapping` are always `not-performed`. The API does not build a Core event program, authorize source timing edits, produce animated SVG or implement playback. Unrelated admitted text saves preserve source timing; changing authored timing remains blocked by the existing preservation-save boundary.

[Parser tests](../../packages/pptx-parser/test/timing.test.mjs) cover eight focused groups, including duplicates/aliases, raw-only scope, part isolation, exact budgets and three text-save rounds. [Browser parity tests](../../scripts/timing-inventory-parity.test.mjs) compare Node/ESM/global with an independent Chromium DOM oracle and test cancellation, invocation capture and concurrent retry. Installed runtime/types consumers exercise public exports.

The optional [independent runner](../../scripts/inspect-native-timing-inventory.mjs) accepts source and previously resaved PPTX paths. Python stdlib ZIP/ElementTree checks raw indices, expanded paths, decoded fields, ancestry and candidate identities. It starts no native application. The [recorded manifest](../testing/native-timing-inventory-evidence.json) checks two four-slide files (129 timing nodes each). Original native application/version and artifact provenance are recorded separately in [animation evidence](../testing/animation-mapping-evidence.json); the inventory runner does not certify that provenance or actual playback. PowerPoint interactive playback remains pending.


The separate [paragraph target save profile](paragraph-target-save.md) now admits proven structural index preservation during existing scalar text edits. The read-only inventory itself still grants no write authority or authored mapping.
