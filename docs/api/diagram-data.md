# SmartArt source data graphs

`inspectPptx(bytes, options)` now includes optional `diagramInventory.dataModels`. Both import projections expose it through `session.report.inspection`. This experimental source profile reads the located data parts discovered by the [component inventory](smartart-inventory.md), once per distinct part. It does not add Core diagram authoring or permit SmartArt text/node edits.

<!-- doc-test: docs/examples/diagram-data.ts -->
```ts
import { inspectPptx, type PptxDiagramDataModelInfo } from "@pptkit/pptx-parser";

export async function readDiagramModels(bytes: Uint8Array): Promise<readonly PptxDiagramDataModelInfo[]> {
  const inspection = await inspectPptx(bytes, {
    diagramLimits: { maxWorkUnits: 100_000, maxFindings: 1_000 },
  });
  return inspection.diagramInventory?.dataModels ?? [];
}
```

## Source identity and graph

Each `PptxDiagramDataModelInfo` has `part`, `coverage: "identities-references-direct-text"`, `status`, `points`, `connections` and located `findings`. Arrays retain direct XML order; they are not sorted by connection order or transformed into a tree. One-based raw XML preorder `elementIndex` is scoped to the model's data part, including for resolved reference targets and text. These indices are source locations, not Core IDs.

The recognized identity profile is signed int32 or uppercase braced GUID, with XML edge whitespace removed only for interpretation. Raw `modelId` and reference spellings remain unchanged. Numeric aliases such as `+01` and `1` share identity. Points and connections share one identity space: cross-kind collisions count as duplicates. Each entity reports `unique`, `duplicate`, `missing` or `unrecognized`. Lowercase GUIDs and other identifiers remain literal but unrecognized by this bounded profile.

Points recognize `node` (omitted type default), `asst`, `doc`, `pres`, `parTrans`, `sibTrans`; connections recognize `parOf` (omitted default), `presOf`, `presParOf`, `unknownRelationship`. Unknown kinds produce findings. Empty type attributes are not treated as omission. Known kinds describe source records, without establishing a family, valid hierarchy or layout input. Cycles, multiple parents and parallel edges are retained without repair or a DAG-validity claim.

| Source field | Inspection behavior |
| --- | --- |
| `cxn/@srcId`, `destId` | Required point references. Missing values have `absent` status and a finding. |
| `cxn/@parTransId`, `sibTransId` | Optional point references, without proving the target has the corresponding transition kind. |
| Transition `pt/@cxnId` | Optional connection reference for `parTrans`/`sibTrans`. On other point kinds the literal value has `uninterpreted` status. |
| `pt/prSet/@presAssocID` | Optional point reference; direct `prSet` attributes also remain namespace/local/value bags. |
| `cxn/@srcOrd`, `destOrd` | Raw spelling plus parsed unsigned-32-bit value when recognized; missing/invalid values produce findings. XML order remains unchanged. |
| `cxn/@presId` | Uninterpreted presentation/layout URI spelling; not resolved as a point ID. |

Interpreted references report `absent`, `unrecognized`, `missing`, `ambiguous`, `wrong-kind` or `resolved`. Only resolved references include `targetElementIndex`; no arbitrary duplicate is selected. No implicit relationships or unrecognized layout parameter semantics are synthesized. A resolved ID proves a unique entity of the requested broad kind only.

## Direct text and coverage

Point text records direct `dgm:t/a:p` paragraphs, including empty paragraphs and lexical list levels. Runs retain source indices, kind (`text`, `break`, `field`) and decoded stored text. Break text is `\n`; fields retain available ID/type and stored text without evaluating slide numbers, dates or formulas. Direct `a:t` must be one scalar leaf. Ambiguous/mixed text is not guessed and produces a finding. Non-XML-whitespace characters directly inside a text body, paragraph, run, field or break are diagnosed; the affected body, paragraph or run is omitted from the extracted text. XML indentation whitespace (space, tab, CR and LF) is ignored, while characters such as non-breaking space are not treated as indentation. Formatting, body/list styles, inherited/theme text, layout evaluation and visual text ownership are outside this coverage.

`inspected` means the bounded profile found no recorded problem; it is not schema, topology, native validity or completeness certification. `partial` retains available records with findings. A data part declaring/using MCE, missing or multiple direct point lists, or multiple direct connection lists is `source-only` with empty graph arrays; raw branches/lists are never merged. An absent connection list yields an empty connection array. Unknown list/point/text children have located findings where visited; the profile does not enumerate every unhandled attribute or recursively interpret extensions.

Cache resource records additionally expose optional `modelReferences`, containing raw `dsp:sp` indices and lexical `modelId`, including inactive branches. Repeated cache references can be legitimate. They are not duplicate identities in the data graph. The [explicit mapping extension](diagram-mapping.md) adds normalized presentation evidence and recognized cache references per actual data/cache pair; it does not establish complete semantic ownership, freshness or permission for a cache rewrite. Existing [cache previews](diagram-preview.md) remain ordinary Core fallback shapes.

## Limits and save behavior

`PptxInspectionOptions` extends the existing OPC options with `diagramLimits`. `DEFAULT_DIAGRAM_LIMITS` is a frozen defaults record: `maxWorkUnits: 1_000_000`, `maxFindings: 10_000`. Overrides must name known keys with positive safe integer values. Work checkpoints and emitted findings accumulate over all distinct data parts per inspection; the explicit mapping extension also charges presentation/cache mapping work and findings. Shared data is read once; limits are not reset per part. Work units count graph-processing checkpoints rather than bytes or a public stable operation count. Existing ZIP/XML/OPC limits separately bound raw parsing, attributes, text and source-index construction.

Exceeding either graph limit throws `PackageInspectionError` with `package-limit-exceeded` and the current part. No truncated successful graph is returned. Import translates this to `import-limit-exceeded`; actual candidate validation uses save-stage error semantics. Cancellation remains supported. Configuration is captured before the first await, including Node file I/O, and retained for actual-output reprojection during supported unrelated text saves.

Returned records are readonly detached data; runtime deep freezing is not promised. Source bytes are unchanged. Synthetic both-dialect no-op and three unrelated text-save rounds retain graph components and their payloads. SmartArt text has no source-save binding; no semantic graph edits, native SmartArt generation or full preservation/editability guarantee follows.

## Evidence and remaining work

[Unit fixtures](../../packages/pptx-parser/test/diagram-data.test.mjs) cover identity aliases/collisions, typed references, direct text, cycles, ambiguity, cumulative budgets, no-op and cancellation. [Independent browser DOM](../../scripts/smartart-inventory-parity.test.mjs) checks source indices/edges against Node/browser reports and preserves graph payloads over three text saves. [Node file integration](../../scripts/file-save-parity.test.mjs) proves a custom diagnostic budget above the default survives file option capture and candidate reprojection. Installed Node/browser/type consumers use public exports.

[Four pinned source cases](../testing/smartart-data-evidence.json) independently verify 128 points, 129 connections and 399 interpreted references across Process, organization-chart, Cycle and PictureStrips layouts. [Three further source cases](../testing/smartart-data-text-evidence.json) cover nonempty caches and existing independent-text probes. The offline [runner](../../scripts/inspect-smartart-inventory.mjs) uses Python ZIP/ElementTree to check identities, reference targets, orders, property bags, text and raw cache model references; exact no-op output is also verified. Third-party inputs stay local pending redistribution/fixture licence admission. No native application was launched by these experiments.

Full data-to-visual mapping, model/family decisions, cache consistency/layout and actual native node editing remain under ADV-02/03. See [ADR0020](../architecture/decisions/0020-smartart-source-data-graph.md) and the [execution cards](../planning/advanced-capability-todos.md#接续执行卡数据图图表保存与时间线).
