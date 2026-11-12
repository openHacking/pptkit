# Versioned import capability observations

`describeImportCapabilities(report, options?)` converts the public `ImportReport`'s projection/findings into a bounded, deeply readonly seven-operation view. It is synchronous and pure: no package/file/native inspection, source or document mutation, session creation, candidate save validation, network access or runtime test execution occurs. The result is descriptive evidence, not an edit/save authorization token.

<!-- doc-test: docs/examples/import-capabilities.ts -->
```ts
import { describeImportCapabilities, importPptx } from "@pptkit/pptx-parser";

export async function inspectCapabilitySubjects(bytes: Uint8Array) {
  const session = await importPptx(bytes, { projection: "content", diagramPreview: "cache" });
  try {
    return describeImportCapabilities(session.report, {
      filter: { feature: "smartart.cache-preview" },
      limits: { maxSubjects: 100 },
    });
  } finally {
    session.dispose();
  }
}
```

The output has `schemaVersion: "pptkit-import-capabilities-v1"`, `catalogVersion: "pptkit-capability-references-v1"`, `coverage: "reported-source-findings"`, projection, exact filter, counts, operation totals and subjects. `nativeVerification` and `candidateSaveValidation` are always `"not-performed"`. The function does not inventory unreported source objects or establish that the finding producer covered every property. In particular, existing text-run binding eligibility and actual `prepareSave` validation remain separate from generic object-feature findings.

## Exact subjects and complete operations

A subject is one exact tuple of feature, source part/path/object ID/relationship ID/raw element index, Core slide ID and Core element ID. Absent and explicitly empty strings remain distinct; every identity field participates. Locations are retained as supplied and never evaluated as XPath or used for array/label/geometry matching. A source-part preservation finding does not transfer preservation or editability to a different object/property subject.

Subjects sort by the JSON encoding of that fixed tuple using locale-independent UTF-16 code-unit order. Within each operation, observations sort by finding ID, tier and reason using the same ordering. Original finding IDs/reasons/tiers and duplicate observations remain available. Subject `id` is its one-based ordinal in the complete sorted matching set; it is stable for equivalent subject sets but is not a durable cross-document identity or a source binding. JSON object-key order is immaterial.

Every returned subject contains seven rows in this order: `create`, `import`, `edit`, `preview`, `export`, `no-op-save`, `edit-save`.

| Status | Meaning |
| --- | --- |
| `reported` | At least one finding exists and every tier agrees; `tier` is that observed tier. |
| `unreported` | No finding exists for this exact subject/operation; observations are empty and no tier is supplied. This never means supported. |
| `conflicting` | Findings disagree on tier. All observations remain, and no winning tier is selected. |

The original tiers (`editable`, `preserved`, `fallback`, `unsupported`) remain scoped to their original feature/operation/projection. Reasons may describe a bounded mapped subset; a reported editable import never automatically becomes an editable source-save. Conflicting tiers remain conflicts even if one says editable, and equal tiers do not collapse distinct reasons. Do not combine subjects by label or merge generic part findings with property findings.

## Reference index and compatibility

Known feature/category rules attach readonly `references` containing the exact rule, `match: "exact" | "prefix"`, related task IDs and repository-relative artifact paths classified as contract, test or native-observation. Prefix rules are navigation hints for a category, not an exhaustive supported-feature registry. Unknown feature names remain observable and have no fabricated references. Reference presence does not change tiers or establish that an artifact/test applies to, or was executed for, the current input. Even SmartArt subjects linking native observations keep nativeVerification not-performed.

The reference index is maintained with Parser and has its own version. It currently links the reported structural/metadata/section/transition/source/drawing/chart/table/notes/SmartArt categories, as well as first-slide-number, master artwork and background observations. Tests verify referenced files and task identifiers exist and exercise the latter observations from actual source fixtures in both dialects. Broader authored capabilities, exact per-property/native application certification and complete task/evidence coverage remain pending.

Recognized input fields must be own data properties on plain records; known accessors, sparse/accessor finding arrays, malformed projection/operation/tier enums and invalid raw indices reject. Additional future fields on reports/findings/sources/options/filters are ignored without reading their values or executing getters. Unknown resource-limit names reject to avoid silently ignored quotas. Unknown feature strings are allowed; unknown operation/tier/projection values require a compatible newer API rather than guessed mappings. Input reports may be caller-created data; the function does not authenticate their claims and cannot manufacture source-save authority. Consumers of the JSON schema should ignore additive fields and reject incompatible schema versions. Incompatible report/reference meanings require a new version.

## Filters, truncation and resource bounds

`filter` supports exact feature, logical part, slideId and elementId string matches, combined with AND. No regex, prefix query or path evaluation occurs. Validation and finding/text/work bounds still cover the complete input, including nonmatching findings; filtering cannot hide malformed input or bypass resource accounting.

| Limit | Default | Meaning |
| --- | --- | --- |
| `maxFindings` | 150,000 | Maximum input findings before grouping |
| `maxTextCodeUnits` | 16,777,216 | Cumulative free-form finding/source/Core string UTF-16 units, including repeats and filter values; enum constants are validated separately |
| `maxWorkUnits` | 33,554,432 | Traversal, strings, identity keys, comparison/sorting and operation aggregation |
| `maxSubjects` | 1,000 | Complete matching subjects retained; zero requests counts only |

The first three limits require positive safe integers; maxSubjects permits zero. `DEFAULT_IMPORT_CAPABILITY_LIMITS` is readonly. Hard quota exhaustion throws `PptxImportError` with `import-limit-exceeded`, without partial output or input mutation; invalid values/types throw RangeError/TypeError. Sorting charges potential shared-prefix comparison length, preventing long identities/reasons from becoming unaccounted repeated work. These checkpoints are an implementation budget, not a portable performance metric. `workUnits` may differ for differently ordered equivalent inputs or between JavaScript runtimes with different sorting strategies while subject/observation order remains canonical.

Only maxSubjects truncates successfully. The result retains the complete `totalFindings`, `matchedFindings`, `totalSubjects` and per-operation status totals over **all matching subjects**, alongside `returnedSubjects`, `returnedFindings` and explicit `truncated`. A returned subject always includes all seven operation rows and all its observations. Omitted subjects must not be interpreted as absent source content or unsupported/supported operations. A filter with no matches returns zero subjects and truncated false, without claiming complete source coverage. The helper retains no handles/caches and has no asynchronous cancellation hook; its synchronous work is bounded by the supplied quotas.

## CLI view

```sh
pptkit inspect deck.pptx --projection content --capabilities --json
pptkit inspect deck.pptx --projection content --diagram-preview cache --capabilities --capability-feature smartart.cache-preview --json
pptkit inspect deck.pptx --capabilities --capability-part /ppt/slides/slide1.xml --max-capability-subjects 0 --json
```

The opt-in CLI envelope adds `reportKind: "capabilities"` and places this versioned view in `report`, replacing the ordinary ImportReport **only for that explicit flag**. Default inspect output remains compatible. Capability filters/subject limit require `--capabilities` and reject on compare. The CLI accepts subject limits from 0 through 150,000 and retains its existing output byte budget; a too-large complete view returns a valid output-limit error. The human view prints all seven statuses and task links. Neither a successful counts-only report nor a reference index means native or candidate-save validation passed. See [CLI contract](cli-inspection.md).

Evidence: [eight focused groups](../../packages/pptx-parser/test/capabilities.test.mjs), [Node/browser ESM/global integration](../../scripts/capability-report-parity.test.mjs), [CLI tests](../../packages/cli/test/commands.test.mjs), installed Node/shim/declaration consumers and [ADR0033](../architecture/decisions/0033-import-capability-observations.md). Tests retain every source finding through an independent provenance reconstruction, cover 40 seeded shuffled passes against an operation oracle, exact work/text boundaries, conflicts/duplicates, future fields/getters, complete-subject truncation and unchanged no-op saves. Full IMP-04.4/ECO-09 acceptance remains broader than this observation transport.

[Source table-style projection](table-style-import.md) now reports bounded inline/reference/default region choices and per-cell fill/border/text sources. Conditional-region preview remains fallback; resolved style identity does not establish full native appearance or edit-save authority.
