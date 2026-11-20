# Direct SmartArt text observations

Experimental `diagramTextComparison: "direct-paragraphs"` adds readonly `diagramInventory.textComparisons` to `inspectPptx` and `session.report.inspection`. The default omits the field. It is independent of cached preview and works through byte/file imports and source-save candidate validation. The literal profile is required; booleans, null and unknown values reject.

<!-- doc-test: docs/examples/diagram-text-comparison.ts -->
```ts
import { inspectPptx, type PptxDiagramTextComparison } from "@pptkit/pptx-parser";

export async function compareDirectDiagramText(bytes: Uint8Array): Promise<readonly PptxDiagramTextComparison[]> {
  const report = await inspectPptx(bytes, { diagramTextComparison: "direct-paragraphs" });
  return report.diagramInventory?.textComparisons ?? [];
}
```

Each explicitly associated `(dataPart, cachePart)` pair has `coverage: "direct-paragraph-vectors"`, `freshness: "unknown"`, a pair status and shape observations. Shared owners reuse a pair; distinct caches remain separate. Indices are raw XML preorder identities: `cacheElementIndex` belongs to the cache part; optional `presentationElementIndex` and `sourcePointElementIndex` belong to the data part.

A cache shape must explicitly name a unique content/assistant point, either directly or through a presentation point whose recognized association evidence agrees on one target. No visible-label, array-index or nearest-geometry matching occurs. Incomplete/data-graph findings or a source-only cache mapping produce pair status `source-only`, a reason and no guessed shapes. Other observed pairs retain per-shape results:

| Shape status | Observation |
| --- | --- |
| `equal` | The exact source/cache paragraph vectors agree. |
| `different` | Both vectors are available and differ. Neither side is selected as authoritative. |
| `unavailable` | Associations or direct text bodies cannot be interpreted within this profile; the reason is retained. |

Available observations expose `sourceParagraphs`, `cacheParagraphs` and `nonemptyText`. Paragraph boundaries, soft breaks (`\n`), empty paragraphs, XML-decoded characters, Unicode and whitespace are exact. There is no Unicode or whitespace normalization. `nonemptyText` is true if either vector has non-whitespace characters; false is vacuous/whitespace-only evidence, not proof of visual emptiness. Neither value proves visibility.

Coverage requires exactly one direct DiagramML `t` / DSP `txBody`, direct DrawingML paragraphs, scalar text runs and soft breaks. Absent bodies are unavailable rather than inferred empty. Fields, unknown body/paragraph/run structures, duplicate recognized property containers and direct opaque characters are unavailable. Only XML whitespace (space, tab, CR and LF) counts as indentation outside text leaves; non-breaking spaces and other Unicode whitespace are opaque characters. Run formatting, computed fields, full rich-text semantics, inherited/duplicated layout text, fonts, geometry, paint and rendering equality are outside coverage. Property contents are not compared as text.

Every result retains unknown freshness. Equal characters cannot prove the current theme/layout/font/cache dependency version; different characters can arise from layout-specific text roles and do not establish universal staleness. The [controlled native corpus](../testing/smartart-stale-cache.md) demonstrates actual data/cache disagreement without admitting a general stale-cache detector. Cached preview continues to display stored cache characters and reports unknown freshness. This option neither blocks that explicitly requested preview nor creates native SmartArt editing or source-write authority.

The existing aggregate `diagramLimits.maxWorkUnits` charges raw indexing, lookup/traversal, scalar characters and every pair comparison, including repeated long-text use. `maxFindings` also counts source-only pairs and unavailable/different shapes. Budget exhaustion rejects with part provenance; no truncated successful text report is returned. Existing XML/archive budgets and cancellation apply. Work-unit thresholds are implementation details. Options are captured before the first await and reused during candidate validation.

Evidence: [seven focused groups](../../packages/pptx-parser/test/diagram-text-comparison.test.mjs), [independent DOM/Node/ESM/global and three-round preservation checks](../../scripts/diagram-text-parity.test.mjs), installed Node/declaration consumers and [ADR0030](../architecture/decisions/0030-diagram-text-observations.md). Data/cache payloads survive unrelated supported text edits byte-for-byte; cached text still has no save binding.
