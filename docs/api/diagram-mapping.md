# Explicit SmartArt data and cache mappings

The experimental [source graph profile](diagram-data.md) now exposes two normalized mapping layers through `inspectPptx` and `session.report.inspection`. `dataModels[i].presentationMappings` records recognized point/connection evidence once per data part. `diagramInventory.cacheMappings` records cache-shape references once per explicitly associated `(dataPart, cachePart)` pair. Both use part-local raw XML indices; neither creates Core node identities or source-edit bindings.

<!-- doc-test: docs/examples/diagram-mapping.ts -->
```ts
import { inspectPptx, type PptxDiagramPresentationMapping } from "@pptkit/pptx-parser";

export async function traceFirstCachedDiagram(bytes: Uint8Array): Promise<readonly {
  cacheElementIndex: number;
  dataElementIndex: number | undefined;
  presentation: PptxDiagramPresentationMapping | undefined;
}[]> {
  const inventory = (await inspectPptx(bytes)).diagramInventory;
  const cache = inventory?.cacheMappings?.[0];
  if (!cache) return [];
  const model = inventory?.dataModels?.find(item => item.part === cache.dataPart);
  const presentations = new Map(model?.presentationMappings?.map(item => [item.elementIndex, item]));
  return cache.shapes.map(shape => ({
    cacheElementIndex: shape.elementIndex,
    dataElementIndex: shape.model.targetElementIndex,
    presentation: shape.model.targetElementIndex === undefined ? undefined : presentations.get(shape.model.targetElementIndex),
  }));
}
```

## Presentation evidence

Each recognized `kind: "pres"` point has a mapping with its raw data-part `elementIndex`, ordered `sources` and `agreement`. Sources contain a typed `reference`, evidence kind and source index. `presAssocID` evidence uses the **point** index and its inspected direct `prSet/@presAssocID`. `presOf` evidence uses the **connection** index and its source reference, only when the connection destination resolves uniquely to that presentation point. Property evidence comes first; connections retain XML order. Duplicate evidence is retained. Nonpresentation/unresolved destinations remain in the graph without an invented owner.

| Agreement | Meaning |
| --- | --- |
| `none` | No recognized explicit evidence. |
| `unresolved` | At least one recorded source reference does not resolve. |
| `different-targets` | Recorded references resolve to more than one point; no winner is selected. |
| `same-target` | All recorded references resolve to the same point, named by `targetElementIndex`. |

Agreement describes recorded evidence, not complete semantic ownership or validity. Two `different-targets` points occur in the pinned process3 source. Referenced points may be document, assistant or transition points as well as content nodes. The inspector does not choose a text owner, follow chains, merge text, evaluate layout `presOf` expressions or assign styling. Unknown/ambiguous structures retain the underlying graph findings and coverage.

## Cache pairs and references

Only located explicit cache associations contribute pairs. Shared owners reuse one pair; one data part with two slide-local caches yields two pairs. Distinct data parts using the same cache resolve separately against their own identity spaces. Unassigned candidates never acquire data owners. There is no Cartesian product of resource parts; direct owners remain in the existing inventory.

`PptxDiagramCacheMapping` contains `dataPart`, `cachePart`, `coverage: "explicit-model-references"`, `status`, `shapes` and `findings`. Shapes follow source order through one direct `dsp:drawing/dsp:spTree` and direct nested `dsp:grpSp` children. Each shape has a raw **cache-part** index and `model` reference; its resolved target index is scoped to **dataPart**. The graph's int32/GUID rules apply across point and connection identities. Only unique points resolve; `absent`, `unrecognized`, `missing`, `ambiguous` and `wrong-kind` produce cache findings. Repeated shape references to one point remain separate shapes and do not create data identity collisions.

`inspected` means this reference pass produced no cache finding, not full data validity. `partial` retains references and findings. An explicit empty tree is inspected with zero shapes, without proving preview availability. MCE caches, missing/duplicate trees or source-only data produce `source-only` maps with a reason and no inferred shapes. Extension-wrapped shapes are excluded from recognized paths while remaining in raw `parts[].modelReferences`; raw and recognized counts can differ.

## Limits and evidence

`diagramLimits.maxWorkUnits` also charges model indexing, presentation evidence, recognized cache traversal, diagram owners and pair resolution. Cache trees are scanned once per cache; models are indexed once per data part. Mapping findings share `maxFindings` with graph findings. Limits fail explicitly with part provenance; no successful truncated maps are returned. Existing option capture through byte/file import and actual candidate validation, XML/OPC limits and cancellation remain in force. Work-unit thresholds are implementation details.

Three unrelated text-save rounds retain exact graph/component payloads and mappings. No new diagram write bindings are added. Cache presence, source-text agreement or successful references do not prove freshness, native validity, visual equality or editability. Semantic model/family selection, cache regeneration and native node edits remain ADV-02/03 work.

[Eight focused groups](../../packages/pptx-parser/test/diagram-mapping.test.mjs) cover both dialects, one-to-many and differing/unresolved evidence, identity collisions, tree scope, MCE, multi-page/multi-model sharing, unassigned candidates, budgets and no-op. [Independent DOM integration](../../scripts/smartart-inventory-parity.test.mjs) derives paths from XML and checks Node/browser results and three saves. Installed Node/browser ESM/global and readonly declarations exercise public distribution.

Python ZIP/ElementTree independently checks **103 presentation mappings**, including two `different-targets` points, across [four original files](../testing/smartart-mapping-evidence.json) and [three nonempty-cache files](../testing/smartart-mapping-text-evidence.json). The latter resolve **10 cached shapes**; original controls include three empty caches and one missing cache. Exact no-op, source graph and seven text-frame probes remain checked. Third-party fixtures stay local pending licensing admission; no native application result is inferred. See [ADR0021](../architecture/decisions/0021-smartart-explicit-model-mapping.md).

The opt-in [direct text comparison](diagram-text-comparison.md) adds bounded paragraph-vector observations over these explicit associations. It retains unknown freshness and does not turn association agreement into semantic ownership.
