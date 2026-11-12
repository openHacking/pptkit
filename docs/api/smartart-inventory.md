# SmartArt source component inventory

`inspectPptx(bytes, options)` optionally adds `slides[i].diagrams` when the effective slide shape tree contains diagram graphic frames. Both import projections expose the same inspection in `session.report.inspection`. This experimental profile identifies source objects and their four component associations. It does not create Core diagram nodes or grant editing permission.

<!-- doc-test: docs/examples/smartart-inventory.ts -->
```ts
import { inspectPptx, type PptxDiagramInfo } from "@pptkit/pptx-parser";

export async function listDiagrams(bytes: Uint8Array): Promise<readonly { part: string; diagram: PptxDiagramInfo }[]> {
  const inspection = await inspectPptx(bytes);
  return inspection.slides.flatMap(slide => (slide.diagrams ?? []).map(diagram => ({ part: slide.part, diagram })));
}
```

Objects follow effective shape-tree order, recursively through `p:grpSp`. Discovery requires a direct `p:graphicFrame/a:graphic/a:graphicData` whose URI is the diagram namespace for the entry dialect. Only direct `dgm:relIds` participates. Arbitrary descendants, extension payloads, master artwork and inactive MCE branches are outside this discovery profile; absence of `diagrams` does not prove the package contains no diagram data.

Each `PptxDiagramInfo` contains a one-based raw XML preorder `elementIndex`, outermost-first `groupElementIndices`, optional lexical `drawingId`/`name`, `identityStatus`, and four `components`. Use `(slide.part, elementIndex)` to locate source objects. Drawing IDs are not Core IDs; missing, invalid or repeated unsigned-32-bit numeric identities are reported separately. Ambiguity includes ordinary supported shape-tree object kinds and the tree root; `02` and `+2` refer to the same numeric identity. Identity status does not validate references from unknown content.

Components appear in data/layout/colors/style order and use `r:dm`, `r:lo`, `r:cs`, `r:qs` respectively. The style relationship is `diagramQuickStyle`, with a `drawingml.diagramStyle` content type. Associations are checked by relationship, content type and the raw dialect-specific component root, never by a filename convention.

| Status | Meaning |
| --- | --- |
| `missing-reference` | No recognized role attribute, including absent `dgm:relIds`. |
| `missing-relationship` | The referenced ID is not declared on the slide. |
| `wrong-relationship-type` | The declared relationship has another role. |
| `external` | Expected role, external target; no resource is fetched. |
| `fragment-target` | Expected role, internal target with a fragment. |
| `wrong-content-type` | Target content type is not the role's component type. |
| `wrong-root` | Target raw root name/namespace does not match the component. |
| `located` | Role, whole-part target, content type and raw root match. No schema/data/layout correctness follows. |

Where available, records include relationship ID, raw target spelling, logical target part and content type. `layoutIdentifier` is the uninterpreted `layoutDef/@uniqueId` of a located layout; unknown/custom values remain literal and do not infer a supported family. Shared components retain the same logical part across object records; the direct resource evidence below remains separate from complete transitive ownership proof.

Status precedence follows the table order after reference lookup: wrong role is reported before mode, fragment, type and root. Global ZIP/XML/OPC invalidity still fails the inspection before a diagram report is returned. In particular, a declared internal relationship to an absent part is a `package-invalid` error, not `missing-relationship`. Duplicate structural children such as two direct `relIds` reject with `presentation-invalid` and a source part. Existing inspection resource limits and cancellation apply to traversal and raw source indexing; no network or media execution is introduced.

Returned types are readonly detached data, without XML trees or preservation handles. Runtime deep freezing is not promised. Reading leaves input bytes intact. Exact no-op and three unrelated existing-text save rounds are tested for the named synthetic fixtures; this is not blanket diagram preservation or a native SmartArt compatibility result.

Each discovered diagram now also has a `cache` association report. Optional package-level `diagramInventory` inventories cache candidates and directly associated resources. The following extension remains an inspection profile; `located` must never be displayed as “editable SmartArt.”

## Drawing-cache association

The recognized path is raw `dgm:dataModel/dgm:extLst/a:ext/dsp:dataModelExt`, with extension URI and `dsp` namespace both `http://schemas.microsoft.com/office/drawing/2008/diagram`. Only the block's unqualified `relId` and `minVer` attributes participate. `relId` resolves against the **owning slide's** relationship table. It does not resolve against the data part, and sharing one data part across slides does not share the relationship-ID scope.

An association is `located` only when the slide relationship has the Microsoft `diagramDrawing` role, an internal whole-part target, the `application/vnd.ms-office.drawingml.diagramDrawing+xml` content type, and a raw `dsp:drawing` root. The report includes `dataPart`, raw extension-block `references` with one-based XML `elementIndex`, and available relationship/target/type information. `minimumVersion` is an uninterpreted URI; the inspector does not decide whether a particular application can lay out the diagram.

The component failure statuses above also apply. Additional statuses are `unresolved-data` when the data component was not located, `ambiguous-reference` for multiple recognized blocks or multiple raw extension lists, and `source-only` when the data part declares/uses MCE. This deliberately conservative MCE policy avoids choosing or borrowing an inactive extension reference. `missing-reference` means no recognized direct reference, including a block lacking the unqualified attribute; it does not prove no producer-specific reference exists. A namespaced `r:relId` is not substituted for `relId`. Arbitrary nested extension descendants are not scanned for a cache owner.

Every association has `freshness: "unknown"`, including located caches. No data/cache consistency, native validity or freshness is inferred from part presence. Inspection itself does not render or convert the cache. The separate explicit [preview option](diagram-preview.md) can project a bounded cached drawing into ordinary Core shapes, without semantic nodes.

## Candidates and shared resources

`diagramInventory.coverage` is `direct-source-associations`. Its `cacheCandidates` list enumerates slide-level relationships with the Microsoft cache role, even on slides without discovered diagrams. Each candidate reports target validation and `referencedBy`, the raw diagram element indices whose located associations use that exact relationship ID. Empty means no recognized assignment, not an orphan or permission to delete. Several candidate IDs may point to the same resource part.

`parts` deduplicates located component, cache and direct image targets by logical part name. Each record has content type, roles, recognized direct diagram `owners`, `incomingRelationships` and `imageRelationships`. Owners are `(slidePart, elementIndex, role)` tuples. Shared images are reached through the component/cache's image edges; image owners are not multiplied into one copy per diagram. Incoming edges include unrelated source parts in the effective OPC inventory, so the report can reveal outside users. Inactive relationship branches and arbitrary non-relationship source references still require the separate preservation-closure profile.

Image records cover declared image relationships and raw dialect-specific `a:blip/@r:embed` or `@r:link` references in the inventoried component/cache XML. Each reference has its raw element index and attribute. Missing declarations, wrong relationship roles, external/fragment targets and non-image content types are distinct statuses. Image `located` checks only the relationship and `image/*` content type, without fetching, decoding, signature validation or rendering. Declared but unreferenced images have an empty reference list. Opaque/inactive XML branches are included as raw evidence; these are not effective preview-usage claims. Image-only targets are inventoried without recursively scanning their payloads.

The report's work is proportional to the effective relationship graph, direct owner edges and scanned raw component/cache XML, plus output size. Shared part scans and cache-reference discovery are memoized; existing ZIP/XML/OPC budgets and cancellation remain in force. It is not a complete transitive ownership/closure proof and cannot authorize cleanup, structural saves or diagram editing.

The separate [cached preview profile](diagram-preview.md) covers a bounded explicit drawing projection. A [source graph profile](diagram-data.md) additionally reads direct point/connection identities, references and stored text, with raw cache model references and cumulative graph budgets. Cache freshness, full preview fidelity, complete image/resource closure, inherited/master diagrams, semantic authoring graphs, native editing and family layout remain scheduled under ADV-01–03.

Evidence: [object/component tests](../../packages/pptx-parser/test/smartart.test.mjs), [cache/resource tests](../../packages/pptx-parser/test/diagram-resources.test.mjs), [independent DOM/browser and save integration](../../scripts/smartart-inventory-parity.test.mjs), [ADR0016](../architecture/decisions/0016-smartart-component-inventory.md) and [ADR0017](../architecture/decisions/0017-smartart-cache-resource-associations.md). Portable test fixtures are repository-authored structural samples, not producer-certified native diagrams.

An additional [source experiment](../testing/smartart-inventory-evidence.json) pins four LibreOffice regression-corpus files by repository commit, Git blob and SHA-256. Independent Python ZIP/ElementTree checks verify slide order, source identities, component/cache links, incoming edges and raw image-reference inventories, alongside exact no-op saved files. The Process, organization-chart and PictureStrips cases have explicit caches; the cycle2 case has no recognized cache reference. PictureStrips inventories three image parts. This is source evidence, with no application launched and no native editing/preview claim. Third-party source files remain local experiment inputs pending redistribution/fixture licensing admission. After obtaining the pinned inputs and their provenance manifest, rerun `node scripts/inspect-smartart-inventory.mjs <corpus-directory>`; the ordinary test suite remains offline and independent of those files.

The optional source runner flag `--preview` now probes that projection too. Independent raw XML confirms that the three located source caches in this corpus have zero drawing children, while cycle2 has no reference; all four therefore receive unavailable placeholders. This is not evidence for nonempty native cache rendering or native visual parity.

A separate [nonempty-cache text experiment](../testing/smartart-text-transform-evidence.json) pins three further public files and verifies seven independent text rectangles/angles/strings. It extends source evidence only; unsupported visual presets/styles remain partial, and native application rendering/editing remains unverified.
