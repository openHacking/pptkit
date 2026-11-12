# Placeholder matching and dimension projection

Experimental `importPptx(bytes, { projection: "content" })` now resolves a bounded placeholder ancestry and dimension profile. It reuses the verified slide → layout → master relationships and retained effective XML trees. It does not add source XML, source identities, or preservation handles to Core.

```ts
import { importPptx } from "@pptkit/pptx-parser";

const session = await importPptx(bytes, {
  projection: "content",
  sessionLimits: { maxPlaceholderEntries: 100_000 },
});
try {
  for (const ph of session.report.placeholderInheritance?.placeholders ?? []) {
    console.log(ph.source, ph.parent?.status, ph.geometryStatus);
    console.log(ph.dimensions.width?.valueEmu);
    console.log(ph.dimensions.width?.source);
  }
} finally {
  session.dispose();
}
```

The complete checked example is [placeholder-inheritance.ts](../examples/placeholder-inheritance.ts).

## Matching profile

| Step | Rule |
| --- | --- |
| Slide → layout | Numeric `p:ph/@idx`, including schema default `0`. Type, drawing ID, object name and drawing order do not choose the parent. |
| Layout → master | Placeholder type: `ctrTitle` → `title`; `body`, `subTitle`, `obj`, `chart`, `tbl`, `dgm`, `media`, `clipArt`, `pic` → `body`; `title`, `dt`, `ftr`, `sldNum` match their own type. The omitted source type is `obj`. |
| Missing relationship/match | `missing`, with no invented parent. |
| Multiple candidates | `ambiguous`, with count and matching key; no arbitrary first candidate. |
| Grouped placeholders or unsupported master rule | `unsupported`, with a reason. `hdr` and `sldImg` are recognized source types but have no master matching rule in this profile. |

Matches are scoped to the owning source parts. Two masters with the same indices/types remain distinct. Unrelated masters/themes are never searched as fallbacks. Ordinary shapes do not acquire placeholder semantics from coincident IDs or names. The report inventories each declaration once in visited slide/layout/master parts; it is not an inventory of every unused template part.

## Dimensions and geometry

Offset (`x`, `y`) and extent (`width`, `height`) inherit separately, with local values taking precedence over the matched layout and master. Each resolved dimension stores original EMUs, source part/path/object identity and source level. An existing `off` or `ext` must contain valid required scalar attributes. Zero and negative offsets are values, not missing-data sentinels.

The content projection converts resolved EMUs to points and stores local Core boxes. Existing complete local transforms retain their direct behavior. A partial local transform may use inherited dimensions while retaining its own explicit rotation/flips. An absent local transform uses neutral rotation/flips only when ancestor transforms are neutral. Inherited non-neutral rotation/reflection is reported as unsupported until the transform cascade has native evidence. Grouped placeholder inheritance remains unsupported; existing complete local group geometry continues through the direct content reader.

For shapes without a local geometry choice, the nearest matched ancestor may supply one of the eight supported preset names. An explicit unsupported preset or custom geometry blocks that fallback. Inherited fills, effects and text, and adjustments outside the parallelogram subset, are not mapped by the preset-name projection and receive a fallback finding. A matched ancestor may supply one bounded literal parallelogram slant; `drawing.placeholder-slant` records the source guide and target. The nearest explicit geometry wins, unsupported adjustments block projection, and no source-edit permission is implied. A [PowerPoint observation](../testing/inherited-slant-powerpoint-evidence.json) covers one layout value and one master value before and after one save/reopen; other placeholder combinations remain unverified. A [cross-package test](../../scripts/parallelogram-parity.test.mjs) edits a bound shape-text run on a modified version of that source through three programmatic preservation rounds: the inherited slant survives and every non-slide package part remains byte-identical. This does not permit changing the slant in a preservation save. Source part-level findings still apply.

`geometryStatus: "resolved"` means the bounded dimensions can be projected; it does not assert that the full drawing, style or text is supported. An unsupported preset can still keep the object source-only despite resolved dimensions. Successful objects receive normal Core identities; omitted drawings receive no fabricated Core identity.

## Report and resource contracts

Content reports have `placeholderInheritance.placeholders`, whose entries include declaration location, effective type/index, whether each attribute was explicit, grouped state, parent match, dimension provenance and geometry status/reason. These are detached, deeply frozen snapshots. Structure reports have no `placeholderInheritance` property.

The optional `maxPlaceholderEntries` session limit defaults to **100,000** and counts distinct declarations across visited source parts. A shared layout/master counts once, regardless of page count. Ambiguous candidate lists are represented by count/key rather than expanded repeatedly. Shared parent resolution is cached once per part. Traversal uses cooperative cancellation and the existing ZIP/XML/OPC budgets; invalid numbers, booleans, singleton declarations and duplicate drawing identities fail content import before publishing a session. Capability rows remain separately bounded by `maxReportFindings`.

## Operation boundaries

The matched dimensions become editable local Core values, with the format-specific source provenance retained in the Parser report. Master/layout prompt text and artwork are not copied into slide contents. This prevents inherited prompt duplication, while leaving unimplemented inherited artwork explicitly source-only.

Fresh export produces a new PPTKit document from those local values. It does not recreate the source template chain. A bounded frame, paragraph, list and run style profile is also available; see [Text style inheritance](text-style-inheritance.md). Complete theme fonts/colors, backgrounds, non-neutral inherited transforms, dynamic Core multi-master authoring and template rebinding remain pending under AUT-07 and IMP-05.1.2–05.1.4. Default edit and all rebuild requests remain blocked. The opt-in [text-run profile](text-edit-save.md) permits only existing bound run characters and retains unchanged source inheritance; it does not rebind or write geometry/style overrides. Original-byte no-op save retains the complete unchanged source.

## Verification

- [Independent source fixtures and tests](../../packages/pptx-parser/test/placeholder-inheritance.test.mjs): both dialects, aliases/MCE, default/explicit indices, type families, multiple masters, ambiguity, local zero/false overrides, grouped and rotated ancestors, malformed input, budgets/cancellation, shared parents and session isolation.
- [Cross-package tests](../../scripts/placeholder-import-parity.test.mjs): three fresh rounds, independent output slide XML, stable dimensions/resource size, no repeated master text, local edits and unaffected pages.
- [Browser consumer](../../scripts/browser-bundles.test.mjs) compares Node/Chromium reports and bytes without network requests; [installed consumers](../../tests/consumers/node.mjs) exercise packed public APIs, with [declaration checks](../../tests/consumers/types.ts).

These tests do not establish native PowerPoint/LibreOffice layout fidelity or interactive template editing. Those results remain **untested** and the corresponding parent/child TODOs remain unchecked. Matching-policy rationale and source implementation references are recorded in [ADR 0003](../architecture/decisions/0003-placeholder-dimension-projection.md).
