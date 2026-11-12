# Experimental OPC source comparison

`compareOpcPackages(before, after, options?)` is exported by the root and `/browser` entries of `@pptkit/pptx-parser`, including `PPTKit.compareOpcPackages` in the global bundle. Both inputs must be `Uint8Array`. It validates both OPC inputs independently and compares source parts and relationship identities without importing a Core document or using Exporter.

The result establishes equality under one named profile. It does not establish complete OOXML schema validity, native editability, visual parity, a preservation-write implementation, or semantic equality of fresh exports. In particular, fresh export can materialize an inherited theme color to RGB while dropping unsupported source structures; source comparison reports those differences. Effective-value comparison remains a separate IMP-08.1.1 requirement.

## Profiles and allowed differences

The default profile is **`opc-payload-v1`**. Empty ZIP directory entries, compression method, local/central item order, timestamps and ZIP comments do not participate. Every logical OPC part, including unreferenced parts, and the complete `[Content_Types].xml` stream participate. Payloads are compared byte for byte. Binary equality does not rely on CRC32, length or a noncryptographic digest; a regression fixture uses two distinct equal-length/equal-CRC32 values. Embedded workbooks and other nested packages remain binary payloads at this layer.

**`opc-xml-v1`** requires explicit selection. Parts declared as XML under the OPC profile are compared through a conservative XML 1.0 source tree. Other payloads retain exact byte comparison.

| XML difference | `opc-payload-v1` | `opc-xml-v1` |
| --- | --- | --- |
| Encoding/BOM, XML declaration presence/encoding spelling | Changed bytes differ | Allowed when decoded source structures agree |
| Attribute order, quote style, empty-element spelling | Changed bytes differ | Allowed |
| Character references versus decoded characters, contiguous CDATA/text boundaries | Changed bytes differ | Allowed |
| XML line-ending/attribute-value normalization | Changed bytes differ | Uses the XML reader's decoded values |
| Declaration `standalone` absent/yes/no | Changed bytes differ | Remains significant |
| Element/attribute prefix names and namespace declarations, including unused bindings | Changed bytes differ | Remain significant |
| Ordinary attribute values, numeric spellings and units | Changed bytes differ | Remain significant |
| Child order, text spaces, indentation and `xml:space` | Changed bytes differ | Remain significant |
| Comments and processing instructions, including before/after the root | Changed bytes differ | Remain significant |
| Unknown extensions and unselected MCE branches | Changed bytes differ | Remain significant |

This is not XML canonicalization or a schema-aware equivalence algorithm. Prefixes are deliberately retained because arbitrary attribute/text values can contain QNames; blindly removing namespace declarations or replacing prefixes could hide changed source meaning. No XML children are reordered, no default values are synthesized, and no unknown markup is filtered. XML version/DTD/encoding restrictions remain those of the [XML reader](xml-inspection.md).

Logical part names and content-type strings are compared exactly. ASCII-case-equivalent OPC lookup does not authorize a source rename. Relationship keys are exact `(source part, source-local ID)` pairs. Type, mode, raw target, resolved internal target and fragment are compared; external targets are never fetched. Equivalent relative target spellings, relationship renames and whole-part renames still differ. Allowing those requires a separately verified identity bijection, which this profile does not implement.

Both profiles also compare the original control-stream payload/tree. Changing an unused type declaration or otherwise harmless relationship ordering can therefore produce a conservative source difference even when the effective graph agrees.

## Result

`OpcPackageComparison` is a frozen detached record containing:

- `profile` and `equal`.
- `before` / `after`: ordinary OPC part count, effective relationship count and expanded bytes. The part count excludes `[Content_Types].xml`; expanded bytes include it.
- `differenceCount`: the complete number of reported categories/parts/relationships. A changed XML part reports its first unequal token, not every changed property.
- `differences`: frozen records with `code`, logical `part`, optional relationship identity, first unequal byte offset or XML position.
- `truncated`: whether `maxDifferences` limited the retained records. Comparison/counting continue, and `equal` cannot become true because the report was truncated.
- `workUnits`: charged comparison work for byte/string code units and XML tokens.

Codes are `part-added`, `part-removed`, `content-type-changed`, `payload-changed`, `xml-changed`, `relationship-added`, `relationship-removed`, and `relationship-changed`. Part/control findings are sorted by exact logical name, followed by relationship findings sorted by source/ID. Sorting is independent of locale and ZIP ordering. A changed relationship XML part and its changed graph edge can produce separate findings.

`byteOffset` is zero-based in expanded payload bytes; for unequal prefix lengths it is the shorter length. XML `path` uses `$declaration/standalone` or `$element[n]` with optional child/attribute position. Element indices are one-based preorder indices in the original retained document; they are diagnostic locations, not XPath expressions or source mutation handles. Findings contain no source text or binary payload.

## Options, budgets and failure

`OpcComparisonOptions` extends `OpcInspectionOptions`. ZIP/XML/OPC options apply to both inputs. The additional `comparisonLimits` defaults are exported as frozen `DEFAULT_COMPARISON_LIMITS`:

| Limit | Default | Meaning |
| --- | --- | --- |
| `maxTotalExpandedBytes` | 128 MiB | Combined expanded payload bytes across both inputs, including control XML |
| `maxWorkUnits` | 512 Mi units | Comparison byte/string code units and tokens examined |
| `maxDifferences` | 1,000 | Maximum retained findings |

Limits must be positive safe integers. Both input buffers and all limit/profile options are captured before the first await. A signal stays live. Parsing validates both full packages even when inputs are identical or a finding limit would be reached. No partial comparison is returned on failure.

The combined expanded budget clips each input's ZIP expansion allowance before decompression: a clipped expansion throws `archive-limit-exceeded`. When the first input leaves no allowance for the second, comparison throws `package-limit-exceeded`. Exhausted comparison work also throws `package-limit-exceeded`. Other archive/XML/OPC errors propagate as `PackageInspectionError`, rather than being misreported as an ordinary difference. A broken relationship target therefore rejects the candidate before any equality result. A well-formed package can still contain invalid application-level IDs or schemas; equality of such input is not a validity guarantee.

Traversal and byte comparisons check cancellation and yield periodically. `operation-aborted` returns no result; there is no shared mutable session, output writer, file replacement or implicit network access. Per-input inspection limits still bound parsing; `maxWorkUnits` is a comparison budget rather than a total CPU/memory meter. Retained trees/payloads are private to the call and no preservation handles are created.

<!-- doc-test: docs/examples/package-comparison.ts -->
```ts
import {
  compareOpcPackages,
  type OpcPackageComparison,
} from "@pptkit/pptx-parser";

export async function comparePreservedCandidate(
  original: Uint8Array,
  candidate: Uint8Array,
  signal?: AbortSignal,
): Promise<OpcPackageComparison> {
  return compareOpcPackages(original, candidate, {
    profile: "opc-payload-v1",
    comparisonLimits: {
      maxTotalExpandedBytes: 64 * 1024 * 1024,
      maxWorkUnits: 256 * 1024 * 1024,
      maxDifferences: 100,
    },
    ...(signal ? { signal } : {}),
  });
}
```

## Evidence and remaining work

- [Parser tests](../../packages/pptx-parser/test/comparison.test.mjs): binary CRC collision, unknown/unreferenced content, control declarations, graph targets/identities, XML normalization and hidden-source changes, prolog/epilog, input capture, deterministic bounded reports, budgets, cancellation and deep iterative traversal.
- [Independent package integration](../../scripts/package-comparison-parity.test.mjs): separately read ZIP payloads and SHA256 oracles, generated charts/notes plus opaque content, source loss during fresh output, single-fault mutations and three reopened no-op rounds.
- [Browser bundles](../../scripts/browser-bundles.test.mjs): ESM/global comparison against independent Chromium DOM structure with no retrieval.
- [Installed Node consumer](../../tests/consumers/node.mjs) and [declarations](../../tests/consumers/types.ts): exported package entry points, readonly findings and explicit profile types.
- [Retained native-producer artifact comparison](../testing/package-comparison-evidence.json): SHA256-verified inputs from the earlier LibreOfficeDev color experiment, independently checked through Python `zipfile`. Exact-payload findings match one added part, 23 removed parts and 30 changed payloads; both profiles report inequality. This reinspection does not launch an application or classify every producer rewrite as loss.

For a local pair after building Parser, run `node scripts/inspect-package-comparison.mjs before.pptx after.pptx`. The [runner](../../scripts/inspect-package-comparison.mjs) emits both profiles and input SHA256 values without writing either input.

IMP-08.1.1–.1.3 remain incomplete: effective theme/font/numeric comparison for selected fresh-output semantics, validated renaming bijections, deeper workbook/object invariants, the complete corruption corpus and native reopening gates still need evidence. This API does not open IMP-09 edit/rebuild saves.
