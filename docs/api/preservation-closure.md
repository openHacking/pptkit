# Experimental OPC source relationship closure

`inspectOpcClosure(bytes, options)` is exported by `@pptkit/pptx-parser` root and `/browser` entries, and as `PPTKit.inspectOpcClosure` in the global bundle. It validates the OPC package and analyzes dependencies from explicitly selected logical roots. Its fixed profile is `source-relationships-v1`.

This is a source relationship analysis, not an editable graph, garbage collector, standalone package extractor or preservation writer. It does not establish that all implicit/application-level references inside arbitrary XML/binary payloads are understood. Complete relationship coverage does not authorize deletion or edit-save. Sessions support original-byte no-op; the separate opt-in [text-run profile](text-edit-save.md) applies additional source/candidate guards before admitting that bounded edit. General rebuild remains blocked.

## Roots and traversal

`OpcClosureOptions` extends `OpcInspectionOptions` with required nonempty `roots` and optional `closureLimits`.

- `/` denotes the package source and follows its root relationship set.
- Other roots must identify existing ordinary parts. Logical-name grammar and ASCII-case-insensitive identity lookup match the [OPC profile](opc-inspection.md); returned identities use actual source spelling. Relative paths, fragments, absent parts and relationship-set roots reject.
- Multiple equivalent roots deduplicate after the raw root-count limit. Canonical roots sort independently of caller order and ZIP order.
- Each included ordinary part brings its own `.rels` part, when present. Recognized internal target parts are followed transitively. Cyclic/shared/self references do not duplicate selected parts or repeatedly traverse owners.
- External references remain exact strings and are not loaded, executed or resolved against an invented external base.
- Nested workbook/embedded package bytes are ordinary target payloads. There is no recursive embedded-package interpretation.

The content-types control stream is reported separately and charged to the selected expanded-byte footprint. Its declarations can still refer to parts outside the closure. Copying this list of parts and the original control stream does **not** establish a valid extracted package.

## All source branches participate

The effective OPC graph contains only relationships selected by its MCE consumer. Preservation analysis instead scans the original relationship XML, including all `Choice` and `Fallback` branches and ignored source markup. Both inactive and effective parseable dependencies participate in the union. The ordinary XML payloads are validated under existing budgets without retaining their trees for this analysis; only raw relationship documents are retained privately.

Each `OpcSourceRelationship` contains the owner `source`, `relationshipPart`, one-based raw XML `elementIndex`, available `id`/`type`/`target`/`targetMode`, and resolved `targetPart`/`fragment` when available. Invalid or missing fields in inactive branches remain represented with findings; a parseable target can still participate even if its ID/type is invalid.

`matchesEffectiveRelationship` means that the source fields match an effective graph edge. It does not identify the selected source declaration: identical definitions in mutually exclusive branches can both match.

Known declarations inside opaque markup are collected conservatively. Unknown elements/attributes, unsupported placement/fields and missing inactive internal targets produce findings. Their presence prevents a complete-coverage conclusion for the affected direction. Standard OPC owner-relative resolution does not claim to implement semantics introduced by opaque attributes such as an unsupported base override.

Repeated relationship IDs in different branches of the **same** `AlternateContent` can be mutually exclusive. Duplicates that may coexist under the recorded branch scopes produce `ambiguous-source-identity`; ordinary/unconditional declarations can coexist with branch declarations. Cross-correlations between separate MCE capability expressions are not solved, so uncertain identity combinations remain conservative findings.

Active package errors, including selected missing targets or duplicate selected relationship IDs, still reject during OPC validation. A missing inactive target can pass effective OPC inspection but result in incomplete source coverage. Analysis never silently treats those two situations as equivalent.

## Result and coverage

The frozen detached `OpcClosureInspection` contains:

| Field | Meaning |
| --- | --- |
| `roots` | Canonical ordinary roots/package source |
| `parts` | Selected unique part metadata, sorted by exact logical name |
| `relationships` | Raw source relationships whose owners were reached, including inactive branches |
| `incomingRelationships` | Raw source relationships from outside owners into selected target parts |
| `unreachedParts` | Parts not reached from these roots; these are **not** proven unused |
| `findings` | Source findings across the package, including outside owners |
| `coverage.outgoingComplete` | No detected source relationship findings within reached owners |
| `coverage.incomingComplete` | No detected source relationship findings in outside owners |
| `controlStream` | Content-types stream name and expanded bytes |
| `expandedBytes` | Unique selected part bytes plus the control stream |
| `sourceRelationshipCount` | All recognized raw declarations across the entire package |
| `effectiveRelationshipCount` | Entire effective OPC graph count |
| `workUnits` | Charged analysis work |

`parts` records expose name, ZIP entry name, content type, expanded bytes, discovery `reason` and optional first-discovery `via`. Reasons are `root`, `dependency`, and `relationship-set`. Explicit root ownership takes precedence; implicit relationship sets omit `via`. Dependency provenance identifies owner, optional relationship ID and raw element index.

Relationships/findings follow sorted relationship-part order and raw preorder. Parts and unreached names sort by code-unit order, independently of locale. Both source counts are global; they need not equal the length of the selected `relationships` array.

Finding codes are `opaque-source-markup`, `invalid-source-relationship`, `missing-source-target`, and `ambiguous-source-identity`. Every finding names its owner, relationship part, raw element index and reason. Missing/opaque information outside the selected closure affects incoming coverage because outside ownership may be incomplete. All findings remain available to explain the result.

Both completeness flags concern the recognized relationship-source profile. Ordinary XML/binary application references, source subtree ownership, object IDs, copies/deletes and source reconciliation still need IMP-07.2/07.3 and IMP-09. Even `outgoingComplete: true` and `incomingComplete: true` never prove that an unreached part may be deleted. Whole-source no-op continues to preserve such parts.

## Budgets and failure

`DEFAULT_CLOSURE_LIMITS` is frozen and exported:

| Limit | Default | Accounting |
| --- | --- | --- |
| `maxRoots` | 1,000 | Raw requested roots before deduplication |
| `maxSourceRelationships` | 100,000 | All raw `Relationship` declarations, including inactive branches |
| `maxClosureParts` | 20,000 | Unique selected ordinary and relationship-set parts |
| `maxClosureExpandedBytes` | 128 MiB | Unique selected expanded bytes plus control XML |
| `maxFindings` | 10,000 | All source findings, including outside the closure |
| `maxWorkUnits` | 16,000,000 | Root strings, raw traversal/field work, branch checks and graph analysis |

Every limit must be a positive safe integer. Existing per-input ZIP/XML/OPC budgets still govern package validation/decompression, including unselected XML. The closure byte limit accounts for the selected preservation footprint; it does not replace archive limits or claim to be a peak-memory meter. Inactive declarations are independently bounded instead of charging only effective edges.

Root arrays/limits and input bytes are captured before yielding; the abort signal remains live. Iterative XML/graph traversal and periodic cancellation checks avoid recursive graph loading. Exceeding a closure budget throws `PackageInspectionError` with `package-limit-exceeded`; existing ZIP/XML/OPC errors propagate. An aborted call throws `operation-aborted`.

Findings are never truncated into an apparently complete graph. Failure/cancellation returns no partial result, creates no session/writer state and changes no source bytes. Returned records contain metadata only; private trees, buffers and mutable maps do not cross the API.

<!-- doc-test: docs/examples/preservation-closure.ts -->
```ts
import {
  inspectOpcClosure,
  type OpcClosureInspection,
} from "@pptkit/pptx-parser";

export async function inspectSlideDependencies(
  source: Uint8Array,
  slidePart: string,
  signal?: AbortSignal,
): Promise<OpcClosureInspection> {
  return inspectOpcClosure(source, {
    roots: [slidePart],
    closureLimits: {
      maxSourceRelationships: 10000,
      maxClosureParts: 5000,
      maxClosureExpandedBytes: 64 * 1024 * 1024,
    },
    ...(signal ? { signal } : {}),
  });
}
```

## Evidence and remaining gates

- [Parser tests](../../packages/pptx-parser/test/closure.test.mjs): shared/cyclic/self dependencies, incoming ownership, all MCE branches, inactive failures/identity collisions, canonical roots, deterministic provenance, full limits, cancellation, immutable capture, 32 independently generated graph-model cases, long cycles and deep source markup.
- [Cross-package integration](../../scripts/preservation-closure-parity.test.mjs): generated images/charts/workbooks/notes, shared image outside owners, inactive resource targets, exact selected-byte accounting through a separate ZIP reader and original-byte sessions. Unreached data remains preserved; incomplete inactive dependencies do not grant edit permission.
- [Browser ESM/global tests](../../scripts/browser-bundles.test.mjs): independent DOM target enumeration, incoming ownership, missing inactive targets, budgets and no retrieval.
- [Installed consumer](../../tests/consumers/node.mjs) and [declarations](../../tests/consumers/types.ts): public API, readonly results and required root options.
- [Retained native artifact evidence](../testing/preservation-closure-evidence.json): SHA256-verified source and LibreOfficeDev output from the earlier color experiment, each inspected from `/` and its first slide. Independent Python `zipfile`/ElementTree traversal confirms selected parts, incoming references, unreached parts, source counts and expanded-byte accounting for these named ASCII-path artifacts. This reinspection launches no application and provides no new native edit/save evidence.

The [inspection runner](../../scripts/inspect-preservation-closure.mjs) accepts a source filename followed by one or more logical roots:

```sh
node scripts/inspect-preservation-closure.mjs source.pptx /ppt/slides/slide1.xml
```

It prints the source SHA256 and full inspection without writing the input.

IMP-07.1.1 remains open for complete preservation-state retention/reconciliation and operation-level evidence. This source union is a prerequisite for that work; it does not close unknown XML editing, cleanup ownership, native save or the full backlog.
