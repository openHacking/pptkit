# Experimental OPC package inspection

`inspectOpcPackage(bytes: Uint8Array, options?: OpcInspectionOptions): Promise<OpcPackageInspection>` is available from the root and `/browser` entries of `@pptkit/pptx-parser`, and as `PPTKit.inspectOpcPackage` in the global bundle. It combines the [ZIP inspector](archive-inspection.md), [XML reader](xml-inspection.md), content-type mapping and relationship graph checks.

A successful result establishes the documented package profile. It does not establish that the package is a presentation, validate every OOXML schema or embedded object, import a Core document, or promise preservation saves. Empty and non-presentation OPC packages are supported; `inspectPptx` and semantic import/session APIs remain future work.

The implementation follows the package model and mappings in [ECMA-376 Part 2, fifth edition](https://ecma-international.org/publications-and-standards/standards/ecma-376/), especially clauses 6.2, 6.5 and 7.2–7.3. Its explicit restrictions below are not a universal OPC conformance claim.

## Content types and XML coverage

The exact `[Content_Types].xml` stream is mandatory and is inspected as XML, but is not returned as an OPC part. `Types`, `Default` and `Override` markup is namespace-aware, with required attributes, no unknown control elements/attributes and no nested content in declarations. Namespace declarations, comments and processing instructions are inert. Empty `Types` is accepted for an empty package.

Default extensions use the OPC extension spelling profile and ASCII case-insensitive lookup. Override names use logical part identity and take precedence. Duplicate equivalent defaults/overrides, absent override targets and uncovered ordinary parts reject. Unused defaults are allowed. MIME type/subtype tokens and ASCII parameter syntax are checked; original content-type spelling is retained in the result. This profile does not validate charset parameters against payload encodings or interpret arbitrary MIME-specific parameters.

Relationship parts have their defined relationship content type even when the types stream does not declare it. A conflicting declaration, parameters on that type, or that type assigned to an ordinary part rejects. Application/XML, text/XML and media types whose subtype ends in `+xml` are parsed regardless of filename or reachability. Binary payloads are CRC/size checked but are not interpreted or recursively opened. An extension `.xml` alone does not override a binary content type.

Every declared XML part must pass the per-part XML profile and cumulative budgets. OPC-defined control XML (the types stream, relationships, core properties and signature XML) additionally restricts explicit encoding declarations to UTF-8 or UTF-16; BOM-marked UTF-16LE/BE bytes remain supported, but declarations naming UTF-16LE/BE are rejected for these OPC-defined streams. Namespace URIs in ordinary XML are inventoried, not interpreted as known/unknown semantic capabilities. XML-level relationship-ID references in slide markup, Strict/Transitional semantics, AlternateContent branch selection, embedded ZIP/workbook contents and content-specific schema rules are not validated by this package API yet. Corpus tests inspect embedded workbook packages separately to exercise the API; production inspection does not implicitly recurse.

## Part names and references

Returned part names are absolute logical names such as `/ppt/slides/slide1.xml`; `entryName` retains the actual ZIP item name. Non-ASCII logical characters are decoded from UTF-8 percent sequences in ASCII ZIP names. Escaped ASCII reserved characters, such as `%20` for a space, retain their escaped spelling. Direct non-ASCII ZIP item names are outside this OPC profile, although the generic ZIP inspector accepts UTF-8 names.

Lookup folds only ASCII A–Z. It does not fold non-ASCII case, decode escaped ASCII reserved characters into alternate spellings, or normalize relationship type identifiers. Empty segments, segments ending in a dot, invalid UTF-8/percent encodings, encoded unreserved ASCII or slashes, and characters outside the supported part-name grammar reject. The ZIP inspector's additional NFC/case/percent alias restrictions still apply, so this is a conservative intersection of the two profiles. Some otherwise legal OPC names and aliases are therefore rejected before graph inspection.

Directory entries are included in the ZIP inventory but are not OPC parts. Interleaved piece items and `[trash]` physical items are explicitly unsupported. Other unmappable physical items reject rather than silently disappearing from an apparently complete inventory. Full physical-item preservation remains future work.

## Relationship graph

Relationship sets are located by the reserved `_rels`/`.rels` naming convention. The source is `/` for the package set, otherwise the actual logical source part. A source must exist; relationship parts cannot be source or target endpoints. Relationship IDs are source-local xsd:ID/NCName values with XML whitespace collapse; duplicates within a set reject. Original relationship type strings are retained and compared case-sensitively.

Internal targets are resolved relative to the **source part**, not the relationship file's directory. The resolver supports root-relative paths, dot/parent segments and fragments, resolves ASCII-case-equivalent spellings to the actual part name, and rejects missing targets. Shared assets, self-links and cycles are allowed; no recursive graph walk is needed. The raw target is always retained; `targetPart` is the actual resolved part name and `fragment`, when present, is the uninterpreted spelling after `#`.

Internal targets with schemes, authorities, queries, excess parent traversal above the package root, empty/directory paths or prohibited escaped part-name characters reject. Fragment-only references may address the source part, but cannot address the package itself. The excess-parent and escaped-name restrictions are deliberate profile choices; the implementation does not silently clamp or normalize them into different targets.

External targets retain their original string and never receive a `targetPart`. Relative external targets are not resolved against an invented base URI; absolute targets, file URIs and other schemes are not loaded or executed. Basic lexical screening rejects empty targets, controls/whitespace, malformed percent escapes and malformed first-segment schemes. This is not a general IRI validator, URL allowlist or permission to navigate/fetch the returned reference. The caller owns any later resource policy.

Relationship control XML must use the Relationships namespace and expected attributes. `xml:base` and unknown ordinary attributes reject. Relationship XML that declares/uses MCE is projected with the [markup compatibility processor](markup-compatibility.md) before these checks; `OpcPartInfo.markupCompatibility` reports branch/ignore decisions. Foreign extensions without MCE directives still produce `package-unsupported`. Active unsupported MCE capabilities produce `mce-unsupported`; malformed supported constructs produce `mce-invalid`. Relationship element character content permitted by its string content model is counted as XML content but not interpreted as a graph edge. No source markup is rewritten.

## Result, budgets and failure behavior

The result contains detached readonly archive inventory, part inventory (content type, source ZIP spelling, relationship-part flag, byte count and optional XML summary), relationship records, explicit default/override declarations, and cumulative XML totals. Parts follow archive inventory order; relationships follow each relationship part's XML order. Compare names/identities rather than assuming ZIP reordering preserves array order. Binary/XML payloads, mutable trees and internal maps are not exposed.

Options have separate `zipLimits`, `xmlLimits`, aggregate `limits` and `signal`. Existing ZIP/per-part XML defaults apply; aggregate defaults are:

| Limit | Default |
| --- | --- |
| `maxXmlParts` | 10,000, including the types stream |
| `maxXmlCharacters` | 32 Mi UTF-16 source units |
| `maxXmlNodes` | 1,000,000 XML nodes |
| `maxXmlTextCharacters` | 16 Mi UTF-16 content units |
| `maxRelationships` | 100,000 graph edges |

`DEFAULT_OPC_LIMITS` is frozen. Limits are validated and snapshotted before asynchronous work, as are input bytes and signal identity. Shared XML counters are enforced **during** each part's parsing rather than summed only after a potentially large tree has been allocated. XML totals use the units defined in the XML contract. Relationship counts are checked as validated records are added; XML/node budgets independently bound the control tree built before that stage.

ZIP payloads are privately retained during inspection, subject to the ZIP expanded-byte limits; XML parsing adds its own bounded strings/trees. These limits do not promise a hard process-memory cap. There are cooperative yields during ZIP/XML processing, every 16 XML parts and every 512 relationship records. Failure/abort returns no partial package inventory, mutates no caller bytes, and leaves independent/retried calls usable.

ZIP/XML failures retain their existing typed codes. Package checks add `package-invalid`, `package-unsupported` and `package-limit-exceeded`, with an entry label when available. The first detected failure is reported; error ordering is not a stable multi-error diagnostic contract. Invalid option shapes/limits raise TypeError/RangeError.

## Checked consumer

<!-- doc-test: docs/examples/opc-inspection.ts -->
```ts
import { inspectOpcPackage } from "@pptkit/pptx-parser";

export async function inspectPackage(bytes: Uint8Array, signal?: AbortSignal) {
  const result = await inspectOpcPackage(bytes, {
    zipLimits: { maxInputBytes: 32 * 1024 * 1024 },
    xmlLimits: { maxDepth: 128 },
    limits: { maxXmlNodes: 500000, maxRelationships: 20000 },
    ...(signal ? { signal } : {}),
  });
  return {
    parts: result.parts,
    externalReferences: result.relationships.filter((reference) => reference.targetMode === "External"),
    xmlTotals: result.xmlTotals,
  };
}
```

## Verification

Independent ZIP builders exercise content-type precedence, malformed control markup, URI spelling, encoded Unicode, source/target identity, cycles, fragments, external references, all resource budgets and ownership/cancellation. Generated PPTX fixtures and embedded chart workbooks run through this inspector alongside the existing independent Chromium package verifier. Browser tests verify graph results, dangling-reference and aggregate-limit failures with zero resource requests. Installed Node/TypeScript/ESM/global consumers use the public API.

The [native inspection evidence](../testing/xml-inspection-evidence.json) also records OPC part/relationship counts for the seven hash-matched LibreOfficeDev save artifacts. Reproduce with `pnpm inspect:native-xml` when those local artifacts and Python 3 are available. This re-inspection does not replace application editability, playback or preservation-save verification. IMP-02 remains incomplete pending full presentation dialect/extension policy, broader physical/container policies and their conformance evidence.
