# Experimental XML part inspection

`inspectXmlPart(bytes: Uint8Array, options?: XmlInspectionOptions): Promise<XmlPartInspection>` is available from the root and `/browser` entries of `@pptkit/pptx-parser` and as `PPTKit.inspectXmlPart` in the global bundle. It verifies a complete XML document and returns detached readonly metadata. It does not create a Core document, validate OOXML schemas, select AlternateContent branches, inspect package relationships, or guarantee preservation saves.

The actual typed-array view length is checked before copying, without invoking subclass getters or iterators. Shared-memory-backed views are rejected, including buffers from another JavaScript realm; callers can supply an ordinary detached copy instead.

## Accepted XML profile

The reader uses namespace-aware [XML 1.0](https://www.w3.org/TR/xml/) and [Namespaces in XML 1.0](https://www.w3.org/TR/xml-names/) well-formedness checks. It rejects malformed nesting, multiple roots, undeclared prefixes, invalid characters/references, duplicate attributes including expanded-name aliases, and invalid namespace bindings. Default namespaces apply to element names; unprefixed ordinary attributes retain an empty namespace URI. Namespace URI strings are recorded, not fetched or interpreted as schema support.

Supported byte encodings are UTF-8 with an optional BOM and UTF-16LE/UTF-16BE with a required BOM. Decoding is fatal on malformed byte sequences. A declaration, if present, must specify XML 1.0 and agree with the detected encoding: `UTF-8`, the matching `UTF-16LE`/`UTF-16BE`, or `UTF-16` for either BOM-marked UTF-16 order, case-insensitively. XML 1.1, other encoding declarations, UTF-32 and BOM-less UTF-16 are outside this profile. Repeated leading BOMs reject; U+FEFF inside character content remains legal.

DTD declarations are rejected, including internal subsets and external identifiers. Only predefined XML entities and legal numeric character references are decoded. There is no entity resolver, schema loader, stylesheet evaluator, filesystem access or implicit network request. Processing instructions, comments and CDATA remain inert content. A DTD-like string inside a comment or CDATA is not a declaration.

## Metadata and units

The result includes root qualified name/prefix/local name/namespace URI, detected encoding, input bytes, decoded source characters, node/element/attribute counts, maximum depth, content characters and sorted unique nonempty namespace URIs. URIs include those used in names and namespace declarations, including the namespace declaration URI itself when present.

`characters` counts decoded source UTF-16 code units after removing the encoding BOM, before XML line-ending/entity normalization. `textCharacters` counts normalized text, CDATA, attribute values (including namespace declaration values), comments and processing-instruction bodies in UTF-16 code units. Names and PI targets are covered by source limits, not counted as content text.

`nodes` counts elements, contiguous ordinary text runs, individual CDATA sections, comments and processing instructions. The XML declaration and attributes are not separate nodes. Whitespace outside the root participates in text/node budgets. `elements` includes the root; root depth is one. Inventory does not expose a DOM or mutable parser. Internal ordered tree projection is reserved for package parsing, excludes prolog/epilog nodes, and is not a lossless serialization format; future preservation must retain source bytes separately.

## Limits, ownership and cancellation

| Limit | Default |
| --- | --- |
| `maxInputBytes` | 8 MiB |
| `maxCharacters` | 4 Mi UTF-16 source units |
| `maxNodes` | 100,000 |
| `maxDepth` | 256 |
| `maxAttributesPerElement` | 256, including namespace declarations |
| `maxTextCharacters` | 2 Mi UTF-16 content units |
| `maxPendingCharacters` | 128 Ki UTF-16 source units between parser events |

`DEFAULT_XML_LIMITS` is frozen. Overrides must be positive safe integers. Input bytes, limits, part label and signal identity are captured before the first asynchronous yield. Mutating caller bytes/options cannot change the in-flight interpretation. The part label is diagnostic metadata, never a path or URL to load.

Decoding uses chunks of at most 1 KiB and yields a task every 32 chunks, allowing cancellation during large documents. Source/structure/content limits are enforced as parsing proceeds. `maxPendingCharacters` also bounds unfinished attribute values, names, comments, CDATA, declarations, PI and text before a final delimiter arrives. It measures a conservative source span between public parser events rather than a semantic token length; an event gap may include delimiters. A single long valid text run can exceed this limit even while total content remains below `maxTextCharacters`; callers needing such runs must raise both relevant limits explicitly.

These are logical input/work limits, not a hard process-memory sandbox. A parser token can temporarily exceed its pending limit by at most one decoded input chunk before rejection, and parser state/decoded strings incur overhead. Node/attribute/depth limits apply before retaining a further corresponding structure. No successful partial inventory is returned on failure or cancellation.

Errors use `PackageInspectionError` with `xml-invalid`, `xml-unsupported`, `xml-limit-exceeded`, or `operation-aborted`. `entry` contains the supplied part label. When available, `location` provides a frozen parser cursor with one-based line and column and a UTF-16 source offset; it is not a promise of the offending token's start. Errors detected before decoding need not have a location. Invalid options/input types use TypeError/RangeError.

## Checked consumer

<!-- doc-test: docs/examples/xml-inspection.ts -->
```ts
import { inspectXmlPart, PackageInspectionError } from "@pptkit/pptx-parser";

export async function inspectPresentationXml(bytes: Uint8Array, signal?: AbortSignal) {
  try {
    return await inspectXmlPart(bytes, {
      partName: "ppt/presentation.xml",
      limits: { maxDepth: 64, maxNodes: 20000 },
      ...(signal ? { signal } : {}),
    });
  } catch (error) {
    if (error instanceof PackageInspectionError) {
      return { code: error.code, part: error.entry, location: error.location };
    }
    throw error;
  }
}
```

## Dependency and verification boundary

The XML implementation uses pinned `saxes@6.0.0`, already present in the workspace's development dependency graph, behind a private adapter. Its [upstream repository](https://github.com/lddubeau/saxes) is archived; PPTKit owns the wrapper's supported profile and regression evidence. No saxes types/instances cross the public API. A CommonJS build plugin bundles saxes/xmlchars for browser consumers; distributed browser artifacts and the package include third-party notices.

The wrapper uses only public events/cursors. In saxes 6 the position read immediately after `write` double-counts that chunk; post-write pending checks use the wrapper's own decoded source count. Events use their event-time cursor. A dedicated many-short-tags/chunk-boundary test prevents this dependency behavior from becoming a false resource-limit failure.

Tests cover encoding, normalization, namespaces, malformed input, DTD/entity rejection, all limits, long unfinished tokens, exact boundaries, input snapshots and cancellation. Chromium DOMParser provides an independent namespace/count and well-formedness comparison; real PPTKit PPTX/XML and embedded-workbook parts plus installed Node/browser consumers exercise integration. This evidence covers the named profile, not universal XML or OOXML conformance. Package-wide XML totals, content-type dispatch and a bounded OPC relationship profile are now available through [inspectOpcPackage](opc-inspection.md). Strict/Transitional/Markup Compatibility semantics remain under IMP-02.


## Native-producer sample evidence

[Recorded results](../testing/xml-inspection-evidence.json) cover the seven retained LibreOfficeDev save artifacts described in [the native evidence](../testing/libreoffice-evidence.json). The verification script first matches each file's SHA-256 against that earlier record, runs the public ZIP and OPC inspectors, then independently checks XML/relationship parts extracted by Python's independent ZIP reader. It does not rerun LibreOffice or establish semantic/visual round-trip fidelity.

With the retained `.tmp/compatibility/roundtrip` artifacts, built Parser package and Python 3 available, run `pnpm inspect:native-xml`. The script fails if a source hash, ZIP inspection, independent extraction or XML inspection fails and writes the JSON report only after all fixtures pass. CI corpus tests use generated artifacts and do not assume the local native files exist.
