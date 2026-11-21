# PPTX structural inspection

`inspectPptx(bytes, options?)` is an experimental structural reader in `@pptkit/pptx-parser`, available from root/browser/global entries. It reads actual PresentationML ordering, dimensions, source identities, structural relationships, document metadata and native section membership. It does not create a Core document, import drawing content or create a preservation-save session.

## Result

`PptxInspection` includes [presentation entry identification](presentation-inspection.md) and its full OPC inventory, plus:

| Field | Meaning |
| --- | --- |
| `metadata` | Optional core/extended/custom groups, each with source part, package relationship ID and ordered properties; absent groups are not fabricated. See the metadata profile below. |
| `sectionList` | Optional native MS-PPTX section collection with source GUID/name, ordered slide references and uninterpreted nested extensions. No collection is synthesized from slide labels. |
| `presentationExtensions` | Direct main-presentation extension descriptors: normalized URI, interpreted/uninterpreted status and direct payload root names; no raw XML. |
| `slideSize` | Optional `{ widthEmu, heightEmu }` from the effective `sldSz`. No size is invented when the element is absent. |
| `notesSize` | Required notes dimensions in exact EMUs. |
| `firstSlideNumber` | Signed numbering setting; defaults to 1. This is not a source slide ID or array index. |
| `firstSlideNumberDeclared` | Whether the source explicitly declares `firstSlideNum`; `false` distinguishes the default 1 from an explicit 1. Optional in the TypeScript interface for compatibility with older inspection producers. |
| `slides` | Presentation order from `sldIdLst`; each record contains numeric source `id`, relationship ID, logical source part, hidden flag, optional cSld name, layout part and notes part. |
| `slideMasters` | Listed master source identities, their ordered layout identities/references, and optional theme part. |
| `notesMasters`, `handoutMasters` | Listed auxiliary master relationship IDs and parts; each list permits at most one entry. |
| `layouts` | Unique encountered layouts, optional cSld names and master part references. Shared layouts appear once. |
| `findings` | Located unresolved/missing inheritance and unlisted structural references. These do not imply the package is fully valid or ready for semantic import. |
| `markupCompatibility` | Effective MCE reports for interpreted presentation and extended/custom metadata parts, keyed by logical part name. Relationship reports remain in `package.parts`. |

Source numeric IDs and relationship IDs remain distinct from future Core IDs. Names are `cSld/@name`, including an explicitly empty name; no title is inferred from text. Slide `show` accepts the XML boolean spellings `true`, `false`, `1`, `0`, with surrounding XML whitespace; absent means shown. `hidden` is the inverse.

ZIP record order, relationship enumeration, filenames, numbering settings, hidden flags and unused slide relationships never determine page order. An absent or empty slide list yields an empty ordered list even when unlisted slide parts exist. Custom shows and native sections never redefine this order; native section membership is reported separately.

## Validated structure and failures

The reader checks these contracts after the bounded OPC pass:

- Recognized main type and entry dialect, with matching effective roots for each interpreted slide/master/layout/notes/theme target.
- Exact source-local `r:id` resolution, correct relationship type, internal whole-part targets and appropriate target content types. Relationships with external targets or fragments cannot stand in for structural parts.
- Unique numeric IDs within each identity list and unique target parts within each list. Source aliases resolving to the same part cannot bypass target uniqueness.
- Slide IDs in `[256, 2147483647]`; master/layout IDs in `[2147483648, 4294967295]`. Integer lexical validation accepts signs and leading zeros, rejects fractions/exponents/non-XML whitespace, and checks numeric rather than string uniqueness.
- Slide dimensions in `[914400, 51206400]` EMUs; notes dimensions in `[0, 27273042316900]`. These follow the fifth-edition Strict/Transitional schemas. Values remain integers; no pixel conversion, rounding or layout inference occurs.
- At most one of each interpreted direct child list/size element, valid identity-list child names and no non-whitespace list text. Notes size is required. The reader does not validate every attribute, ordering rule or drawing schema.
- At most one implicit slide-layout, layout-master, master-theme, slide-notes or notes-backlink relationship per source where inspected. A listed layout pointing back to another master, notes pointing back to another slide, or multiple slides sharing a notes part rejects.

Malformed interpreted structure throws `PackageInspectionError` with `presentation-invalid` and the source part. Lower-layer ZIP/XML/MCE errors retain their codes. Missing implicit inheritance links produce `missing-inheritance` findings instead of invented defaults. Layouts or notes pointing to unlisted masters, layouts absent from their listed master, and presentation slide relationships absent from `sldIdLst` produce explicit findings. Consumers must inspect these findings before relying on resolved inheritance. This reader does not calculate inherited colors, placeholders, text or geometry.

Listed masters/layouts/auxiliary masters and used slide/notes/theme targets are checked. Unlisted main slide relationships also receive target type/root checks but are not added to page order. Metadata content types also receive package cardinality/linkage checks; other unreferenced parts retain OPC/XML inventory only. The inspector is not a whole-package OOXML schema validator.

## MCE and unknown content

MCE is processed before interpreting fields in parts that declare/use its namespace. The fixed structural configuration recognizes the entry dialect's PresentationML, officeDocument relationships and base DrawingML main/chart/diagram/picture/lockedCanvas/compatibility namespaces. Namespace recognition supports structural preprocessing; it does not claim those objects are editable. PresentationML `ext` and DrawingML `ext` are opaque extension islands, except the recognized main-presentation section extension described below. Vendor extensions are not granted capabilities simply by being declared.

Unknown active namespaces outside applicable ignorable rules/islands reject through the [MCE profile](markup-compatibility.md). Unselected branches cannot contribute slide IDs or trigger missing relationship errors. Decisions are reported; discarded or opaque content is not claimed as preserved. Source XML budgets include every branch.

Unlike the cheaper entry-only API, this reader can process a root AlternateContent before checking its effective main root. Every interpreted part must ultimately have exactly one expected root with no non-whitespace text outside it. Main entry relationship/type constraints still apply.

Text, shapes, images, charts, notes content, custom shows, animations and extension payloads are not semantically imported here. The raw package inventory is not a save guarantee. IMP-03 remains partial until content mapping and valid Core/session behavior are delivered.

## Native sections and presentation extensions

`sectionList` recognizes the MS-PPTX PowerPoint 2010 section extension in the main presentation's `extLst`, with URI `{521415D9-36F7-43E2-AB2F-B90AF26B5E84}` and payload namespace `http://schemas.microsoft.com/office/powerpoint/2010/main`. URI comparison uses XML token whitespace normalization and otherwise exact spelling. Both Strict and Transitional presentation containers are covered by synthetic fixtures; native evidence currently covers Transitional only.

The collection records its main `part`, `extensionUri` and source-ordered `sections`. Each section has optional source GUID `id`, optional `name`, ordered `slides` containing numeric slide ID, actual part and relationship ID, and nested extension descriptors. The source schema permits absent names/GUIDs; neither is fabricated. Empty names remain distinct from absent names. Repeated names do not merge sections. Present GUIDs must match the uppercase schema spelling and be unique after token whitespace normalization.

An absent extension yields no `sectionList`. A present section list must contain at least one section; a section can have zero slides. Members resolve by source numeric ID in the effective presentation order, never by filename, array position, section name or ZIP ordering. Invalid numeric identities, unresolved IDs and repeated membership within/across sections reject with `presentation-invalid`. A physical slide absent from the effective presentation order cannot satisfy a section reference.

Partial or differently ordered membership is retained with findings instead of silently repaired: `section-order` includes `sectionIndex` (zero-based) and `slideId` when consecutive membership does not follow presentation order; `unsectioned-slide` identifies each listed slide omitted from the section collection. Empty sections do not advance the position. These findings do not reorder pages or assign missing members to a guessed section.

The reader checks the known extension envelope, section/list child names, child sequence/cardinality, scalar attributes and forbidden text/nested ID content. Duplicate known extension instances reject. A section's optional `p14:extLst` follows the MS-PPTX schema's Transitional `p:CT_ExtensionList` child namespace even inside a Strict container. Nested payloads remain uninterpreted, with URI/root descriptors. Other main-presentation extensions likewise expose descriptors. Descriptor status `interpreted` means the recognized section fields were read, not that every nested extension is understood or preserved. Unknown payloads and missing unknown-extension URIs are not fully schema-validated.

For MCE, only the main part additionally recognizes the section namespace. A private format-reader hook opens the known section `ext` only under the presentation → extLst ownership path; generic MCE callers retain their existing opaque-element contract. Ancestor namespace/ignorable scopes, AlternateContent wrappers and ProcessContent wrappers remain effective. Unknown active content in the interpreted section grammar rejects; unknown nested islands remain opaque, including islands reusing the section URI under `p14:extLst`. An opaque ancestor always takes precedence. Non-main slide/master/layout parts do not acquire the section namespace capability.

This inspection API does not create a Core document. Separate [import projections](import-session.md#native-section-member-labels) map named members onto Core `slide.section` labels with native-collection limitations reported. Core fresh export currently writes only PPTKit slide-extension label data. Native section collection editing, cross-document identity mapping and structural-save reconciliation remain open under IMP-03.5/IMP-09/AUT-07.

References: Microsoft [section extension placement](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/1f21a089-944d-410b-bd47-4f5e692c2532), [section fields](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/5bb610d5-ced8-44ed-a4b6-cd3d5f938906), [section-list cardinality](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/b499c39b-7c50-4a58-88e6-f86e6046be50), [slide-list cardinality](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/7bde1772-1da3-4311-b7f1-dcf5f7fe96c4) and [complete extension schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/cc915728-85a4-46a0-a07c-eb5416a5762e). See [native section evidence](../testing/native-sections-evidence.json) for the scoped LibreOfficeDev save/reopen experiment; PowerPoint and interactive editing remain untested.

## Document metadata profile

For how inspected values map into Core during import, see the [metadata projection contract](metadata-projection.md).

`metadata.core`, `metadata.extended` and `metadata.custom` are optional `MetadataPartInfo<T>` records with `part`, `relationshipId` and source-ordered `properties`. Each recognized metadata content type permits at most one part and must be linked from the package root by the matching internal whole-part relationship. Wrong types, duplicate links, orphan typed parts, fragments, external targets, wrong effective roots and mixed metadata dialects reject with `presentation-invalid`. Conventional `docProps` filenames are not required. MIME comparisons are case-insensitive; namespace and relationship identities are exact.

- **Core:** all 15 properties are read by expanded XML name, not prefix. Duplicate/unknown properties, nested scalar elements and unrecognized attributes reject. `created`/`modified` require an `xsi:type` QName resolving to `dcterms:W3CDTF`, including namespace aliases and local rebinding. Ordinary core properties disallow `xml:lang`; keywords and their `value` children allow it. `keywordSegments` retains text/value boundaries, language inheritance and explicit empty language resets. The aggregate keyword `value` concatenates source text; it does not split a keyword list on punctuation. Structural import reports keyword language or segment information as unsupported by Core even when the keyword text is empty; the source inspection retains the annotation.
- **Extended:** all 23 scalar fields retain their XML-decoded lexical strings. Signed int32 ranges and XML boolean spellings are checked. Counts are source metadata, not recalculated or asserted to match actual slides. `HeadingPairs`, `TitlesOfParts`, `HLinks` and `DigSig` require their vector/blob wrapper and return an `uninterpreted` descriptor.
- **Custom:** source `pid` (int32), `fmtid` (schema GUID), optional name/linkTarget, and variant element type are recorded. Duplicate format-ID/property-ID pairs and duplicate present names reject. Distinct format IDs may reuse an ID. Missing names remain absent. Names such as `__proto__` and `constructor` remain array values, never object keys. The schema GUID spelling is uppercase hexadecimal inside braces; surrounding XML whitespace is collapsed. No Office-specific PID minimum or restricted variant subset is imposed.
- **Values:** string, integer, boolean, date, real, decimal, empty and null variants have `status: "scalar"`, `kind` and a string `value`. Integer widths/ranges, boolean, real/decimal lexical forms and empty/null content are checked. `i8`/`ui8`, high-precision decimals and real exponents never pass through JavaScript number conversion. Dates remain lexical values: **calendar/date lexical validity is not checked**, and no timezone conversion occurs. XML decoding/line-ending normalization has already happened; OOXML string escape sequences are not decoded a second time.
- **Uninterpreted values:** vector, array, binary/stream/storage, currency, error and class-ID values expose `status: "uninterpreted"`, `namespaceUri` and `local`. Their internal schemas, lexical values and semantics are not validated or returned. This is visible missing interpretation, not opaque source preservation. No signature validity or document protection claim follows from `DigSig`/`DocSecurity`. `linkTarget` and `HyperlinkBase` are recorded without resolving or loading resources.

Core properties prohibit MCE attributes/elements. An unused MCE namespace declaration alone is allowed. Extended/custom properties use a separate MCE configuration recognizing their own property namespace and the matching variant namespace; projection decisions join `markupCompatibility`. Unknown active content rejects outside applicable ignorable rules. This structural/lexical profile is not full ECMA schema validation, semantic import, or preservation save.

Structural import reports unmapped recognized core and extended properties and separately reports each custom property as mapped or unsupported by the Core projection. Custom findings use the format GUID and numeric property ID as source identity; the inspection record retains scalar XML-normalized text and marks complex values uninterpreted. These findings count against the session report budget. Source-preserving no-op saves still retain the original package bytes.

The structural/content import maps selected source values to editable Core metadata as follows:

| Source property | Core field | Projection |
| --- | --- | --- |
| `dc:title`, `dc:creator`, `dc:subject`, `dc:description`, `dc:language` | `title`, `author`, `subject`, `description`, `language` | XML-normalized text, including explicit empty strings. |
| Extended `Company` | `company` | XML-normalized text, including an explicit empty string. |
| `cp:keywords` | `keywords` | One array entry containing the aggregate text, including an empty entry; language and mixed-content boundaries remain in inspection only. |
| `cp:revision` | `revision` | Only canonical nonnegative decimal text within JavaScript's safe integer range; other forms remain in inspection with an unsupported finding. |
| Named, unlinked custom `lpwstr`, `bool`, canonical `i4`/`r8` and canonical UTC `filetime` with seconds or milliseconds | `customProperties` | Typed ordered entries; source IDs and original boolean spelling remain in inspection. Literal escaped underscores are decoded; other escape codes, overlapping decoded sequences and decoded-name collisions are unsupported. Unsupported custom properties are individually reported and omitted from the projection. |
| Other recognized core/extended and custom properties | None | Inspection only; complex custom/extended values may be uninterpreted. |

Absent source properties stay absent in `document.metadata`; they are distinct from present empty values. `normalizePresentation` separately defaults missing title to `PPTKit Presentation`, author to `PPTKit`, language to `en-US`, keywords to an empty array and revision to `1`; optional company, subject and description remain absent. These normalization defaults are not evidence of source values and may appear in a fresh export. Source-preserving no-op save retains the original bytes. Editable Core projection also does not imply that a field is admitted by the separately opt-in [source-preserving metadata save profile](core-metadata-save.md).

The metadata reader reuses the original bounded XML pass. Long custom-property and mixed-keyword lists cooperate with cancellation. Failed reads expose no partial result. Exported metadata interfaces are readonly; returned records remain detached mutable JavaScript objects at runtime.

## Resource and ownership behavior

Options are the existing `OpcInspectionOptions`. One OPC read captures caller input/options and applies ZIP, per-part XML and package-wide limits. The private reader retains bounded XML trees for structural interpretation, avoiding a second inflate/parse and avoiding double-counting XML budgets. This uses more memory than envelope-only inspection: the source XML trees plus effective projections may coexist. Configure limits for the available environment.

Relationship and part maps avoid repeated package scans; inverse master-layout membership uses sets. Cyclic layout/master and notes/slide edges are checked as links, not recursively expanded. Processing yields cooperatively and uses the originally captured abort signal. Failure returns no partial inspection. Results expose detached metadata only, are readonly in TypeScript and are not recursively frozen.

## Example and evidence

```ts
import { inspectPptx } from "@pptkit/pptx-parser";

export async function readSlideOrder(bytes: Uint8Array) {
  const result = await inspectPptx(bytes, {
    limits: { maxXmlNodes: 500000, maxRelationships: 20000 },
  });
  return {
    metadata: result.metadata,
    sections: result.sectionList,
    extensions: result.presentationExtensions,
    size: result.slideSize,
    slides: result.slides,
    unresolved: result.findings,
  };
}
```

See the [checked example](../examples/pptx-inspection.ts). Independent ZIP fixtures test Strict/Transitional source identities, shuffled record/relationship order, size/ID boundaries, reference failures, hidden states, notes back-links, missing inheritance, MCE order selection, ownership and cancellation. Generated corpus checks compare dimensions, hidden flags and exported document properties with authoring/Layout state. Metadata fixtures exercise both dialects, all core/extended fields, integer widths, GUID/name identities, mixed-language keywords, QName rebinding, malformed links/values, MCE, budgets, ownership and traversal cancellation. Browser and installed-package consumers exercise the public surface. Native inspection evidence now records document properties, ordered slides and structural findings for hash-matched LibreOfficeDev outputs; core/extended scalar values are also checked with independent Python ElementTree parsing. This is not a new native application round-trip or a Strict-native claim.

References: [ECMA-376 Parts 1 and 4](https://ecma-international.org/publications-and-standards/standards/ecma-376/) and their fifth-edition `pml.xsd` / `dml-main.xsd` schemas, [Microsoft presentation structure](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-presentations), and [slide show/hidden semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.slide).

Metadata references: ECMA-376 Part 1 §§15.2.12, 22.2–22.4 and the Strict/Transitional `shared-documentProperties*.xsd` schemas; OPC (Part 2, 2021) §11 for core properties and mixed-language keywords.
