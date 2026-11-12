# Presentation package inspection

`inspectPresentationPackage(bytes, options?)` is an experimental Node/browser API in `@pptkit/pptx-parser`. It identifies a PresentationML package entry point after the complete [OPC inspection profile](opc-inspection.md) succeeds. It does not import slides, validate the PresentationML schema, process MCE, execute macros or establish round-trip fidelity.

## Entry contract

The package must have exactly one root `officeDocument` relationship of either recognized dialect. The relationship must target an internal whole part, without a fragment. The target can have any OPC-valid name; there is no assumption that it is `ppt/presentation.xml`, and no filename extension is read from a caller path. Nested `officeDocument` relationships are not package entry points.

The target must have one of these main content types (MIME type matching is case-insensitive; parameters are unsupported):

| Main content type | kind | macroEnabled | suggestedExtension |
| --- | --- | --- | --- |
| application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml | presentation | false | pptx |
| application/vnd.openxmlformats-officedocument.presentationml.template.main+xml | template | false | potx |
| application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml | slideshow | false | ppsx |
| application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml | presentation | true | pptm |
| application/vnd.ms-powerpoint.template.macroEnabled.main+xml | template | true | potm |
| application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml | slideshow | true | ppsm |

`macroEnabled` describes only the main content type. It does not assert that a VBA project exists or that a non-macro-enabled package contains no executable content. `suggestedExtension` is a classification label, not permission to rename, save or execute the package. Macro structure validation and preservation policy remain ADV-11 work.

The main root must be `presentation` in the namespace matching the entry relationship:

| dialect | officeDocument relationship type | root namespace |
| --- | --- | --- |
| transitional | http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument | http://schemas.openxmlformats.org/presentationml/2006/main |
| strict | http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument | http://purl.oclc.org/ooxml/presentationml/main |

Prefixes are irrelevant. Namespace and relationship URIs remain case-sensitive. An unused declaration of another dialect does not change the classification. The `dialect` field identifies the entry relationship and main root only; it is not a claim that every part conforms to that dialect. Mixed or unknown descendant markup remains in the package inventory without interpretation.

## Result, limits and failures

`PresentationPackageInspection` returns `package` (the complete detached `OpcPackageInspection`), `mainPart` (logical absolute part name), `mainRelationshipId`, `dialect`, `kind`, `macroEnabled`, and `suggestedExtension`. Readonly TypeScript declarations prevent ordinary consumer mutations; results are not recursively frozen and changes to one result cannot affect later calls.

Options are `OpcInspectionOptions`, including all ZIP, per-part XML, aggregate XML/relationship budgets and cancellation. The implementation performs one bounded OPC read and classifies its inventory; it does not inflate or parse the source again. The same synchronous input/options capture, failure atomicity and cancellation identity apply. No resource fetching or filesystem I/O occurs.

`PackageInspectionError` retains lower-layer error codes. `presentation-invalid` reports a missing/ambiguous entry, an external/fragment entry, or a root/relationship mismatch. `presentation-unsupported` reports an unrecognized main content type, including Word/Excel packages and parameterized main types. `entry` locates the relationship or main part. Lower-layer failures take precedence because the whole OPC package is checked before entry classification.

A minimal root, unknown child, invalid schema attribute or unprocessed AlternateContent can still receive an entry classification. No object is reported as imported, editable or preserved. Namespace inventories include declarations and do not prove feature usage.

## Example

```ts
import { inspectPresentationPackage } from "@pptkit/pptx-parser";

export async function identifyPresentation(bytes: Uint8Array) {
  const result = await inspectPresentationPackage(bytes, {
    limits: { maxRelationships: 20000 },
  });
  return {
    mainPart: result.mainPart,
    dialect: result.dialect,
    kind: result.kind,
    suggestedExtension: result.suggestedExtension,
  };
}
```

The [checked example](../examples/presentation-inspection.ts) uses only public exports. Independent synthetic packages cover both dialects and all six content types; generated/native fixtures and installed consumers provide additional evidence. Synthetic Strict fixtures identify the namespace contract, not native Strict application conformance.

Format references: [Microsoft PresentationML structure](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document), [Microsoft main-part relationship rule](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/32be961f-d71a-4812-913f-b675c79aa88a), [Office macro-enabled format specification](https://officeprotocoldoc.z19.web.core.windows.net/files/MS-OFFMACRO/%5BMS-OFFMACRO%5D.pdf).

For effective MCE-aware ordering, dimensions and master/layout/notes/theme references, use [PPTX structural inspection](pptx-inspection.md). This entry-only API retains its cheaper root/type identification contract.
