# Markup compatibility inspection

`inspectMarkupCompatibility(bytes, options)` applies an explicit consumer configuration to an XML part using the experimental MCE processor in `@pptkit/pptx-parser`. It returns source XML inventory, effective output counts and located decisions. The private projection is available to Parser readers; it is not a public AST, serializer or preservation-save API.

## Configuration and processing

`MarkupCompatibilityOptions` extends [XML inspection options](xml-inspection.md). `understoodNamespaces` is a required array of exact namespace URIs. Declarations in an input do not grant understanding. URI comparison is case-sensitive; prefixes are resolved in the source scope. XML and XMLNS are understood automatically. `understandNoNamespace` defaults to true for unqualified elements/attributes and can be set to false. Do not pass an empty string or the MCE namespace in the namespace array.

`extensionElements` optionally names application-defined extension islands by `{ namespaceUri, local }`. MCE is suspended on the entire named element, its attributes and descendants, including unknown markup. Use only extension points specified by the consuming format; this configuration is not a mechanism for claiming arbitrary content is supported. XML well-formedness, encoding and resource limits still apply inside islands. The MCE namespace cannot be an island.

The implemented profile follows the constructs in ECMA-376 Part 3, fifth edition (2015):

| Construct | Behavior |
| --- | --- |
| `mc:Ignorable` | Add resolved namespace URIs to inherited ignorable scope. Unknown ignorable elements and attributes are removed from the effective projection. Understood namespaces are retained. |
| `mc:ProcessContent` | Resolve `prefix:local` or `prefix:*` against an ignorable namespace; unwrap matching unknown elements and process their children in order. |
| `mc:MustUnderstand` | Reject if an actively processed element names an unsupported namespace. Ignored subtrees and unselected branches do not trigger semantic mismatches. |
| `mc:AlternateContent` | Select the first Choice whose entire Requires list is understood; otherwise use Fallback if present; otherwise emit nothing. |
| `mc:Choice` / `mc:Fallback` | Validate placement, ordering and attributes. Only selected content enters the effective projection. |

Ignorable and ProcessContent scopes are additive, bind to namespace URIs rather than mutable prefixes, and end when their declaring element ends. Empty Ignorable/ProcessContent/MustUnderstand lists are legal; Requires must contain at least one bound prefix. Only XML whitespace separates tokens. ProcessContent cannot refer to a namespace that is not ignorable. Unwrapped elements carrying `xml:base`, `xml:lang` or `xml:space` reject because discarding that wrapper would alter inherited context.

Outside configured islands, structural/list syntax is checked throughout the source, including discarded branches. Unknown MCE constructs, including older `PreserveElements`/`PreserveAttributes` hints, reject explicitly instead of being silently treated as implemented preservation behavior. This strict rejection profile is separate from selecting an active branch. Non-MCE unknown namespaces in active ordinary markup also reject unless removed by an applicable ignorable rule; namespace understanding is not a schema-validity check.

AlternateContent permits ignorable foreign children, but a foreign child that survives ignoring (including an understood or unwrapped child) causes a mismatch. At least one Choice is required; Fallback is optional, unique and last among MCE branches. A root AlternateContent can yield zero or multiple output roots; this XML-part inspector reports the resulting forest. Consumers that require one document root must enforce that separately.

## Reports and ownership

`MarkupCompatibilityInspection` contains:

- `source`: the complete `XmlPartInspection`, counting all input including discarded branches.
- `output`: top-level element names, total elements/attributes and text/CDATA length in UTF-16 code units. Attributes include retained namespace declarations; text length excludes attribute values, comments and PI bodies.
- `decisions`: effective traversal decisions for ignored elements/attributes, unwrapped elements, alternate selection and opaque extension islands. Each contains the source expanded name and a one-based element index in the full original XML preorder, including discarded and opaque descendants. Ignored attributes additionally identify the attribute; alternate decisions identify the selected Choice/Fallback index when one exists.

Only reached decisions are reported. A dropped subtree or unselected branch is represented by its enclosing removal/selection, not an expanded per-descendant loss list. Removing MCE directives from effective markup is implicit. These are processing observations, not source-preservation guarantees or permission to save a lossy document.

Source bytes, limits, signal identity and capability configuration are captured before asynchronous work. Projection is iterative, has cooperative cancellation and cannot duplicate source subtrees. Scoped maps use reversible writes instead of copying inherited dictionaries for every element. XML budgets bound the entire source before projection; discarded or opaque content cannot bypass them. A failure returns no partial report. Reports are detached between calls and readonly in TypeScript, not recursively frozen.

The internal projection retains expanded names and source namespace declarations; removing wrappers can remove lexical declaration sites. It must not be serialized directly as namespace-correct XML. Future preservation saves must retain source bytes and reconcile edits independently. No source tree or payload escapes through this API.

## Errors and package integration

`PackageInspectionError` adds `mce-invalid` for malformed supported constructs and `mce-unsupported` for unsupported MCE policy or consumer capability mismatches. `entry` is the supplied `partName`; XML failures retain their existing source locations. MCE failures currently identify the part and reason, without byte/line coordinates. XML/budget/cancellation errors retain their lower-layer codes.

[OPC inspection](opc-inspection.md) now processes relationship XML that declares or uses the MCE namespace with only the OPC Relationships namespace understood and no opaque islands. `OpcPartInfo.markupCompatibility` records the effective report. The effective result must contain exactly one Relationships root, then pass ordinary relationship validation. Only selected relationship edges are resolved/count against `maxRelationships`; all branches still count toward XML source budgets. Unselected relationships can refer to absent targets because they are not effective package relationships.

Content-type control XML retains its existing strict schema/profile. Other XML parts, including PresentationML slide bodies, are not automatically processed by package inspection. [Presentation inspection](presentation-inspection.md) still classifies the main entry only. [PPTX structural inspection](pptx-inspection.md) applies a documented fixed configuration to interpreted PresentationML parts. Full semantic import, broader extension policy and preservation remain open.

## Example and verification

```ts
import { inspectMarkupCompatibility } from "@pptkit/pptx-parser";

export async function inspectCompatibilityPart(bytes: Uint8Array) {
  return inspectMarkupCompatibility(bytes, {
    understoodNamespaces: ["http://schemas.openxmlformats.org/presentationml/2006/main"],
    partName: "ppt/slides/slide1.xml",
    limits: { maxDepth: 256, maxNodes: 100000 },
  });
}
```

This example declares only one consumer namespace; real importers must explicitly supply every namespace their implementation understands. See the [checked example](../examples/markup-compatibility.ts).

Focused tests verify actual projected text/order, namespace scope exit/rebinding, a fixed eight-case nested-choice truth table, inactive MustUnderstand, opaque islands, deep iterative traversal, malformed constructs, ownership and cancellation. OPC tests verify effective edges and raw-source budgets. Browser and installed-package consumers exercise public exports. Format references are the [official ECMA-376 standard page](https://ecma-international.org/publications-and-standards/standards/ecma-376/) and its Part 3 fifth-edition download, clauses 7–9 and Annex A. No native Office round-trip or universal MCE implementation parity is claimed.
