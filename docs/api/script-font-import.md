# Explicit source script-font projection

Content import can opt into `textFontProfile: "script-families-v1"`. It projects supported Latin, East Asian and complex-script declarations independently into Core's `fontFamily`, `eastAsianFontFamily` and `complexScriptFontFamily`. Omitting the option retains the existing single-family projection, including its unequal-script-font blocker and report shape.

See the [typed example](../examples/script-font-import.ts), [authoring contract](script-fonts.md), [text inheritance](text-style-inheritance.md), [table style projection](table-style-import.md) and [ADR0047](../architecture/decisions/0047-source-script-font-profile.md). `ImportReport.textFontProfile` records the explicit profile. Invalid values and structure projection with this option reject before package/file work. Byte and Node file imports capture it before asynchronous work; candidate validation retains that captured value.

## Independent declarations and provenance

| Source declaration | Selected property |
| --- | --- |
| Run/default-run `a:latin/@typeface` | `run.fontFamily` |
| Run/default-run `a:ea/@typeface` | `run.eastAsianFontFamily` |
| Run/default-run `a:cs/@typeface` | `run.complexScriptFontFamily` |
| Table style `font` collection | Its three independent named slots |
| Table style `fontRef/@idx` major/minor | Explicit lt/ea/cs slots in the target font component |
| Direct notes-body font declarations | Same three properties with notes-part/source-object provenance |

Source `+mj/+mn-lt/ea/cs` references resolve against the verified target slide/layout/master font component. Supported direct or theme families materialize as local strings. Selected declaration paths and actual theme dependencies remain format-owned report data; no source theme identity enters Core. A missing/empty slot in a present nearer font scheme blocks an older scheme and records that selected scheme or slot dependency. No unsupported slot is filled from a different master.

Each property cascades independently through the existing source text levels. For example, a master supplying Latin/East/Complex, a slide list overriding East, a paragraph overriding Complex and a run overriding Latin produce the three independently selected values and sources. A local Latin declaration does not clear inherited East/Complex values.

An explicit empty, missing-required, invalid or unsupported typeface creates an unsupported blocker for that property. It does not uncover a lower script value. The projection omits that unsupported Core slot and retains its diagnostic; ordinary Core/default-family fallback is a projection value, not a successful source inheritance result. Empty source fonts do not mean the authoring API's null reset, and language/host substitution is not guessed. Duplicate interpreted font slots reject. MCE selection precedes effective slot interpretation; report paths describe that effective tree.

An explicit table font collection replaces its collection choice: missing required slots block rather than borrowing older collection slots. A table font reference resolves each named slot separately. Supplemental script-font rules remain outside this explicit-slot profile: an explicit collection with nonuniform/unresolved supplemental entries blocks its font projection. Uniform supplemental declarations that agree with all three explicit families may retain dependency evidence. This is not automatic language/Han variant or arbitrary script-map selection.

Referenced theme collections with supplemental entries emit located source-only findings while the named-slot projection remains explicit. They do not silently acquire language-specific selection. Those findings are emitted during bounded traversal and remain separate from a mapped literal lt/ea/cs value. The native collection distinction is documented by Microsoft's [MinorFont](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.minorfont?view=openxml-3.0.1) and [SupplementalFont](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.supplementalfont?view=openxml-3.0.1) descriptions.

## Notes, scope and resource boundaries

Direct notes-body font slots now receive selected properties in `textStyleInheritance`. A notes list style can use the `notes-direct` source level; direct runs retain their run level. Notes-part/object provenance distinguishes them from drawings. Notes master/theme inheritance remains unimplemented: notes theme references report unsupported and never borrow the slide theme. Source bindings and note ownership remain independent of style records.

Inherited declaration parsing/audits are isolated by target theme scope in this profile. One source node interpreted under a missing font component on one slide cannot leave audit notes on another valid slide. Shared table definitions retain their existing scope isolation.

Drawing and notes properties/dependencies share the existing session-wide `maxTextStyleProperties` / `maxTextStyleDependencies` limits. Table style properties/dependencies/work use their existing cumulative limits. General source/XML/path/MCE/report limits and cancellation continue to apply; there is no new unbounded registry or native font query. Limits charge actual selected records and retained dependency locations, not total process memory. A failed/aborted import returns no usable partial session; independent retry remains possible.

## Saving and output

The option selects source projection, not edit permission. No-op retains original bytes. Existing explicit text-run profiles can still admit bound scalar character changes; V3 covers admitted text boxes, shape text, direct notes and table anchors. The saved candidate is reprojected with the captured font profile and compared to the full approved Core snapshot. Original font declarations, styles, themes, relationships and unmodified payloads are retained.

Changing any font family, adding/removing a script field or mixing that with text edits remains atomically blocked. Restoring the baseline permits retry. New font fields are not written back from materialized Core values. Fresh output is a separate operation: it can serialize the projected families into native slots but cannot reconstruct original theme/collection abstractions.

## Evidence and open gates

[Fourteen Parser groups](../../packages/pptx-parser/test/script-font-import.test.mjs) cover both dialects, default compatibility, fieldwise inherited/local overrides, unsupported blockers, explicit theme components and missing-slot dependencies, tables/notes, MCE/duplicates, two-master and shared-source scope isolation, exact budgets, mixed-owner three-round scalar saves, font mutation/retry, byte/file capture and cancellation. [Three cross-package/browser groups](../../scripts/script-font-import-parity.test.mjs) add Node/ESM/global agreement, independent source XPath/theme-slot DOM checks, three mixed-owner fresh rounds and capture/failure retries. Installed runtime/types consumers use public exports and a real prepared scalar save.

Native font selection, actual installed-font coverage, language/Han/supplemental maps, complete notes/theme inheritance, shaping/bidi/vertical text and PowerPoint interactive editing remain independent gates. This profile does not close the complete script-font or source-style-write tasks.

An [independent Python/native experiment](../testing/script-font-import.md) verifies 18 source font properties and six scalar leaves in each original/three-save package, exact non-character byte retention and rejection of a deliberately wrong font. Native output remains observational: run splitting, synthetic-font substitution and invisible CJK despite extractable text are retained compatibility gaps.

The separate [v2 supplemental profile](supplemental-font-import.md) now transports admitted collection script rules after effective slot inheritance. All v1 behavior above remains unchanged; native font/shaping and complete source semantics remain separately gated.

The independent [notes text profile](notes-style-import.md) now enables bounded notes-master/theme style resolution for default/v1/v2 font modes. Omitting it preserves the direct-notes boundary above.
