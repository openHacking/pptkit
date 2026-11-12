# Text style inheritance in content import

Experimental content import resolves a bounded subset of DrawingML text frame, paragraph, list and run properties. It writes resolved values into ordinary Core text values and records the selected source declaration in `ImportReport.textStyleInheritance`. Core does not retain source XML, master relationships or a live template binding.

```ts
import { importPptx, type PptxTextStylePropertyInfo } from "@pptkit/pptx-parser";

export async function inspectTextStyles(bytes: Uint8Array) {
  const session = await importPptx(bytes, {
    projection: "content",
    sessionLimits: { maxTextStyleProperties: 500_000 },
  });
  try {
    const properties: readonly PptxTextStylePropertyInfo[] =
      session.report.textStyleInheritance?.properties ?? [];
    return properties.map(({ property, level, status, value, source, target, resolvedFrom, reason }) => ({
      property, level, status, value, source, target, resolvedFrom, reason,
    }));
  } finally {
    session.dispose();
  }
}
```

The complete checked example is [text-style-inheritance.ts](../examples/text-style-inheritance.ts).

## Cascade profile

Text frame margins, vertical alignment, wrapping and auto-fit are resolved independently from text character/paragraph styles. For a supported slide drawing, frame values cascade from matching master placeholder to matching layout placeholder to the local body properties. Local explicit values, including zero and false, win for that property.

Character and paragraph styles cascade per property. The selected local text paragraph level is applied exactly; the parser does not borrow a missing value from a different level:

1. Master named text style selected from the placeholder family; plain shapes use `otherStyle`.
2. Presentation `defaultTextStyle` for plain shapes.
3. Matching master placeholder list style.
4. Matching layout placeholder list style.
5. Local `txBody/lstStyle` defaults and the selected `lvl1pPr`…`lvl9pPr`.
6. Paragraph `pPr/defRPr`, direct paragraph properties, and direct run properties.

An explicit unsupported higher declaration blocks an older inherited value for that same property. It never silently falls through to a lower source. Supported explicit clearing values such as bold false, no underline, no bullet, zero spacing, zero margin and wrap none are retained.

Ancestor prompt paragraphs and runs are not style parents. Only their selected style declarations participate; prompt text and artwork are not copied into the slide. Notes and table cells use the shared local list/property parser, but do not receive slide/layout/master text-style cascade or `textStyleInheritance` target records in this profile.

## Supported values and limits

The profile maps Core-supported frame, paragraph and run values: four margins, vertical alignment, wrapping, auto-fit, alignment, paragraph base direction, indentation, line/paragraph spacing, character bullets, four automatic numbering styles with start value, font family, size, bold, italic, single underline/strike, direct sRGB/alpha, language and run-level RTL. Theme colors resolve through source color-map overrides and the theme color scheme; direct RGB and system-color fallback slots are supported. The [text-color transform profile](text-color-transforms.md) evaluates ordered tint/shade, HSL luminance/saturation and alpha operations into local RGB and `colorOpacity`, with transform provenance. Theme overrides and major/minor Latin font slots resolve to concrete local color/font values. Different source masters may therefore resolve the same slot to different values without collapsing their themes into Core's single document theme.

Explicit paragraph tab lists map as a whole when they contain at most 256 strictly increasing nonnegative positions and explicit supported alignments. Empty lists remain explicit; missing position/alignment and unsupported or oversized lists block older declarations. Each mapped stop contributes a source dependency to the existing provenance budget. This does not interpret default tab intervals or establish native visual equivalence; see [text and styles](core/text-and-styles.md).

The report's `resolvedFrom` locations identify the color-map/theme palette or theme font used by a selected declaration. A slide color-map override applies to its text, including inherited layout/master declarations. A slide with no color-map override uses its matching layout override before the master map; an explicit master mapping delegates directly to the master.

Theme color and font schemes are selected independently in the target slide's verified relationship chain: slide override, layout override, master override, then the master theme. This order also applies to a local slide run whose theme declaration does not occur in the layout. An override with only a font scheme inherits the color scheme; one with only a color scheme inherits the font scheme. A present scheme replaces that component as a unit: an unsupported or missing slot blocks older values instead of merging individual palette/font slots. This distinction follows the optional scheme components in the [Open XML SDK ThemeOverride schema](https://github.com/dotnet/Open-XML-SDK/blob/main/data/schemas/schemas_openxmlformats_org_drawingml_2006_main.json). Theme and override targets require their matching content types and dialect roots; external or duplicate override relationships are rejected, never fetched.

`pPr/@rtl` maps to paragraph `direction`; explicit false values override inherited RTL. The selected false presentation-level default is retained in provenance but omitted from the projected style after cascade resolution, preserving the implicit LTR baseline. The `a:rtl` child of run properties maps independently to optional `rightToLeft`, including omitted `val` as true and explicit false overrides. Neither mapping adds direction style writes to source-preserving text-run profiles. See [run direction evidence](../testing/bidi-first-open-observation.md#run-level-direction-is-a-separate-source-feature) for the preview and font limitations.

Explicit unknown attributes, unsupported encodings and unsupported color/font choices are source-located findings. The selected property report records mapped values or an unsupported blocker with its source, target, cascade level and any resolution dependencies.

Color modifiers outside the named transform profile, automatic script/font selection, complex bullets, application-default tabs, advanced line spacing, vertical writing and unmodeled effects remain source-only. Explicit major/minor East Asian and complex-script theme slots now resolve when their concrete families agree with the explicit Latin family; absent/empty slots and different families remain blockers. The default source profile projects one font family per run, so differing script-specific fonts block that projection rather than silently applying the Latin font. Explicit [script-families-v1](script-font-import.md) now maps Core's three family fields with independent cascade, blockers and provenance; the default blocker remains compatible. This profile does not certify Office visual equivalence.

`maxTextStyleProperties` defaults to **500,000** selected property records. `maxTextStyleDependencies` separately defaults to **1,000,000** `resolvedFrom` locations across those records, bounding shared-transform provenance amplification. Shared source lists are parsed and audited once; target records are bounded separately. Existing XML/archive/report budgets and cooperative cancellation also apply. Reports are detached, deeply frozen snapshots; structure-only reports omit `textStyleInheritance`.

Fresh export materializes the imported effective values as local text styling and does not rebuild the source template chain. Reimport therefore sees local declarations instead of the original inherited levels. Original-byte no-op save retains the complete unchanged source. Source style/override writes and rebuild stay blocked; the separate [text-run profile](text-edit-save.md) can replace bound characters while retaining the original style/template chain. Full theme resolution, template rebinding and style-preservation writes remain under IMP-05 and IMP-09.

## Verification

- [Parser fixtures](../../packages/pptx-parser/test/text-inheritance.test.mjs) cover both dialects, all nine paragraph levels, per-property precedence and clearing, unsupported blockers, named master styles, ambiguity, multiple masters, source locations, notes/table local parsing, budgets and cancellation.
- [Partial theme override fixtures](../../packages/pptx-parser/test/theme-overrides.test.mjs) cover slide/layout/master component inheritance, direct and inherited declarations, unsupported higher components, MCE, invalid relationships/types/roots and session isolation.
- [Cross-package parity tests](../../scripts/text-inheritance-parity.test.mjs) cover three fresh rounds, independent exported slide XML, stable package size, no prompt duplication, local edits, no-op bytes and blocked source saves.
- [Effective-color integration checks](../../scripts/import-content-parity.test.mjs) independently check source palette/map XML and concrete imported colors. Fresh-export comparisons accept theme references and equal RGB values, but still detect wrong colors and unrelated style changes. Exact-byte no-op checks separately retain source-theme identity.
- [Installed public consumer checks](../../tests/consumers/node.mjs) exercise runtime reports; [declaration checks](../../tests/consumers/types.ts) check the exported readonly contract.

The optional [native theme experiment](../../scripts/inspect-theme-text-import.mjs) uses a normal exported deck with a partial layout font override and a partial slide color override. Its [evidence](../testing/theme-text-import-evidence.json) records source/projection hashes, exact no-op, a LibreOfficeDev save and independently extracted PDF text/color/font spans. The tested development build renders the source using the base color/font and writes different direct styles during save; the fresh materialized projection renders the selected override values. These observed differences remain a compatibility limitation, not a native-fidelity pass. Reproduce after building with `node scripts/inspect-theme-text-import.mjs [soffice-executable]`; Python/PyMuPDF is required for the independent PDF inspection.

PowerPoint text layout and template interaction remain unverified. The evidence here does not close IMP-05.1, AUT-07 or IMP-09.
