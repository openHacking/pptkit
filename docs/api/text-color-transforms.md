# Text color transforms and alpha

Experimental content import evaluates the following DrawingML text-color profile for direct `srgbClr`, system `sysClr` with `lastClr`, and resolved `schemeClr` values. Theme palette entries and partial theme overrides may also contain these transforms. Unknown color kinds or transforms remain explicit blockers; this is not a complete DrawingML paint engine.

| Transform | Evaluation |
| --- | --- |
| `tint` | Mix linear RGB with white: `1 - (1 - channel) × amount`. Zero gives white; one preserves the input. |
| `shade` | Multiply linear RGB by amount. Zero gives black; one preserves the input. |
| `lum`, `lumMod`, `lumOff` | Set, multiply or offset HSL lightness. |
| `sat`, `satMod`, `satOff` | Set, multiply or offset HSL saturation. |
| `alpha`, `alphaMod`, `alphaOff` | Set, multiply or offset alpha independently from RGB. |

Transforms run in source order, first on the selected palette value and then on the color reference. RGB stays floating point until final 8-bit quantization. Tint/shade use the IEC sRGB transfer function to enter/leave linear scRGB; HSL operates on encoded sRGB. Consecutive HSL operations retain their intermediate hue/saturation even when the rendered color becomes gray or black. Every component is clamped after its operation. Alpha does not force a color-space conversion.

Values accept signed integer thousandths of a percent and decimal percentage strings, with XML whitespace at their boundaries. Fixed tint/shade/alpha values are in `0..100%`, alpha offsets in `-100..100%`, and alpha modulation is nonnegative. General HSL percentage types may be signed or exceed 100%, with the resulting component clamped. Integer encodings must fit the signed 32-bit schema value. Missing/invalid values, unknown attributes, foreign transforms, nested transforms and non-whitespace leaf text block the whole declaration; no partially transformed RGB or alpha is returned.

The element descriptions and percentage types follow the Open XML SDK's ISO/IEC excerpts for [Tint](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tint), [Shade](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shade), [LuminanceModulation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.luminancemodulation), [SaturationOffset](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.saturationoffset) and [AlphaOffset](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphaoffset). The color-space/precision rules above specify PPTKit's evaluated profile; native evidence is recorded separately below.

## Core and output contract

`TextRunStyleInput.colorOpacity` is an optional finite number in `0..1`. It inherits like other run style fields; explicit `1` clears inherited translucency and `0` retains transparent text without removing its text, geometry or actions. It works in text elements, shape text, presets/placeholders, table cells and notes. It does not affect shape fills, image paint or other runs.

Normalized IR v2 carries optional `colorOpacity`; absence means `1`. Normalization omits opaque alpha so existing opaque snapshots retain their shape. Older normalized input without the field remains valid. Layout/measurement does not use alpha to remove or resize text.

<!-- doc-test: docs/examples/text-color-transforms.ts -->
```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function createTranslucentText() {
  const document = createPresentation();
  document.addSlide().addElement({
    type: "text",
    box: { x: 40, y: 40, width: 320, height: 80 },
    opacity: 0.5,
    content: [{
      runs: [
        { text: "Quarter opacity", style: { color: "336699", colorOpacity: 0.5 } },
        { text: "Element opacity", style: { color: "336699", colorOpacity: 1 } },
      ],
    }],
  });
  return {
    normalized: normalizePresentation(document),
    output: await generatePptx(document),
  };
}
```

PPTX output multiplies text paint alpha by the existing element opacity and emits DrawingML `alpha` at integer `1/100000` precision. It materializes the effective color/alpha; fresh output does not reproduce the original transform expression. Native SVG uses `fill-opacity`; mixed-run HTML uses per-run CSS opacity, under the element's existing opacity. Transparent text still occupies its normal layout area. Notes with empty but explicitly translucent runs remain authored content.

## Import provenance and limits

A supported fill supplies both `run.color` and `run.colorOpacity`. An explicit opaque fill resets lower-level alpha. An unsupported fill blocks both properties so neither an old RGB nor old alpha leaks into the projected style.

`resolvedFrom` records the palette/map and source transform locations, in addition to each property's declaration and target. Reports remain detached and frozen. `maxTextStyleProperties` bounds the number of selected records; `maxTextStyleDependencies`, default **1,000,000**, independently bounds the sum of their dependency locations, including repeated uses of shared source transforms. Notes/table local parsing retains its existing source findings and does not add shape ancestry records. Existing XML/archive/report limits still apply.

Original-byte no-op save retains all transform expressions and unknown content. Source style/transform edits and rebuild remain blocked. The separate opt-in [text-run profile](text-edit-save.md) can replace existing run characters while retaining unchanged source style expressions; these color-projection/fresh-output tests alone do not establish preservation-aware editing.

## Evidence and remaining limits

- [Focused transforms/import tests](../../packages/pptx-parser/test/color-transforms.test.mjs): independent HSL constants, linear mixing boundaries, transform order, gray/black intermediates, signed/percent lexical forms, schema/unknown blockers, 256 gray identities, both dialects, theme overrides, source locations, exact no-op and dependency budgets.
- [Core tests](../../packages/core/test/text-opacity.test.mjs): immutable authoring state, preset inheritance/reset, shape/cell/notes paths, zero/default alpha, validation and old IR shape.
- [Cross-package tests](../../scripts/text-inheritance-parity.test.mjs): independent slide/notes XML, alpha composition, all local text owners and three fresh export/import rounds without cumulative color transformation or resource growth.
- [Chromium tests](../../packages/svg-renderer/test/measurement.spec.mjs): independently sampled native SVG pixels, combined element/run alpha, mixed-run computed styles and unchanged text widths.
- [Installed consumers](../../tests/consumers/node.mjs) and [declaration consumers](../../tests/consumers/types.ts): public alpha authoring, preview, import and exported types.

The optional [native runner](../../scripts/inspect-color-transform-import.mjs) and its [evidence](../testing/color-transform-import-evidence.json) cover nine named samples under the recorded LibreOfficeDev build. Eight preserve the selected style through native save and produce identical source/fresh PDF rasters. The gray shade sample differs: source/native yields `5E5E5E`, while PPTKit's IEC evaluation yields `5C5C5C`. The runner records this discrepancy, source/output hashes and independent PyMuPDF spans/raster hashes; it does not change expected Core values to make the application match. PDF text-span alpha alone is not a paint-opacity oracle; native saved styles and raster comparisons are separate evidence.

Reproduce after building with `node scripts/inspect-color-transform-import.mjs [soffice-executable]`. It requires Python/PyMuPDF and uses a fresh application profile/output directory. This evidence does not establish interactive editing, PowerPoint compatibility, arbitrary theme inheritance, complete transforms or preservation writes. IMP-05.1.3.3.1 and its parent gates remain open.
