# `@pptkit/svg-renderer`

Browser-oriented SVG preview rendering for PPTKit presentations. The renderer consumes
the public Core and Layout contracts and returns one standalone SVG string per slide.
It also exposes an explicit browser-only PNG/JPEG adapter over resolved Layout; see
[browser raster output](../../docs/api/browser-raster.md) for resource checks and limitations.

For detached slide-space selection frames and rectangular hit testing, see the
[selection contract](../../docs/api/svg-selection.md).

```ts
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

const result = await renderPresentationToSvg(presentation);
preview.innerHTML = result.slides[0]?.svg ?? "";
```

The output uses SVG geometry, images, and native `<text>/<tspan>` layout for simple
text. Bar, line, and pie charts render as native SVG geometry. XHTML `foreignObject` remains a fallback for mixed-run rich text and tables. It
targets modern browsers and is intended for QA previews, not as a
pixel-identical PowerPoint renderer or a portable SVG interchange format. URL assets
work by default; local path assets require a custom `SvgAssetResolver`.
Wrapping text containing right-to-left or Arabic-number bidi classes, or
directional marks and controls, also uses XHTML so the browser can resolve
direction across the whole paragraph. Routing uses fixed macOS ICU Unicode 17
data, including some conservative private-use values; browser Unicode data and
font support may differ.
Line breaking and bidi rendering can still vary between browsers and PowerPoint.
Native SVG text lines retain soft-wrap spaces and hard-break characters in the DOM,
including the bounded explicit-tab path.
Browser text selection may add or omit line breaks and is not a source-text
round-trip interface. List markers and explicit tab layouts also add generated
text; paragraph boundaries are not represented as DOM text separators.

Run styles with `rightToLeft: true` request DrawingML selective run-level direction and complex-script font selection. SVG preview reports `text-run-rtl-approximation`; the built-in browser text measurer rejects these runs because browser layout cannot reproduce the selective override. Explicit `false` does not trigger this limitation.
