# `@pptkit/svg-renderer`

Browser-oriented SVG preview rendering for PPTKit presentations. The renderer consumes
the public Core and Layout contracts and returns one standalone SVG string per slide.

```ts
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

const result = await renderPresentationToSvg(presentation);
preview.innerHTML = result.slides[0]?.svg ?? "";
```

The output uses SVG geometry, images, and native `<text>/<tspan>` layout for simple
text. XHTML `foreignObject` remains a fallback for mixed-run rich text and tables. It
targets modern browsers and is intended for QA previews, not as a
pixel-identical PowerPoint renderer or a portable SVG interchange format. URL assets
work by default; local path assets require a custom `SvgAssetResolver`.
