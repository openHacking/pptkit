# Use PPTKit Directly in a Browser

PPTKit publishes browser bundles for Core, Layout, PPTX export, and SVG preview. A
static web page can use these bundles to create a presentation, preview it, and
download an editable PPTX without a Node.js runtime or application build step.

Serve the page over HTTP or HTTPS. Browsers may reject module imports and asset fetches
from a page opened with `file://`.

## ESM with an import map

Map all four packages to browser ESM files from the same PPTKit release:

```html
<script type="importmap">
{
  "imports": {
    "@pptkit/core": "https://cdn.jsdelivr.net/npm/@pptkit/core@latest/dist/browser/index.js",
    "@pptkit/layout": "https://cdn.jsdelivr.net/npm/@pptkit/layout@latest/dist/browser/index.js",
    "@pptkit/pptx-exporter": "https://cdn.jsdelivr.net/npm/@pptkit/pptx-exporter@latest/dist/browser/index.js",
    "@pptkit/svg-renderer": "https://cdn.jsdelivr.net/npm/@pptkit/svg-renderer@latest/dist/browser/index.js"
  }
}
</script>
<script type="module">
  import { createPresentation } from "@pptkit/core";
  import { generatePptx } from "@pptkit/pptx-exporter";
  import { renderPresentationToSvg } from "@pptkit/svg-renderer";

  const presentation = createPresentation();
  const slide = presentation.addSlide();
  slide.addElement({
    type: "text",
    content: "Generated in the browser",
    box: { x: 64, y: 64, width: 600, height: 72 }
  });

  const preview = await renderPresentationToSvg(presentation);
  document.querySelector("#preview").innerHTML = preview.slides[0].svg;

  const result = await generatePptx(presentation);
  const blob = new Blob([result.bytes.slice().buffer], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "deck.pptx";
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 0);
</script>
```

The `@latest` URLs make the example easy to try. Pin all four URLs to the same exact
version in production so a new release cannot change one package independently of its
dependencies.

## Global scripts

Pages that do not use modules can load the global bundles in dependency order. Every
bundle extends the same `globalThis.PPTKit` object:

```html
<script src="https://cdn.jsdelivr.net/npm/@pptkit/core@latest/dist/browser/global.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@pptkit/layout@latest/dist/browser/global.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@pptkit/pptx-exporter@latest/dist/browser/global.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@pptkit/svg-renderer@latest/dist/browser/global.js"></script>
<script>
  const presentation = PPTKit.createPresentation();
  presentation.addSlide();

  const previewPromise = PPTKit.renderPresentationToSvg(presentation);
  const exportPromise = PPTKit.generatePptx(presentation);
</script>
```

Core must load before Layout; both must load before the exporter or SVG renderer.

## Runtime and asset constraints

- The browser bundles require modern `fetch`, `Blob`, `URL`, `structuredClone`, and
  typed-array APIs. They do not include Node polyfills.
- The default exporter accepts URL assets, including `data:` and `blob:` URLs. It does
  not accept local filesystem paths.
- Remote images must permit the page origin through CORS. A CDN URL does not bypass an
  image server's CORS policy.
- Keep object URLs alive until PPTX generation or SVG rendering finishes, then revoke
  them to release memory.
- A Content Security Policy must allow the chosen CDN in `script-src` and any remote
  assets in `img-src` or `connect-src`. Blob downloads may require `blob:` in the
  relevant policy.
- Browser generation provides package bytes and structural warnings. It does not
  replace visual verification in PowerPoint or LibreOffice.

The complete runnable source is in the [Browser ESM example](../../examples/browser-esm/README.md).
