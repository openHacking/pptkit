# Browser PNG and JPEG output

`renderLayoutToBrowserRaster(layout, options)` is an explicit browser adapter over
the existing resolved Layout and SVG renderer. It returns a promise of
`{ slides, warnings }`, where each slide has its original `slideId`, zero-based
`index`, `hidden` flag, output `width`, `height`, `mimeType` and detached
`Uint8Array` bytes. It requires browser `document`, `Image`, `DOMParser`, Blob URL
and 2D canvas APIs. Importing the package remains safe in Node; calling this adapter
without those APIs rejects with `BrowserRasterError` code
`raster-environment-unavailable`. It does not write files.

```ts
const result = await renderLayoutToBrowserRaster(layout, {
  format: "png",
  scale: 2,
  backgroundMode: "transparent",
  tableMode: "native",
  mixedTextMode: "native",
});
```

`format` is `"png"` or `"jpeg"`; `scale` is positive and finite and defaults to 1.
Each output dimension is `Math.round(cropDimension * scale)`, must be between 1 and
16,384 pixels, and each image is capped at 16 million pixels. `crop` is an optional
positive rectangle `{ x, y, width, height }` in slide coordinates, wholly inside
the slide; omission uses the complete slide. Rasterization maps this rectangle through
the decoded SVG image's intrinsic pixel size, so fractional slide sizes and browser
resampling can introduce rounding differences. `slideIndices` and hidden-slide
behavior match SVG: selected indices retain their original order and IDs; hidden
slides are returned and marked, not silently removed. Empty selection returns no
images.

`backgroundMode: "transparent"` omits the slide background rectangle, leaving
unpainted PNG pixels transparent. Inherited artwork and elements still render.
The default `"slide"` paints the resolved slide background as before. JPEG is
opaque: `jpegBackground` is a `#RRGGBB` matte, default `#FFFFFF`, and `jpegQuality`
is a finite number from 0 through 1, default 0.92. JPEG-only options reject for PNG.
Actual encoding, compression, color handling and interpolation are browser-owned;
no ICC conversion or pixel-identical PowerPoint output is promised.

`maxTotalPixels` defaults to 64 million and cannot exceed that value;
`maxOutputBytes` defaults to 64 MiB and covers cumulative encoded bytes. Both accept
zero for an empty result. Pixel limits are checked before decoding slide images;
the encoded byte limit is checked after each browser encoding. These limits are not
overall CPU or peak-memory bounds: SVG construction, source image decoding, font
programs, browser buffers, and the current encoded Blob are outside them. Existing
`maxSvgBytes` and `embeddedFonts` limits apply independently.

The adapter always collects SVG portability warnings. Missing, rejected or external
image URLs and inline SVG image assets cause `raster-resource-unavailable`; they
are not silently substituted in the returned result. Each emitted inline raster
image is independently decoded with positive intrinsic dimensions before the outer
SVG is decoded. A successful decode still does not prove visible pixels, complete
image validity or correct colors. This is relevant, for example, to the documented
12-bit JPEG case. SVG-as-image support for `foreignObject`, font loading and
complex content varies by browser; warnings remain in the result and are not a
fidelity certification. By default (`fontPolicy: "allow-fallback"`) font
availability and glyph coverage are not checked.

`jpegDecodePolicy: "require-opaque"` additionally checks every unique emitted
`data:image/jpeg`, `image/jpg` or `image/pjpeg` resource after browser decoding.
It draws the decoded source into a temporary canvas and rejects if any decoded pixel
has alpha below 255, since JPEG itself has no alpha channel. This catches the named
Chromium 12-bit failure in which decode succeeds but every pixel is transparent.
The policy checks the declared data-URL MIME, not the source signature or complete
JPEG structure; it does not detect incorrect but opaque colors, blank opaque output,
orientation errors, or differences between decoding the source and rendering it
inside SVG. The default `"browser"` retains prior behavior.

`maxJpegInspectionPixels` defaults to 16 million, has a hard maximum of 64 million,
and counts decoded pixels across unique data URLs in the whole operation. Each
decoded dimension must also be at most 16,384 pixels. It may be zero when no JPEG
inspection is needed, and is valid only with `"require-opaque"`. Inspection scans
64-row bands and rejects budget overflow with `raster-limit-exceeded`; unreadable or
transparent pixels reject with `raster-resource-unavailable`. It can allocate a
source-sized canvas and runs synchronously between abort checks, so it is not a
peak-memory bound or fully preemptible. No color conversion or 12-to-8-bit
transcoding is performed. [Named cross-engine evidence](../testing/jpeg-raster-opacity.md)
records different behavior in Chromium, Firefox and WebKit.
When a slide includes images, `raster-image-frame-unverified` reports that any
animation is sampled by the browser; the adapter does not choose or guarantee a
particular animation frame.

`imageAnimationPolicy: "require-static"` instead requires a synchronous
`isStaticImage(request)` callback for every emitted image use on selected slides.
The frozen request contains the exact `dataUrl`, its declared `mimeType`, `slideId`
and `slideIndex`; it can include the full image payload and the caller owns its
handling. The MIME value comes from the URL and is not proof of the byte format.
The callback must independently verify the exact bytes and return a boolean; a
false result or exception rejects the operation with `raster-image-not-static`,
and a non-boolean (including a promise) throws `TypeError`. It must not mutate
rendering inputs. A callback with the default `"allow"` policy is invalid.
All emitted images on a selected slide are checked, even outside the crop and
when an identical URL appears more than once. Browser decoding still runs after
the assertion. A `raster-image-static-attested` warning makes the trust boundary
explicit: the renderer has **not** independently verified the assertion or extracted
a first frame. An application may supply an independently verified still image
through the asset resolver. Slide animation timelines remain static SVG metadata
and are governed separately by `event-timeline-static`.

`fontPolicy: "require-embedded"` requires every emitted font family and
bold/italic combination on selected slides to match a caller-supplied `embeddedFonts`
face, comparing family names case-insensitively. It then loads each distinct
required face with the browser's `FontFace` API; absent declarations or load
failures reject with `raster-font-unavailable`, identifying the first relevant
slide. Unused supplied faces need not load. The check is conservative for crops:
it includes all emitted text on a selected slide, including text outside the crop.
It creates temporary `FontFace` objects without adding them to `document.fonts`.
Success does **not** verify glyph coverage, font identity, actual selection inside
the SVG image, or visual equivalence. Browser CSP can prevent loading even when the
program itself is valid; the existing `svg-embedded-font-unverified` warning remains.

`fontPolicy: "require-covered"` adds a required synchronous `fontCoverage(request)`
callback to the embedded-face and browser-load checks. For every emitted nonempty
text fragment it receives a frozen request containing `family`, `bold`, `italic`,
`text`, nominal SVG `fontSize`, `slideId`, optional `elementId` and optional
`language`. This includes visible list markers and emitted chart labels; chart
labels have no language value. Text can be split at wrapping, run or selected-font
boundaries; formatting-only hidden spans are not checked. XML-forbidden characters
are replaced before the callback, matching the SVG output. It runs for selected
slides, including text outside a crop. The callback must return a synchronous
boolean: `false` or an exception rejects with `raster-font-unavailable`, while a
non-boolean (including a promise) throws `TypeError`. It must not mutate rendering
inputs. Supplying the callback with another policy rejects as invalid configuration.
It may receive presentation text, so the caller owns its handling.

The callback is a caller-supplied assertion, not an internal font parser. Bind it
to the exact embedded bytes and derive it from a trusted cmap and style inventory;
account for whitespace and controls as appropriate. Per-fragment coverage does not
prove cross-run shaping, variation selectors, browser selection of the embedded
face or absence of glyph fallback. Font programs still must pass `FontFace.load`;
the operation remains atomic if either check fails.

`signal` uses the SVG renderer's cancellation contract. Cancellation can stop
waiting for an asset resolver, image decode, Blob encoding or byte read, and
temporary object URLs are revoked. Browser work already started may finish later.
The complete operation rejects without returning partial slide bytes on cancellation,
resource, decode, encoding or budget failure. `BrowserRasterError.code` identifies
environment, resource, font, static-image assertion, decode, encoding and budget errors; `slideIndex` and
`elementId` are present when their locations are known. Invalid options throw `TypeError` or
`RangeError`; cancellation throws `SvgRenderAbortError`. The output is a static
image and does not preserve editability or interaction.

Browser regression tests cover PNG transparency, JPEG matte, scaling and crop,
hidden-slide order, table and chart pixels, alpha compositing, malformed/remote
images, nested SVG rejection, budgets, strict font declaration/load failures,
font-loading and encoding cancellation and retry, coverage rejection for text,
list markers, tables and charts, JPEG opacity inspection and failure/retry,
static-image assertions against a known two-frame
APNG and a static PNG, plus
no partial result when a later image is corrupt. The selected cases also passed in
[Firefox and Playwright WebKit](../testing/raster-cross-browser.md); the sample
shows small compositing differences. True font coverage, animation timing, color
profiles, other browser versions and independent-viewer acceptance remain open.
