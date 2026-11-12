# Chromium PDF output

`@pptkit/pdf-renderer` provides `renderLayoutToPdf(layout, options)` as a Node-side
adapter over the resolved Layout and SVG renderer. Supply an already launched
Playwright-compatible **Chromium** browser; PPTKit neither launches nor closes that
browser. The adapter creates and closes a new browser context per call, blocks
network requests in that context, and returns `{ bytes, pages, warnings }` without
writing files. `bytes` is a detached `Uint8Array`; `pages` identifies each original
slide by ID, zero-based index and hidden flag, in the selected output order.

```ts
import { chromium } from 'playwright';
import { renderLayoutToPdf } from '@pptkit/pdf-renderer';

const browser = await chromium.launch();
try {
  const result = await renderLayoutToPdf(layout, { browser, slideIndices: [1, 0] });
  // The caller chooses whether and where to persist result.bytes.
} finally {
  await browser.close();
}
```

`layout` above is a `LayoutResult` produced by `@pptkit/layout`. The caller owns
the Playwright installation, browser executable, updates and its launch options.
The adapter uses only the documented structural subset of a Playwright browser;
passing another browser implementation is the caller's responsibility. Firefox
and WebKit do not implement the required Playwright PDF operation.

All selected slides have the layout's width and height in PDF points. Dimensions
must be finite and within 1–14,400 points. Chromium can round the requested page
size; it emits `pdf-page-size-rounding` rather than promising exact media boxes.
Selection and hidden-slide behavior match SVG. An empty selection rejects with
`pdf-empty-selection`. `slideIndices`, `idNamespace`, `resolveAsset`, `signal`,
`embeddedFonts`, text/table modes, `backgroundMode` and other SVG options retain
their SVG meanings; slide actions use `document-links` automatically. A selected
target produces a PDF destination when the browser supports it; unselected
targets retain the SVG warning. `backgroundMode: "transparent"` omits the SVG
slide background, but PDF transparency and printer compositing depend on Chromium
and the consumer. There is no PDF crop option.

Portability diagnostics are always enabled. Missing, unsafe, external or nested
SVG image resources reject with `pdf-resource-unavailable`; emitted inline raster
images must also decode with nonzero dimensions in Chromium. Network requests are
blocked, and a restrictive content security policy is installed. A successful
decode does not prove correct colors, frame selection, or actual painted pixels:
in particular some JPEGs can decode successfully yet render transparently. Animated
images are sampled by Chromium without a deterministic frame policy, and any
emitted image produces `pdf-image-frame-unverified`. No remote
fonts are fetched; supplied embedded fonts and system fallback remain
browser-controlled. `pdf-font-selection-unverified` is always returned: it does not
prove font identity, embedding, glyph coverage, shaping or accessibility. SVG
warnings, including `foreignObject` and font-dependency warnings, are retained.

`jpegDecodePolicy: "require-opaque"` adds an opt-in browser-side scan for
declared `image/jpeg`, `image/jpg` and `image/pjpeg` data URLs. JPEG has no alpha;
the adapter rejects decoded transparent or unreadable pixels with
`pdf-resource-unavailable`. `maxJpegInspectionPixels` defaults to 16 million,
cannot exceed 64 million, and counts decoded pixels once per distinct data URL;
each dimension is limited to 16,384 pixels. It is valid only with the strict
policy. Exceeding the budget rejects with `pdf-limit-exceeded`. The check uses a
temporary source-sized canvas and scans in 64-row bands, after browser decoding;
it is not a peak-memory bound or fully preemptible. The default `"browser"`
retains prior behavior.

This check trusts the declared MIME and detects alpha only. It does not validate
JPEG structure, opaque-but-wrong colors, orientation, source validity or all
differences between direct image decoding and PDF printing. In a
[named Chromium/MuPDF observation](../testing/pdf-jpeg-opacity-observation.md),
two 12-bit JPEGs were passed through into PDFs whose image dictionaries declared
8 bits; MuPDF could not decode them. The strict policy rejected those samples.

`jpegPrecisionPolicy: "require-8-bit"` is a separate, opt-in source-header
check. It reads at most the first 65,536 decoded bytes of every distinct emitted
image data URL and identifies JPEG by its SOI bytes, even when its declared image
MIME differs. A discovered JPEG frame precision other than 8 rejects with
`pdf-jpeg-unsupported-precision`; a JPEG without a readable frame header within
the limit, or an image data URL whose prefix cannot be decoded by this bounded
profile, rejects with `pdf-jpeg-precision-unavailable`. The profile accepts
ordinary Base64 or percent-encoded data, but rejects whitespace or percent
escapes inside Base64 and non-ASCII or literal `#` in a raw prefix. Non-JPEG
images with a readable prefix continue to the normal browser decode checks.
The default `"browser"` does not perform this inspection.
The browser context is isolated but caller-owned browser code must not mutate it;
the adapter checks that the image count remains stable, not that every data URL
is unchanged across the separate inventory and decode steps.

The precision policy checks the first frame header, not complete JPEG structure,
entropy data, all later headers or the actual PDF image dictionary. It cannot
certify color, orientation or reader compatibility. Its prefix scan and the
browser's later image decoding have separate memory and cancellation behavior;
neither is a whole-operation resource bound. The precision and opacity policies
may be enabled together. Both reject the two named 12-bit samples, for different
reasons; the opacity policy depends on the browser's decoded alpha whereas the
precision policy inspects source bytes.

`jpegStructurePolicy: "require-valid-8-bit"` applies the shared
[bounded JPEG structure profile](media-inspection.md) to complete source bytes
and rejects supported-profile JPEGs whose precision is not 8. Unlike the prefix
policy it can examine a frame header beyond 65,536 bytes and rejects duplicate
frames, malformed marker sequences and trailing data. Unsupported or invalid
JPEG structure, a declared JPEG without its signature, or a data URL outside the
capture profile rejects with `pdf-jpeg-structure-unavailable`; valid 12-bit
structure rejects with `pdf-jpeg-unsupported-precision`.

`maxStructureInspectionBytes` defaults to 16 MiB and cannot exceed 64 MiB. It
counts decoded bytes once per distinct image data URL, **including non-JPEG
images**, so MIME-spoofed JPEGs are inspected; overflow rejects with
`pdf-limit-exceeded`. Base64 must be canonical and padded; raw data must use
ASCII and valid percent escapes without a literal `#`. The operation copies
bytes and is not a peak-memory or CPU bound. Cancellation is checked during
structure scanning but synchronous work is not fully preemptible. The profile
does not decode entropy, prove pixel correctness, validate all JPEG families or
inspect the PDF writer's final image dictionary. It can be combined with the
other JPEG policies; enabling the prefix policy too retains its shorter bound.

`imageAnimationPolicy: "require-static"` requires an `isStaticImage(request)`
callback that returns a synchronous boolean for every emitted image use on the
selected pages, including repeated uses. The frozen request carries the exact
`dataUrl`, declared `mimeType`, `slideId` and original zero-based `slideIndex`;
the caller is responsible for independently inspecting the bytes and handling
their potentially sensitive content. The callback must not mutate the rendering
inputs or browser context. A false result or callback exception rejects
with `pdf-image-not-static`; a promise or other non-boolean throws `TypeError`.
Supplying the callback with the default `"allow"` policy is invalid. Successful
attestation adds `pdf-image-static-attested`, but the adapter does not verify the
assertion itself; the image-frame/pixel warning remains. Slide animation timelines
are separate and remain static SVG metadata. See the [named PNG/APNG observation](../testing/pdf-static-image-observation.md).

`imageAnimationPolicy: "require-structurally-static"` instead uses the shared
[static image structure profile](media-inspection.md) on each distinct emitted
data URL. It accepts only supported still PNG, single-frame GIF and standalone
JPEG structures; declared APNG or multi-frame GIF rejects with
`pdf-image-not-static`, while unknown, ambiguous, invalid or unsupported input
rejects with `pdf-image-static-unverified`. MIME declarations do not decide the
format. The callback is not used and must not be supplied in this mode. The
adapter retains `pdf-image-frame-unverified` and adds
`pdf-image-static-structure-checked`: structure does not prove pixel decoding,
reader fidelity or behavior of every consumer-specific extension. A structurally
static 12-bit JPEG may still be unsuitable for PDF; use a JPEG policy separately.

`maxStaticInspectionBytes` defaults to 16 MiB, has a 64 MiB ceiling, and counts
decoded source bytes once per unique image URL. It is valid only with this mode;
zero is useful for a selection with no images. Data-URL capture has the same
canonical Base64 and raw URL restrictions as the full JPEG structure policy.
When both policies are enabled, their independent byte limits each apply. The
operation can make multiple bounded copies and is not a peak-memory or CPU bound.

`fontPolicy: "require-embedded"` requires a caller-supplied `embeddedFonts` face
for each emitted family and bold/italic combination on selected slides, with
case-insensitive family matching. Chromium must then load each requested face
from the emitted document; a missing declaration or load failure rejects with
`pdf-font-unavailable`. The default `"allow-fallback"` retains browser fallback.
The check includes all emitted text on selected pages, including off-page text;
unused supplied faces need not load. It checks declarations and loading, not glyph
coverage, actual selection for every run, shaping, or the resulting PDF font
program. Declared weight and style are not checked against font metadata and do
not rule out synthetic styling. The supplied font array is captured before asynchronous rendering.
[A named native-SVG and XHTML sample](../testing/pdf-font-observation.md) verifies
the embedded face in the produced PDF using MuPDF; other fonts and scripts are
not covered by that observation.

`maxSvgBytes` defaults to 64 MiB for this adapter. `maxOutputBytes` defaults to
64 MiB and rejects with `pdf-limit-exceeded` after Chromium has produced the PDF.
These bounds do not cap peak browser memory, source decoding or CPU. The adapter
checks the PDF signature, not the full PDF structure; use an independent reader
where output verification matters. Abort signals cancel waits and close a created
context; context creation itself is awaited so that a late-created context can be
closed. Browser work may continue until context closure; callers should set their
own deadline through an abort signal. `pdf-print-failed` covers browser and context-close failures. SVG errors
and `SvgRenderAbortError` retain their original contracts.

Text selection, extraction and link annotations depend on SVG and Chromium
printing. A named two-page text, external-link and internal-link example was
checked with MuPDF; see [PDF evidence](../testing/pdf-print-observation.md).
Nested links, other browsers, independent viewers, complex scripts, font identity,
animation timing, actual printing and general PowerPoint visual parity are not
established by that sample. The PDF output is not editable presentation content.
