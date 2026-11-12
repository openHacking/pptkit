# Browser raster cross-engine observation

On the same Mac, Playwright Chromium 149.0.7827.55, Firefox 151.0 and WebKit 26.5
each passed an initial 11 selected raster tests: output size and crop, PNG transparency, JPEG
matte, image rejection, budgets, cancellation and retry, font declaration/load
and coverage callbacks, selected/inherited artwork, table/chart rendering, and
atomic failure on a later corrupt image. These are browser-engine observations,
not testing of installed Firefox or Safari applications.

Three later static-image assertion tests also passed in each engine. The test
whitelists one exact static PNG data URL, rejects a separate APNG, checks repeated
selected uses, cancellation and retry, and verifies the callback is captured before
an asynchronous resolver completes. The APNG fixture is independently checked as
two frames by the [exporter fixture test](../../packages/pptx-exporter/test/inspect-png.test.mjs).
This tests the assertion boundary; the raster adapter does not itself parse or
classify animation.

An additional [named specimen](raster-cross-browser-v3-evidence.json) compares the
same native SVG text and alpha-overlaid shapes, with and without a supplied WOFF2.
Each engine rejects the absent required font, produces transparent corners, and
changes more than 4,000 pixels when that face is embedded. The sampled overlap is
RGBA `(127, 0, 128, 255)` in Chromium, `(126, 0, 128, 255)` in Firefox and
`(128, 0, 128, 255)` in WebKit. Output PNG hashes differ between engines; no
cross-engine pixel equality is claimed. The runner reports no HTTP(S) requests;
that observation is from its request interception, not independently witnessed by
the PNG oracle. WebKit exposes local `blob:` URLs to Playwright routing, so the
test permits local Blob URLs while blocking HTTP(S).

The [independent PNG oracle](../../scripts/research/oracles/raster-cross-browser.py)
checks producer, font and PNG hashes, sampled pixels, changed/dark pixel counts and
nine in-memory negative controls. The third capture is in `.tmp/raster-cross-browser-v3`;
the earlier [first](raster-cross-browser-evidence.json) and [second](raster-cross-browser-v2-evidence.json)
captures remain under `.tmp/raster-cross-browser` and `.tmp/raster-cross-browser-v2`.
All six PNG hashes matched across these three same-machine captures. The first
manifest predates producer-hash binding.
The changed-pixel control shows that the font input affects this sample; it does
not identify the actual selected face, prove complete glyph coverage, or establish
identical shaping or Office appearance. See also the [MuPDF fallback observation](embedded-font-mupdf.md).

```sh
rtk proxy pnpm test:raster-cross-browser
rtk proxy node scripts/research/raster-cross-browser-observation.mjs
rtk proxy python3 scripts/research/oracles/raster-cross-browser.py CAPTURE_DIRECTORY
```

The observation command prints a fresh capture directory; pass that directory to
the oracle. The optional test requires the corresponding Playwright browser binaries. Full
font coverage, animation timing, color profiles, independent application viewers,
more scripts and versions remain unverified.
