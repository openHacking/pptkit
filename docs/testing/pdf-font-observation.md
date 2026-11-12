# Chromium PDF embedded-font observation

The [producer](../../scripts/research/pdf-font-observation.mjs) renders two pages:
one native SVG text run and one linked XHTML `foreignObject` text run. Both use a
caller-defined family alias. The plain run supplies no font; the strict run embeds
the repository's Roboto Mono WOFF2 fixture and requests `fontPolicy:
"require-embedded"`. A separate signature-correct but invalid WOFF2 input is
rejected with `pdf-font-unavailable`.

On Chromium 149.0.7827.55 and MuPDF 1.28.0, both plain pages extracted as
Times-Roman; both strict pages extracted as RobotoMono-Regular and contained an
embedded PDF font program. Text and the XHTML page's external link remained
extractable. The first run's text width changed from approximately 282.60 to
323.98 points, and the second from 200.98 to 259.19 points. This is one named
fallback environment and one Latin font fixture, not a universal fallback rule or
proof of glyph coverage, shaping or source-font equivalence.

A separate strict-mode page uses the same fixture for `ABC 汉字`.
The operation succeeds and MuPDF extracts the whole text, but reports Roboto Mono
only for `ABC ` and separate Type3 fonts for the two Chinese characters. The
`pdf-font-selection-unverified` warning remains present. This concrete fallback
is why a successful face-load check must not be treated as glyph-coverage proof.

The [manifest](pdf-font-evidence.json) binds the producer, adapter source, fixture
and actual PDF hashes. Captures are retained in `.tmp/pdf-font-20261112`. The
independent [MuPDF oracle](../../scripts/research/oracles/pdf-font.py) checks both
pages' extracted text, font names, embedded font bytes, text widths and the link;
it rejects swapped pages and a PDF reopened after its link was deleted. The
browser's invalid-font rejection is producer-reported.
An independent repeat is retained in `.tmp/pdf-font-repeat-20261112` and passed
the same checks and negative controls.

```sh
rtk proxy pnpm build
rtk proxy node scripts/research/pdf-font-observation.mjs
rtk proxy python3 scripts/research/oracles/pdf-font.py CAPTURE_DIRECTORY
```

The producer prints its capture directory. PDFs and font-subset names can differ
between browser versions or hosts. Exact font-program identity, scripts beyond
this fixture, raster comparison, PDF readers other than MuPDF and physical printing
remain unverified.
