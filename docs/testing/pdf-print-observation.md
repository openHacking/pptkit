# Exploratory Chromium PDF printing

This is a bounded observation, not completion of ECO-04.4. A
Playwright-controlled Chromium 149.0.7827.55 printed two actual PPTKit SVG slides
in reverse selection order, using inline SVG and CSS pages requested at 400 × 200
points. One slide contains mixed-run linked text in `foreignObject`; the other
contains native SVG text. The retained PDF and source HTML are under
`.tmp/pdf-print-20261112`, with hashes in the [manifest](pdf-print-evidence.json).

Independent MuPDF 1.28.0 inspection found two pages in the selected order. It
extracted `Bold linked next` and its `https://example.com/beta` link from the first,
and `Alpha example` with no links from the second. Both measured approximately
400.08 × 199.92 points: the requested dimensions were not exact. The producer
recorded no HTTP requests; this request observation is not independently proven
by the PDF oracle. Text extraction does not establish font identity, correct
shaping, visual fidelity or interactive behavior in other viewers.
An independent repeat in `.tmp/pdf-print-repeat-20261112` passed the same oracle
and negative controls; its PDF bytes differed while its HTML hash matched.

The same producer now also calls the [PDF adapter](../api/pdf-renderer.md).
MuPDF checks its retained `adapter.pdf` independently for both page sizes, order,
extracted text, external URI, two internal destination pages and red/white sample
pixels. Covering the red shape causes the check to fail. The adapter also
reports foreignObject, system-font, font-selection and page-rounding warnings.

The first slide also contains text and shape actions targeting the second slide.
The baseline SVG reports `slide-action-not-linked` for each, and the printed PDF
has no internal links. The producer then renders with
`slideActionMode: "document-links"` and prints `linked.pdf`. MuPDF finds the
external link and two named destinations pointing to page 2. Other shapes and
link types, destinations in other selections and other readers remain unverified.

```sh
rtk proxy pnpm build
rtk proxy node scripts/research/pdf-print-observation.mjs
rtk proxy python3 scripts/research/oracles/pdf-print.py CAPTURE_DIRECTORY
```

The producer prints its capture directory; pass that directory to the oracle.
The oracle uses PyMuPDF, checks the producer and captured-file hashes, page count,
order, text, URI annotations and dimensions with a 0.2-point tolerance, and checks
the modified sample's destination page; it also
rejects swapped pages and a PDF reopened after its link was deleted. The PDF
bytes can vary between runs. Page dimensions, links and selection were checked
only for this named example, not for arbitrary documents. Real printing, other
PDF readers and browser versions, embedded fonts and fallback, script shaping,
accessibility, internal slide links, clipping, transparency, crop, arbitrary page
sizes, network policy and error/cancellation behavior remain unverified. Browser
printing also requires an external controlled browser; it is not a byte-producing
API available to ordinary browser JavaScript.
