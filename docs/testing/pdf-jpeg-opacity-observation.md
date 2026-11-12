# PDF JPEG opacity observation

The [producer](../../scripts/research/pdf-jpeg-opacity-observation.mjs) renders
one named 8-bit JPEG and two named 12-bit JPEGs with the default browser policy,
`jpegDecodePolicy: "require-opaque"`, `jpegPrecisionPolicy: "require-8-bit"`,
and `jpegStructurePolicy: "require-valid-8-bit"`. Chromium 149.0.7827.55 returned
PDFs for all three under the default; the strict policy returned a PDF for the
8-bit control and rejected both 12-bit inputs with `pdf-resource-unavailable`.
The precision policy also returned the 8-bit control and rejected both 12-bit
inputs, with `pdf-jpeg-unsupported-precision`.
The structure policy also accepted the 8-bit control and rejected both 12-bit
samples. A modified 8-bit source with two identical frame headers was rejected
as `pdf-jpeg-structure-unavailable`; the oracle confirms both headers are present,
while the rejection itself is producer-reported.

In the two default 12-bit PDFs, MuPDF 1.28.0 independently found the original
JPEG bytes as unmodified DCT streams, while each PDF image dictionary declared
`BitsPerComponent` 8. MuPDF reported unsupported 12-bit JPEG precision and the
sampled image area rendered black. The oracle separately reads each source JPEG's
frame-header precision, checks stream hashes and dictionaries, renders the named
pixel, and rejects a PDF whose image dictionary is altered. It does not establish
how other readers display the files. The strict rejection is producer-reported;
the independent reader cannot inspect a PDF that was not returned.

The [manifest](pdf-jpeg-opacity-evidence.json) binds producer, adapter, prefix
probe, data-URL capture and shared structure-check source, fixtures and returned
PDF hashes. Captures are retained in
`.tmp/pdf-jpeg-opacity-20261112` and `.tmp/pdf-jpeg-opacity-repeat-20261112`.

```sh
rtk proxy pnpm build
rtk proxy node scripts/research/pdf-jpeg-opacity-observation.mjs
rtk proxy python3 scripts/research/oracles/pdf-jpeg-opacity.py CAPTURE_DIRECTORY
```

The producer prints its capture directory. This check does not certify colors,
orientation, other JPEG variants, complete PDF validity, other browser versions,
or other PDF readers.
