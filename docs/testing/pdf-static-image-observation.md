# PDF static-image attestation observation

The [producer](../../scripts/research/pdf-static-image-observation.mjs) uses a
static PNG and a known two-frame APNG. With `require-static`, its callback accepts
only the exact PNG data URL: two selected uses in reverse page order are accepted,
and the APNG is rejected with `pdf-image-not-static`. As a deliberate trust-boundary
control, a callback that always returns true also accepts the APNG. The resulting
PDF carries both `pdf-image-static-attested` and `pdf-image-frame-unverified`;
the adapter has not verified that assertion or fixed an animation frame.

The producer also runs `require-structurally-static`: it accepts the named PNG
and a one-frame GIF, and rejects the named APNG and two-frame GIF. Pillow
independently counts both PNG/APNG and GIF input pairs as one and two frames; MuPDF opens the
accepted GIF PDF and checks the PNG output pixels. The rejection results are
producer-reported. Other formats, ambiguous GIF controls and private PNG chunks
are tested as unverified in package tests, not claimed as animation.

The [manifest](pdf-static-image-evidence.json) binds fixture, producer, adapter,
shared inspection source and PDF hashes. On Chromium 149.0.7827.55 and MuPDF 1.28.0, the independent
[oracle](../../scripts/research/oracles/pdf-static-image.py) checks two static PDF
pages, their red and white sample pixels, the exact callback data-URL hashes and
both output warning sets. A PDF with the red area covered is rejected. Callback
order and APNG rejection are producer-reported; neither PDF pixel samples nor the
oracle establish animation timing or classify arbitrary image bytes.

Captures are retained in `.tmp/pdf-static-image-20261112` and
`.tmp/pdf-static-image-repeat-20261112`. Reproduce with:

```sh
rtk proxy pnpm build
rtk proxy node scripts/research/pdf-static-image-observation.mjs
rtk proxy python3 scripts/research/oracles/pdf-static-image.py CAPTURE_DIRECTORY
```
