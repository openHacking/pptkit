# JPEG raster opacity observation

Three browser engines rendered a named 8-bit JPEG and two named 12-bit JPEGs into
6 × 4 transparent PNG outputs, with default decoding and with
`jpegDecodePolicy: "require-opaque"`. On Chromium 149.0.7827.55, the 8-bit sample
was opaque, but both 12-bit files returned successful default output with all 24
pixels transparent. The opt-in check rejected both. Firefox 151.0 rejected the two
12-bit files at decoding under both policies. Playwright WebKit 26.5 rendered
opaque pixels for all three files under both policies. Opacity alone does not prove
correct colors, orientation or correspondence with PowerPoint.

The [manifest](jpeg-raster-opacity-evidence.json) binds the three fixture bytes,
producer, per-engine results and saved PNG hashes. The independent
[Pillow oracle](../../scripts/research/oracles/jpeg-raster-opacity.py) checks the
output alpha counts and one altered-pixel negative control per retained raster.
The runner's no-HTTP request record is not independently established by the PNG
oracle. The retained capture is `.tmp/jpeg-raster-opacity-20261111`.

```sh
rtk proxy node scripts/research/jpeg-raster-opacity-observation.mjs
rtk proxy python3 scripts/research/oracles/jpeg-raster-opacity.py CAPTURE_DIRECTORY
```

The producer prints a fresh capture directory; pass it to the oracle. The 8-bit
fixture has EXIF orientation metadata, but this experiment checks only opacity,
not orientation. Browser versions, other JPEG families, opaque incorrect pixels,
color profiles, source validity and PowerPoint fidelity remain outside this result.
