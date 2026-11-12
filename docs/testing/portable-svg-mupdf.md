# Independent SVG rendering observation

One 400 × 260 PPTKit SVG was rendered by Chromium 149.0.7827.55 and independently
by MuPDF 1.29.0 through PyMuPDF 1.28.0 on the same Mac. The specimen contains
native SVG text, mixed-run XHTML text, an XHTML table, an embedded PNG with vertical
crop, and overlapping opaque red and half-opacity blue rectangles. The browser was
prevented from making external requests. The renderer's optional portability
diagnostics identified the two XHTML objects and the three objects using fonts;
the image was inline.

In this observation MuPDF displayed native text but omitted all visible ink in the
mixed-run region and the complete table region. Chromium displayed both. Both
engines rendered the sampled image colors and overlap colors: the three selected
overlap samples agree within two 8-bit channel values and the two crop-image samples
within eight. This is not a complete pixel comparison, a font identity check, or
proof of crop correctness for other images. XML inspection separately checks the
embedded image and its actual transformed geometry. The isolated HTML regions were
otherwise white; a missing font alone cannot explain the missing yellow table fill.

The [manifest](portable-svg-mupdf-evidence.json) binds the generated SVG, warning
report, raster outputs, browser version and empty request log by SHA-256. The
independent [oracle](../../scripts/research/oracles/portable-svg-probe.py) checks
region ink, known colors, geometry and warning identities, and rejects five blanked
controls. These controls alter the browser raster in memory; they are not separate
viewer captures. Raster artifacts remain in `.tmp/portable-svg-probe`.
An independent second run in `.tmp/portable-svg-probe-repeat` produced the same
manifest and file hashes on this machine; this is not cross-machine reproducibility.

Reproduce from the repository root after building the packages, with Playwright
Chromium, PyMuPDF and Pillow installed:

```sh
rtk proxy node scripts/research/portable-svg-probe.mjs
rtk proxy python3 scripts/research/oracles/portable-svg-probe.py
```

The result demonstrates a compatibility failure in this renderer and engine version,
not a claim about all standalone viewers or all features. It does not establish a
portable output mode, embedded font policy, semantic text preservation, independent
GUI viewer behavior, or pixel-equivalent PowerPoint rendering. ECO-04.2.1 remains open.
