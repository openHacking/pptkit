# Embedded SVG font in MuPDF

One SVG text sample requests a unique family alias and embeds the test-only Roboto
Mono WOFF2 through `embeddedFonts`. A control uses the same SVG content without the
font. MuPDF 1.29.0 through PyMuPDF 1.28.0 reported `Times-Roman` for both, and the
two raster buffers were byte-identical. The font file, SVGs, reports and PNGs are
bound by the [evidence manifest](embedded-font-mupdf-evidence.json). This observation
shows that this MuPDF configuration did not select the supplied face for this
sample; it does not prove all MuPDF fonts, formats or versions behave identically.

```sh
rtk proxy node scripts/research/embedded-font-mupdf-probe.mjs
rtk proxy python3 scripts/research/oracles/embedded-font-mupdf-probe.py
```

The local artifacts are in `.tmp/embedded-font-mupdf-probe`. Separately, a
Chromium test observes a custom face for inline SVG and different isolated-image
pixels when the WOFF2 is omitted. Neither result establishes PowerPoint rendering
or a general cross-viewer embedding guarantee.
