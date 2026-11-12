# Mixed-run native SVG observation

One no-wrap text box with red regular `Alpha` and blue bold `Beta` was generated
with the default XHTML path and with `mixedTextMode: "native-no-wrap"`. MuPDF 1.29.0
through PyMuPDF 1.28.0 omitted all visible text and extracted no text from the
default SVG. It rendered 344 red and 427 blue pixels from the native SVG and
extracted `Alpha Beta`. This observes visible colors and extractable text for one
sample; it does not prove matching fonts, text positions, shaping or broad viewer
compatibility. The SVGs retain text, not glyph outlines, and require viewer fonts.

The [evidence](native-mixed-svg-evidence.json) records hashes of both SVGs, warnings
and MuPDF images. The generator and independent observer can be rerun with:

```sh
rtk proxy node scripts/research/native-mixed-svg-probe.mjs
rtk proxy python3 scripts/research/oracles/native-mixed-svg-probe.py
```

The generated artifacts remain in `.tmp/native-mixed-svg-probe`.
The original evidence predates the expanded native wrapping mode and its revised
approximation-warning text; rerunning with current code produces a different warning
report hash while retaining the historical capture unchanged. Subsequent font
dependency details also change warning reports without changing the sampled SVG.

A second specimen places the same colored runs inside an eight-degree rotated
rectangle. MuPDF displays 352 red and 436 blue pixels and extracts `Alpha Beta`
with a measured text direction of 8.0000004 degrees; the default XHTML output
again contains no extracted text or colored pixels in MuPDF. The [shape evidence](native-mixed-shape-svg-evidence.json)
binds the SVGs, warnings and images; the orientation check uses MuPDF's extracted
line direction and verifies the shape's SVG transform. This does not prove exact
PowerPoint geometry, clipping, font selection, flip behavior, or all shapes.

```sh
rtk proxy node scripts/research/native-mixed-shape-svg-probe.mjs
rtk proxy python3 scripts/research/oracles/native-mixed-svg-probe.py .tmp/native-mixed-shape-svg-probe
rtk proxy python3 scripts/research/oracles/native-mixed-shape-svg-probe.py .tmp/native-mixed-shape-svg-probe
```

The previous [broader MuPDF observation](portable-svg-mupdf.md) remains unchanged.
Tables, wrapping, vertical text, run links and direction-sensitive text are not
converted by `native-no-wrap`. The separate `native` text and table modes cover
additional cases; complete portable SVG acceptance remains open.
