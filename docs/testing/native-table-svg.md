# Native SVG table observation

A two-row, two-column PPTKit table has a merged yellow header and separate cyan and
pink body cells. The same document was rendered with default XHTML and with
`tableMode: "native"`, then opened by MuPDF 1.29.0 through PyMuPDF 1.28.0. In this
one sample, the default SVG yielded white at three cell samples and no extracted
text. The native SVG yielded the three authored colors and extracted `Header`,
`Left` and `Right`. The independent observer also checks merged-cell role metadata,
cell clip references and expected warning codes. The [manifest](native-table-svg-evidence.json)
binds both SVGs, warnings and MuPDF raster outputs by SHA-256.

```sh
rtk proxy node scripts/research/native-table-svg-probe.mjs
rtk proxy python3 scripts/research/oracles/native-table-svg-probe.py
```

Artifacts remain under `.tmp/native-table-svg-probe`. This is a headless engine
observation, not an independent viewer UI or PowerPoint comparison. It does not
establish font identity, full pixel equivalence, border fidelity, screen-reader
behavior, arbitrary merges or general clipping correctness. Those remain open.

A separate [quarter-turn specimen](native-table-rotation-evidence.json) places
red clockwise and blue counterclockwise text in adjacent cells. MuPDF extracts the
two words at +90° and −90°, respectively, and each color's observed pixels stay
inside its corresponding cell. The default XHTML version has no visible colored
text in this engine. These checks do not establish Office placement, wrapping or
direction for other writing modes. Reproduce with:

```sh
rtk proxy node scripts/research/native-table-rotation-probe.mjs
rtk proxy python3 scripts/research/oracles/native-table-rotation-probe.py
```

After the cell accessibility-name change, the specimen was generated again in
`.tmp/native-table-svg-ax-probe`; the [new manifest](native-table-svg-ax-evidence.json)
records the new SVG hashes. Both MuPDF raster hashes match the earlier capture.
The earlier manifest and artifacts are retained as historical evidence. The
current producer reproduces the newer SVG, not the historical SVG. A separate
Chromium Playwright test checks that a narrow visually wrapped header and cell
remain exposed as `Heading` and `Value` in the browser accessibility tree. This
does not establish screen-reader or other-browser behavior.
Later font-dependency reporting adds structured details to warning reports, so
rerunning the producer does not reproduce the historical warning hashes.
