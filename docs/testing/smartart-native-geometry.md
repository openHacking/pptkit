# SmartArt native geometry baseline

Contract and limits: [ADR0024](../architecture/decisions/0024-smartart-native-geometry-observation.md).
This baseline advances layout-input evidence, not a layout engine or public
native SmartArt API. The source/producer artifacts remain local pending licence
admission.

## Reproduction

Portable pure checks are included in `pnpm test:diagram-research` and `pnpm test`:

```sh
rtk proxy pnpm test:diagram-research
```

The optional extractor uses Python's standard library. Its owned-fixture tests
need no native application or third-party documents:

```sh
rtk proxy python3 scripts/research/test_diagram_geometry_extractor.py
```

Generate the 24 presentation-mode candidates using the
[minimum-generation protocol](smartart-native-generation.md), then use the printed
manifest path and an explicit installed native runtime:

```sh
rtk proxy node scripts/inspect-smartart-native-geometry.mjs /absolute/candidate-run/manifest.json /absolute/soffice
```

The optional runner also needs Python/PyMuPDF. No code or source document is
downloaded. Every invocation creates a fresh artifact directory, converts all 24
inputs to flat ODF and PDF, checks unchanged inputs, independently compares
rectangles and records [baseline evidence](smartart-native-geometry-evidence.json).
Command failures are recorded per case; an `observed` result means the observation
ran, not that the diagram passed display/fidelity/editing gates.

## Reading the evidence

| Field | Meaning and limits |
| --- | --- |
| `groupAssociation` | Explicit source frame name matched one native drawing group; missing/ambiguous is not treated as empty. |
| `geometry.shapes` | Every recognized drawing occurrence, with raw paragraphs, preset/geometry strings and source-name association evidence. |
| `boundsPoints` | Explicit axis-aligned rectangle in page points; does not include evaluated text areas, effects or clipping. |
| `transform-source-only` | Own/ancestor transforms prevent resolving a box; raw geometry stays available. |
| `association` | Exact native name evidence against source presentation IDs, retaining repeated occurrences and ambiguous IDs. No stable node binding is created. |
| `storedDrawingText` | Authored marker observations in native drawing paragraphs; nested lists/text boxes are included. |
| `text` | Authored marker observations in PDF extraction; separate from stored text and visual review. Empty graphs retain markerCount zero. |
| `pdf` | Page bounds, text spans/fonts, path paint bounds and hashed screenshots; no fabricated per-node ownership. |
| `cachePolicyComparisons` | Controlled absent/empty-cache geometry comparison under its stated exclusion policy; neither cache freshness nor regeneration is established. |

Each source input retains its SHA256. Each converted FODP/PDF and picture has an
independent artifact hash; do not silently replace an input or attribute a change
to PPTKit when it came from the native producer. Font-face declarations and PDF
font names describe this run; they are not a verified glyph-coverage inventory.

Review long-text pictures for actual crop/overlap. In the observed Cycle long
case, the top oval clips its label even though the shape box fits the frame.
PDF marker omissions therefore remain failures/limitations to investigate.
Stored text can remain complete while the rendered result is unusable.
Record picture-specific visual findings separately from automated numeric checks.

[Picture-specific reviews](smartart-native-geometry-review.json) record two Cycle
standard/long pictures and one Process standard picture with exact input/image
hashes. Cycle long-label clipping is visible; Process's right card extends beyond
the source frame. Numeric observations flag three out-of-frame boxes in each
Process standard/long case (both cache policies). These findings are not
extrapolated to every layout or environment.

Next cases are wide/deep trees, multiple assistants, varying frame sizes and
theme/direction parameters. Actual glyph/text-area fitting and deterministic
PPTKit geometry require their own independent expected results; they cannot be
inferred from matching native cache-policy controls.

## Process overflow attribution update — 2026-10-25

[Six controlled variants](smartart-process-layout-controls.json) each show three out-of-frame boxes in the named LibreOfficeDev runtime: original source, original data/fresh shell, generated baseline, root variables, title variables and all presentation variables. Original source hashes remain unchanged. Reproduce with:

```sh
rtk proxy node scripts/inspect-smartart-process-layout-controls.mjs /absolute/candidate/manifest.json /absolute/source-controls/manifest.json /absolute/soffice
```

Independent FODP/Decimal checks and the pure observer agree. The failure is not isolated to generated data; do not call it a proven generator defect or fix it by copying unproven variables. Separate [PowerPoint standard Process observations](smartart-generated-native-editing.md) fit the source frame. Neither result generalizes across applications or long-text cases.
