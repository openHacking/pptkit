# Existing chart numeric source-save evidence

This repository-authored experiment advances IMP-09.2.6.4 for the fixed-topology [existing-values-v1 profile](../api/chart-value-save.md). It combines a separate Python ZIP/XML reader with named LibreOffice headless save/reopen. Interactive Edit Data in PowerPoint and LibreOffice remains **untested**.

## Reproduce

Build the packages, then run:

```sh
pnpm build
pnpm test:chart-save-evidence
node scripts/inspect-chart-value-save.mjs /path/to/soffice
```

The optional evidence tests require Python 3; the native runner additionally requires LibreOffice and Python PyMuPDF (`fitz`) to inspect the PDF page count/text. These optional tools are not added as runtime dependencies or prerequisites for the standard test command. The runner writes a unique local `.tmp/chart-value-save-native/run-*` directory and updates [the evidence manifest](chart-value-save-evidence.json). It never edits a user document or asserts a GUI action from a headless result.

The [fixture](../../tests/fixtures/chart-save.mjs) contains one bar, line and pie chart on separate pages, plus a fourth untouched bar chart. Bar/line also have an untouched second series. Titles, notes, chart frames, labels and source part relationships remain unchanged during PPTKit saves.

## Observed results

Application: **LibreOfficeDev 26.8.0.0.alpha0, commit 2c87e51eeaa2b413ff4ae097b2705eea1995d8e5**, macOS arm64. These observations apply to this named version and sample.

| Round | Bar selected point | Line selected point | Pie selected point | PPTKit changed outer parts | Other outer payloads retained exactly |
| --- | --- | --- | --- | --- | --- |
| 1 | 12.5 | 13.5 | 14.5 | 6 | 39 |
| 2 | 0 | 0 | 0 | 6 | 39 |
| 3 | 37.25 | 38.25 | 39.25 | 6 | 39 |

Each round imports the previous **PPTKit source-save** result and writes only the first series' second point in the first three charts. It never substitutes a fresh export for a preservation save. The outer inventory remains 45 parts. Three chart payloads and their three embedded workbooks change; all other payloads, including the control chart/workbook, remain byte-identical. Every changed inner package retains its exact part inventory and all unapproved payloads. XML structure, attributes and non-approved scalar text are independently compared, including worksheet cell formatting.

Every saved result was separately opened/resaved by headless LibreOffice. Python independently read back all four chart families and every series' numeric caches, matching the selected values and controls. LibreOffice output contains 48 parts and **no referenced embedded XLSX workbooks** for these charts. Its numeric caches remain readable, but every value is `unmapped` under the cache+XLSX preservation profile. A byte-exact no-op remains available after import. An actual numeric mutation followed by `prepareSave({ mode: "edit" })` returns blocked with `save-edit-unsupported` and no save plan in each native result. Native cache retention therefore supplies no permission for another numeric preservation save.

The final saved source also reopens to a four-page PDF with all four expected page titles. The page/text observation proves document rendering occurred; it does not prove visual parity, chart labels, repair-dialog absence or interactive data editing.

## Independent oracle and adversarial checks

[The Python oracle](../../scripts/research/chart-value-save-oracle.py) uses only standard-library ZIP/XML/Decimal operations. It follows presentation, slide, chart and XLSX worksheet relationships; resolves the explicit local one-column numeric ranges used by these fixtures; checks dense cache indices, source series identities/order and bound worksheet cells; verifies proposed values in both storage layers, including signed zero; and compares all approved/unapproved outer and inner payloads. It does not call Parser, read its binding report or use its normalized snapshots. Input/archive/XML/node/depth bounds and DTD/entity rejection keep the observation profile finite.

[Seven evidence-test groups](../../scripts/chart-value-save-evidence.test.mjs) verify both OOXML dialects and three chart families over three rounds; multi-slide/multi-series controls; stale XLSX and source disagreement; wrong/duplicate/invalid proposals; unapproved numeric/style/part mutations; signed-zero disagreement; duplicate ZIP entries, DTD/depth rejection and multiple proposals in one worksheet. Decimal comparisons use the original XML text, including values beyond binary floating-point precision; the numeric values in the public inventory remain floating-point summaries. A separate group checks the stored manifest's rounds, numeric observations and explicit native-write eligibility boundary. These are independent evidence checks, not a complete SpreadsheetML validator or an opaque-extension dependency proof.

## Open gates

- PowerPoint and LibreOffice actual Edit Data, followed by save/close/reopen and independent cache/workbook checks.
- Three rounds beginning from admitted native-produced XLSX-backed sources, including formatting/reference behavior.
- Series/category structural operations under IMP-09.2.6.3, shared-edit choices and complete dependency closure.
- Full appearance, numeric/date formats, labels and supported implicit references.

After the Mac was unlocked, the hash-verified round-3 candidate was opened in PowerPoint 16.104. The selected chart’s **Edit Data in Excel** command was disabled, so no interactive data edit or native save/reopen was performed. Excel was not found at `/Applications/Microsoft Excel.app`; that observation does not establish the cause of the disabled command. A separate attempt to launch the bundled LibreOffice headless application as a graphical app timed out. The captured [PowerPoint observation](chart-value-powerpoint-observation.json) records the sample and screenshot hashes; its UI statements are manual observations, not automated assertions. Run `python3 scripts/research/verify-chart-powerpoint-observation.py` to recheck the retained local files, hashes and XML facts.

The same candidate reveals a display limitation: the source chart fixes its value axis to 0–9, and the saved selected value is 37.25. PowerPoint visibly clips the bar at the axis maximum. Source and candidate chart XML independently confirm the fixed maximum and changed cache value. Preserving the axis is consistent with the current numeric-only contract, but numeric preservation is not a guarantee that edited values fit an existing fixed axis. This observation is limited to the named sample; the earlier independent oracle separately checks its cache and workbook. IMP-09.2.6.4 and its parents remain unchecked.

`validatePresentation` reports `chart-value-outside-fixed-axis` at each finite value outside a valid explicit fixed axis. It covers line, clustered bar, standard area, combo series on their respective primary/secondary axes, and scatter/bubble X/Y coordinates with a complete finite pair. The warning includes owner identity and the value path, does not alter the scale or data, and does not block normalization or a preservation save. Its bounds are inclusive. It is a conservative risk indication, not a pixel visibility test: auto, invalid, stacked and percent-stacked axes, incomplete pairs and blank values are outside this check; zero-display blanks are not converted to zeros. Bubble radius crossing the plot edge is not assessed. The check uses the projected Core axis, so an unmapped or partial source axis cannot be diagnosed by it. [Source-save](../../scripts/chart-value-save-parity.test.mjs) and [fresh combo](../../scripts/chart-combo-parity.test.mjs) regressions verify warnings across save/reimport; neither is a native Edit Data test.
