# Combined starting-number, text and chart save: PowerPoint observation

On 2026-10-29, a repository-authored four-slide source was generated and saved by PPTKit with three changes in one preservation plan: insert starting number 17, change the first slide title to `Combined chart and page number`, and change the first bar chart's second primary-series point from 5 to 7 in both the chart cache and its workbook.

The retained files are under `.tmp/combined-number-chart-powerpoint/`:

| File | SHA-256 |
| --- | --- |
| `source.pptx` | `8a427afdd6f635ac410ed474d5cc57cd888a571e3101142c697eff8d78f2ef85` |
| `candidate.pptx` | `b1ab7a4c1a74092a967ce77895a4014e8e5bd7221804500055773c2dfa5f7422` |
| `round1.pptx` after first native save | `ce629c0a769abdb96a6e4f309645918689f445f99c87ddcdca9349c9fb1eeab6`
| `round2.pptx` after second native save | `4adc4fd9a79de25552dd935eb073f05c8be110d16494be26d352c19d5bcc1049`
| `round3.pptx` after third native save | `3cef264f633472aad73f891cf66d81e191928ad6ff5c94e61e46b0d2dff18092` |
| `open.png` | `14db05ec7206161c7cf25f03925ef9e522e86630d33233d43655e61282ccca6e` |
| `reopen.png` | `22e916958587662e1e7d5152fcccab7b0e022bc08890eca16948dfcb95ddeb95` |
| `reopen2.png` | `4435fa149e199c5fd3514a04bafc6dc1b10e1516b450bd695b57ec21b8457cff` |
| `reopen3.png` | `bc1a97bd93d4ca1f4d9816d1ef0ef668cb8281829bce1320a0fcb85211286cc7` |

Microsoft PowerPoint for Mac 16.104 opened the disposable copy. The initial screenshot shows the title, slide numbers 17–20, and the selected bar reaching 7. The first slide was cut and pasted after the second, then the file was saved, closed and reopened. The second screenshot shows the moved chart on slide 18 with the same title and bar height; slide 17 is the line chart. Two further rounds cut and pasted the same chart slide after the following slide, saved and reopened it. The target appeared on pages 19 and 20 with the same title and visible bar height; screenshots `reopen2.png` and `reopen3.png` capture those states. Screenshots and UI observations are manual, not automated pixel assertions.

A separate Python ZIP/XML oracle checks the original candidate against the source: only presentation, first slide, chart and workbook payloads changed; only the target worksheet changed within the workbook; the four XML trees differ only at the intended attribute, title, cache value and B3 cell. The existing independent chart oracle follows the native-resaved slide relationships and reports the moved bar chart as the second slide with cache values `[2, 7, 8]`; independent XLSX XML inspection finds B3 = 7. Parser inspection of the first native save also finds start number 17 and all 21 chart point bindings available. The [round verifier](../../scripts/research/verify-combined-number-chart-native.py) resolves actual presentation relationships and checks all four titles, chart inventories and control primary values, target chart cache and its referenced B3 cell in each of the three retained native files. PowerPoint rewrote the package and renamed chart/workbook parts; source-byte preservation is claimed only for the PPTKit candidate, not the native resave.

Reproduce the source/candidate comparison and negative checks with `pnpm test:chart-save-evidence`; the preparation and narrow oracle are [here](../../scripts/prepare-combined-number-chart-native.mjs) and [here](../../scripts/research/verify-combined-number-chart.py). The Python oracle is scoped to this generated sample and is not a general OOXML validator.

This is one application/version, one sample and three native save/reopen rounds after page reordering. It does not validate native Edit Data, other chart types or values, visual pixel equivalence.
