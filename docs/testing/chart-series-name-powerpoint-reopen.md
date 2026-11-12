# PowerPoint series-name reopen observation

On 2026-11-12, Microsoft PowerPoint for Mac 16.104 opened three successive Parser v5 source-save outputs from the generated [two-series source](../../tests/fixtures/chart-label-source.pptx). The PowerPoint accessibility tree reported one clustered-column chart, both series, the three category labels and values `1, 2, 3` and `4, 5, 6`. The first series' legend and each of its three data-point containers reported the expected name in each round:

| Round | Input opened in PowerPoint | First series name | SHA-256 |
| --- | --- | --- | --- |
| 1 | [round 1](../../tests/fixtures/chart-series-name-native-round-1.pptx) | Native revised | `028a10b4a11c1a644568f026627774ecadebf473c8463b4a02c0071571a4c359` |
| 2 | [round 2](../../tests/fixtures/chart-series-name-native-round-2.pptx) | Native second | `b24156b7f05f4db9507c2b6f879cd04fcbf1deb183cf1b93faa67824668554df` |
| 3 | [round 3](../../tests/fixtures/chart-series-name-native-round-3.pptx) | Native third | `6838385c792b73ec55d5c9c9a01a1c3363a49f9c9af85608e925d9b215d72484` |

The three inputs were produced by consecutive Parser saves, not by PowerPoint resaves. The observed evidence is the application's accessibility tree, not a screenshot or a pixel comparison. PowerPoint was not used to edit the workbook or save these three inputs; these reopens alone do not establish Excel editing, PowerPoint resave behavior, other chart families, Unicode rendering or other producers' compatibility.

Separately, PowerPoint 16.104 saved a copy of round 3 as [this retained file](../../tests/fixtures/chart-series-name-powerpoint-resave.pptx), SHA-256 `8fe55fb2dc2d4b6a25b852dc46b6ecbf34562cd6036a0a8c91e254de7041414b`. Closing and reopening that file showed the same two series, legend entries, category labels and point values in the accessibility tree. An independent ZIP/XML check found `Native third` in the first series cache and embedded workbook B1, and `Second & series` in the second cache and C1. Parser v5 then rebound the retained file and saved a further name and numeric change in an automated test. PowerPoint changed the embedded workbook part name to `Microsoft_Excel____.xlsx`; the check does not require byte-identical resaving. This is one resave of one generated chart, without interactive workbook editing.

On 2026-11-12, the retained resave was reopened in PowerPoint and its chart selected. The Chart Design ribbon reported **Edit Data in Excel**, **Select Data**, **Switch Row/Column** and **Refresh Data** as disabled. Spotlight returned no application with bundle identifier `com.microsoft.Excel`, and `/Applications/Microsoft Excel.app` was absent. These observations explain why this attempt did not reach an interactive data edit; they do not prove the cause of the disabled controls or establish that Excel is unavailable everywhere on the machine. No Edit Data result is claimed.

`node scripts/research/generate-series-name-native-rounds.mjs` reproduces the inputs in a temporary directory after building Parser. `pnpm test:chart-save-evidence` checks retained hashes and independently compares the ZIP/XML source changes for each round; it cannot reproduce or certify the interactive PowerPoint observation.
