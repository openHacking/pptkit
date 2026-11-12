# Field projection and chart source-save: PowerPoint observation

On 2026-10-29 a repository-authored four-slide presentation was generated with a slide-number field whose cached text was `STALE`. PPTKit imported it with field projection and chart, text and starting-number save profiles. One preservation plan changed the adjacent ordinary text from `Page: ` to `Edited page: `, the first bar chart's second primary point from 5 to 7 in the cache and workbook, and the starting number to 17. The field itself was not edited.

Retained files under `.tmp/field-chart-native-powerpoint/`:

| Artifact | SHA-256 |
| --- | --- |
| `source.pptx` | `2bf7806729416bdba7b51c5056fe4791521678c698e11d5d789aedba0ced5782` |
| `candidate.pptx` | `5e9e4c5e13fb4fbb203097d83ded4f429f8662759c6efd4def9326d967c23eb8` |
| `round1.pptx` after first native save | `aee4296e3c2589aa34d99de29c6901e31ee82aa552855ec70e1c937498115958`
| `round2.pptx` after second native save | `ae7a89211153afe8d562a46fe6420dc0988f8a413fd0062a79dbcef754bf1ff2`
| `round3.pptx` after third native save | `fa55024c67eb36b65e31a40d1979eaf9b9776ff418841b2a74de869f611c6a69` |
| `open.png` | `35aba7d195d57807167a8e650a0d60fc28df96ff633d290b7d86e24ff2173e49` |
| `reopen.png` | `df9d87c4898d0dfde89a9116f186c52de8e91849ab5b3585a5fc4cda02df3385` |
| `reopen2.png` | `96e50109293c35e4207f5bb43fe8792005cdef58ecc40eae7ffa7439cd4407a5` |
| `reopen3.png` | `2faf5492237985d2a4a9f93aa40a23d2f85b7b2305361756a126e097f789ec8b` |

PowerPoint for Mac 16.104 displayed `Edited page: 17 / note` and the bar at 7 on initial open. The first slide was cut and pasted after the second, then the disposable copy was saved, closed and reopened. The moved page displayed `Edited page: 18 / note` and the same bar. Two further cut/paste, save and reopen rounds moved it to pages 19 and 20, where the field displayed the respective page number and the bar remained visible at 7. Screenshots record these manual observations, not automated visual comparisons.

The independent [candidate verifier](../../scripts/research/verify-field-chart.py) checks that only the four expected outer payloads and one workbook worksheet changed, and that their XML trees differ only at the approved attribute, text and two numeric cells. The source field's UUID, type and `STALE` cache remain unchanged in the PPTKit candidate. Negative tests reject modified field IDs, caches, chart data and unrelated slides. After the native save, independent ZIP/XML inspection finds the same field UUID and `slidenum` type on the moved second slide, with cache `18`; the independent chart relationship reader finds the second slide's bar values `[2, 7, 8]` and its associated workbook B3 contains 7. The [native verifier](../../scripts/research/verify-field-chart-native.py) follows the presentation relationships in all three retained rounds and checks the field UUID, type, page-specific cache 18/19/20, neighboring text, chart cache, workbook cell and control titles/values. PowerPoint rewrites package parts, so the source-preservation claim applies to the PPTKit candidate only.

Reproduce the candidate and negative checks with `pnpm test:chart-save-evidence`. This is one application/version, one slide-number field, one sample and three native save/reopen rounds. It does not validate other field types, interactive chart data editing or complete visual parity.
