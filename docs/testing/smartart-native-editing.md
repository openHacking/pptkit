# SmartArt native source editing and content-model evidence

This is a named-source experiment, not a general compatibility or native authoring guarantee. PowerPoint **16.104**, build **16.104.25121423**, on macOS **26.5** was operated through Computer Use on local copies of the pinned LibreOffice corpus inputs. Originals remain unchanged. The [evidence JSON](smartart-content-model-evidence.json) records input/output SHA-256, source commit/URLs, raw content identities, paragraphs/children, cache counts, UI observations and screenshot/AX hashes.

| Layout / source | Actual native observations | Independent saved-file evidence | Still untested |
| --- | --- | --- | --- |
| process3 / smartart-accent-process | Opened without an observed repair dialog; SmartArt Design showed Accent Process. Appended `PPTKit Process` to `a`, added `PPTKit Added Process`, moved the added root earlier, saved, closed and reopened. Edited text and three top-level cards were visible after reopening. | Four → five content nodes; original content IDs/control text retained, added node text and parent order match the research projection; zero → eight cached shapes, each referring to a model point. | Native deletion, theme change, repeated edit rounds, calibrated visual comparison, a minimal PPTKit-generated native diagram. |
| orgChart1 / smartart-org-chart | Opened without an observed repair dialog; SmartArt Design showed Organization Chart. Appended `PPTKit Hierarchy` to Manager, added `PPTKit Added Hierarchy` as a child, reordered it and saved. Separate assistant placement remained visible. | Five → six visible content/assistant nodes; Assistant role and second Manager paragraph retained; zero → ten cached shapes. Original content IDs remain and the new child has the expected text. | Reopen was not observed: the Mac locked during path selection. Deletion, theme change, repeated native rounds and generated-output admission also remain open. |
| cycle2 / smartart-cycle | Native UI not reached before the Mac lock. | Original source has five flat document children, ordered by `srcOrd`, with no recognized cache. The candidate records layout-generated closure. | All native open/edit/save/reopen and family admission operations. |

Native addition triggered visible automatic re-layout of existing source SmartArt. That does not prove PPTKit can generate or regenerate those caches. Saved outputs still contain diagram data/layout/style/color parts and model-linked cached shapes; these were not conversions to ordinary shape groups. Native source IDs were independently compared rather than inferred from visible labels. Process and hierarchy evidence must not be extrapolated to Cycle, other layouts or other application versions.

UI actions were manually observed by the agent; the automated verifier checks the resulting files and recorded artifact hashes. It does not prove a manual observation by itself. Screenshots and AX captures are retained under `.tmp/smartart-native-model-20261025`; third-party fixture redistribution admission remains pending, so binary inputs/outputs and source screenshots are not included as portable test fixtures. The ordinary research suite uses owned synthetic records and runs offline.

After obtaining the pinned inputs and retaining the native output/observation artifacts, reproduce the file/model checks with:

```sh
node scripts/inspect-smartart-content-model.mjs [source-corpus-directory] [native-artifact-directory]
pnpm test:diagram-research
```

The source corpus requires its provenance manifest; the native artifact directory requires `observations.json` and the named saved files/artifacts. The runner validates hashes for original sources, then compares Parser/research projections with an independent Python ZIP/XML path. Exact original-byte no-op passes for all three sources and both saved native outputs. UI observation fields remain separate from file-verification fields.

[ADR0022](../architecture/decisions/0022-smartart-content-model-research.md) records the provisional model: ordered content forest, explicit assistants, source bindings outside the candidate, plain-text-only coverage and projection-local IDs. It also records failure cases and why public Core/native authoring/layout support remains unadmitted. Remaining work follows ADV-02.1.3, ADV-02.2 and ADV-03; existing parent tasks remain incomplete.


## Reopen update after unlock — 2026-10-25

After the user unlocked the Mac, the saved source Hierarchy reopened in the same PowerPoint environment. Manager edits, its second paragraph, the assistant and added ordinary child were visible. `hierarchy-source-reopened.png/.ax.txt` are retained in `.tmp/native-diagram-powerpoint-kmjuxfmg` and referenced by the [new generated-output evidence](smartart-generated-powerpoint-evidence.json). The earlier table describes the original interrupted observation; this update supplies the later reopen observation. Source edits and [generated-output editing](smartart-generated-native-editing.md) remain separate experiments.
