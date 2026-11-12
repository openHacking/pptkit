# Controlled SmartArt data/cache disagreement

The [optional runner](../../scripts/inspect-smartart-stale-cache.mjs) creates controlled copies from hash-pinned retained PowerPoint Process (`process3`), Hierarchy (`orgChart1`) and Cycle (`cycle2`) outputs. Originals and unmodified package payloads are independently checked by Python ZIP/ElementTree. Selected targets require explicit association and initially equal nonempty single-paragraph text. Selection uses raw source indices/model IDs, not label search. ElementTree rewrites selected test parts; it is not a preservation-save implementation.

Each family has baseline, data-only text change, cache-only text change, identical text changes on both sides, and empty cache tree. Empty controls retain group properties. All fifteen pass exact no-op and the opt-in character comparison. Cached previews follow stored cache text, including the disagreement controls, and retain unknown freshness. Empty caches use placeholders. This corpus concerns text; it does not cover arbitrary stale geometry/styles/fonts.

## Reproduction

```sh
rtk proxy node scripts/inspect-smartart-stale-cache.mjs docs/testing/smartart-generated-powerpoint-evidence.json /absolute/path/to/soffice
```

Omit the executable for offline controls only. Each run gets a unique local output directory. Native calls use an isolated profile, a per-operation timeout, and separate direct PDF / saved PPTX / reopened PDF directories. Failures are retained as failures. Inspect `native-progress.json` while running. The final [evidence manifest](smartart-stale-cache-evidence.json) records input/transformed/native/PDF hashes, selected source/cache identities, direct/native text, preview reports and command outputs. It references local binaries that remain unavailable for redistribution pending admission.

## Observed 2026-10-26

Source application: PowerPoint 16.104 build 16.104.25121423. Native headless observer: LibreOfficeDev 26.8.0.0.alpha0, commit `2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`. All 45 conversion operations complete across fifteen cases. Python independently rechecks ZIP/XML and extracts direct/reopened PDF text with PyMuPDF. Marker/label matching ignores PDF whitespace fragmentation; it does not establish glyph or geometry fidelity.

The following observations hold separately for all three selected layouts:

| Control | Parser text observation | Direct and reopened PDF | Saved source data/cache |
| --- | --- | --- | --- |
| Baseline | selected vectors equal | original label | original text |
| Data-only | selected vectors different | original cache label, marker absent | new data marker, old cache text |
| Cache-only | selected vectors different | new cache marker, original selected label absent | old data text, new cache marker |
| Both | selected vectors equal | new marker | marker in both |
| Empty | no cached shapes | original label | source data retained, no cached shapes |

Thus this LibreOffice version reuses nonempty stored cache characters even when the selected source text differs, while empty controls render from another path. Do not infer that empty caches were serialized as regenerated DSP shapes: the saved empty controls still have zero cache shapes. These headless observations do not establish native UI editing/repair/playback, comprehensive text roles, style/layout/font freshness or family-wide fidelity. The subsequent PowerPoint UI matrix below supplies separate named-application evidence. Keep ADV-03.7 and layout/native admission open.


## PowerPoint UI continuation — 2026-10-26

[Retained PowerPoint evidence](smartart-stale-cache-powerpoint-evidence.json) separately records PowerPoint 16.104 build 16.104.25121423. Actual GUI open → Command-S → close → reopen was performed for all fifteen controls. Source markers appear on opening the data-only/both controls; the cache-only marker is not displayed and the original selected data label appears instead. Empty controls display original data labels. No repair dialog was observed. Original controls remain immutable; operations use separate local copies. This is one standard specimen per named layout, not a certification for all producers/layouts.

| Control | PowerPoint open and reopen | Saved DSP cache | LibreOffice direct/reopened PDF |
| --- | --- | --- | --- |
| Data-only | new data marker | rebuilt to new data text | old cache label |
| Cache-only | original data label | rebuilt to original data text | cache-only marker |
| Both | new marker | same new marker | new marker |
| Empty | original data label | regenerated: Process/Hierarchy 8 shapes, Cycle 6 | original label, saved cache still empty |

Three additional fresh data-only copies exercise native re-layout and return: Process 重点流程 → 基本流程 → 重点流程; Hierarchy 标准 → 左悬挂 → 标准; Cycle 基本循环 → 块循环 → 基本循环. Intermediate/restored/reopened screenshots retain visible source markers and the SmartArt design controls. These operations are separate from the ordinary-save controls. They retain the original source content identities, roles, text and ordered parent edges according to independent ZIP/XML checks. Presentation identities and layout parameters are not promised stable.

A further six controls extend each family's data-only case to three consecutive rounds. Each round changes only the selected data point's scalar text in the preceding PowerPoint save, deliberately retaining its old nonempty cache; Python verifies the sole changed package payload is the data part and confirms an actual text disagreement before opening. All six new markers appear on opening/reopening, and all saved directly comparable source/cache vectors agree. These are controlled source mutations followed by native rebuilds, **not** three rounds of native text editing through the UI.

Across the three-round chain, content IDs/roles/order are retained. Process/Hierarchy each keep 46 package parts, 43 relationships, 5 diagram-directory parts and zero media parts; Cycle keeps 26/23/5/0. Native ZIP sizes are respectively 43,590/43,589/43,591 bytes; 45,818/45,825/45,824; and 19,312/19,312/19,313. These observed counts do not prove arbitrary long-run resource bounds.

The [optional retained-output verifier](../../scripts/inspect-smartart-stale-cache-powerpoint.mjs) checks all **24** native outputs with public Parser/no-op APIs and an independent Python ZIP/ElementTree path. It hashes every initial/reopened image and AX artifact; re-layout cases also hash intermediate/restored artifacts. It verifies exact source/control hashes, saved/reopened byte identity, direct cached text, content identity/roles/paragraphs/ordered edges, generated cache shape counts and round input/payload/resource constraints. It does not drive or replay the native UI:

```sh
rtk proxy node scripts/inspect-smartart-stale-cache-powerpoint.mjs /absolute/path/to/retained/powerpoint-run
```

An optional second argument selects an alternate evidence output file. The [optional adversarial runner](../../scripts/test-smartart-stale-cache-powerpoint.mjs) and [three retained rejection results](smartart-stale-cache-powerpoint-negative-evidence.json) confirm the verifier rejects saved/reopened snapshot mismatch, cached text corruption, and a consistent content-GUID rewrite that still has valid references and equal cached text. Run `rtk proxy node scripts/test-smartart-stale-cache-powerpoint.mjs /absolute/path/to/retained/powerpoint-run` to create separate mutated copies; originals and positive evidence stay hash-identical. The first validation attempt incorrectly required intentionally conflicting **input** vectors to agree; the verifier now requires agreement only for native **outputs**, while proving the repeated-round inputs do disagree.

All caches/previews still report unknown freshness. Long marker strings trigger native fitting and sometimes smaller text/boxes. Exact font/geometry/paint fidelity, arbitrary rich text, other layouts, image nodes, shared-source variants, native node deletion/editing and layout-generation contracts remain independent gates. [ADR0031](../architecture/decisions/0031-native-stale-cache-policy.md) records the cross-application policy consequence; ADV-03.7 remains open.
