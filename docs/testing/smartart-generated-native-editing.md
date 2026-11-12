# Generated SmartArt PowerPoint editing observations

This is scoped research under [ADR0025](../architecture/decisions/0025-generated-smartart-native-edit-evidence.md), not a public native-authoring guarantee. Environment: PowerPoint **16.104**, build **16.104.25121423**, macOS **26.5**. Computer Use operated local copies; original generated candidates remain unchanged. Source components remain local pending licensing admission.

| Generated layout | Observed native operations and reopened result | Immutable stages |
| --- | --- | --- |
| process3 / Accent Process | Edited root; added and moved a root; changed SmartArt colors; Basic Process then back; Office theme; save/close/reopen. Three cards retain both detail children. | 2; five content nodes, three roots, eight cache shapes |
| orgChart1 / Organization Chart | Edited manager; added/reordered ordinary child; Horizontal Organization Chart then back; Office theme; save/close/reopen. Assistant stays separately connected. | 2; five nodes, two roots, one assistant, eight cache shapes |
| cycle2 / Basic Cycle, integer input | Edited label; Add Shape plus paste unexpectedly made two nodes, one blank; reordered labeled node; Text Cycle then back; Office theme; save/close/reopen. Five nodes and closure remain. | 3; final deletion stage has four roots and eight cache shapes |
| cycle2 / GUID input | Three separate text-edit/save/close/reopen cycles, changing top, B, then C. Three original content GUIDs retained. | 3; three roots; original data bytes match current GUID writer |

No repair prompt was observed on initial opening. Do not reinterpret this as automated absence detection or a universal repair-free guarantee. The integer Cycle root actually became `cycle  PPTKit Cyclelead`; the added blank is a recorded outcome, not a passed four-node expectation. After explicit action-specific user authorization, the blank third text-pane entry was deleted once with the native Delete button. Save/close/reopen shows four labeled nodes and four closing arrows. The prior five-node result remains an immutable stage; unlocking itself was not treated as deletion authorization.

## Reproduce saved-file checks

Retain `.tmp/native-diagram-powerpoint-kmjuxfmg` with `inputs.json`, `observations.json`, immutable PPTX snapshots, screenshots and AX captures, then run:

```sh
rtk proxy node scripts/inspect-smartart-generated-native-edits.mjs /absolute/native-artifact-directory
```

The verifier writes [evidence JSON](smartart-generated-powerpoint-evidence.json). It checks original hashes, all ten stage graphs, ordered paragraphs/roles/roots, cache model references, exact no-op and independent Python ZIP/XML expectations. Three additional supported PPTKit text-preservation saves per final recorded family stage (Cycle now selects the post-deletion stage) alter only slide1.xml and retain diagram payloads. These nine API saves are **not** PowerPoint UI cycles; their native replay is untested. Screenshot/AX hashes link manual records to retained artifacts; the runner does not launch PowerPoint.

Remaining matrix: native delete for Process/Hierarchy and broader Cycle deletion boundaries; Process/Hierarchy three native cycles; structural GUID insertion/deletion/reparent/subtree/reorder identity behavior; empty/single/long/CJK/emoji and both cache policies; font/text fitting, frame/direction variation, regenerated cache consistency and redistribution admission. Record each independently under ADV-02.2.1.3/.1.4; no parent task is closed.


## Authorized Cycle deletion update

The retained post-deletion snapshot is `cycle-after-native-delete.pptx`. The verifier's optional deletionControl checks the before/after files as hash-verified stages, independently proves exactly one blank node removed, no content nodes added, and all four retained IDs/text/roles/order unchanged. Regular stage verification checks cache references and no-op. The reopened screenshot supplies the manually observed four arrows; XML parent topology does not itself prove visual closure. This native operation used the earlier integer-generated/PowerPoint-saved candidate, not the later [stateful GUID writer](smartart-identity-state.md). Process/Hierarchy deletion remains untested.
