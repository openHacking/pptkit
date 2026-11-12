# ADR 0025: Generated SmartArt native edit evidence and specimen identities

Status: research evidence; public authoring admission remains open.

## Decision and scope

Retain private generation research and add explicit `modelIds: "int32" | "guid"` specimen serialization. The integer default stays unchanged. Remap document, content, transition, presentation and connection identities and every typed reference together; charge the expanded output against existing budgets. These traversal-derived GUIDs are specimen-local. They do not provide managed author identities across regeneration, reorder or insertion, and must not become a Core identity contract.

On PowerPoint 16.104 (16.104.25121423), macOS 26.5, the three standard absent-cache generated layouts were recognized as native SmartArt. Actual operations included text edits, addition, reorder, layout switching/back, Office theme change, save/close/reopen. Hierarchy retained its assistant role. Integer identities were replaced by GUIDs on native save. A separate generated GUID Cycle control retained its three content IDs through three native text-edit/save/close/reopen cycles. Its original data XML equals the current GUID writer's standard Cycle output, independently checked byte for byte.

These observations admit neither all native operations nor a public family implementation. Deletion is untested; empty/single/long and empty-cache PowerPoint UI cases remain untested. Process/Hierarchy do not yet have three native cycles. GUID Cycle text cycles do not establish structural identity maintenance. Full rich text, deterministic layout, generated caches, freshness policy and source-component licensing remain separate gates.

## Evidence and controls

See the [native-edit protocol](../../testing/smartart-generated-native-editing.md), [nine-stage file evidence](../../testing/smartart-generated-powerpoint-evidence.json), [24 GUID candidate checks](../../testing/smartart-native-guid-candidate-evidence.json) and [six GUID headless producer cases](../../testing/smartart-native-guid-producer-evidence.json). Manual UI observations and independently verified resulting file properties are separate fields; the offline verifier never replays or certifies UI.

Twelve native research test groups cover both identity policies, topology, typed-reference closure, determinism, invalid policy and expanded-output budgets. Independent browser DOM coverage spans three families, two caches, two presentation policies and both identity policies. Unrelated supported text-save rounds check source payload retention; they are not native edit rounds.

The [six Process geometry controls](../../testing/smartart-process-layout-controls.json) show three out-of-frame boxes for original source, original data in a fresh shell, generated data and three presentation-variable variants in the named LibreOfficeDev runtime. This does not isolate a generator defect. The PowerPoint standard Process observation fits its source frame; cross-application results remain separate. No speculative presentation-variable patch is admitted.

Continue ADV-02.2.1.2–.1.4 and ADV-02.1.3 before the per-layout admission decision ADV-02.2.2 and public ADV-03 work. Parent tasks remain incomplete.


## Subsequent evidence

The user later authorized the specific added blank Cycle node deletion. It was deleted with the native text-pane Delete control, saved, closed and reopened: four labeled nodes and four closing arrows remain. The [live evidence](../../testing/smartart-generated-powerpoint-evidence.json) now has ten stages and an independent deletionControl; exactly one blank removed, retained IDs/text/roles/order unchanged. Process/Hierarchy deletion stays untested. [ADR0026](0026-smartart-identity-state-research.md) separately advances stateful generator identities; it does not inherit native admission from the stateless Cycle experiment.
