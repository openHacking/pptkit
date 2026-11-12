# ADR 0026: Explicit SmartArt identity-state reconciliation research

Status: private executable research, 2026-10-25. Advances ADV-02.2.1.2.1; no public Core authoring or source-preservation writer is admitted.

## Problem and decision

The ordinal GUID specimens in [ADR0025](0025-generated-smartart-native-edit-evidence.md) survive named native text edits, but regenerate different IDs when content is inserted or reordered. Introduce a pure, explicit identity-state snapshot owned by the caller. Reconcile stable semantic keys against reserved GUID assignments before writing the data model. No module-global counter, filesystem access, random source or hidden document state participates.

The [state module](../../../scripts/research/diagram-identity-state.mjs) creates a versioned family-bound state with an explicit uppercase UUID v4/variant **prefix** (`8-4-4-4` hex digits), a monotonic 48-bit suffix counter and ordered key/ID reservations. The caller must allocate a distinct prefix for each independent diagram identity domain and persist the returned state only after its complete package transaction succeeds. The implementation does not generate namespaces or assert global uniqueness.

The private writer accepts `modelIds: "guid"` plus `identityState`, and returns the next detached `specimen.identityState`. Default integer and stateless GUID output remain unchanged. Package failure never mutates the supplied state. The caller must serialize commits: the pure snapshot does not provide compare-and-swap, merge or concurrent allocation authority. Two branches from the same old state are not safe to merge independently.

## Semantic key policy

| Object | Identity anchor | Behavior under edits |
| --- | --- | --- |
| Virtual document | Fixed document key | Retained within the same state |
| Content/assistant point | Candidate managed node ID | Text, ordering and reparenting preserve the ID; role may change only where the existing writer admits it |
| Parent connection and parent/sibling transition points | Uniquely parented child ID plus distinct kind | Parent/order can change while connection/transition identity follows that child |
| Presentation point | Source semantic key plus exact presentation role name | Retains identity while ownership and role stay the same; depth/role changes can create a different presentation identity |
| Presentation connection | Type and both semantic endpoint keys | Retained for the same typed association; changed topology creates a different connection |

All points and connections share one allocated identity space; endpoint, transition, reciprocal connection and presentation association references are serialized through the same mapping. Sibling indices, text, style and traversal ordinals never enter the identity keys. Cycle closure remains layout-generated, not a cyclic parent graph.

Deleted keys remain reserved. An explicit undo using the same candidate node ID restores its identity; a distinct newly authored object requires a new candidate ID. No compaction/recycling occurs. Stale presentation reservations may accumulate when layout roles change; capacity exhaustion rejects atomically. A layout or source binding is not inferred from a label. Parser projections still have traversal-local IDs and must not be fed into this state as persistent managed identity without an explicit source-binding adapter.

## Validation and resource contract

Validate exact family/version, known plain fields, uppercase namespace, monotonic distinct reserved IDs, serial ranges and unique active keys. Reject wrong-family state, conflicting assignments, duplicate active keys and exhausted serial space. Serial gaps below nextSerial remain reserved. Returned assignments/state are detached. Inputs are trusted JavaScript research records, not a hostile-accessor sandbox or authenticated ledger; callers own storage integrity.

Pure defaults: 200,000 retained or active entries, 8,000,000 cumulative UTF-8 key bytes and 600,000 work checkpoints; individual keys are at most 4,096 code units. Retained keys and requested active keys both consume work/bytes. In the writer, reconciliation consumes the same maxWork budget as topology and XML generation, with its key-byte bound set separately to maxOutputBytes. Existing XML output/text/depth budgets continue to apply; the output-byte budget covers emitted XML, not the sum of retained state and XML. No partial state or package is returned on failure.

## Evidence and remaining gates

Five pure groups test immutable reservations, JSON persistence, insertion/reorder/removal/restoration, namespace/family/collision rejection, exact budgets, serial exhaustion and 150 seeded edit sequences against an independent allocation oracle. Two new native-research groups generate eight structural stages for three families under both cache policies (48 packages), test text/insert/reorder/reparent/delete/restore, and exercise failed package retry and the exact cumulative-work boundary. Independent browser DOM verifies each structural stage's actual XML text/roles, parent topology/order, typed references and identity uniqueness. Thirty-six initial browser combinations cover three families, both caches, both presentation policies and integer/stateless-GUID/stateful-GUID identity modes. Unrelated source-text saves cover stateful GUID output as well.

These are generated-file and automated API checks. The new **stateful** writer has no native application admission yet. The separate native Cycle blank-node deletion is evidence about the earlier integer-generated/PowerPoint-saved specimen, not this stateful writer. Public managed IDs, imported-source identity adoption, full rich text/layout/cache policy, state commit concurrency and native structural multi-round admission remain open. ADV-02.2.1.2.1 and its parent tasks stay incomplete.


Subsequent source-to-managed binding research is recorded in [ADR0027](0027-smartart-source-binding-research.md). It introduces a separate scoped source state and explicit retired-source restoration; native GUID adoption, full lifecycle/concurrency and source writes remain unadmitted.
