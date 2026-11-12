# SmartArt source-binding research protocol

Contract: [ADR0027](../architecture/decisions/0027-smartart-source-binding-research.md). This is private, plain-paragraph/ordered-forest research; source bindings do not authorize diagram source saves.

```js
import { bindDiagramSourceContent } from "../../scripts/research/diagram-source-bindings.mjs";
import { createDiagramIdentityState } from "../../scripts/research/diagram-identity-state.mjs";
import { buildNativeDiagramCandidate } from "../../scripts/research/native-diagram-package.mjs";

const bound = bindDiagramSourceContent(inspectedModel, exactLayoutIdentifier, {
  scope: "document-A/diagram-A", namespace: "12345678-1234-4ABC-8DEF",
});
// Reinspect the same logical source; text/order/element indices may change.
const next = bindDiagramSourceContent(nextInspectedModel, exactLayoutIdentifier, {
  scope: "document-A/diagram-A", state: bound.state,
  // restore: [rawModelId] only for an explicitly restored retired source point.
});
const fresh = await buildNativeDiagramCandidate(next.candidate, admittedComponents, {
  presentation: "layout", modelIds: "guid",
  identityState: createDiagramIdentityState(next.candidate.family, "87654321-4321-4CBA-9FED"),
});
// The caller commits next.state and fresh.specimen.identityState with the result.
// Fresh generation is distinct from source-preservation editing.
```

The source state and writer state use separate identity domains. Persist both across their corresponding operations; never substitute traversal IDs or infer identity from equal labels. Context/root changes and incomplete graphs with findings reject. Explicit restore is required when a previously observed retired source ID returns. A producer reusing an active ID between unobserved snapshots is not detectable; maintain source lineage/baseline integrity outside this adapter. No concurrent-fork merge or publication transaction is provided here.

## Owned automated checks

After building packages:

```sh
rtk proxy node --test scripts/diagram-source-bindings.test.mjs scripts/diagram-content-research.test.mjs scripts/native-diagram-research.test.mjs
rtk proxy pnpm test:diagram-research
rtk proxy pnpm test:import
```

There are seven pure binding groups and fifteen native research groups; the latter includes both existing native generation controls and the new source-binding/fresh-writer integration. The pure suite covers hundred-round identical-label sequences, context/canonical identity collisions, retired source restoration, malformed state, exact work/key thresholds and failed-baseline immutability. Integration covers three families × two cache policies × two source identity policies, with actual mutated source ZIP/XML read by Parser. Independent browser DOM performs the controlled mutation; CR character references remain escaped so the control does not accidentally alter text.

## Optional retained-artifact checks

Retain the hash-pinned local source corpus, prior source/native snapshots and the generated PowerPoint artifacts listed in the two evidence manifests. Third-party components and transformed files remain local pending redistribution admission. No download or Office launch is performed.

```sh
rtk proxy node scripts/inspect-smartart-source-bindings.mjs /absolute/pinned-source-directory
```

The runner creates a new local directory and writes [evidence JSON](smartart-source-bindings-evidence.json). Original/source-native manifest hashes are recorded; each file hash and ZIP is independently verified. Five groups/eleven versions compare source identity continuity through recorded native edits and deletion. Python independently reconstructs parent order, paragraphs, assistants, raw indices, canonical identities and managed allocation/reservations. The explicit integer-to-native-GUID mismatch control is a passed rejection, not a successful binding. Manual native UI evidence remains in the earlier [editing protocol](smartart-generated-native-editing.md); these offline checks do not replay it.
