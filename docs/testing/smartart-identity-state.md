# SmartArt identity-state research protocol

Contract and limitations: [ADR0026](../architecture/decisions/0026-smartart-identity-state-research.md). This module is private research and is not an installed-package API.

```js
import { createDiagramIdentityState } from "../../scripts/research/diagram-identity-state.mjs";
import { buildNativeDiagramCandidate } from "../../scripts/research/native-diagram-package.mjs";

// Allocate a distinct namespace prefix for this diagram in the owning app.
let state = createDiagramIdentityState("cycle", "12345678-1234-4ABC-8DEF");
const candidate = {
  family: "cycle", closure: "layout-generated", roots: ["owned-a", "owned-b"],
  nodes: ["owned-a", "owned-b"].map(id => ({ id, role: "content", paragraphs: [id], children: [] })),
};
// components must be explicitly supplied admitted layout/colors/style XML.
const result = await buildNativeDiagramCandidate(candidate, components, {
  presentation: "layout", modelIds: "guid", identityState: state,
});
// Commit bytes and state together in the caller's serialized transaction.
// This assignment alone is not a filesystem publication transaction.
state = result.specimen.identityState;
```

Reuse the returned state on subsequent edits; do not recreate it or use Parser's traversal-local IDs as managed node IDs. Retired assignments are retained for explicit restoration and never reassigned to different keys. Distinct new objects need distinct candidate IDs. Independent diagrams require independent namespaces. This snapshot is neither a source-identity importer nor concurrent merge authority.

After `pnpm build`, run:

```sh
rtk proxy node --test scripts/diagram-identity-research.test.mjs scripts/native-diagram-research.test.mjs
rtk proxy pnpm test:diagram-research
rtk proxy pnpm test:import
```

The focused command runs five pure identity and fifteen native research groups. Structural edit tests inspect 48 packages through Parser and independent browser DOM; initial DOM coverage spans 36 combinations. Existing stateless controls remain tested. Seeded allocation tests use an independent never-reuse map. Failures include duplicate/foreign/nonmonotonic IDs, corrupt version/family, unknown fields, duplicate active keys, retained/active byte/work/entry budgets, final 48-bit allocation, failed package preparation and retry. Exact work-boundary search checks cumulative reconciliation plus XML work instead of copying implementation checkpoint counts.

No native runtime is launched by these tests. Subsequent native checks must exercise stateful output and structural changes explicitly; earlier stateless/native Cycle observations cannot certify stateful generation.


The later [scoped source-binding adapter](smartart-source-bindings.md) can supply managed candidate IDs from recognized source identities with explicit context/retirement state. It does not turn the older traversal IDs into persistent IDs or grant native source-write authority.
