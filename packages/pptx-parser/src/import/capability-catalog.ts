import type { ImportCapabilityReferences } from "../types/capability.js";
// Reference hints only. Neither this index nor artifact presence determines tiers.
const rules: readonly ImportCapabilityReferences[] = [
  { rule: "source-", match: "prefix", tasks: ["IMP-07.1", "IMP-09.1"], artifacts: [{ kind: "contract", path: "docs/api/import-session.md" }, { kind: "test", path: "packages/pptx-parser/test/session.test.mjs" }] },
  { rule: "metadata.", match: "prefix", tasks: ["IMP-03.4"], artifacts: [{ kind: "contract", path: "docs/api/pptx-inspection.md" }, { kind: "test", path: "packages/pptx-parser/test/metadata.test.mjs" }] },
  { rule: "presentation-dimensions", match: "exact", tasks: ["IMP-03.1"], artifacts: [{ kind: "contract", path: "docs/api/import-session.md" }, { kind: "test", path: "packages/pptx-parser/test/session.test.mjs" }] },
  { rule: "first-slide-number", match: "exact", tasks: ["AUT-08.1.3"], artifacts: [{ kind: "contract", path: "docs/api/first-slide-number-save.md" }, { kind: "test", path: "scripts/first-slide-number.test.mjs" }] },
  { rule: "slide-order-and-visibility", match: "exact", tasks: ["IMP-03.1"], artifacts: [{ kind: "contract", path: "docs/api/import-session.md" }, { kind: "test", path: "packages/pptx-parser/test/session.test.mjs" }] },
  { rule: "section.", match: "prefix", tasks: ["IMP-03.5"], artifacts: [{ kind: "contract", path: "docs/api/pptx-inspection.md" }, { kind: "test", path: "packages/pptx-parser/test/sections.test.mjs" }] },
  { rule: "slide.transition", match: "exact", tasks: ["ADV-05.1"], artifacts: [{ kind: "contract", path: "docs/api/slide-transitions.md" }, { kind: "test", path: "packages/pptx-parser/test/transitions.test.mjs" }] },
  { rule: "smartart.", match: "prefix", tasks: ["ADV-01.1.2", "ADV-03.7"], artifacts: [{ kind: "contract", path: "docs/api/diagram-preview.md" }, { kind: "test", path: "packages/pptx-parser/test/diagram-preview.test.mjs" }, { kind: "native-observation", path: "docs/testing/smartart-stale-cache-powerpoint-evidence.json" }] },
  { rule: "chart.", match: "prefix", tasks: ["IMP-06.2", "AUT-06.2"], artifacts: [{ kind: "contract", path: "docs/api/chart-import.md" }, { kind: "test", path: "packages/pptx-parser/test/charts.test.mjs" }] },
  { rule: "table.", match: "prefix", tasks: ["IMP-06.1", "IMP-09.2.5"], artifacts: [{ kind: "contract", path: "docs/api/table-import.md" }, { kind: "test", path: "packages/pptx-parser/test/tables.test.mjs" }] },
  { rule: "notes.", match: "prefix", tasks: ["IMP-03.3", "IMP-09.2.1.2"], artifacts: [{ kind: "contract", path: "docs/api/notes-import.md" }, { kind: "test", path: "packages/pptx-parser/test/content-notes.test.mjs" }] },
  { rule: "master-artwork.", match: "prefix", tasks: ["IMP-05.1.3.3.3", "AUT-07.1.3.2"], artifacts: [{ kind: "contract", path: "docs/api/master-artwork.md" }, { kind: "test", path: "packages/pptx-parser/test/master-artwork-import.test.mjs" }] },
  { rule: "artwork.background", match: "exact", tasks: ["IMP-05.1.3.3.3"], artifacts: [{ kind: "contract", path: "docs/api/master-artwork.md" }, { kind: "test", path: "packages/pptx-parser/test/master-artwork-import.test.mjs" }] },
  { rule: "artwork.background.", match: "prefix", tasks: ["IMP-05.1.3.3.3"], artifacts: [{ kind: "contract", path: "docs/api/master-artwork.md" }, { kind: "test", path: "packages/pptx-parser/test/master-artwork-import.test.mjs" }] },
  { rule: "drawing.", match: "prefix", tasks: ["IMP-04.4"], artifacts: [{ kind: "contract", path: "docs/api/content-import.md" }, { kind: "test", path: "packages/pptx-parser/test/content.test.mjs" }] },
];
export function capabilityReferences(feature: string): ImportCapabilityReferences | undefined {
  const rule = rules.find(r => r.match === "exact" ? feature === r.rule : feature.startsWith(r.rule));
  return rule ? { ...rule, tasks: [...rule.tasks], artifacts: rule.artifacts.map(a => ({ ...a })) } : undefined;
}
