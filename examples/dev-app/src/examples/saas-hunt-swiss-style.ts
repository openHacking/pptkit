import type { ExampleDefinition } from "../example-types.js";
import { saasHuntSwissStyleInput } from "./saas-hunt-swiss-style-data.js";

export const saasHuntSwissStyleExample: ExampleDefinition = {
  id: "export-saas-hunt-swiss-style",
  title: "Export SaaS Hunt Swiss Style",
  feature: "export-pptx",
  description: "Reconstruct the supplied seven-slide Swiss Style deck with editable text, rectangles, and lines.",
  inputKind: "presentation-config",
  source: {
    label: "SaaS Hunt Swiss Style presentation config",
    content: JSON.stringify(saasHuntSwissStyleInput, null, 2),
  },
  scenarioTags: ["swiss-style", "editable", "reference-fidelity"],
  expectedCapabilities: {
    normalize: "implemented",
    render: "implemented",
    exportPptx: "implemented",
  },
  status: "ready",
  createInput() {
    return structuredClone(saasHuntSwissStyleInput);
  },
};
