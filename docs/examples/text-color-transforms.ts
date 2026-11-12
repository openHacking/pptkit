import { createPresentation, normalizePresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function createTranslucentText() {
  const document = createPresentation();
  document.addSlide().addElement({
    type: "text",
    box: { x: 40, y: 40, width: 320, height: 80 },
    opacity: 0.5,
    content: [{
      runs: [
        { text: "Quarter opacity", style: { color: "336699", colorOpacity: 0.5 } },
        { text: "Element opacity", style: { color: "336699", colorOpacity: 1 } },
      ],
    }],
  });
  return {
    normalized: normalizePresentation(document),
    output: await generatePptx(document),
  };
}
