import { createPresentation } from "@pptkit/core";
import { resolveLayout, type LayoutOptions } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

export async function createStackedCards() {
  const presentation = createPresentation({ size: { width: 600, height: 300 } });
  const slide = presentation.addSlide({ id: "cards" });
  for (const id of ["first", "second", "third"]) {
    slide.addElement({ type: "shape", id, shape: "roundRect", box: { x: 0, y: 0, width: 140, height: 100 } });
  }
  const layout: LayoutOptions = {
    composition: [{
      type: "stack", container: { kind: "slide", id: slide.id },
      elementIds: ["first", "second", "third"], direction: "horizontal",
      padding: 30, gap: 20, justify: "center", align: "center",
    }],
    diagnostics: { slideBounds: true },
  };
  const resolved = resolveLayout(presentation, layout);
  if (resolved.analysis?.diagnostics.length) return { diagnostics: resolved.analysis.diagnostics };
  const [pptx, preview] = await Promise.all([
    generatePptx(presentation, { layout }),
    renderPresentationToSvg(presentation, { layout }),
  ]);
  return { pptx, preview };
}
