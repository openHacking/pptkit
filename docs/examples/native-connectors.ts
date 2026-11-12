import { createPresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function connectedShapes() {
  const document = createPresentation();
  const slide = document.addSlide();
  slide.addElement({ type: "shape", id: "left", shape: "rect", box: { x: 40, y: 40, width: 80, height: 40 } });
  slide.addElement({ type: "shape", id: "right", shape: "ellipse", box: { x: 260, y: 40, width: 80, height: 40 } });
  slide.addElement({ type: "connector", start: { elementId: "left", anchor: "right" },
    end: { elementId: "right", anchor: "left" },
    style: { paint: { type: "solid", color: "225588" }, width: 2, endArrow: "triangle" } });
  return generatePptx(document, { connectorMode: "native" });
}
