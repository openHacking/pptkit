import { createPresentation } from "@pptkit/core";

export const deck = createPresentation();
deck.addSlide().addElement({
  type: "text", box: { x: 40, y: 40, width: 300, height: 120 },
  content: [{
    style: { tabStops: [{ position: 72, align: "left" }, { position: 180, align: "decimal" }] },
    runs: [{ text: "Item\tPrice\t12.50" }],
  }],
});
