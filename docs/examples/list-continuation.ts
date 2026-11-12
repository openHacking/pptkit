import { createPresentation } from "@pptkit/core";

export const deck = createPresentation();
deck.addSlide().addElement({
  type: "text", box: { x: 40, y: 40, width: 300, height: 120 },
  content: [
    { style: { bullet: { type: "number", startAt: 4 } }, runs: [{ text: "First" }] },
    { style: { bullet: { type: "number", startAt: "continue" } }, runs: [{ text: "Next" }] },
  ],
});
