import { createPresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function createFadeDeck() {
  const deck = createPresentation();
  deck.addSlide().addElement({
    type: "text", content: "First page",
    box: { x: 40, y: 40, width: 400, height: 60 },
  });
  const second = deck.addSlide({
    transition: { type: "fade", durationMs: 1234, onClick: true, afterMs: 5000 },
  });
  second.addElement({
    type: "text", content: "Fade into this page",
    box: { x: 40, y: 40, width: 400, height: 60 },
  });
  return generatePptx(deck);
}
