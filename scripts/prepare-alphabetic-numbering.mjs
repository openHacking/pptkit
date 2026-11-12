import { writeFile } from "node:fs/promises";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
const file = process.argv[2]; if (!file) throw Error("Provide a new PPTX output path.");
const deck = createPresentation(), slide = deck.addSlide();
for (const [index, startAt] of [26, 27, 28, 52, 53, 54].entries()) slide.addElement({ id: `number-${startAt}`, type: "text", box: { x: 50, y: 20 + index * 60, width: 500, height: 50 }, content: [{ style: { bullet: { type: "number", style: "alphaLowerPeriod", startAt } }, runs: [{ text: `VALUE${startAt}` }] }] });
await writeFile(file, (await generatePptx(deck)).bytes, { flag: "wx" });
