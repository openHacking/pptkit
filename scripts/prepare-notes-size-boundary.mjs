import { writeFile } from "node:fs/promises";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

const target = process.argv[2], shape = process.argv[3];
const scenarios = { portrait: [72, 4032], landscape: [4032, 72], rounded: [72 - 0.4 / 12700, 4032 + 0.4 / 12700] };
if (!target || !Object.hasOwn(scenarios, shape)) throw new Error("Provide a new PPTX path and portrait, landscape or rounded.");
const [width, height] = scenarios[shape];
const deck = createPresentation({ id: `notes-size-${shape}`, size: { width, height } });
deck.addSlide({ id: "first", notes: `NOTES-${shape}-FIRST`, elements: [{ id: "label", type: "text", box: { x: 10, y: 10, width: 50, height: 25 }, content: `SLIDE-${shape}-FIRST` }] });
deck.addSlide({ id: "second", notes: `NOTES-${shape}-SECOND`, elements: [{ id: "label-2", type: "text", box: { x: 10, y: 10, width: 50, height: 25 }, content: `SLIDE-${shape}-SECOND` }] });
await writeFile(target, (await generatePptx(deck)).bytes, { flag: "wx" });
