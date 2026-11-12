import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

await mkdir(".tmp/native-tab-stops", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/native-tab-stops/powerpoint-position-"));
const deck = createPresentation({ id: "tab-position-matrix", size: { width: 960, height: 540 } });
const controls = [
  { align: "center", words: ["I WORD", "WIDE WORD", "MIXED WORD"] },
  { align: "right", words: ["I WORD", "WIDE WORD", "MIXED WORD"] },
  { align: "decimal", words: ["1.2 WORD", "123.45 WORD", "0.678 WORD"] },
];
const stops = [90, 120, 180, 240], width = 150;
for (const position of stops) for (const control of controls) {
  const prefix = `${control.align}-${position}`, slide = deck.addSlide({ id: prefix });
  slide.addElement({ id: `guide-${prefix}`, type: "shape", shape: "rect", box: { x: 60 + position - 0.5, y: 90, width: 1, height: 420 }, style: { fill: { type: "solid", color: "CC2200" }, stroke: { paint: { type: "none" } } } });
  slide.addElement({ id: `title-${prefix}`, type: "text", box: { x: 60, y: 25, width: 700, height: 45 }, frame: { margin: 0 }, content: `TAB ${control.align.toUpperCase()} — width ${width}, stop ${position}` });
  for (const [index, word] of control.words.entries()) slide.addElement({ id: `${prefix}-${index}`, type: "text", box: { x: 60, y: 100 + index * 130, width, height: 120 }, frame: { margin: 0, wrap: true }, content: [{ style: { tabStops: [{ position, align: control.align }] }, runs: [{ text: "P\t", style: { fontFamily: "Arial", fontSize: 28 } }, { text: word, style: { fontFamily: "Arial", fontSize: 28, bold: index === 2 } }] }] });
}
const result = await generatePptx(deck);
await writeFile(path.join(directory, "positions.pptx"), result.bytes, { flag: "wx" });
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ sha256: createHash("sha256").update(result.bytes).digest("hex"), controls, stops, width, warnings: result.warnings }, null, 2) + "\n", { flag: "wx" });
console.log(directory);
