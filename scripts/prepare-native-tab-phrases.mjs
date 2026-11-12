import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

await mkdir(".tmp/native-tab-stops", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/native-tab-stops/powerpoint-phrases-"));
const deck = createPresentation({ id: "tab-phrase-matrix", size: { width: 960, height: 540 } });
const controls = [
  { align: "left", words: ["I WIDE WORD", "WIDE WORD MORE", "MIXED WORD LONG"] },
  { align: "center", words: ["I WIDE WORD", "WIDE WORD MORE", "MIXED WORD LONG"] },
  { align: "right", words: ["I WIDE WORD", "WIDE WORD MORE", "MIXED WORD LONG"] },
  { align: "decimal", words: ["1.2 WIDE WORD", "123.45 WORD MORE", "0.678 WORD LONG"] },
];
const widths = [150, 220], position = 120;
for (const width of widths) for (const control of controls) {
  const prefix = `${control.align}-${width}`, slide = deck.addSlide({ id: prefix });
  slide.addElement({ id: `guide-${prefix}`, type: "shape", shape: "rect", box: { x: 179.5, y: 90, width: 1, height: 420 }, style: { fill: { type: "solid", color: "CC2200" }, stroke: { paint: { type: "none" } } } });
  slide.addElement({ id: `title-${prefix}`, type: "text", box: { x: 60, y: 25, width: 700, height: 45 }, frame: { margin: 0 }, content: `TAB ${control.align.toUpperCase()} — width ${width}, stop ${position}` });
  for (const [index, word] of control.words.entries()) slide.addElement({ id: `${prefix}-${index}`, type: "text", box: { x: 60, y: 100 + index * 130, width, height: 180 }, frame: { margin: 0, wrap: true }, content: [{ style: { tabStops: [{ position, align: control.align }] }, runs: [{ text: "P\t", style: { fontFamily: "Arial", fontSize: 28 } }, { text: word, style: { fontFamily: "Arial", fontSize: 28, bold: index === 2 } }] }] });
}
const result = await generatePptx(deck);
await writeFile(path.join(directory, "phrases.pptx"), result.bytes, { flag: "wx" });
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ sha256: createHash("sha256").update(result.bytes).digest("hex"), controls, widths, position, warnings: result.warnings }, null, 2) + "\n", { flag: "wx" });
console.log(directory);
