import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

await mkdir(".tmp/native-tab-stops", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/native-tab-stops/powerpoint-matrix-"));
const deck = createPresentation({ id: "tab-wrap-width-matrix", size: { width: 960, height: 540 } });
const controls = [
  { align: "left", words: ["I WORD", "WIDE WORD", "MIXED WORD"] },
  { align: "center", words: ["I WORD", "WIDE WORD", "MIXED WORD"] },
  { align: "right", words: ["I", "WIDE", "MIXED"] },
  { align: "decimal", words: ["1.2 WORD", "123.45 WORD", "0.678 WORD"] },
];
const widths = [135, 150, 180, 240];
for (const width of widths) for (const control of controls) {
  const slide = deck.addSlide({ id: `${control.align}-${width}` });
  const prefix = `${control.align}-${width}`;
  slide.addElement({ id: `guide-${prefix}`, type: "shape", shape: "rect", box: { x: 179.5, y: 90, width: 1, height: 420 }, style: { fill: { type: "solid", color: "CC2200" }, stroke: { paint: { type: "none" } } } });
  slide.addElement({ id: `title-${prefix}`, type: "text", box: { x: 60, y: 25, width: 700, height: 45 }, frame: { margin: 0 }, content: `TAB ${control.align.toUpperCase()} — width ${width}, stop 120` });
  for (const [index, word] of control.words.entries()) slide.addElement({ id: `${prefix}-${index}`, type: "text", box: { x: 60, y: 100 + index * 130, width, height: 120 }, frame: { margin: 0, wrap: true }, content: [{ style: { tabStops: [{ position: 120, align: control.align }] }, runs: [{ text: "P\t", style: { fontFamily: "Arial", fontSize: 28 } }, { text: word, style: { fontFamily: "Arial", fontSize: 28, bold: index === 2 } }] }] });
}
const result = await generatePptx(deck);
await writeFile(path.join(directory, "matrix.pptx"), result.bytes, { flag: "wx" });
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ sha256: createHash("sha256").update(result.bytes).digest("hex"), controls, widths, tabPosition: 120, warnings: result.warnings }, null, 2) + "\n", { flag: "wx" });
console.log(directory);
