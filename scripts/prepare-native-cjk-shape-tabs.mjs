import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

await mkdir(".tmp/native-tab-stops", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/native-tab-stops/powerpoint-cjk-shape-"));
const deck = createPresentation({ id: "native-shape-tabs", size: { width: 960, height: 540 } });
const controls = [
  { align: "left", words: ["一", "销售", "总收入"] },
  { align: "center", words: ["一", "销售", "总收入"] },
  { align: "right", words: ["一", "销售", "总收入"] },
  { align: "decimal", words: ["1.2", "123.45", "0.678"] },
];
for (const control of controls) {
  const slide = deck.addSlide({ id: control.align });
  slide.addElement({ id: `title-${control.align}`, type: "text", box: { x: 60, y: 30, width: 700, height: 50 }, frame: { margin: 0 }, content: `SHAPE TAB ${control.align.toUpperCase()} — guide at 340 pt` });
  for (const [index, word] of control.words.entries()) {
    slide.addElement({ id: `${control.align}-${index}`, type: "shape", shape: "rect", box: { x: 60, y: 120 + index * 85, width: 700, height: 65 }, style: { fill: { type: "solid", color: "F2F5FA" }, stroke: { paint: { type: "none" } } }, text: { frame: { margin: 0, wrap: false }, content: [{ style: { tabStops: [{ position: 280, align: control.align }] }, runs: [{ text: "合计\t", style: { fontFamily: "Hiragino Sans GB", fontSize: 28 } }, { text: word, style: { fontFamily: "Hiragino Sans GB", fontSize: 28, bold: index === 2 } }] }] } });
  }
  slide.addElement({ id: `guide-${control.align}`, type: "shape", shape: "rect", box: { x: 339.5, y: 105, width: 1, height: 310 }, style: { fill: { type: "solid", color: "CC2200" }, stroke: { paint: { type: "none" } } } });
}
const result = await generatePptx(deck);
await writeFile(path.join(directory, "shapes.pptx"), result.bytes, { flag: "wx" });
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ sha256: createHash("sha256").update(result.bytes).digest("hex"), controls, position: 280, warnings: result.warnings }, null, 2) + "\n", { flag: "wx" });
console.log(directory);
