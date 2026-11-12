import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

await mkdir(".tmp/native-tab-stops", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/native-tab-stops/run-"));
const deck = createPresentation({ id: "native-tab-study", size: { width: 960, height: 540 } });
const samples = [
  { id: "left", align: "left", samples: ["I", "WIDE", "MIXED"] },
  { id: "center", align: "center", samples: ["I", "WIDE", "MIXED"] },
  { id: "right", align: "right", samples: ["I", "WIDE", "MIXED"] },
  { id: "decimal", align: "decimal", samples: ["1.2", "123.45", "0.678"] },
];
for (const item of samples) {
  const slide = deck.addSlide({ id: item.id });
  slide.addElement({ id: `guide-${item.id}`, type: "shape", shape: "rect", box: { x: 339.5, y: 105, width: 1, height: 310 }, style: { fill: { type: "solid", color: "CC2200" }, stroke: { paint: { type: "none" } } } });
  slide.addElement({ id: `title-${item.id}`, type: "text", box: { x: 60, y: 30, width: 700, height: 50 }, frame: { margin: 0 }, content: `TAB ${item.align.toUpperCase()} — red guide at 340 pt` });
  for (const [index, sample] of item.samples.entries()) {
    slide.addElement({ id: `${item.id}-${index}`, type: "text", box: { x: 60, y: 120 + index * 85, width: 700, height: 65 }, frame: { margin: 0, wrap: false }, content: [{
      style: { tabStops: [{ position: 280, align: item.align }] },
      runs: [{ text: "P\t", style: { fontFamily: "Arial", fontSize: 28 } }, { text: sample, style: { fontFamily: "Arial", fontSize: 28, bold: index === 2 } }],
    }] });
  }
}
const output = await generatePptx(deck);
const bytes = output.bytes, filename = "tab-stops.pptx";
await writeFile(path.join(directory, filename), bytes, { flag: "wx" });
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ filename, sha256: createHash("sha256").update(bytes).digest("hex"), guideX: 340, textLeft: 60, tabPosition: 280, slides: samples, warnings: output.warnings }, null, 2) + "\n", { flag: "wx" });
console.log(directory);
