import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

await mkdir(".tmp/native-tab-list", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/native-tab-list/run-"));
const deck = createPresentation({ id: "tab-list-study", size: { width: 960, height: 540 } });
const bullets = [{ id: "symbol", value: { type: "bullet", character: "•" } }, { id: "number3", value: { type: "number", startAt: 3 } }, { id: "number333", value: { type: "number", startAt: 333 } }];
const indents = [{ indent: 36, hanging: 18 }, { indent: 27, hanging: 27 }, { indent: 54, hanging: 18 }];
const cases = [];
for (const bullet of bullets) for (const metrics of indents) for (const align of ["left", "center", "right", "decimal"]) {
  const index = cases.length, text = align === "decimal" ? "12.34" : "WIDE";
  const slide = deck.addSlide({ id: `s-${index}` });
  slide.addElement({ id: `sample-${index}`, type: "text", box: { x: 60, y: 100, width: 800, height: 80 }, frame: { margin: 0, wrap: false }, content: [{
    style: { ...metrics, bullet: bullet.value, tabStops: [{ position: 280, align }] },
    runs: [{ text: "P\t", style: { fontFamily: "Arial", fontSize: 28 } }, { text, style: { fontFamily: "Arial", fontSize: 28 } }],
  }] });
  cases.push({ index, bullet: bullet.id, ...metrics, align, text });
}
const output = await generatePptx(deck), file = path.join(directory, "list-tabs.pptx");
await writeFile(file, output.bytes, { flag: "wx" });
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ sourceSha256: createHash("sha256").update(output.bytes).digest("hex"), cases, warnings: output.warnings }, null, 2) + "\n", { flag: "wx" });
console.log(directory);
