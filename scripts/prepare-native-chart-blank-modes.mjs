import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
await mkdir(".tmp/native-chart-blanks", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/native-chart-blanks/powerpoint-modes-"));
const deck = createPresentation({ id: "chart-blank-modes", size: { width: 960, height: 540 } });
for (const [kind, mode] of [["line", "gap"], ["line", "zero"], ["line", "span"], ["bar", "zero"], ["pie", "zero"]]) {
  const id = `${kind}-${mode}`, slide = deck.addSlide({ id });
  slide.addElement({ id: `title-${id}`, type: "text", box: { x: 40, y: 20, width: 800, height: 45 }, content: `${id.toUpperCase()}: 4, blank, 1, blank, 3` });
  slide.addElement({ id: `chart-${id}`, name: `chart-${id}`, type: "chart", chartType: kind, blankDisplay: mode, box: { x: 50, y: 80, width: 800, height: 400 }, categories: ["A", "BLANK1", "C", "BLANK2", "E"], series: [{ name: "Series", values: [4, null, 1, null, 3] }], legend: { visible: true }, ...(kind === "pie" ? {} : { axes: { value: { scale: { min: 0, max: 5, majorUnit: 1 } } } }) });
}
const generated = await generatePptx(deck); await writeFile(path.join(directory, "blank-modes.pptx"), generated.bytes, { flag: "wx" });
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ sha256: createHash("sha256").update(generated.bytes).digest("hex"), values: [4, null, 1, null, 3], warnings: generated.warnings }, null, 2)+"\n", { flag: "wx" }); console.log(directory);
