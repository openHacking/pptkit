import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

await mkdir(".tmp/native-doughnut", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/native-doughnut/powerpoint-"));
const deck = createPresentation({ id: "doughnut-native", size: { width: 960, height: 540 } });
for (const [id, holeSize, firstSliceAngle, series] of [
  ["rings", 50, 0, [{ name: "INNER", values: [1, 3], pointColors: ["#FF0000", "#0000FF"] }, { name: "OUTER", values: [3, 1], pointColors: ["#00AA00", "#FFAA00"] }]],
  ["rotate", 30, 90, [{ name: "ROTATED", values: [1, 3], pointColors: ["#FF0000", "#0000FF"] }]],
  ["blank", 70, 0, [{ name: "BLANKS", values: [2, null], pointColors: ["#FF0000", "#0000FF"] }]],
]) {
  const slide = deck.addSlide({ id });
  slide.addElement({ id: `title-${id}`, type: "text", box: { x: 30, y: 10, width: 800, height: 45 }, content: id.toUpperCase() });
  slide.addElement({ id: `chart-${id}`, name: `chart-${id}`, type: "chart", chartType: "doughnut", box: { x: 80, y: 70, width: 700, height: 420 }, categories: ["A", "B"], series, holeSize, firstSliceAngle, legend: { visible: false } });
}
const output = await generatePptx(deck); const file = path.join(directory, "doughnut.pptx");
await writeFile(file, output.bytes, { flag: "wx" });
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ sha256: createHash("sha256").update(output.bytes).digest("hex"), warnings: output.warnings }, null, 2) + "\n", { flag: "wx" });
console.log(directory);
