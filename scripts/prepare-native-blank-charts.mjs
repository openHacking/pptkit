import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
await mkdir(".tmp/native-chart-blanks", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/native-chart-blanks/powerpoint-"));
const deck = createPresentation({ id: "blank-chart-native", size: { width: 960, height: 540 } });
for (const kind of ["bar", "line", "pie"]) {
  const slide = deck.addSlide({ id: kind });
  slide.addElement({ id: `title-${kind}`, type: "text", box: { x: 40, y: 20, width: 800, height: 45 }, content: `${kind.toUpperCase()}: 4, blank, 0, 3` });
  slide.addElement({ id: `chart-${kind}`, name: `chart-${kind}`, type: "chart", chartType: kind, box: { x: 50, y: 80, width: 800, height: 400 }, categories: ["A", "BLANK", "ZERO", "D"], series: [{ name: "Series", values: [4, null, 0, 3] }], legend: { visible: true } });
}
const generated = await generatePptx(deck); await writeFile(path.join(directory, "blank-charts.pptx"), generated.bytes, { flag: "wx" });
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ sha256: createHash("sha256").update(generated.bytes).digest("hex"), values: [4, null, 0, 3], warnings: generated.warnings }, null, 2)+"\n", { flag: "wx" });
console.log(directory);
