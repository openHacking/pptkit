// Builds an experimental package using a generated scatter chart's two value
// axes and relationships. No public bubble authoring API is used or implied.
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { bubbleSeriesParts } from "../packages/pptx-exporter/dist/ooxml/bubble-series.js";
import { createZip } from "../packages/pptx-exporter/dist/archive/create-zip.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const publicApi = process.argv.includes("--public");
const [output] = process.argv.slice(2).filter(value => value !== "--public");
const directory = path.resolve(output ?? (publicApi ? ".tmp/bubble-public-probe" : ".tmp/bubble-export-probe"));
await mkdir(directory, { recursive: true });
const series = [
  { name: "Blue", color: "#4472C4", xValues: [1, 4], yValues: [2, 4], sizes: [1, 9] },
  { name: "Orange", color: "#ED7D31", xValues: [7], yValues: [6], sizes: [4] },
];
const deck = createPresentation({ size: { width: 720, height: 405 } });
deck.addSlide({ elements: [{ type: "chart", chartType: publicApi ? "bubble" : "scatter", title: "Bubble size probe", box: { x: 40, y: 30, width: 640, height: 345 },
  axes: { x: { scale: { min: 0, max: 9, majorUnit: 1 } }, y: { scale: { min: 0, max: 8, majorUnit: 1 } } }, series }] });
if (publicApi) {
  await writeFile(path.join(directory, "bubble.pptx"), (await generatePptx(deck)).bytes);
} else {
  const parts = readGeneratedZip((await generatePptx(deck)).bytes);
  const chartName = "ppt/charts/chart1.xml", bookName = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
  const source = parts.get(chartName).toString();
  const { xml, workbook } = bubbleSeriesParts(series);
  const body = `<c:bubbleChart><c:varyColors val="0"/>${xml}<c:bubble3D val="0"/><c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:sizeRepresents val="area"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>`;
  const replaced = source.replace(/<c:scatterChart>[\s\S]*?<\/c:scatterChart>/, body);
  if (replaced === source || !parts.has(bookName)) throw new Error("Scatter carrier no longer matches the probe.");
  parts.set(chartName, new TextEncoder().encode(replaced));
  parts.set(bookName, workbook);
  await writeFile(path.join(directory, "bubble.pptx"), createZip([...parts].map(([name, data]) => ({ name, data }))));
}
