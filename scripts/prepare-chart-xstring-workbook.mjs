import { writeFile } from "node:fs/promises";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const output = process.argv[2];
if (!output) throw new Error("Provide a new output XLSX path.");
const document = createPresentation(), slide = document.addSlide();
slide.addElement({ type: "chart", chartType: "bar", box: { x: 0, y: 0, width: 400, height: 300 },
  categories: ["_x0041_", "line\rbreak", "_x005F_x0041_", "plain & 中文", "_x0041__x0042_"],
  series: [{ name: "name _x000D_", values: [1, 2, 3, 4, 5] }] });
const pptx = await generatePptx(document), parts = readGeneratedZip(pptx.bytes);
await writeFile(output, parts.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"), { flag: "wx" });
