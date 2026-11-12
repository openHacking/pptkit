import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

const output = process.argv[2];
if (!output) throw new Error("Provide an output directory for the native pagination probe.");
const policy = process.argv[3];
if (policy !== undefined && policy !== "source-boundaries") throw new Error("Unknown native row-region policy.");
const presentation = createPresentation({ size: { width: 400, height: 400 } });
for (const id of ["source", "paged"]) {
  const rows = Array.from({ length: 5 }, (_, index) => ({ height: 40, cells: [{ content: `ROW ${index}` }] }));
  presentation.addSlide({ id }).addElement({
    id: `table-${id}`, type: "table", box: { x: 40, y: 60, width: 240, height: 200 }, columns: [240], rows,
    nativeFirstRowStyle: true, nativeTableStyleFlags: { lastRow: true, bandedRows: true }, nativeTableStylePreset: "medium-2-accent-2",
  });
}
const layout = { tablePagination: [{ elementId: "table-paged", area: { x: 40, y: 60, width: 240, height: 80 }, headerRows: 1, ...(policy === undefined ? {} : { nativeRowRegions: policy }) }] };
const result = await generatePptx(presentation, { layout });
if (result.slideCount !== 5 || !result.warnings.some(w => w.code === "pagination-native-row-style-reapplied")) throw new Error("Unexpected pagination probe result.");
await mkdir(output, { recursive: true });
await writeFile(path.join(output, "source.pptx"), result.bytes, { flag: "wx" });
