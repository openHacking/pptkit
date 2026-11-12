import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { authoredNotesPages } from "../tests/fixtures/authored-notes-pages.mjs";

const root = path.resolve(".tmp/authored-notes-pages");
await mkdir(root, { recursive: true });
const directory = await mkdtemp(path.join(root, "run-"));
for (const [name, size] of [["wide", { width: 720, height: 260 }], ["portrait", { width: 360, height: 720 }], ["square", { width: 500, height: 500 }]]) {
  const document = authoredNotesPages(createPresentation, size, name), { bytes, warnings } = await generatePptx(document);
  const file = path.join(directory, name + ".pptx"); await writeFile(file, bytes);
  await writeFile(path.join(directory, name + "-expected.json"), JSON.stringify({
    name, size, file, sha256: createHash("sha256").update(bytes).digest("hex"), warnings,
    sourceAssembly: "Public createPresentation/duplicateSlide/setNotes/generatePptx only; no XML or ZIP patches.",
    controls: document.slides.map(slide => ({ id: slide.id, label: slide.elements[0].content[0].runs[0].text, notes: slide.notes })),
    nativeReopen: "not tested",
  }, null, 2) + "\n");
}
await writeFile(path.join(root, "latest.txt"), directory + "\n");
console.log(directory);
