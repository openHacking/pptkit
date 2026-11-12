import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createHash } from "node:crypto";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { notesInsertionStyles } from "../tests/fixtures/notes-insertion-styles.mjs";

const root = path.resolve(".tmp/notes-insertion-styles"); await mkdir(root, { recursive: true });
const directory = await mkdtemp(path.join(root, "run-"));
const { document, cases } = notesInsertionStyles(createPresentation), { bytes, warnings } = await generatePptx(document);
await writeFile(path.join(directory, "source.pptx"), bytes);
await writeFile(path.join(directory, "expected.json"), JSON.stringify({ cases, warnings, sourceSha256: createHash("sha256").update(bytes).digest("hex"), sourceAssembly: "Public authoring and fresh generatePptx; no XML/ZIP patches.", nativeOperation: "Three no-edit Save As/close/reopen rounds, then append a named word into each empty notes paragraph and three further rounds." }, null, 2) + "\n");
await writeFile(path.join(root, "latest.txt"), directory + "\n"); console.log(directory);
