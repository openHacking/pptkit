// Independent source evidence only. Does not certify native font/shaping behavior.
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { mixedDeclaredSlotBytes } from "../tests/fixtures/declared-slot-font-import.mjs";

if (process.argv.length !== 2) throw new Error("Usage: node scripts/inspect-declared-slot-font-import.mjs");
await mkdir(".tmp/declared-slot-font-import", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/declared-slot-font-import/run-")), snapshots = [], hashes = [];
const options = { projection: "content", textFontProfile: "script-families-v3", notesTextProfile: "master-style-v1", editSaveProfile: "text-runs-v3" };
let bytes = await mixedDeclaredSlotBytes();
for (let round = 0; round <= 3; round++) {
  const file = `round-${round}.pptx`; await writeFile(path.join(directory, file), bytes); hashes.push({ file, sha256: createHash("sha256").update(bytes).digest("hex") });
  const session = await importPptx(bytes, options);
  try {
    snapshots.push({ file, document: normalizePresentation(session.document), fonts: [...session.report.textStyleInheritance.properties, ...session.report.tableStyleInheritance.properties].filter(p => p.property.startsWith("run.scriptFontFamilies.")) });
    if (round === 3) break;
    const slide = session.document.slides[0]; let owner = 0;
    const change = paragraphs => { for (const p of paragraphs) for (const r of p.runs) r.text = `Round${round + 1} owner${++owner} 中عאΩ`; };
    for (const element of [...slide.elements]) {
      const next = structuredClone(element);
      if (next.type === "text") change(next.content);
      else if (next.type === "shape") change(next.text.content);
      else if (next.type === "table") for (const row of next.rows) for (const cell of row.cells) change(cell.content);
      slide.replaceElement(next.id, next);
    }
    const notes = structuredClone(slide.notes); change(notes); slide.setNotes(notes);
    const plan = await session.prepareSave({ mode: "edit" }); if (plan.status !== "ready") throw new Error("Scalar preservation rejected."); bytes = (await session.save(plan.plan)).bytes;
  } finally { session.dispose(); }
}
const expected = path.join(directory, "expected.json"); await writeFile(expected, JSON.stringify(snapshots));
const independent = JSON.parse((await promisify(execFile)("python3", ["scripts/research/oracles/declared-slot-font-import.py", expected, directory], { timeout: 30000, maxBuffer: 4 * 1024 * 1024 })).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), artifactsDirectory: path.relative(process.cwd(), directory), profile: options, hashes, independent, native: "untested", fontCoverageAndShaping: "untested", sourceFontWrites: "blocked", scope: "Named generated mixed-owner source: declared-slot mapped/blocked provenance, actual scalar text and exact non-character bytes across three source saves. No native typography claim." };
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n");
await writeFile("docs/testing/declared-slot-font-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, ...independent }));
