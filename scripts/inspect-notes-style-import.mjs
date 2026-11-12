// Source-only independent notes style/theme evidence; no native appearance claim.
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { notesStyleFixture, addSecondNotes } from "../packages/pptx-parser/test/notes-style-fixture.mjs";
import { encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";

if (process.argv.length !== 2) throw new Error("Usage: node scripts/inspect-notes-style-import.mjs");
await mkdir(".tmp/notes-style-import", { recursive: true }); const directory = await mkdtemp(path.resolve(".tmp/notes-style-import/run-"));
const options = { projection: "content", notesTextProfile: "master-style-v1", textFontProfile: "script-families-v2", editSaveProfile: "text-runs-v3" }, snapshots = [], hashes = [];
let bytes = Uint8Array.from(encodeStructure(addSecondNotes(notesStyleFixture(), true)));
for (let round = 0; round <= 3; round++) {
  const file = `round-${round}.pptx`; await writeFile(path.join(directory, file), bytes); hashes.push({ file, sha256: createHash("sha256").update(bytes).digest("hex") });
  const session = await importPptx(bytes, options);
  try {
    snapshots.push({ file, document: normalizePresentation(session.document), structureFindings: session.report.inspection.findings, properties: session.report.textStyleInheritance.properties.filter(p => p.target.part.includes("/notes/")) });
    if (round === 3) break;
    for (const [index, slide] of session.document.slides.entries()) { const notes = structuredClone(slide.notes); notes[0].runs[0].text = `Round${round + 1} notes${index + 1} 中عא`; slide.setNotes(notes); }
    const slide = session.document.slides[0], element = structuredClone(slide.elements[0]); (element.type === "text" ? element.content : element.text.content)[0].runs[0].text = `Round${round + 1} drawing`; slide.replaceElement(element.id, element);
    const plan = await session.prepareSave({ mode: "edit" }); if (plan.status !== "ready") throw new Error("Mixed inherited notes save rejected."); bytes = (await session.save(plan.plan)).bytes;
  } finally { session.dispose(); }
}
const expected = path.join(directory, "expected.json"); await writeFile(expected, JSON.stringify(snapshots));
const independent = JSON.parse((await promisify(execFile)("python3", ["scripts/research/oracles/notes-style-import.py", expected, directory], { timeout: 30000, maxBuffer: 4 * 1024 * 1024 })).stdout);
const evidence = { schemaVersion: 1, sourceConformance: { classification: "non-conforming-multi-notes-master-control", standardReference: "https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document#notes-master-part", reason: "PresentationML permits at most one Notes Master part. This fixture deliberately keeps a second unlisted actual master; source projection/preservation does not certify package compliance.", findings: snapshots[0].structureFindings.filter(f => f.code === "unlisted-master") }, recordedAt: new Date().toISOString(), artifactsDirectory: path.relative(process.cwd(), directory), profile: options, hashes, independent, nativeAppearanceAndEditing: "untested", fullNotesPageLayout: "unsupported", notesStyleWrites: "blocked", scope: "Non-conforming source control with two actual notes masters/themes, type-based placeholders with unequal indices, selected style/font/color rules and three mixed scalar source saves. No native layout or typography certification." };
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n"); await writeFile("docs/testing/notes-style-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, ...independent }));
