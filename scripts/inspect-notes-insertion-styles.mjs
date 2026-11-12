import assert from "node:assert/strict";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import path from "node:path";
import { createHash } from "node:crypto";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

assert.equal(process.argv.length, 3, "Usage: node scripts/inspect-notes-insertion-styles.mjs <captured-directory>");
const directory = path.resolve(process.argv[2]), hash = b => createHash("sha256").update(b).digest("hex");
const options = { projection: "content", editSaveProfile: "text-runs-v3", textFontProfile: "script-families-v3", notesTextProfile: "master-style-v1" };
const names = ["source", ...["empty", "typed"].flatMap(stage => [1, 2, 3].map(n => stage + "-round" + n))];
const records = [], output = path.join(directory, "parser-inspections"); await mkdir(output, { recursive: true });
for (const name of names) {
  const bytes = await readFile(path.join(directory, name + ".pptx")), session = await importPptx(bytes, options);
  try {
    const snapshot = { document: normalizePresentation(session.document), report: session.report };
    const notes = snapshot.document.slides.map(s => s.notes), runs = notes.map(n => n[0].runs.at(-1));
    assert.deepEqual(runs.map(r => [r.text, r.style.fontFamily, r.style.fontSize, r.style.bold, r.style.italic, r.style.language]), [[name.startsWith("typed") ? "Proof" : "", "Georgia", 14, false, false, "fr-FR"], [name.startsWith("typed") ? "Proof" : "", "Arial", 12, true, true, "en-US"], [name.startsWith("typed") ? "Proof" : "", "Arial", 18, false, false, "en-US"], [name.startsWith("typed") ? "Proof" : "", "Georgia", 16, false, true, "en-US"]]);
    const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready"); assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), bytes);
    const slide = session.document.slides[0], original = structuredClone(slide.notes), mixed = structuredClone(original);
    mixed[0].runs[0].text += "X"; mixed[0].runs[0].style.bold = !mixed[0].runs[0].style.bold;
    slide.setNotes(mixed); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); slide.setNotes(original);
    const restored = await session.prepareSave({ mode: "no-op" }); assert.equal(restored.status, "ready"); assert.deepEqual(Buffer.from((await session.save(restored.plan)).bytes), bytes);
    for (const s of session.document.slides) { const notes = structuredClone(s.notes); notes[0].runs.at(-1).text += "X"; s.setNotes(notes); }
    const candidate = await session.prepareSave({ mode: "edit" });
    let character;
    if (name.startsWith("empty")) {
      assert.equal(candidate.status, "blocked"); character = { status: "blocked", reason: "Native empty paragraphs contain only insertion properties; no bound a:t scalar owner exists." };
    } else {
      assert.equal(candidate.status, "ready"); const saved = (await session.save(candidate.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(saved);
      assert.deepEqual([...after.keys()], [...before.keys()]);
      for (const [part, body] of before) {
        if (/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(part)) {
          const text = name === "source" ? "<a:t></a:t>" : "<a:t>Proof</a:t>", target = text.replace("</a:t>", "X</a:t>");
          // The source two-empty control changes only its final empty run.
          const input = body.toString(), at = input.lastIndexOf(text); assert.ok(at >= 0);
          assert.equal(after.get(part).toString(), input.slice(0, at) + target + input.slice(at + text.length));
        } else assert.deepEqual(after.get(part), body);
      }
      const file = path.join(output, name + "-candidate.pptx"); await writeFile(file, saved);
      character = { status: "ready", file, sha256: hash(saved), nativeReopen: "not tested" };
    }
    const file = path.join(output, name + ".json"), serialized = JSON.stringify(snapshot, null, 2) + "\n"; await writeFile(file, serialized);
    records.push({ name, sha256: hash(bytes), snapshot: file, snapshotSha256: hash(serialized), exactNoOpAndRestore: true, mixedStyleRejected: true, character });
  } finally { session.dispose(); }
}
const result = { schemaVersion: 1, directory, options, records, total: { actualInputs: records.length, exactNoOps: records.length, mixedStyleRejectionsAndRestores: records.length, insertionOnlyCharacterRejections: records.filter(r => r.character.status === "blocked").length, materializedHeadlessCharacterCandidates: records.filter(r => r.character.status === "ready").length }, scope: "Parser verification of captured native files; no native reopening of generated character candidates and no source style permission." };
await writeFile("docs/testing/notes-insertion-parser-evidence.json", JSON.stringify(result, null, 2) + "\n"); console.log(JSON.stringify(result.total));
