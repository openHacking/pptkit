import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

if (process.argv.length !== 3) throw Error("Usage: node scripts/inspect-native-notes-append.mjs <edit-run-directory>");
const root = path.resolve(process.argv[2]), exec = promisify(execFile);
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const options = { projection: "content", notesTextProfile: "master-style-v1", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" };
const groups = [];
for (const variant of ["A", "B"]) {
  const directory = path.join(root, variant), expected = JSON.parse(await readFile(path.join(directory, "expected.json"), "utf8"));
  for (const oracle of ["native-notes-append.py", "native-notes-append-checks.py"]) {
    const result = await exec("python3", [path.resolve("scripts/research/oracles", oracle), directory], { timeout: 120000, maxBuffer: 4 * 1024 * 1024 });
    console.log(variant, oracle, result.stdout.trim());
  }
  const independent = JSON.parse(await readFile(path.join(directory, "append-inspection.json"), "utf8"));
  const negatives = JSON.parse(await readFile(path.join(directory, "append-negative-inspection.json"), "utf8"));
  const inputs = [];
  for (let round = 1; round <= 3; round++) {
    const bytes = await readFile(path.join(directory, `append-round${round}.pptx`)), session = await importPptx(bytes, options);
    try {
      const document = normalizePresentation(session.document), observations = [];
      for (const [index, control] of expected.controls.entries()) {
        const paragraphs = document.slides[index].notes;
        assert.equal(paragraphs.length, 1);
        const runs = paragraphs[0].runs, edited = expected.editTargets.includes(index + 1);
        assert.equal(runs.length, edited ? 2 : 1);
        assert.equal(runs[0].text, control.text);
        assert.equal(runs[0].style.language, control.language);
        assert.ok(runs.every(run => run.style.fontSize === 40 && !run.style.bold && !run.style.italic));
        if (edited) {
          assert.equal(runs[1].text, "Ω");
          assert.equal(runs[1].style.language, "en-US");
          const selected = resolveTextFontSegments("Ω", runs[1].style, document.theme);
          assert.equal(selected.length, 1);
          const actual = independent.rounds[round - 1].pdf.pages[index].actualFont;
          const selectedPostscript = expected.fonts.find(font => font.family === selected[0].family)?.postscriptName;
          observations.push({ slide: index + 1, sourceRun: runs[1], selected, actualNativeFace: actual, declaredSelectionMatchesNative: selectedPostscript === actual });
        }
      }
      const plan = await session.prepareSave({ mode: "no-op" });
      assert.equal(plan.status, "ready");
      assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
      const slide = session.document.slides[2], original = structuredClone(slide.notes), changed = structuredClone(original);
      changed[0].runs[1].style.bold = true;
      slide.setNotes(changed);
      assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
      slide.setNotes(original);
      let previous = bytes;
      const candidates = [];
      for (let edit = 1; edit <= 3; edit++) {
        const next = structuredClone(slide.notes);
        next[0].runs[1].text += "X";
        slide.setNotes(next);
        const prepared = await session.prepareSave({ mode: "edit" });
        assert.equal(prepared.status, "ready");
        const saved = (await session.save(prepared.plan)).bytes, before = readGeneratedZip(previous), after = readGeneratedZip(saved);
        assert.deepEqual([...after.keys()], [...before.keys()]);
        for (const [part, payload] of before) {
          if (part === "ppt/notesSlides/notesSlide3.xml") {
            const from = `<a:t>Ω${"X".repeat(edit - 1)}</a:t>`;
            assert.equal(payload.toString().split(from).length, 2);
            assert.equal(after.get(part).toString(), payload.toString().replace(from, from.replace("</a:t>", "X</a:t>")));
          } else assert.deepEqual(after.get(part), payload);
        }
        const reopened = await importPptx(saved, options);
        try { assert.deepEqual(normalizePresentation(reopened.document), normalizePresentation(session.document)); }
        finally { reopened.dispose(); }
        candidates.push({ round: edit, sha256: hash(saved), nonCharacterPayloadsExact: true, fullReprojectionEqual: true, nativeReopen: "not tested" });
        previous = saved;
      }
      inputs.push({ round, sha256: hash(bytes), exactNoOp: true, styleWriteBlockedAndRestored: true, observations, candidates });
    } finally { session.dispose(); }
  }
  groups.push({ variant, independent, negatives, inputs });
}
const application = JSON.parse((await exec("python3", ["-c", "import json,plistlib,sys; p=plistlib.load(open(sys.argv[1],'rb')); print(json.dumps({k:p[k] for k in ['CFBundleShortVersionString','CFBundleVersion']}))", "/Applications/Microsoft PowerPoint.app/Contents/Info.plist"])).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), root, application, groups,
  total: { nativeRounds: 6, nativeNotesPages: 90, uniqueAppendTargets: 22, exactNoOps: 6, sourceStyleWriteRejections: 6, scalarCandidates: 18, actualCorruptionsRejected: groups.reduce((n, g) => n + g.negatives.rejectedCount, 0), appendedSelectionMatches: groups.reduce((n, g) => n + g.inputs[0].observations.filter(o => o.declaredSelectionMatchesNative).length, 0) },
  scope: "Native append only; existing original traces and outside-append raster preserved. Scalar candidates are headless, not native reopened. Source-only three-consumer tests remain separate. Local styles, theme response, full notes artwork and source style writer remain open." };
await writeFile("docs/testing/native-notes-append-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify(evidence.total));
