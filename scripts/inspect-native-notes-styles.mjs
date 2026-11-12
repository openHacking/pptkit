import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { normalizePresentation } from "../packages/core/dist/index.js";

if (process.argv.length !== 3) throw Error("Usage: node scripts/inspect-native-notes-styles.mjs <edit-run-directory>");
const root = path.resolve(process.argv[2]), exec = promisify(execFile), groups = [];
const options = { projection: "content", notesTextProfile: "master-style-v1", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" };
for (const variant of ["A", "B"]) {
  const directory = path.join(root, variant), expected = JSON.parse(await readFile(path.join(directory, "expected.json"), "utf8"));
  for (const oracle of ["native-notes-styles.py", "native-notes-styles-checks.py"]) {
    const result = await exec("python3", [path.resolve("scripts/research/oracles", oracle), directory], { timeout: 120000, maxBuffer: 4 * 1024 * 1024 });
    console.log(variant, oracle, result.stdout.trim());
  }
  const independent = JSON.parse(await readFile(path.join(directory, "styles-inspection.json"), "utf8"));
  const negatives = JSON.parse(await readFile(path.join(directory, "styles-negative-inspection.json"), "utf8"));
  const inputs = [];
  for (const stage of ["bold", "italic", "regular"]) for (let round = 1; round <= 3; round++) {
    const bytes = await readFile(path.join(directory, `${stage}-round${round}.pptx`)), s = await importPptx(bytes, options);
    try {
      const document = normalizePresentation(s.document);
      for (const [index, control] of expected.controls.entries()) {
        const runs = document.slides[index].notes[0].runs, edited = expected.editTargets.includes(index + 1);
        assert.equal(runs.length, edited ? 2 : 1);
        assert.equal(runs[0].text, control.text);
        assert.equal(runs[0].style.language, control.language);
        assert.equal(runs[0].style.bold, false);
        assert.equal(runs[0].style.italic, false);
        if (edited) {
          assert.equal(runs[1].text, "Ω");
          assert.equal(runs[1].style.language, "en-US");
          assert.equal(runs[1].style.bold, stage === "bold");
          assert.equal(runs[1].style.italic, stage === "italic");
        }
      }
      const plan = await s.prepareSave({ mode: "no-op" });
      assert.equal(plan.status, "ready");
      assert.deepEqual(Buffer.from((await s.save(plan.plan)).bytes), bytes);
      inputs.push({ stage, round, sha256: createHash("sha256").update(bytes).digest("hex"), exactNoOp: true, originalRunRegular: true, omegaOnlyProjectedFlags: true });
    } finally { s.dispose(); }
  }
  groups.push({ variant, independent, negatives, inputs });
}
const application = JSON.parse((await exec("python3", ["-c", "import json,plistlib,sys; p=plistlib.load(open(sys.argv[1],'rb')); print(json.dumps({k:p[k] for k in ['CFBundleShortVersionString','CFBundleVersion']}))", "/Applications/Microsoft PowerPoint.app/Contents/Info.plist"])).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), root, application, groups,
  total: { nativeStages: 6, nativeRounds: 18, nativeNotesPages: 270, uniqueStyledTargets: 22, exactNoOps: 18, actualCorruptionsRejected: groups.reduce((n, g) => n + g.negatives.rejectedCount, 0) },
  scope: "Captured native Ω-only bold, pure italic and restored regular; source style writes remain blocked. No full original-text styling, theme response, complete notes artwork or font-equivalence certification. Cross-consumer source regressions and append candidates are separate evidence." };
await writeFile("docs/testing/native-notes-styles-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify(evidence.total));
