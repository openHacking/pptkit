import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { chromium } from "@playwright/test";
import { normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

if (process.argv.length !== 3) throw Error("Usage: node scripts/inspect-native-table-font-styles.mjs <native-style-run-directory>");
const directory = path.resolve(process.argv[2]), exec = promisify(execFile), hash = bytes => createHash("sha256").update(bytes).digest("hex");
const options = { projection: "content", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" }, groups = [];
const browser = await chromium.launch();
try {
  for (const group of ["reference", "filled"]) {
    const folder = path.join(directory, group), expected = JSON.parse(await readFile(path.join(folder, "expected.json"), "utf8")); let negatives;
    for (const file of ["native-table-font-styles.py", "native-table-font-styles-checks.py"]) {
      const output = (await exec("python3", [path.resolve("scripts/research/oracles", file), folder], { timeout: 120000, maxBuffer: 4 * 1024 * 1024 })).stdout;
      if (file.endsWith("-checks.py")) negatives = JSON.parse(output); else console.log(output.trim());
    }
    const independent = JSON.parse(await readFile(path.join(folder, "style-inspection.json"), "utf8")), projections = [];
    const inputs = [{ name: "baseline" }, ...["bold", "italic", "regular"].flatMap(stage => [1, 2, 3].map(round => ({ name: `${stage}-round${round}`, stage, round })))];
    for (const input of inputs) {
      const bytes = await readFile(path.join(folder, input.name + ".pptx")), session = await importPptx(bytes, options); let snapshot, report;
      try {
        snapshot = normalizePresentation(session.document); report = session.report;
        const controls = [];
        for (const [index, control] of expected.controls.entries()) {
          const target = expected.editTargets.includes(index + 1), table = snapshot.slides[index].elements.find(e => e.name === "ProbeTable"), runs = table.rows[0].cells[0].content[0].runs;
          assert.equal(runs.length, target ? 2 : 1); assert.equal(runs[0].text, control.text); assert.equal(runs[0].style.language, control.language);
          for (const run of runs) {
            assert.deepEqual(run.style.scriptFontFamilies, expected.maps[control.kind]);
            assert.equal(run.style.bold ?? false, target && input.stage === "bold"); assert.equal(run.style.italic ?? false, target && input.stage === "italic");
          }
          assert.deepEqual([...new Set(resolveTextFontSegments(runs[0].text, runs[0].style, snapshot.theme).map(s => s.family))], [control.targetFamily]);
          if (target) { assert.equal(runs[1].text, "Ω"); assert.equal(runs[1].style.language, "el-GR"); assert.deepEqual([...new Set(resolveTextFontSegments(runs[1].text, runs[1].style, snapshot.theme).map(s => s.family))], ["Georgia"]); }
          const rules = report.tableStyleInheritance.properties.filter(p => p.target.part === `/ppt/slides/slide${index + 1}.xml` && p.property.startsWith("run.scriptFontFamilies."));
          assert.equal(rules.length, runs.length * 8);
          for (const rule of rules) { assert.equal(rule.status, "mapped"); assert.equal(rule.value, expected.maps[control.kind][rule.property.split(".").at(-1)]); }
          controls.push({ id: control.id, runs, mappedRules: rules.length });
        }
        const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready"); assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), bytes); assert.deepEqual(normalizePresentation(session.document), snapshot);
        const result = { ...input, sha256: hash(bytes), noOpExact: true, controls, scalarCandidates: [] }; projections.push(result);
        if (input.stage) {
          const first = session.document.slides[expected.editTargets[0]-1], table = first.elements.find(e => e.name === "ProbeTable"), wrong = structuredClone(table);
          wrong.rows[0].cells[0].content[0].runs[0].style.bold = input.stage !== "bold"; first.replaceElement(wrong.id, wrong);
          assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); first.replaceElement(table.id, table); result.sourceStyleWriteBlockedAndRestored = true;
          let previous = bytes;
          for (let scalarRound = 1; scalarRound <= 3; scalarRound++) {
            for (const index of expected.editTargets) {
              const slide = session.document.slides[index-1], table = slide.elements.find(e => e.name === "ProbeTable"), next = structuredClone(table);
              next.rows[0].cells[0].content[0].runs[1].text += "X"; slide.replaceElement(next.id, next);
            }
            const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready"); const saved = (await session.save(prepared.plan)).bytes, before = readGeneratedZip(previous), after = readGeneratedZip(saved);
            assert.deepEqual([...after.keys()], [...before.keys()]);
            for (const [part, body] of before) {
              if (expected.editTargets.some(i => part === `ppt/slides/slide${i}.xml`)) {
                const from = `<a:t>Ω${"X".repeat(scalarRound-1)}</a:t>`, to = `<a:t>Ω${"X".repeat(scalarRound)}</a:t>`;
                assert.equal(body.toString().split(from).length, 2); assert.equal(after.get(part).toString(), body.toString().replace(from, to));
              } else assert.deepEqual(after.get(part), body);
            }
            const candidate = await importPptx(saved, options);
            try { assert.deepEqual(normalizePresentation(candidate.document), normalizePresentation(session.document)); assert.equal(candidate.report.textFontProfile, "script-families-v3"); } finally { candidate.dispose(); }
            result.scalarCandidates.push({ round: scalarRound, sha256: hash(saved), nonCharacterBytesExact: true, reprojectionEqual: true, nativeReopen: "not tested" }); previous = saved;
          }
        }
      } finally { session.dispose(); }
      if (input.round === 3) {
        const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
        for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
        const result = await page.evaluate(async ({ input, options }) => {
          const session = await PPTKit.importPptx(Uint8Array.from(input), options);
          try {
            const document = PPTKit.normalizePresentation(session.document), report = session.report, prepared = await session.prepareSave({ mode: "no-op" });
            if (prepared.status !== "ready") throw Error("Browser no-op not ready"); return { document, report, bytes: [...(await session.save(prepared.plan)).bytes] };
          } finally { session.dispose(); }
        }, { input: [...bytes], options });
        assert.deepEqual(result.document, snapshot); assert.deepEqual(result.report, report); assert.deepEqual(result.bytes, [...bytes]); assert.deepEqual(requests, []); projections.at(-1).browserGlobalParityAndExactNoOp = true; await page.close();
      }
    }
    groups.push({ group, independent, negatives, projections });
  }
} finally { await browser.close(); }
const info = JSON.parse((await exec("python3", ["-c", "import json,plistlib,sys; p=plistlib.load(open(sys.argv[1],'rb')); print(json.dumps({k:p[k] for k in ['CFBundleShortVersionString','CFBundleVersion']}))", "/Applications/Microsoft PowerPoint.app/Contents/Info.plist"])).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application: { name: "Microsoft PowerPoint for Mac", ...info }, groups, total: { styledTargets: 15, stages: 3, nativeRoundsPerStage: 3, nativePDFPageRecords: 324, parserExactNoOps: 20, browserExactNoOps: 6, styleWriteRejections: 18, scalarCandidates: 54, rejectedControls: groups.reduce((n,g) => n+g.negatives.rejectedCount,0) }, sourceStyleWrites: "blocked", themeAndNotes: "not tested", scope: "Named native styles and declared source transport; no Windows/theme/notes/full bidi/mark-outline admission. Scalar candidates are not native-reopened edits." };
await writeFile("docs/testing/native-table-font-style-evidence.json", JSON.stringify(evidence, null, 2) + "\n"); console.log(JSON.stringify(evidence.total));
