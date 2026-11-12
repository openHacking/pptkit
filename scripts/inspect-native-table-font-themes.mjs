import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { chromium } from "@playwright/test";
import { normalizePresentation, resolveTextFontSegments, TEXT_FONT_SCRIPT_CODES } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

if (process.argv.length !== 3) throw Error("Usage: node scripts/inspect-native-table-font-themes.mjs <native-theme-run-directory>");
const directory = path.resolve(process.argv[2]), exec = promisify(execFile), hash = bytes => createHash("sha256").update(bytes).digest("hex");
const options = { projection: "content", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" }, groups = [];
const browser = await chromium.launch();
try {
  for (const group of ["reference", "filled", "reference-local", "filled-local"]) {
    const folder = path.join(directory, group), expected = JSON.parse(await readFile(path.join(folder, "expected.json"), "utf8")); let negatives;
    for (const file of ["native-table-font-themes.py", "native-table-font-themes-checks.py"]) {
      const output = (await exec("python3", [path.resolve("scripts/research/oracles", file), folder], { timeout: 120000, maxBuffer: 4 * 1024 * 1024 })).stdout;
      if (file.endsWith("-checks.py")) negatives = JSON.parse(output); else console.log(output.trim());
    }
    const independent = JSON.parse(await readFile(path.join(folder, "theme-inspection.json"), "utf8")), projections = [];
    const inputs = [{ name: "baseline" }, ...[1, 2, 3].map(round => ({ name: `office-round${round}`, stage: "office", round }))];
    for (const input of inputs) {
      const bytes = await readFile(path.join(folder, input.name + ".pptx")), session = await importPptx(bytes, options); let snapshot, report;
      try {
        snapshot = normalizePresentation(session.document); report = session.report;
        const controls = [];
        const nativeRows = input.stage ? independent.rounds[input.round-1].package.controls : null;
        for (const [index, control] of expected.controls.entries()) {
          const target = expected.editTargets.includes(index + 1), local = expected.localOverrideTargets.includes(index + 1), table = snapshot.slides[index].elements.find(e => e.name === "ProbeTable"), runs = table.rows[0].cells[0].content[0].runs;
          assert.equal(runs.length, target ? 2 : 1); assert.equal(runs[0].text, control.text); assert.equal(runs[0].style.language, control.language);
          const rulesMap = input.stage ? Object.fromEntries(Object.entries(nativeRows[index].declared.rules).filter(([script]) => TEXT_FONT_SCRIPT_CODES.includes(script))) : expected.maps[control.kind];
          for (const run of runs) {
            assert.deepEqual(run.style.scriptFontFamilies, local ? undefined : rulesMap);
            assert.equal(run.style.bold ?? false, false); assert.equal(run.style.italic ?? false, false);
            if (local) { assert.equal(run.style.fontFamily, expected.localFonts.latin); assert.equal(run.style.eastAsianFontFamily, expected.localFonts.ea); assert.equal(run.style.complexScriptFontFamily, expected.localFonts.cs); }
          }
          // Compare declaration transport; native face is independently recorded and may differ.
          if (!input.stage && !local) assert.deepEqual([...new Set(resolveTextFontSegments(runs[0].text, runs[0].style, snapshot.theme).map(s => s.family))], [control.targetFamily]);
          if (target) { assert.equal(runs[1].text, "Ω"); assert.equal(runs[1].style.language, "el-GR"); }
          const rules = report.tableStyleInheritance.properties.filter(p => p.target.part === `/ppt/slides/slide${index + 1}.xml` && p.property.startsWith("run.scriptFontFamilies."));
          assert.equal(rules.length, local ? 0 : runs.length * Object.keys(rulesMap).length);
          for (const rule of rules) {
            assert.equal(rule.status, "mapped"); assert.equal(rule.value, rulesMap[rule.property.split(".").at(-1)]);
            assert.equal(rule.source.part, control.kind.startsWith("collection") ? "/ppt/tableStyles.xml" : "/ppt/theme/theme1.xml");
            assert.ok(rule.source.path.includes("@typeface"));
          }
          controls.push({ id: control.id, runs, mappedRules: rules.length });
        }
        const sourcePage = await browser.newPage();
        try {
          const checked = await sourcePage.evaluate(({ parts, properties }) => {
            const docs = Object.fromEntries(Object.entries(parts).map(([name, xml]) => ["/"+name, new DOMParser().parseFromString(xml, "application/xml")]));
            const select = source => { const doc = docs[source.part]; if (!doc) throw Error("Invented source part"); return doc.evaluate(source.path, doc, p => doc.documentElement.lookupNamespaceURI(p), XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; };
            const fonts = properties.filter(p => p.status === "mapped" && (p.property.startsWith("run.scriptFontFamilies.") || (p.level === "run" && ["run.fontFamily", "run.eastAsianFontFamily", "run.complexScriptFontFamily"].includes(p.property))));
            for (const p of fonts) {
              if (select(p.source)?.value !== p.value || !select(p.target)) throw Error("Actual source font/provenance mismatch");
              for (const d of p.resolvedFrom ?? []) if (!select(d)) throw Error("Invented actual source dependency");
            }
            if (!fonts.length) throw Error("No font provenance checked"); return fonts.length;
          }, { parts: Object.fromEntries([...readGeneratedZip(bytes)].filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels")).map(([name, body]) => [name, body.toString()])), properties: report.tableStyleInheritance.properties });
          assert.ok(checked > 0);
        } finally { await sourcePage.close(); }
        const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready"); assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), bytes); assert.deepEqual(normalizePresentation(session.document), snapshot);
        const result = { ...input, sha256: hash(bytes), noOpExact: true, independentSourceXPathValidated: true, controls, scalarCandidates: [] }; projections.push(result);
        if (input.stage) {
          const first = session.document.slides[expected.editTargets[0]-1], table = first.elements.find(e => e.name === "ProbeTable"), wrong = structuredClone(table);
          wrong.rows[0].cells[0].content[0].runs[0].style.fontFamily = "Forbidden"; first.replaceElement(wrong.id, wrong);
          assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); first.replaceElement(table.id, table); result.sourceFontWriteBlockedAndRestored = true;
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
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application: { name: "Microsoft PowerPoint for Mac", ...info }, groups, total: { groups: 4, nativeRoundsPerGroup: 3, baselinePDFPageRecords: 72, themedPDFPageRecords: 216, parserExactNoOps: 16, browserExactNoOps: 4, sourceFontWriteRejections: 12, scalarCandidates: 36, rejectedControls: groups.reduce((n,g) => n+g.negatives.rejectedCount,0) }, sourceFontWrites: "blocked", notes: "not tested", scope: "Named native Office theme/declaration/appearance slice; local overrides are explicitly source-patched. No Windows/notes/general typography parity. Scalar candidates are not native reopened." };
await writeFile("docs/testing/native-table-font-theme-evidence.json", JSON.stringify(evidence, null, 2) + "\n"); console.log(JSON.stringify(evidence.total));
