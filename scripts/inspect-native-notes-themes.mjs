import assert from "node:assert/strict";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { chromium } from "@playwright/test";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

if (process.argv.length !== 3) throw Error("Usage: node scripts/inspect-native-notes-themes.mjs <theme-run-directory>");
const root = path.resolve(process.argv[2]), exec = promisify(execFile), hash = bytes => createHash("sha256").update(bytes).digest("hex"), groups = [];
const options = { projection: "content", notesTextProfile: "master-style-v1", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" };
const browser = await chromium.launch();
try {
  for (const group of ["A", "B", "override"]) {
    const directory = path.join(root, group), expected = JSON.parse(await readFile(path.join(directory, "expected.json"), "utf8"));
    for (const oracle of ["native-notes-themes.py", "native-notes-themes-checks.py"]) {
      const result = await exec("python3", [path.resolve("scripts/research/oracles", oracle), directory], { timeout: 120000, maxBuffer: 4 * 1024 * 1024 });
      console.log(group, oracle, result.stdout.trim());
    }
    const independent = JSON.parse(await readFile(path.join(directory, "theme-inspection.json"), "utf8")), negatives = JSON.parse(await readFile(path.join(directory, "theme-negative-inspection.json"), "utf8"));
    const controls = group === "override" ? JSON.parse(await readFile(path.join(directory, "controls.json"), "utf8")) : expected.controls.map(c => ({ ...c, variant: group }));
    const names = group === "override" ? ["control-distinct-overrides", "native-baseline-round1", "native-baseline-round2", "native-baseline-round3"] : ["baseline"];
    names.push("office-round1", "office-round2", "office-round3");
    const inputs = [];
    await mkdir(path.join(directory, "inspections"), { recursive: true });
    for (const name of names) {
      const bytes = await readFile(path.join(directory, name + ".pptx")), s = await importPptx(bytes, options);
      try {
        const document = normalizePresentation(s.document), report = s.report;
        assert.equal(document.slides.length, controls.length);
        assert.equal(report.inspection.notesMasters.length, 1);
        assert.equal(report.inspection.findings.filter(f => f.code === "unlisted-master").length, 0);
        for (const [i, c] of controls.entries()) {
          const runs = document.slides[i].notes[0].runs;
          assert.equal(runs[0].text, c.text);
          assert.equal(runs[0].style.language, c.language);
          assert.ok(runs.every(r => r.style.fontSize === 40 && !r.style.bold && !r.style.italic));
          assert.equal(runs.length, group !== "override" && expected.editTargets.includes(i + 1) ? 2 : 1);
          if (runs.length === 2) { assert.equal(runs[1].text, "Ω"); assert.equal(runs[1].style.language, "en-US"); }
          const map = c.map ?? expected.map;
          const effective = c.id.startsWith("local-") ? undefined : Object.fromEntries(Object.entries(map).filter(([script]) => !(["empty-ea-tc", "bad-ea-tc"].includes(c.id) && ["Hans", "Hant", "Jpan", "Hang"].includes(script))));
          for (const run of runs) {
            assert.deepEqual(run.style.scriptFontFamilies, effective);
            assert.equal(run.style.fontFamily, c.id.startsWith("local-") ? "Arial" : map.Latn);
            if (c.id.startsWith("local-")) { assert.equal(run.style.eastAsianFontFamily, "Arial Unicode MS"); assert.equal(run.style.complexScriptFontFamily, "Arial Unicode MS"); }
          }
        }
        const snapshot = JSON.stringify({ document, report }, null, 2) + "\n", snapshotFile = path.join(directory, "inspections", name + ".json");
        await writeFile(snapshotFile, snapshot);
        const page = await browser.newPage();
        try {
          const checked = await page.evaluate(({ parts, properties }) => {
            const docs = Object.fromEntries(Object.entries(parts).map(([name, text]) => ["/" + name, new DOMParser().parseFromString(text, "application/xml")]));
            const select = s => {
              const doc = docs[s.part]; if (!doc) throw Error("Missing source part");
              return doc.evaluate(s.path, doc, p => doc.documentElement.lookupNamespaceURI(p), XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            };
            const relation = (owner, kind, optional = false) => {
              const slash = owner.lastIndexOf("/"), doc = docs[owner.slice(0, slash + 1) + "_rels/" + owner.slice(slash + 1) + ".rels"];
              const nodes = [...doc.documentElement.children].filter(n => n.getAttribute("Type").endsWith("/" + kind));
              if (optional && !nodes.length) return undefined;
              if (nodes.length !== 1) throw Error("Ambiguous relationship");
              return new URL(nodes[0].getAttribute("Target"), "https://fixture.invalid" + owner).pathname;
            };
            let count = 0;
            for (const p of properties.filter(p => p.target.part.includes("/notesSlides/"))) {
              if (!select(p.target) || !select(p.source)) throw Error("Missing actual XPath");
              for (const d of p.resolvedFrom ?? []) if (!select(d)) throw Error("Missing source dependency");
              if (p.status !== "mapped") continue;
              const master = relation(p.target.part, "notesMaster"), theme = relation(p.target.part, "themeOverride", true) ?? relation(master, "theme");
              if (p.property.startsWith("run.scriptFontFamilies.")) {
                const n = select(p.source);
                if (p.source.part !== theme || n.value !== p.value || n.ownerElement.getAttribute("script") !== p.property.split(".").at(-1)) throw Error("Borrowed font source");
                count++;
              }
              if (p.property === "run.fontSize") {
                if (p.level !== "notes-master-placeholder" || p.source.part !== master || select(p.source).value !== "4000" || p.value !== 40) throw Error("Wrong master-size provenance");
                count++;
              }
            }
            return count;
          }, { parts: Object.fromEntries([...readGeneratedZip(bytes)].filter(([n]) => n.endsWith(".xml") || n.endsWith(".rels")).map(([n, b]) => [n, b.toString()])), properties: report.textStyleInheritance.properties });
          assert.ok(checked > controls.length);
        } finally { await page.close(); }
        const noOp = await s.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready");
        assert.deepEqual(Buffer.from((await s.save(noOp.plan)).bytes), bytes);
        const slide = s.document.slides[0], original = structuredClone(slide.notes), wrong = structuredClone(original);
        wrong[0].runs[0].style.fontFamily = "Forbidden"; slide.setNotes(wrong);
        assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); slide.setNotes(original);
        const candidates = []; let previous = bytes;
        for (let round = 1; round <= 3; round++) {
          for (const child of s.document.slides) { const notes = structuredClone(child.notes); notes[0].runs[0].text += "X"; child.setNotes(notes); }
          const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready");
          const saved = (await s.save(plan.plan)).bytes, before = readGeneratedZip(previous), after = readGeneratedZip(saved);
          assert.deepEqual([...after.keys()], [...before.keys()]); let changed = 0;
          for (const [part, body] of before) {
            if (after.get(part).equals(body)) continue;
            assert.match(part, /^ppt\/notesSlides\/notesSlide\d+\.xml$/);
            const matches = controls.filter(c => body.toString().includes(`<a:t>notes-${c.variant}-${c.id}</a:t>`));
            assert.equal(matches.length, 1);
            const from = `<a:t>${matches[0].text}${"X".repeat(round - 1)}</a:t>`; assert.equal(body.toString().split(from).length, 2);
            assert.equal(after.get(part).toString(), body.toString().replace(from, from.replace("</a:t>", "X</a:t>"))); changed++;
          }
          assert.equal(changed, controls.length);
          const candidate = await importPptx(saved, options);
          try { assert.deepEqual(normalizePresentation(candidate.document), normalizePresentation(s.document)); }
          finally { candidate.dispose(); }
          candidates.push({ round, sha256: hash(saved), nonCharacterPayloadsExact: true, fullReprojectionEqual: true, nativeReopen: "not tested" }); previous = saved;
        }
        inputs.push({ name, sha256: hash(bytes), snapshotFile, snapshotSha256: hash(snapshot), exactNoOp: true, styleWriteBlockedAndRestored: true, independentXPathChecked: true, candidates });
        if (name === "office-round3") {
          const page = await browser.newPage(), requests = []; await page.route("**/*", r => { requests.push(r.request().url()); return r.abort(); });
          try {
            for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
            const result = await page.evaluate(async ({ bytes, options }) => {
              const s = await PPTKit.importPptx(Uint8Array.from(bytes), options);
              try { const document = PPTKit.normalizePresentation(s.document), report = s.report, plan = await s.prepareSave({ mode: "no-op" }); if (plan.status !== "ready") throw Error("Browser no-op blocked"); return { document, report, bytes: [...(await s.save(plan.plan)).bytes] }; }
              finally { s.dispose(); }
            }, { bytes: [...bytes], options });
            assert.deepEqual(result.document, document); assert.deepEqual(result.report, report); assert.deepEqual(result.bytes, [...bytes]); assert.deepEqual(requests, []);
            inputs.at(-1).browserGlobalParityAndExactNoOp = true;
          } finally { await page.close(); }
        }
      } finally { s.dispose(); }
    }
    const admissions = [];
    if (group === "override") for (const [kind, observation] of Object.entries(expected.nativeAdmission)) admissions.push({ kind, ...observation, sha256: hash(await readFile(path.join(directory, observation.file))) });
    groups.push({ group, independent, negatives, inputs, admissions });
  }
} finally { await browser.close(); }
const application = JSON.parse((await exec("python3", ["-c", "import json,plistlib,sys; p=plistlib.load(open(sys.argv[1],'rb')); print(json.dumps({k:p[k] for k in ['CFBundleShortVersionString','CFBundleVersion']}))", "/Applications/Microsoft PowerPoint.app/Contents/Info.plist"])).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), root, application, groups,
  total: { themedNativeRounds: 9, themedNotesPages: 180, overrideNativeBaselineRounds: 3, overrideNativeBaselineNotesPages: 90, overrideFirstOpenNotesPages: 30, nodeExactNoOps: groups.reduce((n, g) => n + g.inputs.length, 0), browserGlobalExactNoOps: 3, scalarCandidates: groups.reduce((n, g) => n + g.inputs.length * 3, 0), actualCorruptionsRejected: groups.reduce((n, g) => n + g.negatives.rejectedCount, 0) },
  scope: "Office slide theme, A/B independent notes themes and one-master distinct per-notes overrides. Shared override part yields a native repair prompt; no repair performed. Source-patched controls are not fresh author round trips or full schema certificates. Exact source notes payloads remain; final-header-A kerning differs. Scalar candidates not native reopened. Full notes artwork/fields, general font equivalence and source style writes remain open." };
await writeFile("docs/testing/native-notes-themes-evidence.json", JSON.stringify(evidence, null, 2) + "\n"); console.log(JSON.stringify(evidence.total));
