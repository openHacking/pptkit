import assert from "node:assert/strict";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createHash } from "node:crypto";
import { chromium } from "@playwright/test";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const hash = b => createHash("sha256").update(b).digest("hex"), evidenceFile = "docs/testing/notes-image-domains-evidence.json";
const nativeBytes = await readFile(evidenceFile), native = JSON.parse(nativeBytes);
const options = { projection: "content", textFontProfile: "script-families-v3", notesTextProfile: "master-style-v1", editSaveProfile: "text-runs-v3" };
const browser = await chromium.launch(), records = [];
try {
  for (const group of native.groups) {
    const baseline = await importPptx(await readFile(path.join(group.directory, "baseline.pptx")), options);
    const unchanged = normalizePresentation(baseline.document); baseline.dispose();
    const output = path.join(group.directory, "parser-inspections"); await mkdir(output, { recursive: true });
    for (const name of ["rect", "round1", "round2", "round3"]) {
      const bytes = await readFile(path.join(group.directory, name + ".pptx"));
      const expected = name === "rect" ? group.rectSha256 : group.rounds.find(r => r.name === name).sha256;
      assert.equal(hash(bytes), expected);
      const session = await importPptx(bytes, options);
      try {
        const document = normalizePresentation(session.document), report = session.report;
        assert.deepEqual(document, unchanged);
        const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready"); assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), bytes);
        const first = session.document.slides[0], original = structuredClone(first.notes), mixed = structuredClone(original);
        mixed[0].runs[0].text += "X"; mixed[0].runs[0].style.fontFamily = "Forbidden";
        first.setNotes(mixed); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); first.setNotes(original);
        const restored = await session.prepareSave({ mode: "no-op" }); assert.equal(restored.status, "ready"); assert.deepEqual(Buffer.from((await session.save(restored.plan)).bytes), bytes);
        for (const slide of session.document.slides) { const notes = structuredClone(slide.notes); notes[0].runs.at(-1).text += "X"; slide.setNotes(notes); }
        const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready");
        const saved = (await session.save(plan.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(saved);
        const targets = new Map(group.sourceDomains.map((row, i) => [row.notesPart, document.slides[i].notes[0].runs.at(-1).text]));
        assert.deepEqual([...after.keys()], [...before.keys()]);
        for (const [part, body] of before) {
          if (targets.has(part)) {
            const escape = value => value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
            const from = "<a:t>" + escape(targets.get(part)) + "</a:t>", to = from.replace("</a:t>", "X</a:t>");
            assert.equal(body.toString().split(from).length, 2); assert.equal(after.get(part).toString(), body.toString().replace(from, to));
          } else assert.deepEqual(after.get(part), body);
        }
        const candidateFile = path.join(output, name + "-candidate.pptx"); await writeFile(candidateFile, saved);
        const snapshot = JSON.stringify({ document, report }, null, 2) + "\n", snapshotFile = path.join(output, name + ".json"); await writeFile(snapshotFile, snapshot);
        let browserCheck;
        if (name === "round3") {
          const page = await browser.newPage(), requests = []; await page.route("**/*", r => { requests.push(r.request().url()); return r.abort(); });
          try {
            for (const p of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${p}/dist/browser/global.js`, "utf8") });
            const actual = await page.evaluate(async ({ bytes, options }) => {
              const source = Uint8Array.from(bytes), s = await PPTKit.importPptx(source, options);
              try {
                const document = PPTKit.normalizePresentation(s.document), report = s.report, plan = await s.prepareSave({ mode: "no-op" });
                if (plan.status !== "ready") throw Error("Image-domain browser no-op blocked");
                return { document, report, bytes: [...(await s.save(plan.plan)).bytes] };
              } finally { s.dispose(); }
            }, { bytes: [...bytes], options });
            assert.deepEqual(actual, { document, report, bytes: [...bytes] }); assert.deepEqual(requests, []); browserCheck = { format: "global", completeSnapshotAndBytesEqual: true, requests };
          } finally { await page.close(); }
        }
        records.push({ group: group.name, name, sha256: hash(bytes), snapshotFile, snapshotSha256: hash(snapshot), exactNoOpAndRestore: true, styleMixedWriteBlocked: true, candidate: { file: candidateFile, sha256: hash(saved), nativeReopen: "not tested", exactOutsideTargetCharacters: true }, ...(browserCheck ? { browser: browserCheck } : {}) });
      } finally { session.dispose(); }
    }
  }
} finally { await browser.close(); }
const evidence = { schemaVersion: 1, nativeEvidence: path.resolve(evidenceFile), nativeEvidenceSha256: hash(nativeBytes), options, records, total: { nodeInputs: records.length, exactNodeNoOpsAndRestores: records.length, mixedStyleRejections: records.length, headlessCharacterCandidates: records.length, browserGlobalNoOps: records.filter(r => r.browser).length }, scope: "Captured rect and native states only. Source semantics remain equal to the original blank-image baseline; scalar candidates are not native reopened. No source style-write or general font equivalence claim." };
await writeFile("docs/testing/notes-image-domains-parser-evidence.json", JSON.stringify(evidence, null, 2) + "\n"); console.log(JSON.stringify(evidence.total));
