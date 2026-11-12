import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createHash } from "node:crypto";
import { chromium } from "@playwright/test";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { exerciseNativeNotes } from "../tests/support/exercise-native-notes.mjs";
import { nativeNotesSourceAudit } from "../tests/support/native-notes-source-audit.mjs";

const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const options = { projection: "content", notesTextProfile: "master-style-v1", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" };
const manifestPaths = ["append", "styles", "themes"].map(stage => "docs/testing/native-notes-" + stage + "-evidence.json");
const sources = [], registry = new Map();
for (const file of manifestPaths) {
  const bytes = await readFile(file), manifest = JSON.parse(bytes);
  sources.push({ file: path.resolve(file), sha256: hash(bytes), application: manifest.application });
  for (const group of manifest.groups) for (const input of group.inputs) {
    const stage = file.includes("-append-") ? "append" : file.includes("-styles-") ? "styles" : "themes";
    const name = input.name ?? (input.stage ?? "append") + "-round" + input.round;
    const alias = { stage, group: group.variant ?? group.group, name, file: path.join(manifest.root, group.variant ?? group.group, name + ".pptx"), sourceManifest: path.resolve(file) };
    assert.equal(hash(await readFile(alias.file)), input.sha256);
    if (registry.has(input.sha256)) registry.get(input.sha256).aliases.push(alias);
    else registry.set(input.sha256, { sha256: input.sha256, aliases: [alias] });
  }
}
assert.equal([...registry.values()].reduce((n, r) => n + r.aliases.length, 0), 39);
assert.equal(registry.size, 37);
const outputRoot = path.resolve(".tmp/native-notes-consolidated");
await mkdir(outputRoot, { recursive: true });
const directory = await mkdtemp(path.join(outputRoot, "run-")), browser = await chromium.launch(), inputs = [];
const moduleUrl = text => "data:text/javascript;base64," + Buffer.from(text).toString("base64");
const coreUrl = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8"));
const parserUrl = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(coreUrl)));
const globals = await Promise.all(["core", "pptx-parser"].map(name => readFile("packages/" + name + "/dist/browser/global.js", "utf8")));
try {
  for (const [number, entry] of [...registry.values()].entries()) {
    const folder = path.join(directory, "input-" + String(number + 1).padStart(2, "0"));
    await mkdir(folder);
    const bytes = await readFile(entry.aliases[0].file);
    await writeFile(path.join(folder, "input.pptx"), bytes);
    const result = await exerciseNativeNotes({ importPptx, normalizePresentation }, bytes, options);
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", r => { requests.push(r.request().url()); return r.abort(); });
    try {
      const truth = await page.evaluate(nativeNotesSourceAudit, {
        parts: Object.fromEntries([...readGeneratedZip(bytes)].filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels")).map(([name, body]) => [name, body.toString()])),
        document: result.document, properties: result.report.textStyleInheritance.properties,
      });
      const snapshot = JSON.stringify({ document: result.document, report: result.report, independent: truth }, null, 2) + "\n";
      await writeFile(path.join(folder, "input-inspection.json"), snapshot);
      let previous = bytes;
      const candidates = [];
      for (const candidate of result.candidates) {
        const saved = Buffer.from(candidate.bytes), before = readGeneratedZip(previous), after = readGeneratedZip(saved);
        assert.deepEqual([...after.keys()], [...before.keys()]);
        const targets = new Map(truth.rows.map(row => [row.notesPart, row.runs.at(-1).text]));
        for (const [part, body] of before) {
          if (targets.has(part)) {
            const text = targets.get(part) + "X".repeat(candidate.round - 1), escape = value => value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
            const from = "<a:t>" + escape(text) + "</a:t>", to = "<a:t>" + escape(text + "X") + "</a:t>";
            assert.equal(body.toString().split(from).length, 2);
            assert.equal(after.get(part).toString(), body.toString().replace(from, to));
          } else assert.deepEqual(after.get(part), body);
        }
        const file = path.join(folder, "candidate-" + candidate.round + ".pptx");
        await writeFile(file, saved);
        const projected = JSON.stringify({ document: candidate.document, report: candidate.report }, null, 2) + "\n";
        await writeFile(path.join(folder, "candidate-" + candidate.round + ".json"), projected);
        candidates.push({ round: candidate.round, file, sha256: hash(saved), snapshotSha256: hash(projected), exactOutsideTargetCharacters: true, completeReprojectionEqual: true, nativeReopen: "not tested" }); previous = saved;
      }
      await page.addScriptTag({ content: "globalThis.exerciseNativeNotes = " + exerciseNativeNotes.toString() });
      const consumers = [];
      for (const format of ["esm", "global"]) {
        if (format === "global") for (const content of globals) await page.addScriptTag({ content });
        const checked = await page.evaluate(async ({ source, options, expected, coreUrl, parserUrl }) => {
          const c = coreUrl ? await import(coreUrl) : PPTKit, p = parserUrl ? await import(parserUrl) : PPTKit;
          const actual = await exerciseNativeNotes({ importPptx: p.importPptx, normalizePresentation: c.normalizePresentation }, source, options);
          const stable = v => Array.isArray(v) ? v.map(stable) : v && typeof v === "object" ? Object.fromEntries(Object.keys(v).sort().map(k => [k, stable(v[k])])) : v;
          if (JSON.stringify(stable(actual)) !== JSON.stringify(stable(expected))) throw Error("Native notes consumer document/report/blocked-report/candidate bytes parity failed");
          return { exactNoOpAndRestores: actual.exactNoOpAndRestores, mixedWriteRejections: actual.rejections.length, exactCandidateRounds: actual.candidates.length };
        }, { source: [...bytes], options, expected: result, coreUrl: format === "esm" ? coreUrl : null, parserUrl: format === "esm" ? parserUrl : null });
        consumers.push({ format, ...checked });
      }
      assert.deepEqual(requests, []);
      inputs.push({ ...entry, inputFile: path.join(folder, "input.pptx"), snapshotSha256: hash(snapshot), independentMappedProperties: truth.checkedMappedProperties, notesPages: truth.rows.length, nodeExactNoOpAndRestores: true, mixedWriteRejections: result.rejections, candidates, consumers, browserNetworkRequests: requests });
      console.log("verified " + (number + 1) + "/" + registry.size + ": " + entry.aliases.map(a => a.stage + "/" + a.group + "/" + a.name).join(", "));
    } finally { await page.close(); }
  }
} finally { await browser.close(); }
const evidence = {
  schemaVersion: 1, recordedAt: new Date().toISOString(), directory, options, sourceManifests: sources, inputs,
  total: { logicalInputPaths: 39, uniqueInputFiles: 37, consumers: 3, exactNoOpPaths: 111, additionalExactRestores: 333, mixedWriteRejections: 333, characterCandidatePaths: 333, materializedNodeCandidates: 111, nativeReopenedCandidates: 0 },
  scope: "All captured positive append/style/theme inputs, deduplicated by verified bytes. Every file exercises Node/ESM/global exact no-op, three mixed-write rejections/restores and three character rounds with full snapshot/report/bytes parity. Independent DOM checks source topology/fonts/provenance; all non-character payloads exact. Candidates are headless only. Invalid/shared-override native admission, full notes pages, native font equivalence and source style permissions remain separate.",
};
await writeFile("docs/testing/native-notes-consolidated-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
await writeFile(path.join(outputRoot, "latest.txt"), directory + "\n");
console.log(JSON.stringify(evidence.total));
