import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { chromium } from "@playwright/test";
import { inspectPptxNotes, importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { notesInventorySourceAudit } from "../tests/support/notes-inventory-source-audit.mjs";
const hash = b => createHash("sha256").update(b).digest("hex");
const inputs = [], manifests = [];
async function manifest(name) {
  const file = path.resolve(`docs/testing/${name}-evidence.json`), bytes = await readFile(file);
  manifests.push({ file, sha256: hash(bytes) }); return JSON.parse(bytes);
}
const pages = await manifest("native-notes-pages");
const control = pages.files.find(f => f.name === "native-control.pptx");
inputs.push({ file: path.join(pages.directory, control.name), sha256: control.sha256 });
for (const r of [...pages.rounds.slice(-1), ...pages.copiedSlideRounds.slice(-1)]) inputs.push({ file: path.join(pages.directory, r.name), sha256: r.pptxSha256 });
const domains = await manifest("notes-image-domains");
for (const g of domains.groups) { const r = g.rounds.at(-1); inputs.push({ file: path.join(g.directory, r.name + ".pptx"), sha256: r.sha256, domains: r.domains }); }
const authored = await manifest("authored-notes-pages");
for (const g of authored.groups) { const r = g.rounds.at(-1); inputs.push({ file: path.join(authored.directory, r.file), sha256: r.sha256 }); }
const insertion = await manifest("notes-insertion-styles");
for (const g of insertion.groups) { const r = g.rounds.at(-1); inputs.push({ file: path.join(insertion.directory, r.name + ".pptx"), sha256: r.sha256 }); }
await mkdir(".tmp/notes-source-inventory", { recursive: true });
const directory = path.resolve(await mkdtemp(".tmp/notes-source-inventory/run-")), records = [];
const browser = await chromium.launch(), page = await browser.newPage(), requests = [];
try {
  await page.route("**/*", r => { requests.push(r.request().url()); return r.abort(); });
  for (const p of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${p}/dist/browser/global.js`, "utf8") });
  for (const input of inputs) {
    const bytes = await readFile(input.file); assert.equal(hash(bytes), input.sha256);
    const report = await inspectPptxNotes(bytes), xml = Object.fromEntries([...readGeneratedZip(bytes)].filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels")).map(([name, b]) => [name, b.toString("utf8")]));
    const audit = await page.evaluate(notesInventorySourceAudit, { parts: xml, report });
    const actual = await page.evaluate(async bytes => PPTKit.inspectPptxNotes(Uint8Array.from(bytes)), [...bytes]); assert.deepEqual(actual, report);
    assert.ok(report.parts.length); assert.equal(report.geometryPolicy, "unresolved-source-only"); assert.equal(report.fieldEvaluation, "not-performed");
    if (input.domains) assert.deepEqual(report.pages.map(p => ({ slidePart: p.slidePart.slice(1), notesPart: p.notesPart.slice(1), masterPart: p.masterPart.slice(1), masterThemePart: p.masterThemePart?.slice(1), overridePart: p.overridePart?.slice(1) ?? null })), input.domains.map(d => ({ slidePart: d.slidePart, notesPart: d.notesPart, masterPart: d.masterPart, masterThemePart: d.masterThemePart, overridePart: d.overridePart })));
    const session = await importPptx(bytes, { projection: "structure" });
    try { const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); const saved = (await session.save(plan.plan)).bytes; assert.deepEqual(Buffer.from(saved), bytes); assert.deepEqual(await inspectPptxNotes(saved), report); } finally { session.dispose(); }
    const snapshot = JSON.stringify(report, null, 2) + "\n", snapshotFile = path.join(directory, records.length + ".json"); await writeFile(snapshotFile, snapshot);
    records.push({ file: input.file, sha256: hash(bytes), snapshotFile, snapshotSha256: hash(snapshot), audit, pages: report.pages.length, roles: [...new Set(report.parts.flatMap(p => p.drawings.map(d => d.role)))], exactNoOp: true, globalReportEqual: true, capturedDomainMapEqual: Boolean(input.domains) });
  }
  assert.deepEqual(requests, []);
} finally { await browser.close(); }
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), directory, manifests, records, total: { inputs: records.length, pages: records.reduce((n, r) => n + r.pages, 0), fields: records.reduce((n, r) => n + r.audit.fields, 0), sourceReferences: records.reduce((n, r) => n + r.audit.checkedSourceReferences, 0) }, scope: "Current read-only inventory and exact no-op against hash-verified existing captured files; no new native app action, field evaluation, authored geometry or editing admission." };
await writeFile("docs/testing/notes-source-inventory-evidence.json", JSON.stringify(evidence, null, 2) + "\n"); console.log(JSON.stringify(evidence.total));
