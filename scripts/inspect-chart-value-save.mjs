// Optional source-save/native experiment. Actual Edit Data is separately observed through UI.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { createChartSaveFixture } from "../tests/fixtures/chart-save.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-chart-value-save.mjs [soffice-executable]");
const execute = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/chart-value-save-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/chart-value-save-native/run-"));
const nativeDirectory = path.join(directory, "native"), pdfDirectory = path.join(directory, "pdf");
await mkdir(nativeDirectory); await mkdir(pdfDirectory);
const generated = await generatePptx(createChartSaveFixture(createPresentation)); assert.deepEqual(generated.warnings, []);
let bytes = generated.bytes, beforeFile = path.join(directory, "source.pptx"); await writeFile(beforeFile, bytes);
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const oracle = async (...args) => JSON.parse((await execute("python3", ["scripts/research/chart-value-save-oracle.py", ...args], options)).stdout);
const semantics = inventory => inventory.charts.map(c => ({ family: c.family, series: c.series.map(s => ({ sourceId: s.sourceId, order: s.order, values: s.values })) }));
const rounds = [], profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
for (let round = 0; round < 3; round++) {
  const edits = [0, 1, 2].map(chartIndex => ({ chartIndex, seriesIndex: 0, pointIndex: 1, value: round === 1 ? 0 : round === 0 ? 12.5 + chartIndex : 37.25 + chartIndex }));
  const session = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
  const sourceHash = hash(bytes); let report;
  try {
    assert.ok(session.report.chartValues.length > 0 && session.report.chartValues.every(b => b.status === "bound"));
    for (const edit of edits) {
      const slide = session.document.slides[edit.chartIndex], chart = slide.elements.find(e => e.type === "chart"), index = slide.elements.indexOf(chart), updated = structuredClone(chart);
      updated.series[edit.seriesIndex].values[edit.pointIndex] = edit.value;
      slide.removeElement(chart.id); slide.insertElement(index, updated);
    }
    const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
    const saved = await session.save(prepared.plan); bytes = saved.bytes; report = saved.report;
  } finally { session.dispose(); }
  const filename = path.join(directory, `round-${round + 1}.pptx`), editsFile = path.join(directory, `round-${round + 1}-edits.json`);
  await writeFile(filename, bytes); await writeFile(editsFile, JSON.stringify(edits, null, 2) + "\n");
  const independent = await oracle("verify", beforeFile, filename, editsFile);
  assert.equal(independent.changedOuterParts.length, 6); assert.equal(independent.charts.length, 4);
  const nativeLog = await execute(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, filename], options);
  const nativeFile = path.join(nativeDirectory, path.basename(filename)), nativeBytes = await readFile(nativeFile), nativeInventory = await oracle("inspect", nativeFile);
  assert.deepEqual(semantics(nativeInventory), semantics(independent), "Native resave must retain selected and control numeric chart series");
  const reopened = await importPptx(nativeBytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
  let nativeBindings, numericEditTrial;
  try {
    nativeBindings = reopened.report.chartValues.map(({ status, reason }) => ({ status, ...(reason ? { reason } : {}) }));
    const plan = await reopened.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
    assert.deepEqual((await reopened.save(plan.plan)).bytes, new Uint8Array(nativeBytes));
    const first = reopened.report.chartValues[0];
    if (first?.status === "unmapped") {
      const slide = reopened.document.slides.find(s => s.id === first.slideId), chart = slide.elements.find(e => e.id === first.elementId), index = slide.elements.indexOf(chart), changed = structuredClone(chart);
      changed.series[first.seriesIndex].values[first.pointIndex] += .5;
      slide.removeElement(chart.id); slide.insertElement(index, changed);
      const trial = await reopened.prepareSave({ mode: "edit" }); assert.equal(trial.status, "blocked"); assert.equal(trial.plan, undefined);
      numericEditTrial = { status: trial.status, diagnostics: trial.report.diagnostics.map(d => d.code) };
      assert.ok(numericEditTrial.diagnostics.includes("save-edit-unsupported"));
    } else numericEditTrial = { status: "not-performed", reason: "The selected native value has a bound source; rejection is not expected." };
  } finally { reopened.dispose(); }
  rounds.push({ round: round + 1, sourceSha256: sourceHash, savedSha256: hash(bytes), edits, independent, native: { sha256: hash(nativeBytes), inventory: nativeInventory, bindingEligibility: nativeBindings, noOpOriginalBytes: true, numericEditTrial, log: nativeLog }, saveChanges: report.changes });
  beforeFile = filename;
}
const pdfLog = await execute(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdfDirectory, beforeFile], options);
const pdfFile = path.join(pdfDirectory, "round-3.pdf"), pdf = await readFile(pdfFile); assert.ok(pdf.subarray(0, 5).equals(Buffer.from("%PDF-")));
const pdfObservation = JSON.parse((await execute("python3", ["-c", "import sys,json,fitz; d=fitz.open(sys.argv[1]); print(json.dumps({'pages':len(d),'titles':[p.get_text() for p in d]}))", pdfFile], options)).stdout);
assert.equal(pdfObservation.pages, 4); for (const title of ["Bar values", "Line values", "Pie values", "Unchanged control"]) assert.ok(pdfObservation.titles.some(text => text.includes(title)));
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory),
  scope: "PPTKit existing-values-v1 three-round source preservation; independent Python ZIP/XML numeric cache+XLSX checks, exact unrelated payloads, control chart/series; LibreOffice headless resave and PDF reopen. Native resaves are observations, not re-import write admission.",
  source: { sha256: hash(generated.bytes) }, rounds,
  pdf: { sha256: hash(pdf), byteLength: pdf.length, observation: pdfObservation, log: pdfLog },
  interactiveEditData: { powerpoint: "untested", libreOffice: "untested" },
  limitations: ["Headless conversion does not prove interactive Edit Data, repair-dialog absence or visual fidelity.", "LibreOffice may remove XLSX bindings; eligibility is recorded rather than inferred from retained cache data.", "No structural/category/name/format/shared edits are admitted by this experiment."] };
await writeFile("docs/testing/chart-value-save-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, rounds: rounds.length, charts: 4, independentVerified: rounds.every(r => r.independent.verified), application, nativeBindingStatuses: [...new Set(rounds.flatMap(r => r.native.bindingEligibility.map(b => b.status)))] }));
