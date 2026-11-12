// Optional independent four-package source-save validation and named native observations.
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, writeFile, readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { createScriptFontFixture } from "../tests/fixtures/script-fonts.mjs";
if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-script-font-import.mjs [soffice-executable]");
const execute = promisify(execFile), executable = process.argv[2] ?? "soffice", execution = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 }, hash = b => createHash("sha256").update(b).digest("hex");
const application = (await execute(executable, ["--version"], execution)).stdout.trim();
await mkdir(".tmp/script-font-import-native", { recursive: true }); const directory = await mkdtemp(path.resolve(".tmp/script-font-import-native/run-"));
for (const folder of ["native", "source-pdf", "saved-pdf"]) await mkdir(path.join(directory, folder));
const options = { projection: "content", textFontProfile: "script-families-v1", editSaveProfile: "text-runs-v3" }, snapshots = [], hashes = [];
let bytes = (await generatePptx(createScriptFontFixture(createPresentation))).bytes;
for (let round = 0; round <= 3; round++) {
  const file = `round-${round}.pptx`; await writeFile(path.join(directory, file), bytes); hashes.push({ round, sha256: hash(bytes) });
  const session = await importPptx(bytes, options);
  try {
    snapshots.push({ file, document: normalizePresentation(session.document), fonts: [...session.report.textStyleInheritance.properties, ...session.report.tableStyleInheritance.properties].filter(p => ["run.fontFamily", "run.eastAsianFontFamily", "run.complexScriptFontFamily"].includes(p.property)) });
    if (round === 3) break;
    const slide = session.document.slides[0]; let owner = 0;
    const change = paragraphs => { for (const p of paragraphs) for (const r of p.runs) r.text = `Round${round + 1} owner${++owner} 中文العربية`; };
    for (const element of [...slide.elements]) {
      const next = structuredClone(element);
      if (next.type === "text") change(next.content);
      else if (next.type === "shape" && next.text) change(next.text.content);
      else if (next.type === "table") for (const row of next.rows) for (const cell of row.cells) change(cell.content);
      slide.replaceElement(next.id, next);
    }
    const notes = structuredClone(slide.notes); change(notes); slide.setNotes(notes);
    const ready = await session.prepareSave({ mode: "edit" }); if (ready.status !== "ready") throw new Error("Mixed-owner scalar source save was blocked."); bytes = (await session.save(ready.plan)).bytes;
  } finally { session.dispose(); }
}
const expectedFile = path.join(directory, "expected.json"); await writeFile(expectedFile, JSON.stringify(snapshots));
const logs = [], failures = [], profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
async function attempt(stage, args) { try { const result = await execute(executable, [profile, "--headless", ...args], execution); logs.push({ stage, stdout: result.stdout, stderr: result.stderr }); } catch (e) { failures.push({ stage, code: e.code, message: e.stderr || e.message }); } }
const source = path.join(directory, "round-3.pptx"), saved = path.join(directory, "native", "round-3.pptx");
await attempt("resave", ["--convert-to", "pptx:Impress Office Open XML", "--outdir", path.join(directory, "native"), source]);
await attempt("source-pdf", ["--convert-to", "pdf:impress_pdf_Export", "--outdir", path.join(directory, "source-pdf"), source]);
let nativeSha256; try { nativeSha256 = hash(await readFile(saved)); await attempt("saved-pdf", ["--convert-to", "pdf:impress_pdf_Export", "--outdir", path.join(directory, "saved-pdf"), saved]); } catch (e) { if (e.code !== "ENOENT") throw e; }
const independent = JSON.parse((await execute("python3", ["scripts/research/oracles/script-font-import.py", expectedFile, directory], execution)).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory), profile: options, hashes, ...(nativeSha256 ? { nativeSha256 } : {}), logs, failures, independent, scope: "Explicit source font-slot projection and three mixed-owner scalar preservation saves: independent actual source declarations and exact non-character byte masks. Native/PDF are observations only; fonts are synthetic and typography is not certified.", sourceFontWrites: "blocked", interactiveEditing: "untested" };
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n"); await writeFile("docs/testing/script-font-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, failures: failures.length, rounds: independent.rounds, native: independent.native.status, pdf: independent.pdf.map(p => ({ kind: p.kind, status: p.status })) }));
