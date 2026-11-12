// Optional named headless observations. Source oracle is independent of font selection.
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { createScriptFontFixture } from "../tests/fixtures/script-fonts.mjs";
if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-script-fonts.mjs [soffice-executable]");
const execute = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 }, hash = b => createHash("sha256").update(b).digest("hex");
const application = (await execute(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/script-font-native", { recursive: true }); const directory = await mkdtemp(path.resolve(".tmp/script-font-native/run-"));
for (const folder of ["native", "source-pdf", "saved-pdf"]) await mkdir(path.join(directory, folder));
const p = createScriptFontFixture(createPresentation), normalized = normalizePresentation(p), source = path.join(directory, "script-fonts.pptx"), bytes = (await generatePptx(p)).bytes;
await writeFile(source, bytes); const expectedFile = path.join(directory, "expected.json"); await writeFile(expectedFile, JSON.stringify(normalized));
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`, logs = [], failures = [];
async function attempt(stage, args) {
  try { const result = await execute(executable, [profile, "--headless", ...args], options); logs.push({ stage, stdout: result.stdout, stderr: result.stderr }); }
  catch (e) { failures.push({ stage, code: e.code, message: e.stderr || e.message }); }
}
await attempt("resave", ["--convert-to", "pptx:Impress Office Open XML", "--outdir", path.join(directory, "native"), source]);
await attempt("source-pdf", ["--convert-to", "pdf:impress_pdf_Export", "--outdir", path.join(directory, "source-pdf"), source]);
const saved = path.join(directory, "native", "script-fonts.pptx"); let savedSha256;
try { savedSha256 = hash(await readFile(saved)); await attempt("saved-pdf", ["--convert-to", "pdf:impress_pdf_Export", "--outdir", path.join(directory, "saved-pdf"), saved]); } catch (e) { if (e.code !== "ENOENT") throw e; }
const independent = JSON.parse((await execute("python3", ["scripts/research/oracles/script-fonts.py", expectedFile, directory], options)).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory), sourceSha256: hash(bytes), ...(savedSha256 ? { savedSha256 } : {}), failures, logs, independent, scope: "Independent author-to-XML font-slot oracle and deliberate mutation; named native resave run topology and PDF font/text observations only. No full typography/shaping/bidi/interactive or source semantic import claims.", interactiveEditing: "untested", multiFamilySemanticImport: "not-admitted" };
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n"); await writeFile("docs/testing/script-fonts-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, failures: failures.length, sourceVerified: independent.sourceVerified, native: independent.native.status, pdf: independent.pdf.map(p => ({ kind: p.kind, status: p.status })) }));
