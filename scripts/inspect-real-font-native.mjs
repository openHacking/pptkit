// Optional native research. Host font configuration is isolated to each child process.
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, writeFile, readFile } from "node:fs/promises";
import { platform, release, arch } from "node:os";
import { createHash } from "node:crypto";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { buildNativeFontCase, nativeFontCases, nativeFontPrecedenceCases, nativeFontFamilies } from "./research/font-native-cases.mjs";

const arguments_ = process.argv.slice(2), precedence = arguments_.includes("--precedence"), prepareOnly = arguments_.includes("--prepare-only");
const positional = arguments_.filter(argument => !["--precedence", "--prepare-only"].includes(argument));
if (positional.length > 1 || positional.some(argument => argument.startsWith("--"))) throw new Error("Usage: node scripts/inspect-real-font-native.mjs [soffice-executable] [--precedence] [--prepare-only]");
const execute = promisify(execFile), executable = positional[0] ?? "soffice", execution = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 }, hash = bytes => createHash("sha256").update(bytes).digest("hex");
const application = prepareOnly ? "not invoked (prepare-only)" : (await execute(executable, ["--version"], execution)).stdout.trim();
await mkdir(".tmp/real-font-native", { recursive: true }); const directory = await mkdtemp(path.resolve(".tmp/real-font-native/run-"));
for (const folder of ["source", "pdf", "native", "native-pdf", "font-cache"]) await mkdir(path.join(directory, folder));
const escape = s => s.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;");
const fontconfig = path.join(directory, "fonts.conf"); await writeFile(fontconfig, `<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd"><fontconfig><dir>/System/Library/Fonts</dir><dir>/Library/Fonts</dir><cachedir>${escape(path.join(directory, "font-cache"))}</cachedir></fontconfig>`);
const controls = [], fonts = [];
for (const [id, family] of Object.entries(nativeFontFamilies)) {
  const result = (await execute("fc-match", ["-f", "%{family[0]}\t%{postscriptname}\t%{file}\t%{index}\n", family], execution)).stdout.trim().split("\t"), [actualFamily, postscriptName, file, index] = result;
  if (actualFamily !== family) throw new Error(`Requested installed font is unavailable: ${family} -> ${actualFamily}`);
  fonts.push({ id, family, postscriptName, file, index: Number(index), sha256: hash(await readFile(file)) });
}
for (const control of precedence ? nativeFontPrecedenceCases : nativeFontCases) {
  const built = await buildNativeFontCase(control); await writeFile(path.join(directory, "source", `${control.id}.pptx`), built.bytes); controls.push({ ...control, sourceSha256: hash(built.bytes), projection: built.projection });
}
const expected = path.join(directory, "expected.json"); await writeFile(expected, JSON.stringify({ controls, fonts }));
if (prepareOnly) {
  await writeFile(path.join(directory, "prepared.json"), JSON.stringify({ controls: controls.map(control => ({ id: control.id, sourceSha256: control.sourceSha256 })), fonts }, null, 2) + "\n");
  console.log(JSON.stringify({ directory, controls: controls.length, status: "prepared", native: "not invoked" }));
  process.exit(0);
}
const logs = [], failures = [], profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const environment = { ...process.env, FONTCONFIG_FILE: fontconfig };
for (const [stage, type, folder, input] of [["source-pdf", "pdf:impress_pdf_Export", "pdf", "source"], ["resave", "pptx:Impress Office Open XML", "native", "source"], ["saved-pdf", "pdf:impress_pdf_Export", "native-pdf", "native"]]) {
  for (let offset = 0; offset < controls.length; offset += 5) {
    const files = controls.slice(offset, offset + 5).map(c => path.join(directory, input, `${c.id}.pptx`));
    try { const result = await execute(executable, [profile, "--headless", "--convert-to", type, "--outdir", path.join(directory, folder), ...files], { ...execution, env: environment }); logs.push({ stage, offset, stdout: result.stdout, stderr: result.stderr }); }
    catch (e) { failures.push({ stage, offset, code: e.code, message: e.stderr || e.message }); }
  }
}
const independent = JSON.parse((await execute("python3", ["scripts/research/oracles/real-font-native.py", expected, directory], execution)).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory), fontconfig: { file: path.relative(process.cwd(), fontconfig), sha256: hash(await readFile(fontconfig)), directories: ["/System/Library/Fonts", "/Library/Fonts"], scope: "child-process environment only" }, fonts, controls, failures, logs, independent, status: "observed", nativeGui: "untested" };
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n");
await writeFile(precedence ? "docs/testing/real-font-precedence-lo-evidence.json" : "docs/testing/real-font-native-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
await writeFile(".tmp/real-font-native/latest.txt", directory);
console.log(JSON.stringify({ directory, application, failures: failures.length, comparisons: independent.comparisons.map(c => ({ id: c.id, selected: c.projectedFamilies, actual: c.sourcePdfFonts, matches: c.projectedFamiliesMatch })) }));
