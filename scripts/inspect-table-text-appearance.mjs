// Optional independent source/PDF observations; no GUI or native editing claims.
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { TABLE_STYLE_ID } from "../packages/pptx-parser/test/table-style-fixture.mjs";
import { textReference, textColor, fontCollection } from "../packages/pptx-parser/test/table-text-appearance-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-table-text-appearance.mjs [soffice-executable]");
const execute = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const hash = b => createHash("sha256").update(b).digest("hex");
const application = (await execute(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/table-text-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/table-text-native/run-"));
for (const folder of ["native", "source-pdf", "saved-pdf"]) await mkdir(path.join(directory, folder));
const failures = [], logs = [], expected = [], hashes = [];
const attempt = async (name, stage, args) => {
  try { const result = await execute(executable, args, options); logs.push({ name, stage, stdout: result.stdout, stderr: result.stderr }); }
  catch (e) { failures.push({ name, stage, code: e.code, message: e.stderr || e.message }); }
};
for (const name of ["reference", "explicit", "local", "collection"]) {
  const document = createPresentation({ size: { width: 400, height: 300 }, theme: { colors: { accent1: "3579BD", accent2: "CA6420" }, fonts: { body: "Arial", heading: "Arial" } } });
  document.addSlide().addElement({ type: "table", box: { x: 30, y: 30, width: 280, height: 120 }, columns: [140, 140], rows: Array.from({ length: 2 }, (_, r) => ({ height: 60, cells: Array.from({ length: 2 }, (_, c) => ({ content: [{ runs: [{ text: `R${r}C${c}`, style: { fontFamily: "Arial", fontSize: 16 } }] }] })) })) });
  const parts = readGeneratedZip((await generatePptx(document)).bytes), namespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const tx = name === "collection" ? fontCollection("+mn-lt", "+mn-ea", "+mn-cs") + textColor("CA6420", "", "srgbClr")
    : textReference(textColor("accent2")) + (name === "explicit" ? textColor("102030", "", "srgbClr") : "");
  parts.set("ppt/tableStyles.xml", Buffer.from(`<a:tblStyleLst xmlns:a="${namespace}" def="${TABLE_STYLE_ID}"><a:tblStyle styleId="${TABLE_STYLE_ID}" styleName="Appearance"><a:wholeTbl><a:tcTxStyle>${tx}</a:tcTxStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>`));
  let slide = parts.get("ppt/slides/slide1.xml").toString().replace(/<a:tblPr[^>]*>[\s\S]*?<\/a:tblPr>/, `<a:tblPr><a:tableStyleId>${TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`)
    .replace(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g, '<a:rPr sz="1600"/>');
  if (name === "local") slide = slide.replaceAll('<a:rPr sz="1600"/>', `<a:rPr sz="1600"><a:solidFill>${textColor("445566", "", "srgbClr")}</a:solidFill></a:rPr>`);
  parts.set("ppt/slides/slide1.xml", Buffer.from(slide));
  let theme = parts.get("ppt/theme/theme1.xml").toString();
  theme = theme.replace(/<a:(ea|cs) typeface="[^"]*"\s*\/>/g, '<a:$1 typeface="Arial"/>');
  parts.set("ppt/theme/theme1.xml", Buffer.from(theme));
  const bytes = zip([...parts].map(([name, body]) => ({ name, body }))).bytes, source = path.join(directory, `${name}.pptx`);
  await writeFile(source, bytes);
  const s = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
  try { expected.push({ name, table: normalizePresentation(s.document).slides[0].elements[0], provenance: s.report.tableStyleInheritance }); }
  finally { s.dispose(); }
  const item = { name, sourceSha256: hash(bytes) }; hashes.push(item);
  const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, `profile-${name}`)).href}`;
  await attempt(name, "resave", [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", path.join(directory, "native"), source]);
  await attempt(name, "source-pdf", [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", path.join(directory, "source-pdf"), source]);
  const saved = path.join(directory, "native", `${name}.pptx`);
  try {
    item.savedSha256 = hash(await readFile(saved));
    await attempt(name, "saved-pdf", [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", path.join(directory, "saved-pdf"), saved]);
  } catch (e) { if (e.code !== "ENOENT") throw e; }
}
const expectedFile = path.join(directory, "expected.json"); await writeFile(expectedFile, JSON.stringify(expected));
const independent = JSON.parse((await execute("python3", ["scripts/research/oracles/table-text-appearance.py", expectedFile, directory], options)).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory), hashes, failures, logs, independent, scope: "Independent selected source text color/font dependencies and source-control/resaved PDF labels, colors and actual font observations. Full text defaults/script shaping, native interactive editing and source-style writes remain gated.", interactiveEditing: "untested", sourceStyleWrites: "blocked" };
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n");
await writeFile("docs/testing/table-text-appearance-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, failures: failures.length, scenes: independent.scenes.map(s => ({ name: s.name, pdf: s.pdf.map(p => ({ kind: p.kind, status: p.status })) })) }));
