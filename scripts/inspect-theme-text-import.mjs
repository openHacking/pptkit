// Optional native rendering/producer experiment; no interactive editing claim.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdtemp, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { pack } from "../packages/pptx-parser/test/opc-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-theme-text-import.mjs [soffice-executable]");
const run = promisify(execFile), executable = process.argv[2] ?? "soffice";
const options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/theme-text-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/theme-text-native/run-"));
const pdfDirectory = path.join(directory, "pdf"), nativeDirectory = path.join(directory, "native");
await mkdir(pdfDirectory); await mkdir(nativeDirectory);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const document = createPresentation({
  theme: { colors: { accent1: "13579B" }, fonts: { heading: "Liberation Sans", body: "Liberation Sans" } },
});
for (const text of ["Theme first page", "Theme second page"]) document.addSlide().addElement({
  type: "text", name: text, accessibility: { description: text },
  box: { x: 50, y: 50, width: 600, height: 100 },
  content: [{ runs: [{ text, style: { color: { theme: "accent1" }, fontFamily: { theme: "heading" }, fontSize: 24 } }] }],
});
const base = await generatePptx(document);
assert.deepEqual(base.warnings, []);
const parts = readGeneratedZip(base.bytes), theme = parts.get("ppt/theme/theme1.xml").toString();
const fontScheme = /<a:fontScheme\b[\s\S]*?<\/a:fontScheme>/.exec(theme)[0].replaceAll("Liberation Sans", "Liberation Serif");
const colorScheme = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(theme)[0].replace("13579B", "A13B57");
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const addOverride = (name, relPart, component) => {
  const part = `ppt/theme/${name}.xml`;
  parts.set(part, Buffer.from(`<a:themeOverride xmlns:a="${A}">${component}</a:themeOverride>`));
  parts.set(relPart, Buffer.from(parts.get(relPart).toString().replace("</Relationships>",
    `<Relationship Id="native-theme-${name}" Type="${R}/themeOverride" Target="../theme/${name}.xml"/></Relationships>`)));
};
addOverride("fontOverride", "ppt/slideLayouts/_rels/slideLayout1.xml.rels", fontScheme);
addOverride("colorOverride", "ppt/slides/_rels/slide1.xml.rels", colorScheme);
for (const part of ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]) {
  // This is the bounded Latin profile; script-specific theme-font selection is
  // explicitly outside it. Keep the native Latin reference, not direct text.
  parts.set(part, Buffer.from(parts.get(part).toString().replace(/<a:(?:ea|cs)\b[^>]*\/>/g, "")));
}
const types = parts.get("[Content_Types].xml").toString().replace("</Types>",
  ["fontOverride", "colorOverride"].map(name => `<Override PartName="/ppt/theme/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.themeOverride+xml"/>`).join("") + "</Types>");
parts.delete("[Content_Types].xml");
const source = pack(Object.fromEntries(parts), types);
const session = await importPptx(source, { projection: "content" });
try {
  const styles = document => normalizePresentation(document).slides.map(slide => {
    const element = slide.elements.find(element => element.type === "text" || (element.type === "shape" && element.text));
    assert.ok(element, "native sample must retain a text object on each page");
    const run = (element.type === "text" ? element.content : element.text.content)[0].runs[0];
    return { text: run.text, color: run.style.color, fontFamily: run.style.fontFamily, fontSize: run.style.fontSize };
  });
  const expected = [
    { text: "Theme first page", color: "A13B57", fontFamily: "Liberation Serif", fontSize: 24 },
    { text: "Theme second page", color: "13579B", fontFamily: "Liberation Serif", fontSize: 24 },
  ];
  const projection = styles(session.document);
  assert.deepEqual(projection, expected);
  const prepared = await session.prepareSave({ mode: "no-op" });
  assert.equal(prepared.status, "ready");
  assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), source);
  const fresh = await generatePptx(session.document);
  assert.deepEqual(fresh.warnings, []);
  const logs = {}, filenames = {};
  for (const [name, bytes] of [["base", base.bytes], ["source", source], ["projection", fresh.bytes]]) {
    const filename = path.join(directory, `${name}.pptx`); filenames[name] = filename;
    await writeFile(filename, bytes);
    logs[name] = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdfDirectory, filename], options);
  }
  logs.nativeSave = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, filenames.source], options);
  const nativeBytes = await readFile(path.join(nativeDirectory, "source.pptx"));
  const native = await importPptx(nativeBytes, { projection: "content" });
  let nativeProjection;
  try { nativeProjection = styles(native.document); } finally { native.dispose(); }
  const pdfFiles = ["base", "source", "projection"].map(name => path.join(pdfDirectory, `${name}.pdf`));
  const pdfStyles = JSON.parse((await run("python3", ["-c", `
import sys, json, fitz
result=[]
for filename in sys.argv[1:]:
    with fitz.open(filename) as document:
        pages=[]
        for page in document:
            spans=[s for b in page.get_text('dict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans'] if s['text'].startswith('Theme')]
            pages.append([{'text':s['text'],'font':s['font'],'color':format(s['color'],'06X'),'size':s['size']} for s in spans])
        result.append(pages)
print(json.dumps(result))
`, ...pdfFiles], options)).stdout);
  const hash = bytes => createHash("sha256").update(bytes).digest("hex");
  const evidence = {
    schemaVersion: 1, recordedAt: new Date().toISOString(), application,
    platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory),
    scope: "Repository-authored partial layout font override plus slide color override, Latin-only projection, exact no-op, LibreOfficeDev PDF text/color/font inspection and native producer save. Not interactive editing, whole-slide visual parity, arbitrary script fonts, PowerPoint or preservation edit-save.",
    source: { sha256: hash(source), expected }, projection,
    noOpOriginalBytes: true, freshExport: { sha256: hash(fresh.bytes) },
    nativeSource: { sha256: hash(nativeBytes), projection: nativeProjection, sameSelectedStyles: JSON.stringify(nativeProjection) === JSON.stringify(expected) },
    pdf: {
      base: pdfStyles[0], source: pdfStyles[1], freshExport: pdfStyles[2],
      sameSelectedStyles: JSON.stringify(pdfStyles[1]) === JSON.stringify(pdfStyles[2]),
      observedFontChange: JSON.stringify(pdfStyles[0].map(page => page.map(span => span.font))) !== JSON.stringify(pdfStyles[1].map(page => page.map(span => span.font))),
      sourceColorsMatchProjection: pdfStyles[1].every((page, i) => page.length === 1 && page[0].color === expected[i]?.color),
      artifacts: await Promise.all(pdfFiles.map(async filename => {
        const bytes = await readFile(filename);
        return { name: path.basename(filename), sha256: hash(bytes), byteLength: bytes.length };
      })),
    },
    logs,
  };
  await writeFile(path.join(directory, "import-report.json"), JSON.stringify(session.report, null, 2) + "\n");
  await writeFile("docs/testing/theme-text-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
  // Native producer/rendering differences are observed outcomes, not permission
  // to change expected Core values or to claim that this application has parity.
  for (let page = 0; page < 2; page++) {
    assert.equal(pdfStyles[1][page].length, 1);
    assert.equal(pdfStyles[1][page][0].text, expected[page].text);
    assert.equal(pdfStyles[2][page].length, 1);
    assert.equal(pdfStyles[2][page][0].text, expected[page].text);
  }
  console.log(JSON.stringify({ directory, nativeSelectedStylesAgree: evidence.nativeSource.sameSelectedStyles,
    pdfSelectedStylesAgree: evidence.pdf.sameSelectedStyles, observedFontChange: evidence.pdf.observedFontChange,
    sourceColorsMatchProjection: evidence.pdf.sourceColorsMatchProjection }));
} finally { session.dispose(); }
