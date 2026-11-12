// Optional native producer/rendering experiment. No interactive-editing claim.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdtemp, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { pack } from "../packages/pptx-parser/test/opc-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-color-transform-import.mjs [soffice-executable]");
const execute = promisify(execFile), executable = process.argv[2] ?? "soffice";
const options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/color-transform-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/color-transform-native/run-"));
const pdf = path.join(directory, "pdf"), nativeDirectory = path.join(directory, "native");
await mkdir(pdf); await mkdir(nativeDirectory);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const cases = [
  { name: "luminance-half", base: "336699", transforms: '<a:lumMod val="50000"/>', color: "1A334D", opacity: 1 },
  { name: "luminance-offset", base: "336699", transforms: '<a:lumOff val="20000"/>', color: "6699CC", opacity: 1 },
  { name: "saturation-half", base: "336699", transforms: '<a:satMod val="50000"/>', color: "4D6680", opacity: 1 },
  { name: "black-tint-half", base: "000000", transforms: '<a:tint val="50000"/>', color: "BCBCBC", opacity: 1 },
  { name: "gray-shade-half", base: "808080", transforms: '<a:shade val="50000"/>', color: "5C5C5C", opacity: 1 },
  { name: "ordered-hsl", base: "336699", transforms: '<a:lumMod val="50000"/><a:lumOff val="20000"/>', color: "336699", opacity: 1 },
  { name: "alpha-half", base: "000000", transforms: '<a:alpha val="50000"/>', color: "000000", opacity: .5 },
  { name: "alpha-transparent", base: "000000", transforms: '<a:alpha val="0"/>', color: "000000", opacity: 0 },
  { name: "alpha-composition", base: "000000", transforms: '<a:alpha val="50000"/><a:alphaMod val="50000"/><a:alphaOff val="10000"/>', color: "000000", opacity: .35 },
];
const document = createPresentation({ theme: { fonts: { heading: "Liberation Sans", body: "Liberation Sans" } } });
for (const [index, sample] of cases.entries()) document.addSlide().addElement({
  type: "text", box: { x: 50, y: 50, width: 600, height: 100 },
  content: [{ runs: [{ text: `Case ${index}: ${sample.name}`, style: { color: sample.base, fontFamily: "Liberation Sans", fontSize: 36 } }] }],
});
const generated = await generatePptx(document), parts = readGeneratedZip(generated.bytes);
for (const [index, sample] of cases.entries()) {
  const part = `ppt/slides/slide${index + 1}.xml`, original = parts.get(part).toString();
  const modified = original.replace(`<a:srgbClr val="${sample.base}"></a:srgbClr>`, `<a:srgbClr val="${sample.base}">${sample.transforms}</a:srgbClr>`);
  assert.notEqual(modified, original, sample.name);
  parts.set(part, Buffer.from(modified));
}
const contentTypes = parts.get("[Content_Types].xml").toString();
parts.delete("[Content_Types].xml");
const source = pack(Object.fromEntries(parts), contentTypes);
const styles = doc => doc.slides.map(slide => {
  const element = slide.elements.find(e => e.type === "text" || (e.type === "shape" && e.text));
  assert.ok(element, "each sample retains its text object");
  const style = (element.type === "text" ? element.content : element.text.content)[0].runs[0].style;
  return { color: style.color, opacity: style.colorOpacity ?? 1 };
});
const session = await importPptx(source, { projection: "content" });
try {
  const expected = cases.map(({ color, opacity }) => ({ color, opacity }));
  assert.deepEqual(styles(session.document), expected);
  const noOp = await session.prepareSave({ mode: "no-op" });
  assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), source);
  const fresh = await generatePptx(session.document), logs = {};
  for (const [name, bytes] of [["source", source], ["projection", fresh.bytes]]) {
    const file = path.join(directory, `${name}.pptx`);
    await writeFile(file, bytes);
    logs[name] = await execute(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdf, file], options);
  }
  logs.nativeSave = await execute(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, path.join(directory, "source.pptx")], options);
  const nativeBytes = await readFile(path.join(nativeDirectory, "source.pptx")), native = await importPptx(nativeBytes, { projection: "content" });
  let nativeStyles;
  try { nativeStyles = styles(native.document); } finally { native.dispose(); }
  const pdfInspection = JSON.parse((await execute("python3", ["-c", `
import sys, json, hashlib, fitz
result=[]
for filename in sys.argv[1:]:
    pages=[]
    with fitz.open(filename) as doc:
        for page in doc:
            spans=[s for b in page.get_text('dict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans'] if s['text'].startswith('Case')]
            pix=page.get_pixmap(matrix=fitz.Matrix(1,1), alpha=False, colorspace=fitz.csRGB)
            pages.append({'spans':[{'text':s['text'],'color':format(s['color'],'06X'),'alpha':s.get('alpha')} for s in spans],
                          'width':pix.width,'height':pix.height,'rasterSha256':hashlib.sha256(pix.samples).hexdigest()})
    result.append(pages)
print(json.dumps(result))
`, path.join(pdf, "source.pdf"), path.join(pdf, "projection.pdf")], options)).stdout);
  assert.equal(pdfInspection[0].length, cases.length);
  assert.equal(pdfInspection[1].length, cases.length);
  const hash = bytes => createHash("sha256").update(bytes).digest("hex");
  const evidence = {
    schemaVersion: 1, recordedAt: new Date().toISOString(), application, platform: `${platform()} ${release()} ${arch()}`,
    scope: "Nine named direct text-color samples: deterministic Parser projection, exact no-op, LibreOfficeDev headless producer save and independent PDF spans/rasters. Not PowerPoint, interactive editing, arbitrary color transforms or source edit-save.",
    artifactsDirectory: path.relative(process.cwd(), directory),
    sourceSha256: hash(source), freshSha256: hash(fresh.bytes), nativeSha256: hash(nativeBytes), noOpOriginalBytes: true,
    cases: cases.map((sample, index) => ({ ...sample, nativeProjection: nativeStyles[index],
      nativeSelectedStyleAgrees: JSON.stringify(nativeStyles[index]) === JSON.stringify(expected[index]),
      sourcePdf: pdfInspection[0][index], projectionPdf: pdfInspection[1][index],
      sameNativeRaster: pdfInspection[0][index].rasterSha256 === pdfInspection[1][index].rasterSha256,
    })), logs,
  };
  await writeFile("docs/testing/color-transform-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
  console.log(JSON.stringify({ directory, cases: evidence.cases.map(({ name, nativeSelectedStyleAgrees, sameNativeRaster }) => ({ name, nativeSelectedStyleAgrees, sameNativeRaster })) }, null, 2));
} finally { session.dispose(); }
