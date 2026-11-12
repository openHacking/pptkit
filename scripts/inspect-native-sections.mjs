// Optional native evidence runner. Not part of portable CI; requires soffice and Python 3.
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
import { inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
import { createFixture } from "../tests/fixtures/presentations.mjs";
import { section, sectionExtension, SECTION_IDS } from "../packages/pptx-parser/test/sections-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-native-sections.mjs [soffice-executable]");
const run = promisify(execFile), executable = process.argv[2] ?? "soffice";
const options = { timeout: 45000, maxBuffer: 4 * 1024 * 1024 };
const version = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/native-sections", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/native-sections/run-")), output = path.join(directory, "saved"), reopened = path.join(directory, "reopened");
await mkdir(output); await mkdir(reopened);
const generated = await generatePptx(createFixture("text", createPresentation));
assert.equal(generated.slideCount, 2);
const base = path.join(directory, "base.pptx"), input = path.join(directory, "sections.pptx");
await writeFile(base, generated.bytes);
const payload = `<p:extLst>${sectionExtension(section([256], `id="${SECTION_IDS[0]}" name="Introduction"`) + section([], `id="${SECTION_IDS[1]}" name="Empty"`) + section([257], `id="${SECTION_IDS[2]}" name="Conclusion"`))}</p:extLst>`;
await run("python3", ["-c", `
import sys, zipfile
source, target, payload = sys.argv[1:]
with zipfile.ZipFile(source) as src, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as dst:
    for entry in src.infolist():
        data = src.read(entry.filename)
        if entry.filename == 'ppt/presentation.xml':
            text = data.decode('utf-8')
            assert '<p:extLst>' not in text and text.count('</p:presentation>') == 1
            data = text.replace('</p:presentation>', payload + '</p:presentation>').encode('utf-8')
        dst.writestr(entry.filename, data)
`, base, input, payload], options);
const inputBytes = await readFile(input), before = await inspectPptx(inputBytes);
assert.deepEqual(before.sectionList.sections.map((s) => s.slides.map((slide) => slide.id)), [[256], [], [257]]);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const saveLog = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", output, input], options);
const savedPath = path.join(output, "sections.pptx"), savedBytes = await readFile(savedPath), after = await inspectPptx(savedBytes);
const reopenLog = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", reopened, savedPath], options);
const pdf = await readFile(path.join(reopened, "sections.pdf")); assert.ok(pdf.subarray(0, 5).equals(Buffer.from("%PDF-")));
const independent = JSON.parse((await run("python3", ["-c", `
import sys, zipfile, json, xml.etree.ElementTree as ET
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 's': 'http://schemas.microsoft.com/office/powerpoint/2010/main'}
result = []
for filename in sys.argv[1:]:
    with zipfile.ZipFile(filename) as z:
        assert z.testzip() is None
        root = ET.fromstring(z.read('ppt/presentation.xml'))
        sections = root.findall('p:extLst/p:ext/s:sectionLst/s:section', ns)
        result.append([{'id': s.get('id'), 'name': s.get('name'), 'slides': [int(n.get('id')) for n in s.findall('s:sldIdLst/s:sldId', ns)]} for s in sections])
print(json.dumps(result))
`, input, savedPath], options)).stdout);
const summarize = (report) => (report.sectionList?.sections ?? []).map((s) => ({ id: s.id ?? null, name: s.name ?? null, slides: s.slides.map((slide) => slide.id) }));
assert.deepEqual([summarize(before), summarize(after)], independent);
const projected = await importPptx(savedBytes, { projection: "structure" });
const namesBySourceId = new Map(independent[1].flatMap(section => section.slides.map(id => [id, section.name])));
const labels = after.slides.map((slide, index) => ({ sourceSlideId: slide.id, coreSlideId: projected.document.slides[index].id, name: projected.document.slides[index].section }));
assert.deepEqual(labels.map(label => label.name), after.slides.map(slide => namesBySourceId.get(slide.id)));
assert.ok(projected.report.findings.some(f => f.feature === "section.native-collection" && f.operation === "export" && f.tier === "unsupported"));
const noOp = await projected.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready");
assert.equal(Buffer.compare(Buffer.from((await projected.save(noOp.plan)).bytes), savedBytes), 0);
const labelProjection = {
  labels,
  independentMembershipMatch: true,
  originalByteNoOp: true,
  emptySectionsRetainedInReport: projected.report.inspection.sectionList.sections.filter(s => s.slides.length === 0).length,
  nativeCollectionMapping: "unsupported",
  freshExport: "PPTKit slide-label metadata only; native collection is not exported",
};
projected.dispose();
const hash = (bytes) => createHash("sha256").update(bytes).digest("hex");
const equivalent = JSON.stringify(independent[0]) === JSON.stringify(independent[1]);
const evidence = {
  schemaVersion: 1, recordedAt: new Date().toISOString(), application: version, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory),
  scope: "Repository-generated text deck with independently injected MS-PPTX section extension; native headless save/reopen, Core member-label projection and exact no-op bytes, not native collection authoring, interactive editability or PowerPoint verification",
  input: { sha256: hash(inputBytes), bytes: inputBytes.length, sections: independent[0] },
  saved: { sha256: hash(savedBytes), bytes: savedBytes.length, sections: independent[1], findings: after.findings },
  reopenedPdf: { sha256: hash(pdf), bytes: pdf.length }, sectionEquivalence: equivalent ? "matched" : "changed-or-lost",
  independentXmlComparison: "passed for input and saved output", powerPoint: "untested", interactiveEditability: "untested", repairDialogs: "not observable in headless mode", visualFidelity: "unreviewed",
  labelProjection,
  logs: { save: saveLog, reopen: reopenLog },
};
await writeFile("docs/testing/native-sections-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ application: version, sectionEquivalence: evidence.sectionEquivalence, inputSections: independent[0].length, savedSections: independent[1].length, artifactsDirectory: evidence.artifactsDirectory }));
