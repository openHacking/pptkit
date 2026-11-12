// Optional native producer/PDF evidence for the explicit text-runs-v2 profile.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, arch, release } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx, compareOpcPackages } from "../packages/pptx-parser/dist/index.js";
import { createTextOwnerFixture } from "../tests/fixtures/text-owners.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { png } from "../packages/pptx-parser/test/image-fixture.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-text-owner-save.mjs [soffice-executable]");
const executable = process.argv[2] ?? "soffice", execute = promisify(execFile), options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/text-owner-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/text-owner-native/run-")), nativeDirectory = path.join(directory, "native"), pdfDirectory = path.join(directory, "pdf");
await mkdir(nativeDirectory); await mkdir(pdfDirectory);
const parts = readGeneratedZip((await generatePptx(createTextOwnerFixture(createPresentation, png))).bytes);
const opaque = '<p:ext uri="urn:text-owner-evidence"><x:opaque xmlns:x="urn:future">opaque retained owner source</x:opaque></p:ext>';
for (const [part, end] of [["ppt/slides/slide1.xml", "sld"], ["ppt/notesSlides/notesSlide1.xml", "notes"]]) {
  const xml = parts.get(part).toString();
  parts.set(part, Buffer.from(xml.includes("</p:extLst>") ? xml.replace("</p:extLst>", opaque + "</p:extLst>") : xml.replace(`</p:${end}>`, `<p:extLst>${opaque}</p:extLst></p:${end}>`)));
}
parts.set("orphan.bin", Buffer.from([1, 8, 27, 64]));
parts.set("[Content_Types].xml", Buffer.from(parts.get("[Content_Types].xml").toString().replace("</Types>", '<Override PartName="/orphan.bin" ContentType="application/octet-stream"/></Types>')));
const source = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
const expectedShape = "Edited preserved shape 123 <&>", expectedNotes = "Edited preserved notes 456 <&>";
const session = await importPptx(source, { projection: "content", editSaveProfile: "text-runs-v2" });
let edited, saveReport;
try {
  const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "shape" && e.text);
  const shape = structuredClone(slide.removeElement(slide.elements[index].id));
  shape.text.content[0].runs[0].text = expectedShape; slide.insertElement(index, shape);
  const notes = structuredClone(slide.notes); notes[0].runs[0].text = expectedNotes; slide.setNotes(notes);
  const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
  const saved = await session.save(prepared.plan); edited = saved.bytes; saveReport = saved.report;
} finally { session.dispose(); }
const sourceFile = path.join(directory, "source.pptx"), editedFile = path.join(directory, "edited.pptx");
await writeFile(sourceFile, source); await writeFile(editedFile, edited);
const nativeProfile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const logs = {
  nativeSave: await execute(executable, [nativeProfile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, editedFile], options),
  pdf: await execute(executable, [nativeProfile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdfDirectory, editedFile], options),
};
const nativeFile = path.join(nativeDirectory, "edited.pptx"), pdfFile = path.join(pdfDirectory, "edited.pdf");
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as ET,fitz
source,edited,native,pdf,shape,notes=sys.argv[1:]
def package(file):
  with zipfile.ZipFile(file) as archive:
    return {name:archive.read(name) for name in archive.namelist() if not name.endswith('/')}
a,b,c=map(package,[source,edited,native])
assert set(a)==set(b)
changes=sorted(name for name in a if a[name]!=b[name])
assert changes==['ppt/notesSlides/notesSlide1.xml','ppt/slides/slide1.xml']
assert a['orphan.bin']==b['orphan.bin']
for name in changes:
  assert b'opaque retained owner source' in b[name]
ns='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def texts(parts,prefix):
  return [''.join(node.itertext()) for name,data in parts.items() if name.startswith(prefix) and name.endswith('.xml') for node in ET.fromstring(data).iter(ns+'t')]
candidate_slides,native_slides=texts(b,'ppt/slides/'),texts(c,'ppt/slides/')
candidate_notes,native_notes=texts(b,'ppt/notesSlides/'),texts(c,'ppt/notesSlides/')
for values in [candidate_slides,native_slides]:
  assert shape in values
  assert 'Second slide control' in values
  assert 'Nested shape control' in values
  assert 'Cell control A' in values and 'Cell control B' in values
for values in [candidate_notes,native_notes]:
  assert notes in values
  assert 'Second notes control' in values
  assert 'Notes control e\u0301🙂' in values
assert 'Original shape text' not in native_slides
assert 'Original speaker notes' not in native_notes
pdf_doc=fitz.open(pdf)
text='\n'.join(page.get_text() for page in pdf_doc)
assert len(pdf_doc)==2
assert shape in text and 'Second slide control' in text
assert 'Original shape text' not in text
print(json.dumps({
 'oracle':'Python zipfile + ElementTree + PyMuPDF',
 'onlySelectedPayloadsChanged':True,'changedPayloads':changes,'sourcePartSetRetained':True,
 'allUnselectedPayloadsByteIdentical':True,'unknownExtensionsRetainedByPptkit':True,'orphanRetainedByPptkit':True,
 'candidateSlideText':candidate_slides,'candidateNotesText':candidate_notes,
 'nativeSlideText':native_slides,'nativeNotesText':native_notes,
 'pdfPages':len(pdf_doc),'pdfText':text,
 'nativeUnknownExtensionRetained':any(b'opaque retained owner source' in value for value in c.values()),
 'nativeOrphanRetained':'orphan.bin' in c,
 'nativeAddedParts':sorted(set(c)-set(b)),'nativeRemovedParts':sorted(set(b)-set(c))
},ensure_ascii=False))
`;
const independent = JSON.parse((await execute("python3", ["-c", python, sourceFile, editedFile, nativeFile, pdfFile, expectedShape, expectedNotes], options)).stdout);
const nativeBytes = await readFile(nativeFile), comparison = await compareOpcPackages(edited, nativeBytes);
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const evidence = {
  schemaVersion: 1,
  scope: "One shape-text and one direct notes-body run edit in a mixed two-slide deck, independent payload/XML checks, LibreOffice headless producer resave and slide PDF text inspection. Not interactive editing, repair-dialog verification, PowerPoint, notes rendering, complete appearance/native fidelity or full IMP-09 completion.",
  environment: { platform: platform(), release: release(), arch: arch(), application },
  directory, sourceSha256: hash(source), editedSha256: hash(edited), nativeSha256: hash(nativeBytes),
  saveReport, independent, nativeComparison: { equal: comparison.equal, differenceCount: comparison.differenceCount, differences: comparison.differences },
  logs,
};
await writeFile("docs/testing/text-owner-save-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, ...independent }, null, 2));
