// Optional native producer/PDF experiment for the bounded text-runs-v1 profile.
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
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-text-edit-save.mjs [soffice-executable]");
const executable = process.argv[2] ?? "soffice", execute = promisify(execFile), options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/text-edit-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/text-edit-native/run-")), nativeDirectory = path.join(directory, "native"), pdfDirectory = path.join(directory, "pdf");
await mkdir(nativeDirectory); await mkdir(pdfDirectory);
const document = createPresentation({ theme: { fonts: { heading: "Liberation Sans", body: "Liberation Sans" } } }), slide = document.addSlide();
const textStyle = { fontFamily: "Liberation Sans", fontSize: 28, color: "123456" };
slide.addElement({ type: "text", box: { x: 30, y: 30, width: 700, height: 70 }, content: [{ runs: [{ text: "Source text before", style: textStyle }] }] });
slide.addElement({ type: "text", box: { x: 30, y: 130, width: 700, height: 70 }, content: [{ runs: [{ text: "Unchanged control text", style: textStyle }] }] });
slide.setNotes("Unchanged speaker notes");
const parts = readGeneratedZip((await generatePptx(document)).bytes), slidePart = "ppt/slides/slide1.xml";
const extension = '<p:ext uri="urn:text-save-evidence"><x:opaque xmlns:x="urn:evidence" ref="2">opaque retained source</x:opaque></p:ext>';
const xml = parts.get(slidePart).toString();
parts.set(slidePart, Buffer.from(xml.includes("</p:extLst>") ? xml.replace("</p:extLst>", extension + "</p:extLst>") : xml.replace("</p:sld>", `<p:extLst>${extension}</p:extLst></p:sld>`)));
parts.set("orphan.bin", Buffer.from([1, 8, 27, 64]));
parts.set("[Content_Types].xml", Buffer.from(parts.get("[Content_Types].xml").toString().replace("</Types>", '<Override PartName="/orphan.bin" ContentType="application/octet-stream"/></Types>')));
const source = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
const session = await importPptx(source, { projection: "content", editSaveProfile: "text-runs-v1" });
const expectedText = "Edited preserved text 123 <&>";
let edited, saveReport;
try {
  const owner = session.document.slides[0], original = owner.elements.find(element => element.type === "text"), index = owner.elements.indexOf(original);
  const modified = structuredClone(original); modified.content[0].runs[0].text = expectedText;
  owner.removeElement(original.id); owner.insertElement(index, modified);
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
import sys,json,zipfile,hashlib,xml.etree.ElementTree as ET,fitz
source,edited,native,pdf,expected=sys.argv[1:]
def package(file):
  with zipfile.ZipFile(file) as archive:
    return {name:archive.read(name) for name in archive.namelist() if not name.endswith('/')}
a,b,c=map(package,[source,edited,native])
assert set(a)==set(b)
changes=sorted(name for name in a if a[name]!=b[name])
assert changes==['ppt/slides/slide1.xml']
assert a['orphan.bin']==b['orphan.bin']
assert b'opaque retained source' in b['ppt/slides/slide1.xml']
ns='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def texts(parts):
  return [''.join(node.itertext()) for name,data in parts.items() if name.startswith('ppt/slides/') and name.endswith('.xml') for node in ET.fromstring(data).iter(ns+'t')]
edited_texts,native_texts=texts(b),texts(c)
assert expected in edited_texts
assert expected in native_texts
assert 'Unchanged control text' in native_texts
pdf_doc=fitz.open(pdf)
text='\n'.join(page.get_text() for page in pdf_doc)
assert expected in text
assert 'Unchanged control text' in text
assert 'Source text before' not in text
print(json.dumps({
 'oracle':'Python zipfile + ElementTree + PyMuPDF',
 'onlySelectedPayloadChanged':True,'changedPayloads':changes,
 'sourcePartSetRetained':True,'unknownExtensionRetainedByPptkit':True,'orphanRetainedByPptkit':True,
 'candidateText':edited_texts,'nativeProducerText':native_texts,
 'pdfPages':len(pdf_doc),'pdfText':text,
 'nativeUnknownExtensionRetained':any(b'opaque retained source' in value for value in c.values()),
 'nativeOrphanRetained':'orphan.bin' in c,
 'nativeAddedParts':sorted(set(c)-set(b)),'nativeRemovedParts':sorted(set(b)-set(c))
},ensure_ascii=False))
`;
const independent = JSON.parse((await execute("python3", ["-c", python, sourceFile, editedFile, nativeFile, pdfFile, expectedText], options)).stdout);
const nativeBytes = await readFile(nativeFile), comparison = await compareOpcPackages(edited, nativeBytes);
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const evidence = {
  schemaVersion: 1,
  scope: "One named direct text-run preservation edit, independent payload/XML checks, LibreOffice headless producer resave and PDF text inspection. Not interactive editing, repair-dialog verification, PowerPoint, complete appearance/native fidelity or full IMP-09 completion.",
  environment: { platform: platform(), release: release(), arch: arch(), application },
  directory, sourceSha256: hash(source), editedSha256: hash(edited), nativeSha256: hash(nativeBytes),
  saveReport, independent, nativeComparison: { equal: comparison.equal, differenceCount: comparison.differenceCount, differences: comparison.differences },
  logs,
};
await writeFile("docs/testing/text-edit-save-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, ...independent }, null, 2));
