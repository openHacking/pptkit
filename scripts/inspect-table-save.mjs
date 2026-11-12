// Optional native evidence for fixed-topology text-runs-v3 table character saves.
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
import { createTableSaveFixture } from "../tests/fixtures/table-save.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { png } from "../packages/pptx-parser/test/image-fixture.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-table-save.mjs [soffice-executable]");
const executable = process.argv[2] ?? "soffice", execute = promisify(execFile), options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/table-save-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/table-save-native/run-")), nativeDirectory = path.join(directory, "native"), pdfDirectory = path.join(directory, "pdf");
await mkdir(nativeDirectory); await mkdir(pdfDirectory);
const parts = readGeneratedZip((await generatePptx(createTableSaveFixture(createPresentation, png))).bytes), slidePart = "ppt/slides/slide1.xml";
const extension = '<p:ext uri="urn:table-save-evidence"><x:opaque xmlns:x="urn:future">opaque table source</x:opaque></p:ext>', xml = parts.get(slidePart).toString();
parts.set(slidePart, Buffer.from((xml.includes("</p:extLst>") ? xml.replace("</p:extLst>", extension + "</p:extLst>") : xml.replace("</p:sld>", `<p:extLst>${extension}</p:extLst></p:sld>`)).replace("<a:tcPr", '<a:tcPr xmlns:x="urn:future" x:opaque="cell-source"')));
parts.set("orphan.bin", Buffer.from([2, 4, 8, 16]));
parts.set("[Content_Types].xml", Buffer.from(parts.get("[Content_Types].xml").toString().replace("</Types>", '<Override PartName="/orphan.bin" ContentType="application/octet-stream"/></Types>')));
const source = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
const expectedAnchor = "Anchor 123", expectedCell = "Cell 456";
const session = await importPptx(source, { projection: "content", editSaveProfile: "text-runs-v3" });
let edited, saveReport;
try {
  const slide = session.document.slides[0], original = slide.elements.find(e => e.type === "table" && e.columns.length === 3), index = slide.elements.indexOf(original), table = structuredClone(original);
  table.rows[0].cells[0].content[0].runs[0].text = expectedAnchor;
  table.rows[1].cells[0].content[0].runs[0].text = expectedCell;
  slide.removeElement(original.id); slide.insertElement(index, table);
  const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report));
  const saved = await session.save(plan.plan); edited = saved.bytes; saveReport = saved.report;
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
source,edited,native,pdf,anchor,cell=sys.argv[1:]
def package(file):
  with zipfile.ZipFile(file) as archive:
    return {name:archive.read(name) for name in archive.namelist() if not name.endswith('/')}
a,b,c=map(package,[source,edited,native])
assert set(a)==set(b)
changes=sorted(name for name in a if a[name]!=b[name])
assert changes==['ppt/slides/slide1.xml']
assert a['orphan.bin']==b['orphan.bin']
assert b'opaque table source' in b[changes[0]] and b'x:opaque="cell-source"' in b[changes[0]]
ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def tables(parts):
  result=[]
  for name,data in parts.items():
    if name.startswith('ppt/slides/') and name.endswith('.xml'):
      for tbl in ET.fromstring(data).findall('.//a:tbl',ns):
        result.append({'part':name,'columns':[node.attrib for node in tbl.findall('a:tblGrid/a:gridCol',ns)],
          'rows':[{'attributes':row.attrib,'cells':[{'attributes':tc.attrib,'text':''.join(node.text or '' for node in tc.findall('.//a:t',ns))} for tc in row.findall('a:tc',ns)]} for row in tbl.findall('a:tr',ns)]})
  return result
source_tables,candidate_tables,native_tables=map(tables,[a,b,c])
assert len(source_tables)==len(candidate_tables)==2
assert candidate_tables[1]['rows'][0]['cells'][0]['text']==anchor
assert candidate_tables[1]['rows'][1]['cells'][2]['text']==cell
def structure(items):
  return [{'part':t['part'],'columns':t['columns'],'rows':[{'attributes':r['attributes'],'cells':[tc['attributes'] for tc in r['cells']]} for r in t['rows']]} for t in items]
def merges(items):
  return [[[(int(tc['attributes'].get('gridSpan','1')),int(tc['attributes'].get('rowSpan','1')),tc['attributes'].get('hMerge','0') in ['1','true'],tc['attributes'].get('vMerge','0') in ['1','true']) for tc in row['cells']] for row in t['rows']] for t in items]
assert structure(source_tables)==structure(candidate_tables)
def signature(node):
  return (node.tag,sorted(node.attrib.items()),None if node.tag=='{'+ns['a']+'}t' else node.text,node.tail,[signature(child) for child in node])
assert signature(ET.fromstring(a[changes[0]]))==signature(ET.fromstring(b[changes[0]]))
source_atoms=ET.fromstring(a[changes[0]]).findall('.//a:t',ns)
candidate_atoms=ET.fromstring(b[changes[0]]).findall('.//a:t',ns)
assert len(source_atoms)==len(candidate_atoms)
assert [(x.text,y.text) for x,y in zip(source_atoms,candidate_atoms) if x.text!=y.text]==[('Merged anchor control',anchor),('Bottom right source',cell)]
native_cells=[tc['text'] for t in native_tables for row in t['rows'] for tc in row['cells']]
assert anchor in native_cells and cell in native_cells
for control in ['Cell control A','Cell control B','Top right control','Last cell control']:
  assert control in native_cells
pdf_doc=fitz.open(pdf)
text='\n'.join(page.get_text() for page in pdf_doc)
assert len(pdf_doc)==2 and anchor in text and cell in text and 'Second slide control' in text
print(json.dumps({
 'oracle':'Python zipfile + ElementTree + PyMuPDF','changedPayloads':changes,
 'sourcePartSetRetained':True,'allUnselectedPayloadsByteIdentical':True,
 'candidateNonTextXmlStructureEqual':True,'candidateTableGridsAndMergesEqual':True,
 'sourceTables':source_tables,'candidateTables':candidate_tables,'nativeTables':native_tables,
 'nativeTableGridsAndMergesEqual':structure(candidate_tables)==structure(native_tables),
 'nativeMergeSemanticsEqual':merges(candidate_tables)==merges(native_tables),
 'nativeTableCount':len(native_tables),'pdfPages':len(pdf_doc),'pdfText':text,
 'unknownExtensionRetainedByPptkit':True,'orphanRetainedByPptkit':True,
 'nativeUnknownExtensionRetained':any(b'opaque table source' in value for value in c.values()),
 'nativeOrphanRetained':'orphan.bin' in c,
 'nativeAddedParts':sorted(set(c)-set(b)),'nativeRemovedParts':sorted(set(b)-set(c))
},ensure_ascii=False))
`;
const independent = JSON.parse((await execute("python3", ["-c", python, sourceFile, editedFile, nativeFile, pdfFile, expectedAnchor, expectedCell], options)).stdout);
const nativeBytes = await readFile(nativeFile), comparison = await compareOpcPackages(edited, nativeBytes);
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const evidence = {
  schemaVersion: 1,
  scope: "Two existing merged-table anchor run edits in a two-slide mixed deck, independent physical-grid/non-text XML and payload checks, LibreOffice headless producer resave and slide PDF text inspection. Not interactive table editing, repair-dialog verification, PowerPoint, complete layout/native fidelity or full IMP-09.2.5 completion.",
  environment: { platform: platform(), release: release(), arch: arch(), application },
  directory, sourceSha256: hash(source), editedSha256: hash(edited), nativeSha256: hash(nativeBytes),
  saveReport, independent, nativeComparison: { equal: comparison.equal, differenceCount: comparison.differenceCount, differences: comparison.differences }, logs,
};
await writeFile("docs/testing/table-save-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, changedPayloads: independent.changedPayloads, nativeTableCount: independent.nativeTableCount, nativeTableGridsAndMergesEqual: independent.nativeTableGridsAndMergesEqual, nativeMergeSemanticsEqual: independent.nativeMergeSemanticsEqual, pdfPages: independent.pdfPages, nativeUnknownExtensionRetained: independent.nativeUnknownExtensionRetained, nativeOrphanRetained: independent.nativeOrphanRetained }, null, 2));
