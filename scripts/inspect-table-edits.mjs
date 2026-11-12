// Optional native structural evidence for fresh authored table edits, not source-preservation writes.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation, editTable, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { tableEditDeck, tableEditBatches } from "../tests/fixtures/table-edits.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-table-edits.mjs [soffice-executable]");
const execute = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ["--version"], options)).stdout.trim(), hash = bytes => createHash("sha256").update(bytes).digest("hex");
await mkdir(".tmp/table-edits-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/table-edits-native/run-")), nativeDirectory = path.join(directory, "native"), pdfDirectory = path.join(directory, "pdf");
await mkdir(nativeDirectory); await mkdir(pdfDirectory);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`, { document, slide } = tableEditDeck(createPresentation), rounds = [];
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E,copy
file,native,expectedFile=sys.argv[1:];expected=json.load(open(expectedFile));A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def decode(root,strict):
  table=root.find('.//'+A+'tbl');rows=table.findall(A+'tr');columns=[int(n.get('w'))/12700 for n in table.find(A+'tblGrid')]
  slots=[[None]*len(columns) for _ in rows];out=[]
  for r,row in enumerate(rows):
    physical=row.findall(A+'tc');assert len(physical)==len(columns)
    anchors=[]
    for c,node in enumerate(physical):
      old=slots[r][c]
      hm=node.get('hMerge') in ('1','true');vm=node.get('vMerge') in ('1','true')
      if old is not None:
        if strict:
          assert hm==(c>old[1]) and vm==(r>old[0])
          assert not any(n.text for n in node.iter(A+'t'))
        continue
      assert not hm and not vm
      rs=int(node.get('rowSpan','1'));cs=int(node.get('gridSpan','1'))
      assert rs>=1 and cs>=1 and r+rs<=len(rows) and c+cs<=len(columns)
      for y in range(r,r+rs):
        for x in range(c,c+cs):assert slots[y][x] is None;slots[y][x]=(r,c)
      paragraphs=node.find(A+'txBody').findall(A+'p')
      text='\n'.join(''.join(n.text or '' for n in p.iter(A+'t')) for p in paragraphs)
      anchors.append({'rowSpan':rs,'colSpan':cs,'text':text})
    assert all(n is not None for n in slots[r])
    out.append({'height':int(row.get('h'))/12700,'cells':anchors})
  return {'columns':columns,'rows':out}
def inspect(file,strict):
  with zipfile.ZipFile(file) as z:raw=z.read('ppt/slides/slide1.xml')
  return decode(E.fromstring(raw),strict),raw
source,raw=inspect(file,True);assert source==expected
# Deliberately moving a continuation must be visible to the physical-slot oracle.
damaged=E.fromstring(raw);continuation=next((n for n in damaged.iter(A+'tc') if n.get('vMerge')=='1'),None)
mutatedContinuationRejected=None
if continuation is not None:
  continuation.set('vMerge','0')
  try:decode(damaged,True)
  except AssertionError:mutatedContinuationRejected=True
  else:mutatedContinuationRejected=False
  assert mutatedContinuationRejected
nativeModel,_=inspect(native,False)
sameContents=len(nativeModel['rows'])==len(expected['rows']) and len(nativeModel['columns'])==len(expected['columns']) and all(a['cells']==b['cells'] for a,b in zip(nativeModel['rows'],expected['rows']))
widthError=max([abs(a-b) for a,b in zip(nativeModel['columns'],expected['columns'])] or [0]) if len(nativeModel['columns'])==len(expected['columns']) else None
heightError=max([abs(a['height']-b['height']) for a,b in zip(nativeModel['rows'],expected['rows'])] or [0]) if len(nativeModel['rows'])==len(expected['rows']) else None
print(json.dumps({'verified':True,'oracle':'Python zipfile + ElementTree; physical grid, all merge continuations, dimensions and paragraph content','mutatedContinuationRejected':mutatedContinuationRejected,'sourceModel':source,'native':{'status':'pass' if sameContents and widthError is not None and heightError is not None and max(widthError,heightError)<=0.1 else 'partial','sameGridAndParagraphContent':sameContents,'maxColumnPointError':widthError,'maxRowPointError':heightError,'model':nativeModel}},ensure_ascii=False))
`;
for (let index = 0; index < tableEditBatches.length; index++) {
  const result = editTable(slide.elements[0], tableEditBatches[index], { frame: "fit-grid" }); slide.replaceElement("table", result.table);
  const table = normalizePresentation(document).slides[0].elements[0], expected = { columns: table.columns, rows: table.rows.map(row => ({ height: row.height, cells: row.cells.map(cell => ({ rowSpan: cell.rowSpan, colSpan: cell.colSpan, text: cell.content.map(p => p.runs.map(r => r.text).join("")).join("\n") })) })) };
  const generated = await generatePptx(document), filename = path.join(directory, `round-${index + 1}.pptx`), expectedFile = path.join(directory, `round-${index + 1}.expected.json`);
  await writeFile(filename, generated.bytes); await writeFile(expectedFile, JSON.stringify(expected));
  const nativeLog = await execute(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, filename], options);
  const nativeFile = path.join(nativeDirectory, path.basename(filename)), independent = JSON.parse((await execute("python3", ["-c", python, filename, nativeFile, expectedFile], options)).stdout);
  rounds.push({ round: index + 1, operations: tableEditBatches[index], cellMapping: result.cellMapping, sourceSha256: hash(generated.bytes), nativeSha256: hash(await readFile(nativeFile)), independent, nativeLog });
}
const nativeFinal = path.join(nativeDirectory, "round-3.pptx"), pdfLog = await execute(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdfDirectory, nativeFinal], options);
const pdfFile = path.join(pdfDirectory, "round-3.pdf"), pdf = await readFile(pdfFile), pdfObservation = JSON.parse((await execute("python3", ["-c", "import sys,json,fitz; d=fitz.open(sys.argv[1]); print(json.dumps({'pages':len(d),'text':[p.get_text() for p in d]}))", pdfFile], options)).stdout);
assert.equal(pdfObservation.pages, 2); assert.ok(pdfObservation.text.some(text => text.includes("Unchanged table control")));
const tablePdfText = pdfObservation.text[0].replace(/\s+/g, "");
for (const label of ["Anchor A", "Control C", "Control E", "Control F", "Inserted row", "Inserted column", "Tail"]) assert.ok(tablePdfText.includes(label.replace(/\s+/g, "")), label);
assert.ok(!tablePdfText.includes("DeleteB") && !tablePdfText.includes("DeleteD"));
assert.equal((tablePdfText.match(/Tail/g) ?? []).length, 3);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory),
  scope: "Three atomic Core edit batches followed by fresh native-table PPTX export, independent physical-grid ZIP/XML verification, separate LibreOffice headless resaves and saved-output PDF reopening. Not source structural preservation, interactive table editing or full style fidelity.",
  rounds, pdf: { sha256: hash(pdf), observation: pdfObservation, log: pdfLog }, interactiveEditing: "not-performed", sourceStructuralSave: "not-performed" };
await writeFile("docs/testing/table-edit-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, rounds: rounds.length, nativeStatuses: rounds.map(r => r.independent.native.status), pdfPages: pdfObservation.pages }));
