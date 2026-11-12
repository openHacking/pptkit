// Optional independent/native evidence; never opens a GUI or certifies playback.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { nativeTimingSpecimen } from "./research/native-timing-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-paragraph-target-save.mjs [soffice-executable]");
const execute = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ["--version"], options)).stdout.trim(), hash = bytes => createHash("sha256").update(bytes).digest("hex");
await mkdir(".tmp/paragraph-target-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/paragraph-target-native/run-")), nativeDirectory = path.join(directory, "native"), pdfDirectory = path.join(directory, "pdf");
await mkdir(nativeDirectory); await mkdir(pdfDirectory);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const deck = createPresentation();
deck.addSlide().addElement({ type: "text", box: { x: 100, y: 100, width: 400, height: 200 }, content: [
  { runs: [{ text: "Initial paragraph A" }, { text: " retained tail" }] },
  { runs: [{ text: "" }] }, { runs: [{ text: "Initial paragraph C" }] },
] });
deck.addSlide().addElement({ type: "text", box: { x: 100, y: 100, width: 400, height: 100 }, content: "Unchanged paragraph control" });
const parts = readGeneratedZip((await generatePptx(deck)).bytes), slidePart = "ppt/slides/slide1.xml";
assert.match(parts.get(slidePart).toString(), /<p:cNvPr id="2"/);
const timing = nativeTimingSpecimen("fade-in", 2).replace('<p:spTgt spid="2"/>', '<p:spTgt spid="2"><p:txEl><p:pRg st="0" end="2"/></p:txEl></p:spTgt>');
parts.set(slidePart, Buffer.from(parts.get(slidePart).toString().replace('</p:sld>', timing + '</p:sld>')));
let bytes = zip([...parts].map(([name, body]) => ({ name, body }))).bytes, previousFile = path.join(directory, "source.pptx");
await writeFile(previousFile, bytes);
const controlFile = path.join(nativeDirectory, "source.pptx");
const controlLog = await execute(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, previousFile], options);
const sourceHash = hash(bytes), rounds = [];
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E,copy
source,candidate,native,control,expectedA,expectedC=sys.argv[1:]
P='{http://schemas.openxmlformats.org/presentationml/2006/main}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def package(file):
  with zipfile.ZipFile(file) as z:return {n:z.read(n) for n in z.namelist() if not n.endswith('/')}
def owner(root):
  target=next(n for n in root.iter(P+'spTgt') if n.find(P+'txEl/'+P+'pRg') is not None)
  shape=next(n for n in root.iter(P+'sp') if n.find(P+'nvSpPr/'+P+'cNvPr').get('id')==target.get('spid'))
  return target,shape
def fingerprint(n):return (n.tag,tuple(sorted(n.attrib.items())),n.text,n.tail,tuple(fingerprint(c) for c in n))
def verify(a,b):
  assert set(a)==set(b)
  changed=sorted(n for n in a if a[n]!=b[n]);assert changed==['ppt/slides/slide1.xml']
  left,right=map(E.fromstring,[a[changed[0]],b[changed[0]]])
  target,shape=owner(right); oldTarget,oldShape=owner(left)
  assert fingerprint(left.find(P+'timing'))==fingerprint(right.find(P+'timing'))
  paragraphs=shape.find(P+'txBody').findall(A+'p');oldParagraphs=oldShape.find(P+'txBody').findall(A+'p')
  assert len(paragraphs)==len(oldParagraphs)==3
  assert target.find(P+'txEl/'+P+'pRg').attrib=={'st':'0','end':'2'}
  assert paragraphs[0].find('.//'+A+'t').text==expectedA
  assert paragraphs[2].find('.//'+A+'t').text==expectedC
  paragraphs[0].find('.//'+A+'t').text=oldParagraphs[0].find('.//'+A+'t').text
  paragraphs[2].find('.//'+A+'t').text=oldParagraphs[2].find('.//'+A+'t').text
  assert fingerprint(left)==fingerprint(right)
  return changed
a,b,c=map(package,[source,candidate,native]);changed=verify(a,b)
def corrupt(xmlTransform):
  damaged=dict(b);r=E.fromstring(damaged['ppt/slides/slide1.xml']);xmlTransform(r);damaged['ppt/slides/slide1.xml']=E.tostring(r)
  try:verify(a,damaged)
  except (AssertionError,StopIteration):return True
  return False
assert corrupt(lambda r:r.find('.//'+P+'pRg').set('end','1'))
assert corrupt(lambda r:r.find('.//'+P+'cNvPr').set('name','Unapproved mutation'))
def nativeRows(parts):
  rows=[]
  for name,data in parts.items():
    if not name.startswith('ppt/slides/') or not name.endswith('.xml'):continue
    root=E.fromstring(data)
    for t in root.iter(P+'spTgt'):
      ranges=[dict(r.attrib) for r in t.iter(P+'pRg')]
      if not ranges:continue
      shapes=[n for n in root.iter(P+'sp') if n.find(P+'nvSpPr/'+P+'cNvPr').get('id')==t.get('spid')]
      texts=[]
      if len(shapes)==1:
        body=shapes[0].find(P+'txBody')
        if body is not None:texts=[''.join(n.itertext()) for n in body.findall(A+'p')]
      rows.append({'part':name,'spid':t.get('spid'),'ranges':ranges,'paragraphText':texts})
  return rows
rows,controlRows=nativeRows(c),nativeRows(package(control))
expected=[expectedA+' retained tail','',expectedC]
matched=any(row['ranges']==[{'st':'0','end':'2'}] and row['paragraphText']==expected for row in rows)
controlRewrite=bool(controlRows) and all(row['ranges']!=[{'st':'0','end':'2'}] for row in controlRows)
print(json.dumps({'verified':True,'oracle':'Python zipfile + ElementTree; exact nonselected payloads, full XML fingerprint except two approved atoms, timing/target/paragraph identity','changedPayloads':changed,'mutatedRangeRejected':True,'unapprovedAttributeRejected':True,'native':{'status':'pass' if matched else 'partial','rangeAndParagraphsRetained':matched,'rows':rows,'uneditedControlRows':controlRows,'rangeRewriteAlsoPresentInUneditedControl':controlRewrite}},ensure_ascii=False))
`;
for (let round = 1; round <= 3; round++) {
  const expectedA = `Round ${round} paragraph A`, expectedC = `Round ${round} paragraph C`, session = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v1", textTargetSaveProfile: "paragraph-stable-v1" });
  const beforeHash = hash(bytes);
  try {
    const slide = session.document.slides[0], element = structuredClone(slide.elements[0]);
    element.content[0].runs[0].text = expectedA; element.content[2].runs[0].text = expectedC; slide.replaceElement(element.id, element);
    const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready", JSON.stringify(prepared.report)); bytes = (await session.save(prepared.plan)).bytes;
  } finally { session.dispose(); }
  const filename = path.join(directory, `round-${round}.pptx`); await writeFile(filename, bytes);
  const nativeLog = await execute(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, filename], options);
  const nativeFile = path.join(nativeDirectory, path.basename(filename)), independent = JSON.parse((await execute("python3", ["-c", python, previousFile, filename, nativeFile, controlFile, expectedA, expectedC], options)).stdout);
  rounds.push({ round, sourceSha256: beforeHash, candidateSha256: hash(bytes), nativeSha256: hash(await readFile(nativeFile)), independent, nativeLog }); previousFile = filename;
}
const nativeFinal = path.join(nativeDirectory, "round-3.pptx"), pdfLog = await execute(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdfDirectory, nativeFinal], options);
const pdfFile = path.join(pdfDirectory, "round-3.pdf"), pdf = await readFile(pdfFile);
const pdfObservation = JSON.parse((await execute("python3", ["-c", "import sys,json,fitz; d=fitz.open(sys.argv[1]); print(json.dumps({'pages':len(d),'text':[p.get_text() for p in d]}))", pdfFile], options)).stdout);
assert.equal(pdfObservation.pages, 2); assert.ok(pdfObservation.text.some(text => text.includes("Unchanged paragraph control")));
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory), sourceSha256: sourceHash,
  scope: "Explicit paragraph-stable-v1 scalar source saves; three source rounds with independent exact ZIP/XML checks, separate LibreOffice headless range/paragraph observations and saved-output PDF reopening. No GUI, playback, effect editing or producer-wide admission.",
  control: { sha256: hash(await readFile(controlFile)), log: controlLog }, rounds, pdf: { sha256: hash(pdf), observation: pdfObservation, log: pdfLog }, nativePlayback: "not-performed", interactiveEditing: "not-performed" };
await writeFile("docs/testing/paragraph-target-save-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, independentRounds: rounds.length, nativeRangeStatuses: rounds.map(r => r.independent.native.status), pdfPages: pdfObservation.pages }));
