// Optional headless producer observations. Never certify native editing or UI.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { inspectPptx } from "../packages/pptx-parser/dist/index.js";
import { pdfTextEvidence } from "./research/pdf-text-evidence.mjs";

const prepared = JSON.parse(await readFile(new URL("../docs/testing/smartart-native-candidate-evidence.json", import.meta.url), "utf8"));
const sourceDirectory = path.dirname(path.resolve(process.argv[2] ?? path.join(prepared.directory, "manifest.json")));
const manifest = JSON.parse(await readFile(path.join(sourceDirectory, "manifest.json"), "utf8"));
const executable = process.argv[3] && !process.argv[3].startsWith("--") ? process.argv[3] : "soffice", run = promisify(execFile);
const outputIndex = process.argv.indexOf("--output");
if (outputIndex >= 0 && !process.argv[outputIndex + 1]) throw new TypeError("--output requires a filename.");
const commandOptions = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await run(executable, ["--version"], commandOptions)).stdout.trim();
const directory = await mkdtemp(path.join(sourceDirectory, "native-run-")), native = path.join(directory, "native"), pdf = path.join(directory, "pdf");
await mkdir(native); await mkdir(pdf);
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const cases = [], logs = [];
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
for (const sample of manifest.records.filter(row => row.case === "standard")) {
  assert.equal(path.basename(sample.file), sample.file);
  const input = path.join(sourceDirectory, sample.file), original = await readFile(input); assert.equal(hash(original), sample.sha256);
  const log = { file: sample.file };
  try {
    log.save = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", native, input], commandOptions);
    const output = path.join(native, sample.file), saved = await readFile(output);
    log.reopen = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdf, output], commandOptions);
    const rendered = path.join(pdf, sample.file.replace(/\.pptx$/, ".pdf")); await readFile(rendered);
    let parser;
    try {
      const inspection = await inspectPptx(saved);
      parser = { status: "inspected", slides: inspection.slides.length, diagrams: inspection.slides.flatMap(s => s.diagrams ?? []).length,
        dataModels: inspection.diagramInventory?.dataModels?.map(model => ({ part: model.part, status: model.status, points: model.points.length, connections: model.connections.length })) ?? [] };
    } catch (error) { parser = { status: "rejected", code: error.code, message: error.message }; }
    assert.equal(hash(await readFile(input)), sample.sha256);
    cases.push({ ...sample, input, output, rendered, nativeSha256: hash(saved), pdfSha256: hash(await readFile(rendered)), parser });
  } catch (error) {
    log.failure = { message: error.message, code: error.code, stdout: error.stdout, stderr: error.stderr };
    cases.push({ file: sample.file, family: sample.family, cache: sample.cache, status: "native-operation-failed", failure: log.failure });
  }
  logs.push(log);
  console.log(JSON.stringify({ file: sample.file, status: cases.at(-1).status ?? "saved-and-reopened" }));
}
await writeFile(path.join(directory, "operations.json"), JSON.stringify(cases, null, 2) + "\n");
const oracle = String.raw`
import json,sys,zipfile,xml.etree.ElementTree as E,hashlib,fitz
from pathlib import Path
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';P='{http://schemas.openxmlformats.org/presentationml/2006/main}';S='{http://schemas.microsoft.com/office/drawing/2008/diagram}'
def inventory(file):
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None
  parts=sorted(n for n in z.namelist() if n.startswith('ppt/diagrams/') and n.endswith('.xml'))
  models=[];cache_shapes=0
  for name in parts:
   root=E.fromstring(z.read(name))
   if root.tag==D+'dataModel':
    points=root.find(D+'ptLst');connections=root.find(D+'cxnLst')
    texts={p.get('modelId'):'\n'.join(''.join(t.text or '' for t in para.iter(A+'t')) for para in p.findall('./'+D+'t/'+A+'p')) for p in points if p.get('type','node') in ['node','asst']}
    models.append(dict(part=name,points=len(points),connections=len(connections) if connections is not None else 0,texts=texts,assistantIds=[p.get('modelId') for p in points if p.get('type')=='asst']))
   if root.tag==S+'drawing':cache_shapes+=len(list(root.iter(S+'sp')))
  slide=E.fromstring(z.read('ppt/slides/slide1.xml'))
  frames=[n.get('uri') for n in slide.iter(A+'graphicData')]
  return dict(diagramParts=parts,dataModels=models,cacheShapes=cache_shapes,graphicUris=frames,ordinaryShapes=len(list(slide.iter(P+'sp'))))
output=[]
for row in json.loads(Path(sys.argv[1]).read_text()):
 if row.get('status')=='native-operation-failed':output.append(row);continue
 assert hashlib.sha256(Path(row['input']).read_bytes()).hexdigest()==row['sha256']
 assert hashlib.sha256(Path(row['output']).read_bytes()).hexdigest()==row['nativeSha256']
 source=inventory(row['input']);saved=inventory(row['output']);pdf=fitz.open(row['rendered']);text='\n'.join(page.get_text() for page in pdf)
 markers=[paragraph for node in row['authored']['nodes'] for paragraph in node['paragraphs'] if paragraph.strip()]
 normalized=' '.join(text.split());missing=[marker for marker in markers if ' '.join(marker.split()) not in normalized]
 models_retained=len(saved['dataModels'])==len(source['dataModels']) and saved['dataModels']==source['dataModels']
 output.append(dict(file=row['file'],family=row['family'],cache=row['cache'],inputSha256=row['sha256'],nativeSha256=row['nativeSha256'],pdfSha256=row['pdfSha256'],status='saved-and-reopened',source=source,native=saved,parser=row['parser'],pdfPages=len(pdf),pdfExtractedText=text,textMarkers=markers,missingTextMarkers=missing,allTextMarkersInPdf=not missing,controlTextInPdf='PPTKit native diagram control' in normalized,exactDataInventoryRetained=models_retained,diagramGraphicRetained='http://schemas.openxmlformats.org/drawingml/2006/diagram' in saved['graphicUris'],independentZipXmlPdf=True))
print(json.dumps(dict(python=sys.version.split()[0],cases=output)))
`;
const independent = JSON.parse((await run("python3", ["-c", oracle, path.join(directory, "operations.json")], commandOptions)).stdout);
for (const result of independent.cases) if (result.status === "saved-and-reopened") {
  const text = pdfTextEvidence(result.pdfExtractedText, result.textMarkers);
  assert.deepEqual(text.missingTextMarkers, result.missingTextMarkers);
  Object.assign(result, text);
  delete result.textMarkers;
}
const evidence = { profile: "native-diagram-headless-producer-v1", application, directory, sourceDirectory,
  runtime: { node: process.version, python: independent.python }, cases: independent.cases, logs,
  limitations: ["Headless application save/PDF reopening only; no interactive editability, repair-dialog, playback or calibrated visual-fidelity claim.", "A word-spacing marker mismatch is not proof of absent node text: extraction can split words over lines. Whitespace-insensitive matches are separate lexical evidence, not legibility or geometry validation; repeated/overlapping markers do not prove distinct visible occurrences.", "Native source/model losses and unmatched PDF text are observations, not silently accepted compatibility passes.", "The development LibreOffice runtime does not establish stable LibreOffice or PowerPoint support.", "Third-party component payloads and resulting artifacts remain local; native family admission remains incomplete."] };
const evidenceFile = outputIndex >= 0 ? path.resolve(process.argv[outputIndex + 1]) : new URL(manifest.records.some(row => row.modelIdPolicy === "guid") ? "../docs/testing/smartart-native-guid-producer-evidence.json" : manifest.records.some(row => row.presentation === "layout") ? "../docs/testing/smartart-native-presentation-producer-evidence.json" : "../docs/testing/smartart-native-candidate-producer-evidence.json", import.meta.url);
await writeFile(evidenceFile, JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, cases: evidence.cases.map(c => ({ file: c.file, status: c.status, retained: c.diagramGraphicRetained, text: c.allTextMarkersInPdf })) }, null, 2));
