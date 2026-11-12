// Optional native research; saved nodes are evidence, not playback certification.
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
import { animationSpecimens, nativeTimingSpecimen } from "./research/native-timing-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-animation-mapping.mjs [soffice-executable]");
const run = promisify(execFile), options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 }, executable = process.argv[2] ?? "soffice";
const application = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/animation-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/animation-native/run-")), native = path.join(directory,"native"), pdf = path.join(directory,"pdf");
await mkdir(native); await mkdir(pdf);
const deck = createPresentation();
for (const spec of animationSpecimens) deck.addSlide().addElement({ type: "shape", shape: "rect", name: spec.name, text: { content: spec.name }, box: { x: 100, y: 100, width: 240, height: 80 } });
const parts = readGeneratedZip((await generatePptx(deck)).bytes);
for (const [i,spec] of animationSpecimens.entries()) {
  const part = `ppt/slides/slide${i+1}.xml`, xml = parts.get(part).toString();
  // This fixture owns one shape and its source drawing identity, not a general search/remap API.
  assert.match(xml, /<p:cNvPr id="2"/);
  parts.set(part, Buffer.from(xml.replace("</p:sld>", nativeTimingSpecimen(spec.name,2)+"</p:sld>")));
}
const bytes = zip([...parts].map(([name,body]) => ({ name,body }))).bytes, source = path.join(directory,"source.pptx");
await writeFile(source,bytes);
const session = await importPptx(bytes,{ projection:"content", editSaveProfile:"text-runs-v3" });
try { const p = await session.prepareSave({ mode:"no-op" }); assert.equal(p.status,"ready"); assert.deepEqual((await session.save(p.plan)).bytes, new Uint8Array(bytes)); } finally { session.dispose(); }
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory,"profile")).href}`;
const logs = {
  save: await run(executable,[profile,"--headless","--convert-to","pptx:Impress Office Open XML","--outdir",native,source],options),
  reopen: await run(executable,[profile,"--headless","--convert-to","pdf:impress_pdf_Export","--outdir",pdf,path.join(native,"source.pptx")],options),
};
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E,fitz
P='http://schemas.openxmlformats.org/presentationml/2006/main';ns={'p':P}
expected=json.loads(sys.argv[3])
def inspect(file):
 with zipfile.ZipFile(file) as z:
  rows=[]
  for i in range(1,5):
   root=E.fromstring(z.read('ppt/slides/slide%d.xml'%i))
   shapes={n.get('id'):n.get('name') for n in root.findall('.//p:sp/p:nvSpPr/p:cNvPr',ns)}
   behaviors=[]
   for n in root.iter():
    if n.tag not in ['{'+P+'}'+k for k in ['animEffect','animRot','animMotion']]:continue
    targets=[t.get('spid') for t in n.findall('.//p:spTgt',ns)]
    assert all(t in shapes for t in targets)
    behaviors.append({'tag':n.tag.split('}')[-1],'attributes':dict(n.attrib),'targetNames':[shapes[t] for t in targets],'durations':[t.get('dur') for t in n.findall('.//p:cTn',ns)]})
   ids=[n.get('id') for n in root.findall('.//p:timing//p:cTn',ns)]
   assert len(ids)==len(set(ids))
   rows.append({'behaviors':behaviors,'timingIds':ids,'conditions':[dict(n.attrib) for n in root.findall('.//p:timing//p:cond',ns)]})
  return rows
a,b=inspect(sys.argv[1]),inspect(sys.argv[2])
for row,spec in zip(a,expected):
 assert len(row['behaviors'])==1
 n=row['behaviors'][0]
 assert n['tag']==spec['tag'] and n['attributes']==spec['attributes'] and n['targetNames']==[spec['name']] and n['durations']==['500']
doc=fitz.open(sys.argv[4]);assert len(doc)==4
print(json.dumps({'oracle':'Python zipfile + ElementTree + PyMuPDF','source':a,'native':b,'reopenedPdfPages':len(doc),'rawTimingInventoryEqual':a==b}))
`;
const independent = JSON.parse((await run("python3",["-c",python,source,path.join(native,"source.pptx"),JSON.stringify(animationSpecimens),path.join(pdf,"source.pdf")],options)).stdout);
const hash = b => createHash("sha256").update(b).digest("hex");
const evidence = { schemaVersion:1, scope:"Hand-authored entrance/emphasis/exit/path timing specimens, independent source/native inventories, headless save and saved-output PDF reopening. Not schema certification, playback, editability, repair-dialog absence, visual fidelity or public/native animation support.", environment:{platform:platform(),release:release(),arch:arch(),application},directory,specimens:animationSpecimens,sourceSha256:hash(bytes),nativeSha256:hash(await readFile(path.join(native,"source.pptx"))),pdfSha256:hash(await readFile(path.join(pdf,"source.pdf"))),independent,logs };
await writeFile("docs/testing/animation-mapping-evidence.json",JSON.stringify(evidence,null,2)+"\n");
console.log(JSON.stringify({ directory,application,independent },null,2));
