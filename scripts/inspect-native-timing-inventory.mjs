// Optional independent source oracle; does not open an application or observe playback.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { inspectPptxTiming } from "../packages/pptx-parser/dist/index.js";
if (process.argv.length < 3 || process.argv.length > 4) throw Error("Usage: node scripts/inspect-native-timing-inventory.mjs source.pptx [native.pptx]");
const run = promisify(execFile), sha = bytes => createHash("sha256").update(bytes).digest("hex");
await mkdir(".tmp/native-timing-inventory", { recursive: true }); const directory = await mkdtemp(path.resolve(".tmp/native-timing-inventory/run-"));
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E
report=json.load(open(sys.argv[2])); count=0
P=['http://schemas.openxmlformats.org/presentationml/2006/main','http://purl.oclc.org/ooxml/presentationml/main']
def expanded(name):
 if name.startswith('{'): return name[1:].split('}',1)
 return ['',name]
with zipfile.ZipFile(sys.argv[1]) as z:
 assert z.testzip() is None
 for slide in report['slides']:
  root=E.fromstring(z.read(slide['part'].lstrip('/'))); all=list(root.iter()); parents={c:n for n in all for c in n}
  def path(n):
   pieces=[]
   while n is not None:
    parent=parents.get(n); siblings=list(parent) if parent is not None else [n]; same=[c for c in siblings if c.tag==n.tag]
    ns,local=expanded(n.tag); pieces.append('Q{'+ns+'}'+local+'['+str(same.index(n)+1)+']'); n=parent
   return '/'+('/'.join(reversed(pieces)))
  for row in slide['nodes']:
   n=all[row['source']['elementIndex']-1]; ns,local=expanded(n.tag)
   assert [ns,local]==[row['namespaceUri'],row['local']]
   assert row['source']['path']==path(n)
   attrs={tuple(expanded(k)):v for k,v in n.attrib.items()}
   actual={(a['namespaceUri'],a['local']):a['value'] for a in row['attributes'] if a['namespaceUri']!='http://www.w3.org/2000/xmlns/'}
   assert attrs==actual
   direct=(n.text or '')+''.join(c.tail or '' for c in n)
   assert direct==row.get('text','')
   ancestor=parents.get(n); nearest=None
   while ancestor is not None:
    ans,alocal=expanded(ancestor.tag)
    if ans in P and alocal=='cTn': nearest=all.index(ancestor)+1; break
    ancestor=parents.get(ancestor)
   assert nearest==row.get('parentTimeNodeElementIndex')
   for key in ['drawingReference','timeReference']:
    ref=row.get(key)
    if ref:
     if ref['status']=='located': assert len(ref['candidates'])==1 and ref['candidates'][0]['match']=='exact'
     for candidate in ref['candidates']:
      target=all[candidate['source']['elementIndex']-1]
      assert expanded(target.tag)[1]==('cNvPr' if key=='drawingReference' else 'cTn')
      if candidate['match']=='exact': assert target.get('id')==ref.get('rawId')
   count+=1
  for row in slide['drawings']:
   n=all[row['source']['elementIndex']-1]
   assert expanded(n.tag)[1]=='cNvPr' and n.get('id')==row.get('rawId') and n.get('name')==row.get('name')
   assert row['source']['path']==path(n)
print(json.dumps({'verified':True,'oracle':'Python zipfile + ElementTree; raw indices/expanded paths/attributes/text/ancestry/candidate identities','timingNodes':count}))
`;
const artifacts = [];
for (const [index, filename] of process.argv.slice(2).entries()) {
  const bytes = await readFile(filename), report = await inspectPptxTiming(bytes), reportFile = path.join(directory, `report-${index}.json`);
  await writeFile(reportFile, JSON.stringify(report, null, 2) + "\n");
  const independent = JSON.parse((await run("python3", ["-c", python, filename, reportFile], { timeout: 45000, maxBuffer: 4 * 1024 * 1024 })).stdout); assert.equal(independent.verified, true);
  artifacts.push({ input: path.relative(process.cwd(), path.resolve(filename)), sha256: sha(bytes), byteLength: bytes.length, dialect: report.dialect, independent,
    workUnits: report.workUnits, recordCount: report.recordCount, textCharacters: report.textCharacters,
    slides: report.slides.map(s => ({ slideId: s.slideId, part: s.part, timingNodeCount: s.nodes.length, drawings: s.drawings.length,
      behaviors: s.nodes.filter(n => n.role === "behavior").map(n => ({ source: n.source, local: n.local, attributes: n.attributes })),
      drawingTargets: s.nodes.filter(n => n.drawingReference).map(n => ({ source: n.source, kind: n.targetKind, reference: n.drawingReference })) })),
    findings: report.findings, nativePlayback: report.nativePlayback, authoredMapping: report.authoredMapping });
}
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), artifactsDirectory: path.relative(process.cwd(), directory),
  scope: "Public source timing inventory independently checked against raw package XML. Input provenance/application playback must be supplied separately; this tool runs no native application.", artifacts };
await writeFile("docs/testing/native-timing-inventory-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, artifacts: artifacts.map(a => ({ input: a.input, independent: a.independent })) }));
