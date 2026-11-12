// Optional offline verification of retained, hash-pinned source/native artifacts.
// Does not launch Office or replay the separately recorded native observations.
import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { inspectPptx } from "../packages/pptx-parser/dist/index.js";
import { bindDiagramSourceContent } from "./research/diagram-source-bindings.mjs";
const run = promisify(execFile), sha = data => createHash("sha256").update(data).digest("hex");
const sourceDirectory = path.resolve(process.argv[2] ?? ".tmp/smartart-source-corpus/0ea8f79d15a6");
const sourceEvidenceBytes = await readFile(new URL("../docs/testing/smartart-content-model-evidence.json", import.meta.url));
const generatedEvidenceBytes = await readFile(new URL("../docs/testing/smartart-generated-powerpoint-evidence.json", import.meta.url));
const sourceEvidence = JSON.parse(sourceEvidenceBytes), generatedEvidence = JSON.parse(generatedEvidenceBytes);
const provenance = JSON.parse(await readFile(path.join(sourceDirectory, "provenance.json"), "utf8"));
assert.equal(provenance.commit, sourceEvidence.sourceCommit);
await mkdir(".tmp/diagram-source-bindings", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/diagram-source-bindings/run-"));
const specs = sourceEvidence.cases.map(c => ({ name: `source-${c.family}`, family: c.family, versions: c.versions.map(v => {
  const file = path.join(v.kind === "source" ? sourceDirectory : sourceEvidence.retainedDirectory, c.name);
  if (v.kind === "source") assert.equal(provenance.samples.find(p => p.name === c.name)?.sha256, v.sha256);
  return { kind: v.kind, file, sha256: v.sha256 };
}) }));
const stage = name => {
  assert.equal(path.basename(name), name);
  const s = generatedEvidence.stages.find(s => path.basename(s.file) === name); assert.ok(s);
  return { kind: name, file: path.join(generatedEvidence.directory, name), sha256: s.sha256 };
};
const control = generatedEvidence.identityControl;
specs.push({ name: "generated-guid-cycle-text-rounds", family: "cycle", versions: [{ kind: "input", file: control.source, sha256: control.sha256 }, ...control.stages.map(stage)] });
const deletion = generatedEvidence.deletionControl;
specs.push({ name: "native-cycle-blank-deletion", family: "cycle", versions: [stage(deletion.before), stage(deletion.after)] });
const inspectVersion = async version => {
  const bytes = await readFile(version.file); assert.equal(sha(bytes), version.sha256);
  const inspection = await inspectPptx(bytes); assert.equal(inspection.slides[0].diagrams.length, 1);
  const diagram = inspection.slides[0].diagrams[0], dataPart = diagram.components.find(c => c.role === "data").part;
  const model = inspection.diagramInventory.dataModels.find(m => m.part === dataPart), layout = diagram.components.find(c => c.role === "layout").layoutIdentifier;
  assert.ok(model); assert.deepEqual(model.findings, []); return { model, layout };
};
const records = [];
for (const [index, spec] of specs.entries()) {
  let state; const versions = [];
  // Explicit per-case research namespaces, never a global runtime allocator.
  const namespace = `${(index + 1).toString(16).toUpperCase().padStart(8, "0")}-1234-4ABC-8DEF`;
  for (const version of spec.versions) {
    const { model, layout } = await inspectVersion(version);
    const bound = bindDiagramSourceContent(model, layout, { scope: spec.name, namespace, state });
    assert.equal(bound.candidate.family, spec.family); state = JSON.parse(JSON.stringify(bound.state));
    versions.push({ ...version, part: model.part, layout, bound });
  }
  records.push({ name: spec.name, family: spec.family, namespace, versions });
}
const integer = generatedEvidence.inputs.find(i => i.family === "cycle");
const renumbered = stage("cycle-before-native-delete.pptx");
const original = await inspectVersion({ file: integer.input, sha256: integer.sha256 }), saved = await inspectVersion(renumbered);
const initial = bindDiagramSourceContent(original.model, original.layout, { scope: "integer-native-renumber", namespace: "FFFFFFFF-1234-4ABC-8DEF" });
assert.throws(() => bindDiagramSourceContent(saved.model, saved.layout, { scope: "integer-native-renumber", state: initial.state }), /context\/root mismatch/);
const expectedRejections = [{ name: "integer-native-renumber", before: { file: integer.input, sha256: integer.sha256, part: original.model.part }, after: { ...renumbered, part: saved.model.part },
  reason: "Source document root/model IDs changed on native save; no label or positional matching is permitted.", status: "rejected-context-root-mismatch" }];
const manifest = path.join(directory, "inspection.json");
await writeFile(manifest, JSON.stringify({ records, expectedRejections }, null, 2) + "\n");
const independent = JSON.parse((await run("python3", ["-c", String.raw`
import sys,json,zipfile,hashlib,re,xml.etree.ElementTree as E
from pathlib import Path
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def canonical(raw):
 token=raw.strip(' \t\r\n')
 if re.fullmatch(r'[+-]?\d+',token):
  n=int(token);assert -2147483648<=n<=2147483647;return 'int:'+str(n)
 assert re.fullmatch(r'\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}',token)
 return 'guid:'+token
def read(row,part):
 b=Path(row['file']).read_bytes();assert hashlib.sha256(b).hexdigest()==row['sha256']
 with zipfile.ZipFile(row['file']) as z:
  assert z.testzip() is None
  root=E.fromstring(z.read(part.lstrip('/')));raw=list(root.iter());points=[p for p in root.find(D+'ptLst') if p.get('type','node') in ['doc','node','asst']]
  byid={canonical(p.get('modelId')):p for p in points};assert len(byid)==len(points)
  doc=next(canonical(p.get('modelId')) for p in points if p.get('type')=='doc');children={k:[] for k in byid}
  for edge in root.find(D+'cxnLst'):
   if edge.get('type','parOf')=='parOf':children[canonical(edge.get('srcId'))].append((int(edge.get('srcOrd')),canonical(edge.get('destId'))))
  for siblings in children.values():siblings.sort()
  return raw,byid,doc,children
output=[];manifest=json.loads(Path(sys.argv[1]).read_text())
for case in manifest['records']:
 reserved={};serial=1;versions=[];previous=set()
 def allocate(key):
  global serial
  if key not in reserved:reserved[key]='{'+case['namespace']+'-'+format(serial,'012X')+'}';serial+=1
  return reserved[key]
 for row in case['versions']:
  raw,byid,doc,children=read(row,row['part']);allocate(doc)
  order=[];stack=[k for _,k in reversed(children[doc])]
  while stack:
   key=stack.pop();assert key not in order;order.append(key);allocate(key);stack.extend(k for _,k in reversed(children[key]))
  assert len(order)==len(byid)-1
  bound=row['bound'];candidate=bound['candidate'];bindings=bound['bindings'];state=bound['state']
  assert state['rootIdentity']==doc and state['activeSourceIds']==order
  assert state['identities']['nextSerial']==serial
  assert state['identities']['entries']==[dict(key=json.dumps(['source-point',key],separators=(',',':')),modelId=value) for key,value in reserved.items()]
  assert candidate['roots']==[reserved[k] for _,k in children[doc]]
  assert [b['canonicalModelId'] for b in bindings]==order
  assert [n['id'] for n in candidate['nodes']]==[reserved[k] for k in order]
  for binding,node,key in zip(bindings,candidate['nodes'],order):
   source=byid[key];assert raw[binding['elementIndex']-1] is source and binding['sourceModelId']==source.get('modelId')
   assert binding['id']==node['id']==reserved[key] and binding['part']==row['part']
   paragraphs=[]
   for para in source.findall('./'+D+'t/'+A+'p'):
    pieces=[]
    for run in para:
     if run.tag==A+'r':pieces.append(''.join(run.find(A+'t').itertext()))
     elif run.tag==A+'br':pieces.append('\n')
     elif run.tag==A+'fld':raise AssertionError('unadmitted field')
    paragraphs.append(''.join(pieces))
   assert node['paragraphs']==paragraphs and node['role']==('assistant' if source.get('type')=='asst' else 'content')
   assert node['children']==[reserved[k] for _,k in children[key]]
  current=set(order)
  versions.append(dict(kind=row['kind'],file=row['file'],sha256=row['sha256'],part=row['part'],layout=row['layout'],nodes=len(order),roots=len(candidate['roots']),retained=len(current&previous),added=len(current-previous),removed=len(previous-current),bindings=bindings,independentIdentityTopologyTextAndRoles=True))
  previous=current
 output.append(dict(name=case['name'],family=case['family'],versions=versions))
rejections=[]
for case in manifest['expectedRejections']:
 _,_,before,_=read(case['before'],case['before']['part']);_,_,after,_=read(case['after'],case['after']['part']);assert before!=after
 rejections.append(dict(name=case['name'],beforeRootIdentity=before,afterRootIdentity=after,independentRootMismatch=True))
print(json.dumps(dict(python=sys.version.split()[0],cases=output,rejections=rejections)))
`, manifest], { timeout: 60000, maxBuffer: 8 * 1024 * 1024 })).stdout);
const evidence = { profile: "smartart-source-managed-bindings-v1", directory, sourceCommit: provenance.commit,
  sourceEvidenceSha256: sha(sourceEvidenceBytes), generatedEvidenceSha256: sha(generatedEvidenceBytes), runtime: { node: process.version, python: independent.python },
  cases: independent.cases, expectedRejections: expectedRejections.map(row => ({ ...row, ...independent.rejections.find(r => r.name === row.name) })), limitations: ["Offline binding/file evidence only; no new native UI operations replayed or certified.", "Plain paragraph/ordered forest research only; no rich-text/style/cache/freshness or source-save authority.", "Scope/root/part/layout changes and producer-renumbered identities need a separate explicit reconciliation policy; labels/positions are not used.", "Managed binding IDs and fresh-writer native IDs remain separate; neither this adapter nor earlier native observations admits the stateful writer in native applications.", "Third-party source/native artifacts remain local pending redistribution admission."] };
await writeFile(new URL("../docs/testing/smartart-source-bindings-evidence.json", import.meta.url), JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, cases: evidence.cases.map(c => ({ name: c.name, versions: c.versions.map(v => ({ kind: v.kind, nodes: v.nodes, retained: v.retained, added: v.added, removed: v.removed })) })), expectedRejections }, null, 2));
