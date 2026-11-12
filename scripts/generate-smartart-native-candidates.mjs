// Optional offline admission preparation. No native app is launched and no
// third-party component payload is copied into the checked-in fixture corpus.
import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import path from "node:path";
import { inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
import { buildNativeDiagramCandidate } from "./research/native-diagram-package.mjs";
import { projectDiagramContent } from "./research/diagram-content-model.mjs";

const presentation = process.argv.includes("--presentation") ? "layout" : "content-only";
const modelIds = process.argv.includes("--guid") ? "guid" : "int32";
if (modelIds === "guid" && presentation !== "layout") throw new TypeError("The GUID corpus requires the explicit --presentation mode.");
const sourceDirectory = path.resolve(process.argv.slice(2).find(arg => !arg.startsWith("--")) ?? ".tmp/smartart-source-corpus/0ea8f79d15a6");
const provenance = JSON.parse(await readFile(path.join(sourceDirectory, "provenance.json"), "utf8"));
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
await mkdir(".tmp/native-diagram-generated", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/native-diagram-generated/run-")), records = [], componentSources = [];
const n = (id, text, children = [], role = "content") => ({ id, paragraphs: [text], children, role });
function input(family, name) {
  const base = { family, ...(family === "cycle" ? { closure: "layout-generated" } : {}) };
  if (name === "empty") return { ...base, roots: [], nodes: [] };
  if (name === "single") return { ...base, roots: ["one"], nodes: [n("one", `${family} single`)] };
  const label = name === "long" ? "Long CJK 中文 😀 & < > \" ' text with\na soft break and a deliberately lengthy line. ".repeat(6) : `${family} lead`;
  if (family === "cycle") return { ...base, roots: ["a", "b", "c"], nodes: [n("c", "Cycle C"), n("a", label), n("b", "Cycle B")] };
  if (family === "hierarchy") return { ...base, roots: ["lead", "peer"], nodes: [n("peer", "Peer root"), n("assistant", "Assistant", [], "assistant"), n("lead", label, ["worker", "assistant"]), n("worker", "Worker")] };
  return { ...base, roots: ["plan", "ship"], nodes: [n("detail", "Research detail"), n("plan", label, ["detail"]), n("release", "Release detail"), n("ship", "Ship", ["release"])] };
}
for (const [family, name] of [["process", "smartart-accent-process.pptx"], ["hierarchy", "smartart-org-chart.pptx"], ["cycle", "smartart-cycle.pptx"]]) {
  const sample = provenance.samples.find(s => s.name === name); assert.ok(sample);
  const sourcePath = path.join(sourceDirectory, name), bytes = await readFile(sourcePath); assert.equal(sha(bytes), sample.sha256);
  const original = await inspectPptx(bytes), components = original.slides[0].diagrams[0].components;
  const requested = Object.fromEntries(components.filter(c => c.role !== "data").map(c => [c.role, c.part]));
  const extraction = spawnSync("python3", ["-c", String.raw`
import sys,json,zipfile,base64,hashlib,xml.etree.ElementTree as E
p=sys.argv[1];assert hashlib.sha256(open(p,'rb').read()).hexdigest()==sys.argv[3]
with zipfile.ZipFile(p) as z:
 assert z.testzip() is None
 result={}
 for role,part in json.loads(sys.argv[2]).items():
  b=z.read(part.lstrip('/'));r=E.fromstring(b)
  result[role]=dict(base64=base64.b64encode(b).decode(),part=part,sha256=hashlib.sha256(b).hexdigest(),bytes=len(b),uniqueId=r.get('uniqueId'))
 print(json.dumps(result))
`, sourcePath, JSON.stringify(requested), sample.sha256], { encoding: "utf8", timeout: 30000, maxBuffer: 8 * 1024 * 1024 });
  assert.equal(extraction.status, 0, extraction.stderr);
  const extracted = JSON.parse(extraction.stdout), supplied = Object.fromEntries(Object.entries(extracted).map(([role, item]) => [role, new Uint8Array(Buffer.from(item.base64, "base64"))]));
  assert.equal(extracted.style.uniqueId, "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1");
  assert.equal(extracted.colors.uniqueId, "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2");
  componentSources.push({ family, name, source: sample.source, sourceSha256: sample.sha256, components: Object.fromEntries(Object.entries(extracted).map(([role, { base64: omitted, ...item }]) => [role, item])) });
  for (const name of ["empty", "single", "standard", "long"]) for (const cache of ["absent", "empty"]) {
    const authored = input(family, name), built = await buildNativeDiagramCandidate(authored, supplied, { cache, presentation, modelIds }), file = `${family}-${name}-${cache}${presentation === "layout" ? "-presentation" : ""}${modelIds === "guid" ? "-guid" : ""}.pptx`;
    await writeFile(path.join(directory, file), built.bytes);
    const projection = projectDiagramContent(built.inspection.diagramInventory.dataModels[0], built.specimen.layoutIdentifier);
    assert.equal(projection.candidate.nodes.length, authored.nodes.length);
    const session = await importPptx(built.bytes, { projection: "content", diagramPreview: "cache" });
    try {
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
      assert.deepEqual((await session.save(plan.plan)).bytes, built.bytes);
      assert.equal(session.report.diagramPreviews[0].status, "placeholder");
      records.push({ file, family, case: name, cache, presentation, modelIdPolicy: modelIds, authored, identities: built.specimen.identities, sha256: sha(built.bytes), bytes: built.bytes.length, exactNoOp: true,
        preview: { status: session.report.diagramPreviews[0].status, reason: session.report.diagramPreviews[0].reason }, nativeAdmission: "untested" });
    } finally { session.dispose(); }
  }
}
await writeFile(path.join(directory, "manifest.json"), JSON.stringify({ componentSources, records }, null, 2) + "\n");
const verified = spawnSync("python3", ["-c", String.raw`
import sys,json,hashlib,zipfile,xml.etree.ElementTree as E
from pathlib import Path
base=Path(sys.argv[1]);manifest=json.loads((base/'manifest.json').read_text())
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';P='{http://schemas.openxmlformats.org/presentationml/2006/main}';S='{http://schemas.microsoft.com/office/drawing/2008/diagram}';R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
cases=[]
for row in manifest['records']:
 data=(base/row['file']).read_bytes();assert hashlib.sha256(data).hexdigest()==row['sha256']
 authored=row['authored'];nodes={n['id']:n for n in authored['nodes']};order=[]
 def visit(key):
  order.append(key)
  for child in nodes[key]['children']:visit(child)
 for root in authored['roots']:visit(root)
 def identity(serial):return '{00000000-0000-4000-8000-'+format(serial,'012X')+'}' if row.get('modelIdPolicy')=='guid' else str(serial)
 expected_ids={key:identity(i+1) for i,key in enumerate(order)}
 assert row['identities']==[dict(id=key,modelId=expected_ids[key]) for key in order]
 expected_edges=[(identity(0),expected_ids[key],str(i)) for i,key in enumerate(authored['roots'])]
 for key in order:expected_edges.extend((expected_ids[key],expected_ids[c],str(i)) for i,c in enumerate(nodes[key]['children']))
 with zipfile.ZipFile(base/row['file']) as z:
  assert z.testzip() is None
  root=E.fromstring(z.read('ppt/diagrams/data1.xml'));pts=root.find(D+'ptLst');cxns=root.find(D+'cxnLst')
  all_ids=[n.get('modelId') for n in list(pts)+list(cxns)];assert len(all_ids)==len(set(all_ids))
  byid={n.get('modelId'):n for n in pts};assert byid[identity(0)].get('type')=='doc'
  display=[n for n in pts if n.get('type')=='pres']
  if row['presentation']=='content-only':assert not display and len(all_ids)==4*len(nodes)+1
  else:
   assert display
   names={n.get('name') for n in E.fromstring(z.read('ppt/diagrams/layout1.xml')).iter(D+'layoutNode')}
   keys=[(n.find(D+'prSet').get('presAssocID'),n.find(D+'prSet').get('presName')) for n in display];assert len(keys)==len(set(keys))
   assert all(name in names and assoc in byid for assoc,name in keys)
  for key in order:
   p=byid[expected_ids[key]];assert p.get('type','node')==('asst' if nodes[key]['role']=='assistant' else 'node')
   paragraphs=[]
   for para in p.find(D+'t').findall(A+'p'):
    paragraphs.append(''.join('\n' if run.tag==A+'br' else ''.join(run.find(A+'t').itertext()) for run in para))
   assert paragraphs==nodes[key]['paragraphs']
  actual_edges=[]
  for cxn in cxns:
   if cxn.get('type')!='parOf':
    assert cxn.get('type') in ['presOf','presParOf'] and cxn.get('srcId') in byid and cxn.get('destId') in byid
    assert byid[cxn.get('destId')].get('type')=='pres'
    if cxn.get('type')=='presParOf':assert byid[cxn.get('srcId')].get('type')=='pres'
    continue
   actual_edges.append((cxn.get('srcId'),cxn.get('destId'),cxn.get('srcOrd')));assert cxn.get('destOrd')=='0'
   for attr,kind in [('parTransId','parTrans'),('sibTransId','sibTrans')]:
    pt=byid[cxn.get(attr)];assert pt.get('type')==kind and pt.get('cxnId')==cxn.get('modelId')
  assert actual_edges==expected_edges
  slide=E.fromstring(z.read('ppt/slides/slide1.xml'));frames=slide.findall('./'+P+'cSld/'+P+'spTree/'+P+'graphicFrame');assert len(frames)==1
  assert len(slide.findall('./'+P+'cSld/'+P+'spTree/'+P+'sp'))==1 # Ordinary control text only.
  rels={n.get('Id'):n for n in E.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))}
  assert rels['pptkitDiagramStyle'].get('Type').endswith('/diagramQuickStyle')
  has_cache=root.find('./'+D+'extLst/'+A+'ext/'+S+'dataModelExt') is not None;assert has_cache==(row['cache']=='empty')
  if row['cache']=='empty':
   assert root.find('./'+D+'extLst/'+A+'ext/'+S+'dataModelExt').get('relId')=='pptkitDiagramDrawing'
   assert rels['pptkitDiagramDrawing'].get('Type')=='http://schemas.microsoft.com/office/2007/relationships/diagramDrawing'
   assert not list(E.fromstring(z.read('ppt/diagrams/drawing1.xml')).iter(S+'sp'))
  else:assert 'pptkitDiagramDrawing' not in rels and 'ppt/diagrams/drawing1.xml' not in z.namelist()
  source=next(c for c in manifest['componentSources'] if c['family']==row['family'])
  for role in ['layout','colors','style']:assert hashlib.sha256(z.read('ppt/diagrams/'+role+'1.xml')).hexdigest()==source['components'][role]['sha256']
 cases.append(dict(file=row['file'],family=row['family'],case=row['case'],cache=row['cache'],presentation=row['presentation'],modelIdPolicy=row.get('modelIdPolicy','int32'),presentationPoints=len(display),sha256=row['sha256'],bytes=row['bytes'],nodes=len(nodes),connections=len(expected_edges),points=len(pts),exactNoOp=row['exactNoOp'],preview=row['preview'],independentZipXmlTopology=True,nativeAdmission=row['nativeAdmission']))
print(json.dumps(dict(python=sys.version.split()[0],cases=cases)))
`, directory], { encoding: "utf8", timeout: 60000, maxBuffer: 4 * 1024 * 1024 });
assert.equal(verified.status, 0, verified.stderr);
const independent = JSON.parse(verified.stdout);
const evidence = { profile: "native-diagram-minimum-candidates-v1", directory, sourceCommit: provenance.commit, componentSources,
  runtime: { node: process.version, python: independent.python }, cases: independent.cases,
  limitations: ["Fresh research packages with new content/transition data, optional exact-layout presentation data, and explicitly supplied source layout/style/colors; not a public native exporter or preservation writer.", "Cached shapes are never synthesized. Native opening/editing and actual family admission remain separate gates.", "Independent structural/identity/text/package checks are not full OOXML schema or visual-fidelity certification.", "Third-party component bytes and generated artifacts remain local pending fixture redistribution/licensing admission; portable tests use owned synthetic component roots."] };
await writeFile(new URL(modelIds === "guid" ? "../docs/testing/smartart-native-guid-candidate-evidence.json" : presentation === "layout" ? "../docs/testing/smartart-native-presentation-candidate-evidence.json" : "../docs/testing/smartart-native-candidate-evidence.json", import.meta.url), JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, cases: evidence.cases.length, nativeAdmission: "untested", independentChecks: true }, null, 2));
