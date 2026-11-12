// Optional offline cache-box comparison, not a native application replay or
// cache writer. Source/managed/cache identity association is explicit.
import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { inspectPptx } from "../packages/pptx-parser/dist/index.js";
import { bindDiagramSourceContent } from "./research/diagram-source-bindings.mjs";
import { resolveCycleLayout } from "./research/cycle-layout.mjs";
import { observeCycleSpacing } from "./research/cycle-spacing-observation.mjs";
const run = promisify(execFile), sha = bytes => createHash("sha256").update(bytes).digest("hex");
const evidencePath = path.resolve(process.argv[2] ?? "docs/testing/smartart-generated-powerpoint-evidence.json"), evidenceBytes = await readFile(evidencePath), evidence = JSON.parse(evidenceBytes);
await mkdir(".tmp/cycle-layout", { recursive: true }); const directory = await mkdtemp(path.resolve(".tmp/cycle-layout/run-"));
const records = [], states = new Map();
for (const stage of evidence.stages.filter(s => s.family === "cycle")) {
  assert.equal(path.basename(stage.file), stage.file); const file = path.join(evidence.directory, stage.file), bytes = await readFile(file); assert.equal(sha(bytes), stage.sha256);
  const inspection = await inspectPptx(bytes), diagram = inspection.slides[0].diagrams[0], layout = diagram.components.find(c => c.role === "layout").layoutIdentifier;
  const model = inspection.diagramInventory.dataModels.find(m => m.part === diagram.components.find(c => c.role === "data").part);
  const group = stage.file.startsWith("cycle-guid-") ? "guid" : "integer", namespace = group === "guid" ? "88888888-1234-4ABC-8DEF" : "99999999-1234-4ABC-8DEF";
  const bound = bindDiagramSourceContent(model, layout, { scope: `native-cycle/${group}`, namespace, state: states.get(group) }); states.set(group, bound.state);
  const extracted = JSON.parse((await run("python3", ["-c", String.raw`
import sys,json,zipfile,hashlib,xml.etree.ElementTree as E
from decimal import Decimal
P='{http://schemas.openxmlformats.org/presentationml/2006/main}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';S='{http://schemas.microsoft.com/office/drawing/2008/diagram}'
assert hashlib.sha256(open(sys.argv[1],'rb').read()).hexdigest()==sys.argv[2]
def number(value):return float(Decimal(value)/Decimal(12700))
with zipfile.ZipFile(sys.argv[1]) as z:
 slide=E.fromstring(z.read('ppt/slides/slide1.xml'));frames=[f for f in slide.iter(P+'graphicFrame') if any(n.get('uri')=='http://schemas.openxmlformats.org/drawingml/2006/diagram' for n in f.iter(A+'graphicData'))];assert len(frames)==1
 x=frames[0].find(P+'xfrm');assert not x.attrib;off=x.find(A+'off');ext=x.find(A+'ext');frame=dict(x=number(off.get('x')),y=number(off.get('y')),width=number(ext.get('cx')),height=number(ext.get('cy')))
 data=E.fromstring(z.read(sys.argv[3].lstrip('/')));points={p.get('modelId'):p for p in data.find(D+'ptLst')};boxes=[]
 caches=[n for n in z.namelist() if n.startswith('ppt/diagrams/drawing') and n.endswith('.xml')];assert len(caches)==1
 cache=E.fromstring(z.read(caches[0]));tree=cache.find(S+'spTree');assert tree is not None
 assert len(list(cache.iter(S+'spTree')))==1 and not list(tree.iter(S+'grpSp'))
 assert all(not list(n) for n in tree.findall(S+'grpSpPr'))
 for shape in tree.findall(S+'sp'):
  geom=shape.find('./'+S+'spPr/'+A+'prstGeom')
  if geom is None or geom.get('prst')!='ellipse':continue
  pres=points[shape.get('modelId')];assert pres.get('type')=='pres';owner=pres.find(D+'prSet').get('presAssocID');assert points[owner].get('type','node')=='node'
  x=shape.find('./'+S+'spPr/'+A+'xfrm');assert not x.attrib;off=x.find(A+'off');ext=x.find(A+'ext')
  box=dict(x=frame['x']+number(off.get('x')),y=frame['y']+number(off.get('y')),width=number(ext.get('cx')),height=number(ext.get('cy')));assert abs(box['width']-box['height'])<1e-8
  line=shape.find('./'+S+'spPr/'+A+'ln')
  boxes.append(dict(sourceModelId=owner,cacheModelId=shape.get('modelId'),box=box,**(dict(strokeWidth=number(line.get('w')),strokeAlignment=line.get('algn','unspecified')) if line is not None and line.get('w') is not None else {})))
 assert len({b['sourceModelId'] for b in boxes})==len(boxes)
 print(json.dumps(dict(frame=frame,boxes=boxes,cachePart=caches[0])))
`, file, stage.sha256, model.part], { timeout: 30000, maxBuffer: 1024 * 1024 })).stdout);
  const proposed = resolveCycleLayout(bound.candidate, { frame: extracted.frame, layoutIdentifier: layout });
  const bindings = new Map(bound.bindings.map(b => [b.sourceModelId, b.id])); assert.equal(extracted.boxes.length, proposed.nodes.length);
  const deltas = extracted.boxes.map(native => {
    const id = bindings.get(native.sourceModelId), node = proposed.nodes.find(n => n.id === id); assert.ok(node);
    return { id, sourceModelId: native.sourceModelId, cacheModelId: native.cacheModelId, native: native.box, proposed: node.box,
      delta: Object.fromEntries(["x", "y", "width", "height"].map(k => [k, node.box[k] - native.box[k]])) };
  });
  const values = deltas.flatMap(d => Object.values(d.delta)), maxAbsCoordinateDelta = Math.max(...values.map(Math.abs));
  const attribution = observeCycleSpacing({ frame: extracted.frame, circles: bound.candidate.roots.map(id => {
    const native = extracted.boxes.find(b => bindings.get(b.sourceModelId) === id), node = proposed.nodes.find(n => n.id === id);
    return { id, box: native.box, proposed: node.box, ...(native.strokeWidth === undefined ? {} : { strokeWidth: native.strokeWidth, strokeAlignment: native.strokeAlignment }) };
  }) });
  records.push({ file, sha256: stage.sha256, part: model.part, bound, proposed, extracted, deltas, maxAbsCoordinateDelta, rmsCoordinateDelta: Math.sqrt(values.reduce((n, v) => n + v * v, 0) / values.length),
    tolerancePoints: 0.1, matchesWithinTolerance: maxAbsCoordinateDelta <= 0.1, attribution });
}
const manifest = path.join(directory, "comparison.json"); await writeFile(manifest, JSON.stringify(records, null, 2) + "\n");
const independent = JSON.parse((await run("python3", ["-c", String.raw`
import sys,json,math,hashlib
from pathlib import Path
out=[]
for row in json.loads(Path(sys.argv[1]).read_text()):
 assert hashlib.sha256(Path(row['file']).read_bytes()).hexdigest()==row['sha256']
 roots=row['bound']['candidate']['roots'];n=len(roots);assert n in [3,4,5]
 r=1.5/(2*math.sin(math.pi/n));coords=[(r*math.sin(2*math.pi*i/n),-r*math.cos(2*math.pi*i/n)) for i in range(n)]
 xmin=min(x for x,y in coords)-.5;xmax=max(x for x,y in coords)+.5;ymin=min(y for x,y in coords)-.5;ymax=max(y for x,y in coords)+.5
 frame=row['extracted']['frame'];d=min(frame['width']/(xmax-xmin),frame['height']/(ymax-ymin));ox=frame['x']+frame['width']/2-(xmin+xmax)*d/2;oy=frame['y']+frame['height']/2-(ymin+ymax)*d/2
 expected={key:dict(x=ox+x*d-d/2,y=oy+y*d-d/2,width=d,height=d) for key,(x,y) in zip(roots,coords)}
 for node in row['proposed']['nodes']:
  assert all(abs(node['box'][k]-v)<1e-8 for k,v in expected[node['id']].items())
 values=[]
 for item in row['deltas']:
  for key in ['x','y','width','height']:
   delta=expected[item['id']][key]-item['native'][key];assert abs(delta-item['delta'][key])<1e-8;values.append(delta)
 maximum=max(abs(v) for v in values);rms=math.sqrt(sum(v*v for v in values)/len(values));assert abs(maximum-row['maxAbsCoordinateDelta'])<1e-8 and abs(rms-row['rmsCoordinateDelta'])<1e-8
 actual={x['id']:x['native'] for x in row['deltas']};ordered=[actual[k] for k in roots];centers=[(b['x']+b['width']/2,b['y']+b['height']/2) for b in ordered]
 cx=sum(x for x,y in centers)/n;cy=sum(y for x,y in centers)/n;diameter=sum(b['width'] for b in ordered)/n
 radius=sum((x-cx)*math.sin(2*math.pi*i/n)-(y-cy)*math.cos(2*math.pi*i/n) for i,(x,y) in enumerate(centers))/n
 gap=2*radius*math.sin(math.pi/n)/diameter-1;residual=0
 for i,b in enumerate(ordered):
  fit=dict(x=cx+radius*math.sin(2*math.pi*i/n)-diameter/2,y=cy-radius*math.cos(2*math.pi*i/n)-diameter/2,width=diameter,height=diameter)
  residual=max(residual,max(abs(b[k]-fit[k]) for k in b))
 observed=row['attribution'];assert abs(gap-observed['observedGapRatio'])<1e-10 and abs(residual-observed['maxFitResidual'])<1e-8
 rounded=max(abs(math.floor(item['proposed'][key]*12700+.5)/12700-item['native'][key]) for item in row['deltas'] for key in ['x','y','width','height'])
 assert abs(rounded-observed['rounding']['roundedMaxCoordinateDelta'])<1e-8
 out.append(dict(file=Path(row['file']).name,sha256=row['sha256'],nodes=n,frame=frame,diameter=d,maxAbsCoordinateDelta=maximum,rmsCoordinateDelta=rms,tolerancePoints=row['tolerancePoints'],matchesWithinTolerance=maximum<=row['tolerancePoints'],deltas=row['deltas'],attribution=observed,unmeasuredTextNodes=len(row['proposed']['findings']),independentPackingAndDeltaCheck=True,independentSpacingAndQuantizationCheck=True))
print(json.dumps(dict(python=sys.version.split()[0],records=out)))
`, manifest], { timeout: 30000, maxBuffer: 4 * 1024 * 1024 })).stdout);
const svgFiles = [];
for (const row of records) {
  const f = row.extracted.frame, w = 2 * f.width + 60, h = f.height + 70;
  const circles = (nodes, color) => nodes.map((n, i) => `<circle cx="${n.x + n.width / 2}" cy="${n.y + n.height / 2}" r="${n.width / 2}" fill="${color}"/><text x="${n.x + n.width / 2}" y="${n.y + n.height / 2 + 4}" text-anchor="middle" fill="white" font-size="14">${i + 1}</text>`).join("");
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><rect width="100%" height="100%" fill="white"/><text x="20" y="22" font-family="sans-serif" font-size="15">PowerPoint cached circles</text><text x="${40 + f.width}" y="22" font-family="sans-serif" font-size="15">Experimental equal-circle packing</text><g transform="translate(${20 - f.x} ${40 - f.y})">${circles(row.deltas.map(d => d.native), "#4266B1")}</g><g transform="translate(${40 + f.width - f.x} ${40 - f.y})">${circles(row.deltas.map(d => d.proposed), "#087F8C")}</g></svg>`;
  const file = path.join(directory, `${path.basename(row.file, ".pptx")}.svg`); await writeFile(file, svg); svgFiles.push({ file, sha256: sha(svg) });
}
const result = { profile: "equal-circle-cycle-native-box-comparison-v1", directory, nativeApplication: evidence.application, build: evidence.build, os: evidence.os,
  nativeEvidenceSha256: sha(evidenceBytes), runtime: { node: process.version, python: independent.python }, records: independent.records, svgFiles,
  limitations: ["Offline comparisons over retained native outputs; no new UI replay or native authoring admission.", "0.1-point tolerance is an explicit research comparison threshold; mismatches remain mismatches.", "Only circle boxes are compared. Arrows, text fitting/typography, cache freshness and other native parameters are unverified.", "Two-node curved geometry is a proposed policy with no native comparison. LibreOffice ellipse observations remain separate."] };
const beforeTheme = records.find(r => path.basename(r.file) === "cycle-before-native-delete.pptx"), afterTheme = records.find(r => path.basename(r.file) === "cycle-round2-theme.pptx");
if (beforeTheme && afterTheme) {
  const geometry = r => r.extracted.boxes.map(b => ({ sourceModelId: b.sourceModelId, box: b.box })).sort((a, b) => a.sourceModelId.localeCompare(b.sourceModelId));
  const widths = r => [...new Set(r.extracted.boxes.map(b => b.strokeWidth))];
  assert.deepEqual(geometry(beforeTheme), geometry(afterTheme)); assert.notDeepEqual(widths(beforeTheme), widths(afterTheme));
  result.strokeControl = { before: { file: path.basename(beforeTheme.file), sha256: beforeTheme.sha256, widthsPoints: widths(beforeTheme) }, after: { file: path.basename(afterTheme.file), sha256: afterTheme.sha256, widthsPoints: widths(afterTheme) }, exactCircleGeometryRetained: true,
    scope: "Observed theme-change pair only; does not prove arbitrary stroke changes never affect native re-layout." };
}
await writeFile(new URL("../docs/testing/smartart-cycle-layout-evidence.json", import.meta.url), JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ directory, records: result.records.map(r => ({ file: r.file, nodes: r.nodes, maxAbsCoordinateDelta: r.maxAbsCoordinateDelta, matchesWithinTolerance: r.matchesWithinTolerance })) }, null, 2));
