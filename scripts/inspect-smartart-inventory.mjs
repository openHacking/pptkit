import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { normalizePresentation } from "../packages/core/dist/index.js";

// Optional, offline source experiment. Fetching or redistributing third-party files
// is not part of this runner or the ordinary test suite.
const directory = path.resolve(process.argv[2] ?? ".tmp/smartart-source-corpus/0ea8f79d15a6");
const provenance = JSON.parse(await readFile(path.join(directory, "provenance.json"), "utf8"));
const sha256 = bytes => createHash("sha256").update(bytes).digest("hex");
const records = [];
const previewEnabled = process.argv.includes("--preview");
const outputIndex = process.argv.indexOf("--output");
if (outputIndex >= 0 && !process.argv[outputIndex + 1]) throw new TypeError("--output requires an evidence filename.");
const evidenceFile = outputIndex >= 0 ? path.resolve(process.argv[outputIndex + 1]) : new URL("../docs/testing/smartart-inventory-evidence.json", import.meta.url);
for (const sample of provenance.samples) {
  assert.equal(path.basename(sample.name), sample.name);
  const bytes = new Uint8Array(await readFile(path.join(directory, sample.name)));
  assert.equal(sha256(bytes), sample.sha256);
  assert.equal(createHash("sha1").update(Buffer.from(`blob ${bytes.length}\0`)).update(bytes).digest("hex"), sample.gitBlob);
  const inspection = await inspectPptx(bytes), session = await importPptx(bytes, { projection: "structure" });
  try {
    const prep = await session.prepareSave({ mode: "no-op" }); assert.equal(prep.status, "ready");
    const saved = (await session.save(prep.plan)).bytes; assert.deepEqual(saved, bytes);
    await writeFile(path.join(directory, `${sample.name}.no-op.pptx`), saved);
    let cachePreviews, previewTextFrames;
    if (previewEnabled) {
      const preview = await importPptx(bytes, { projection: "content", diagramPreview: "cache" });
      try {
        cachePreviews = preview.report.diagramPreviews;
        previewTextFrames = [];
        for (const slide of normalizePresentation(preview.document).slides) {
          const stack = [...slide.elements];
          while (stack.length) {
            const element = stack.pop();
            if (element.type === "group") stack.push(...element.children);
            else if (element.type === "text" && /-cache-\d+-text$/.test(element.id)) previewTextFrames.push({ id: element.id, box: element.box, transform: element.transform, text: element.plainText });
          }
        }
        const rendered = await renderPresentationToSvg(preview.document);
        for (const [index, slide] of rendered.slides.entries()) await writeFile(path.join(directory, `${sample.name}.preview-${index + 1}.svg`), slide.svg);
      } finally { preview.dispose(); }
    }
    records.push({ name: sample.name, slides: inspection.slides, inventory: inspection.diagramInventory ?? null, ...(cachePreviews ? { cachePreviews, previewTextFrames } : {}) });
  } finally { session.dispose(); }
}
await writeFile(path.join(directory, "inspection.json"), JSON.stringify({ records }, null, 2) + "\n");
const python = String.raw`
import hashlib,json,math,posixpath,re,sys,zipfile
from collections import Counter
from pathlib import Path
import xml.etree.ElementTree as E
base=Path(sys.argv[1]); manifest=json.loads((base/'provenance.json').read_text())
reports=json.loads((base/'inspection.json').read_text())['records']
P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
D='http://schemas.openxmlformats.org/drawingml/2006/diagram'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
DSP='http://schemas.microsoft.com/office/drawing/2008/diagram'
CACHE='http://schemas.microsoft.com/office/2007/relationships/diagramDrawing'
REL='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
def tag(ns,local): return '{'+ns+'}'+local
def target(source,value): return posixpath.normpath(posixpath.join(posixpath.dirname(source),value))
cases=[]
for sample,report in zip(manifest['samples'],reports,strict=True):
 assert sample['name']==report['name']
 data=(base/sample['name']).read_bytes(); saved=(base/(sample['name']+'.no-op.pptx')).read_bytes()
 assert hashlib.sha256(data).hexdigest()==sample['sha256'] and data==saved
 with zipfile.ZipFile(base/sample['name']) as z:
  assert z.testzip() is None
  trees={}; rels={}; incoming={}
  for name in z.namelist():
   if name.endswith('.rels'):
    if name=='_rels/.rels': source='/'
    else:
     parent,filename=posixpath.split(name); source='/'+posixpath.join(posixpath.dirname(parent),filename[:-5])
    rels[source]={}
    for node in E.fromstring(z.read(name)):
     rels[source][node.get('Id')]=node.attrib
     if node.get('TargetMode','Internal')!='External':
      dest=target(source,node.get('Target').split('#')[0])
      incoming.setdefault(dest,[]).append(dict(source=source,relationshipId=node.get('Id'),type=node.get('Type')))
  def tree(part):
   if part not in trees: trees[part]=E.fromstring(z.read(part.lstrip('/')))
   return trees[part]
  content=E.fromstring(z.read('[Content_Types].xml'))
  types={n.get('PartName'):n.get('ContentType') for n in content if n.tag==tag(CT,'Override')}
  defaults={n.get('Extension'):n.get('ContentType') for n in content if n.tag==tag(CT,'Default')}
  def content_type(part): return types.get(part,defaults.get(part.rsplit('.',1)[-1]))
  mainrel=next(r for r in rels['/'].values() if r['Type']==R+'/officeDocument')
  main=target('/',mainrel['Target']); presentation=tree(main)
  order=[target(main,rels[main][n.get(tag(R,'id'))]['Target']) for n in presentation.findall('./'+tag(P,'sldIdLst')+'/'+tag(P,'sldId'))]
  assert order==[s['part'] for s in report['slides']]
  layouts=[]; observed=[]; candidate_count=0
  for slide in report['slides']:
   root=tree(slide['part']); indices={n:i+1 for i,n in enumerate(root.iter())}
   raw=[n for n in root.iter(tag(P,'graphicFrame')) if n.find('./'+tag(A,'graphic')+'/'+tag(A,'graphicData')) is not None and n.find('./'+tag(A,'graphic')+'/'+tag(A,'graphicData')).get('uri')==D]
   assert len(raw)==len(slide.get('diagrams',[]))
   for frame,d in zip(raw,slide.get('diagrams',[]),strict=True):
    assert d['elementIndex']==indices[frame]
    refs=frame.find('./'+tag(A,'graphic')+'/'+tag(A,'graphicData')+'/'+tag(D,'relIds'))
    for component,attr,role,rootname in zip(d['components'],['dm','lo','cs','qs'],['diagramData','diagramLayout','diagramColors','diagramQuickStyle'],['dataModel','layoutDef','colorsDef','styleDef'],strict=True):
     rid=refs.get(tag(R,attr)); relation=rels[slide['part']][rid]; part=target(slide['part'],relation['Target'])
     assert component['status']=='located' and component['relationshipId']==rid and component['part']==part
     assert relation['Type']==R+'/'+role and tree(part).tag==tag(D,rootname)
     assert component['contentType']==content_type(part)
     if attr=='lo':
      assert component.get('layoutIdentifier')==tree(part).get('uniqueId'); layouts.append(component.get('layoutIdentifier'))
    cache=d['cache']; datapart=d['components'][0]['part']; dataroot=tree(datapart)
    extrefs=[]; didx={n:i+1 for i,n in enumerate(dataroot.iter())}
    for ext in dataroot.findall('./'+tag(D,'extLst')+'/'+tag(A,'ext')):
     if ext.get('uri')!=DSP: continue
     for block in ext.findall('./'+tag(DSP,'dataModelExt')):
      ref=dict(elementIndex=didx[block])
      if block.get('relId') is not None: ref['relationshipId']=block.get('relId')
      if block.get('minVer') is not None: ref['minimumVersion']=block.get('minVer')
      extrefs.append(ref)
    assert cache['references']==extrefs and cache['freshness']=='unknown'
    if not extrefs: assert cache['status']=='missing-reference'
    else:
     assert len(extrefs)==1; relation=rels[slide['part']][extrefs[0]['relationshipId']]
     assert relation['Type']==CACHE and cache['status']=='located'
     assert cache['part']==target(slide['part'],relation['Target']) and tree(cache['part']).tag==tag(DSP,'drawing')
    observed.append(cache['status'])
   candidate_count+=sum(r['Type']==CACHE for r in rels[slide['part']].values())
  inv=report['inventory']; assert candidate_count==len(inv['cacheCandidates'])
  graph_evidence=[]
  data_parts={p['part'] for p in inv['parts'] if 'data' in p['roles']}
  assert data_parts=={g['part'] for g in inv.get('dataModels',[])}
  for graph in inv.get('dataModels',[]):
   groot=tree(graph['part']); gidx={n:i+1 for i,n in enumerate(groot.iter())}
   assert graph['coverage']=='identities-references-direct-text'
   point_list=groot.findall('./'+tag(D,'ptLst')); connection_list=groot.findall('./'+tag(D,'cxnLst'))
   assert len(point_list)==1 and len(connection_list)<=1
   pts=point_list[0].findall('./'+tag(D,'pt')); cxns=connection_list[0].findall('./'+tag(D,'cxn')) if connection_list else []
   assert len(pts)==len(graph['points']) and len(cxns)==len(graph['connections'])
   def modelkey(value):
    if value is None:return None
    value=value.strip(' \t\r\n')
    if re.fullmatch(r'[+-]?[0-9]+',value) and -2**31<=int(value)<2**31:return ('int',int(value))
    if re.fullmatch(r'\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}',value):return ('guid',value)
    return None
   identities={}
   for kind,nodes in [('point',pts),('connection',cxns)]:
    for node in nodes:
     key=modelkey(node.get('modelId'))
     if key is not None:identities.setdefault(key,[]).append((kind,node))
   ref_counts=Counter(); text_count=0; paragraphs=0
   def expected_ref(value,kind):
    key=modelkey(value); targets=identities.get(key,[])
    status='absent' if value is None else 'unrecognized' if key is None else 'missing' if not targets else 'ambiguous' if len(targets)>1 else 'wrong-kind' if targets[0][0]!=kind else 'resolved'
    return dict(status=status,**(dict(value=value) if value is not None else {}),**(dict(targetElementIndex=gidx[targets[0][1]]) if status=='resolved' else {}))
   def check_ref(node,attr,kind,record):
    assert record==expected_ref(node.get(attr),kind)
    ref_counts[record['status']]+=1
   for kind,nodes,records in [('point',pts,graph['points']),('connection',cxns,graph['connections'])]:
    for node,record in zip(nodes,records,strict=True):
     value=node.get('modelId'); key=modelkey(value)
     identity_status='missing' if value is None else 'unrecognized' if key is None else 'duplicate' if len(identities[key])>1 else 'unique'
     assert record['elementIndex']==gidx[node] and record.get('modelId')==value and record['identityStatus']==identity_status
     assert record.get('type')==node.get('type')
     rawkind=node.get('type','node' if kind=='point' else 'parOf').strip(' \t\r\n')
     recognized=['node','asst','doc','pres','parTrans','sibTrans'] if kind=='point' else ['parOf','presOf','presParOf','unknownRelationship']
     assert record['kind']==(rawkind if rawkind in recognized else 'unrecognized')
     if kind=='connection':
      for attr,field in [('srcId','source'),('destId','destination'),('parTransId','parentTransition'),('sibTransId','siblingTransition')]:
       if field in ['source','destination'] or node.get(attr) is not None:check_ref(node,attr,'point',record[field])
      for attr,field in [('srcOrd','sourceOrder'),('destOrd','destinationOrder')]:
       lexical=node.get(attr); assert record.get(field)==lexical
       number=int(lexical) if lexical is not None and re.fullmatch(r'[+-]?[0-9]+',lexical.strip(' \t\r\n')) else None
       assert record.get(field+'Value')==(number if number is not None and 0<=number<2**32 else None)
      assert record.get('presentationId')==node.get('presId')
     else:
      if node.get('cxnId') is not None:
       if rawkind in ['parTrans','sibTrans']:check_ref(node,'cxnId','connection',record['connection'])
       else:assert record['connection']==dict(value=node.get('cxnId'),status='uninterpreted')
      props=node.find('./'+tag(D,'prSet'))
      if props is not None:
       expected=[dict(namespaceUri=k[1:].split('}')[0] if k.startswith('{') else '',local=k.split('}')[-1],value=v) for k,v in props.attrib.items()]
       assert record['properties']==expected
       if props.get('presAssocID') is not None:check_ref(props,'presAssocID','point',record['presentationAssociation'])
      textbody=node.find('./'+tag(D,'t'))
      if textbody is not None:
       expected=[]
       for para in textbody.findall('./'+tag(A,'p')):
        item=dict(elementIndex=gidx[para],runs=[]); ppr=para.find('./'+tag(A,'pPr'))
        if ppr is not None and ppr.get('lvl') is not None:item['level']=ppr.get('lvl')
        for run in para:
         if run.tag not in [tag(A,'r'),tag(A,'br'),tag(A,'fld')]:continue
         leaf=run.find('./'+tag(A,'t')); rkind={tag(A,'r'):'text',tag(A,'br'):'break',tag(A,'fld'):'field'}[run.tag]
         assert rkind=='break' or (leaf is not None and len(leaf)==0)
         item['runs'].append(dict(elementIndex=gidx[run],kind=rkind,text='\n' if rkind=='break' else leaf.text or '',**({field:run.get(attr) for attr,field in [('id','fieldId'),('type','fieldType')] if run.get(attr) is not None} if rkind=='field' else {})))
        expected.append(item)
       assert record['text']==expected
       paragraphs+=len(expected); text_count+=sum(len(p['runs']) for p in expected)
   expected_presentations=[]
   for point in pts:
    if point.get('type','node').strip()!='pres':continue
    sources=[]; props=point.find('./'+tag(D,'prSet'))
    if props is not None and props.get('presAssocID') is not None:sources.append(dict(kind='presAssocID',elementIndex=gidx[point],reference=expected_ref(props.get('presAssocID'),'point')))
    for connection in cxns:
     if connection.get('type','parOf').strip()!='presOf':continue
     dest=expected_ref(connection.get('destId'),'point')
     if dest.get('targetElementIndex')==gidx[point]:sources.append(dict(kind='presOf',elementIndex=gidx[connection],reference=expected_ref(connection.get('srcId'),'point')))
    targets={s['reference']['targetElementIndex'] for s in sources if s['reference']['status']=='resolved'}
    agreement='none' if not sources else 'unresolved' if any(s['reference']['status']!='resolved' for s in sources) else 'different-targets' if len(targets)>1 else 'same-target'
    expected_presentations.append(dict(elementIndex=gidx[point],sources=sources,agreement=agreement,**(dict(targetElementIndex=next(iter(targets))) if agreement=='same-target' else {})))
   assert graph.get('presentationMappings',[])==expected_presentations
   for finding in graph['findings']:assert finding['elementIndex'] in gidx.values()
   graph_evidence.append(dict(part=graph['part'],status=graph['status'],points=len(pts),connections=len(cxns),paragraphs=paragraphs,textRuns=text_count,references=dict(ref_counts),findings=len(graph['findings']),presentationMappings=len(expected_presentations),presentationAgreement=dict(Counter(p['agreement'] for p in expected_presentations)),independentIdentitiesReferencesAndText=True))
  expected_pairs=list(dict.fromkeys((d['cache']['dataPart'],d['cache']['part']) for slide in report['slides'] for d in slide.get('diagrams',[]) if d['cache']['status']=='located'))
  assert expected_pairs==[(m['dataPart'],m['cachePart']) for m in inv.get('cacheMappings',[])]
  cache_mapping_evidence=[]
  for mapping in inv.get('cacheMappings',[]):
   assert mapping['coverage']=='explicit-model-references'
   root=tree(mapping['dataPart']); gidx={n:i+1 for i,n in enumerate(root.iter())}; identities={}
   for kind,path in [('point','ptLst/pt'),('connection','cxnLst/cxn')]:
    for n in root.findall('./'+'/'.join(tag(D,p) for p in path.split('/'))):
     key=modelkey(n.get('modelId'))
     if key is not None:identities.setdefault(key,[]).append((kind,n))
   cache_root=tree(mapping['cachePart']); ci={n:i+1 for i,n in enumerate(cache_root.iter())}; cache_trees=cache_root.findall('./'+tag(DSP,'spTree')); assert len(cache_trees)==1
   def recognized_shapes(group):
    for n in group:
     if n.tag==tag(DSP,'sp'):yield n
     elif n.tag==tag(DSP,'grpSp'):yield from recognized_shapes(n)
   expected=[dict(elementIndex=ci[n],model=expected_ref(n.get('modelId'),'point')) for n in recognized_shapes(cache_trees[0])]
   assert mapping['shapes']==expected
   failures=sum(s['model']['status']!='resolved' for s in expected)
   assert mapping['status']==('partial' if failures else 'inspected') and len(mapping['findings'])==failures
   cache_mapping_evidence.append(dict(dataPart=mapping['dataPart'],cachePart=mapping['cachePart'],status=mapping['status'],shapes=len(expected),references=dict(Counter(s['model']['status'] for s in expected)),independentPairAndShapeMapping=True))
  for part in inv['parts']:
   assert part['contentType']==content_type(part['part'])
   key=lambda e:(e['source'],e['relationshipId'],e['type'])
   assert sorted(part['incomingRelationships'],key=key)==sorted(incoming.get(part['part'],[]),key=key)
   if set(part['roles'])=={'image'}: assert part['imageRelationships']==[]; continue
   rtree=tree(part['part']); idx={n:i+1 for i,n in enumerate(rtree.iter())}; image_refs={}
   if 'cache' in part['roles']:
    expected=[dict(elementIndex=idx[n],**(dict(modelId=n.get('modelId')) if n.get('modelId') is not None else {})) for n in rtree.iter(tag(DSP,'sp'))]
    assert part.get('modelReferences',[])==expected
   for blip in rtree.iter(tag(A,'blip')):
    for attr in ['embed','link']:
     rid=blip.get(tag(R,attr))
     if rid is not None: image_refs.setdefault(rid,[]).append(dict(elementIndex=idx[blip],attribute=attr))
   declared={rid for rid,rel in rels.get(part['part'],{}).items() if rel['Type']==R+'/image'}
   assert set(image_refs)|declared=={r['relationshipId'] for r in part['imageRelationships']}
   for img in part['imageRelationships']: assert img['references']==image_refs.get(img['relationshipId'],[])
  previews=[]
  for preview in report.get('cachePreviews',[]):
   item=dict(status=preview['status'],freshness=preview['freshness'],reason=preview.get('reason'))
   if preview.get('cachePart'):
    st=tree(preview['cachePart']).find('./'+tag(DSP,'spTree'))
    count=sum(n.tag not in [tag(DSP,'nvGrpSpPr'),tag(DSP,'grpSpPr')] for n in st)
    item['rawCacheChildren']=count
    if count==0: assert preview['status']=='placeholder' and 'empty' in preview['reason']
   assert preview['freshness']=='unknown'
   previews.append(item)
  checked_frames=[]
  for frame in report.get('previewTextFrames',[]):
   index=int(re.search(r'-cache-(\d+)-text$',frame['id']).group(1))
   owner=next(p for p in report['cachePreviews'] if frame['id'].startswith(p['elementId']+'-cache-'))
   root=tree(owner['cachePart']); nodes=list(root.iter()); node=nodes[index-1]
   tx=node.find('./'+tag(DSP,'txXfrm'))
   if tx is None: continue # Old image-caption output is not an independent txXfrm frame.
   shape=node.find('./'+tag(DSP,'spPr')+'/'+tag(A,'xfrm')); body=node.find('./'+tag(DSP,'txBody'))
   off=tx.find('./'+tag(A,'off')); ext=tx.find('./'+tag(A,'ext')); props=body.find('./'+tag(A,'bodyPr'))
   parents={child:parent for parent in root.iter() for child in parent}
   parent=parents[node]; child_off=parent.find('./'+tag(DSP,'grpSpPr')+'/'+tag(A,'xfrm')+'/'+tag(A,'chOff'))
   ox=int(child_off.get('x')) if child_off is not None else 0; oy=int(child_off.get('y')) if child_off is not None else 0
   box=dict(x=(int(off.get('x'))-ox)/12700,y=(int(off.get('y'))-oy)/12700,width=int(ext.get('cx'))/12700,height=int(ext.get('cy'))/12700)
   assert all(math.isclose(frame['box'][key],value,abs_tol=1e-8) for key,value in box.items())
   rotation=(int(shape.get('rot','0'))+int(tx.get('rot','0'))+(int(props.get('rot','0')) if props is not None else 0))/60000%360
   assert math.isclose(frame['transform']['rotation'],rotation,abs_tol=1e-8)
   text='\n'.join(''.join((run.find('./'+tag(A,'t')).text or '') if run.tag==tag(A,'r') else '\n' if run.tag==tag(A,'br') else '' for run in para) for para in body.findall('./'+tag(A,'p')))
   assert frame['text']==text
   checked_frames.append(dict(elementIndex=index,text=text,box=box,rotation=rotation))
  expected_frames=0
  for preview in report.get('cachePreviews',[]):
   if not preview.get('cachePart') or preview['status']=='placeholder':continue
   for node in tree(preview['cachePart']).iter(tag(DSP,'sp')):
    if node.find('./'+tag(DSP,'txXfrm')) is not None and node.find('./'+tag(DSP,'txBody')) is not None: expected_frames+=1
  assert len(checked_frames)==expected_frames # These pinned source probes have complete unflipped tx rectangles.
  cases.append(dict(name=sample['name'],source=sample['source'],gitBlob=sample['gitBlob'],sha256=sample['sha256'],bytes=len(data),exactNoOp=True,independentZipXml=True,layouts=layouts,cacheStatuses=observed,resourceParts=len(inv['parts']),imageParts=sum('image' in p['roles'] for p in inv['parts']),dataGraphs=graph_evidence,cacheMappings=cache_mapping_evidence,**(dict(cachePreviews=previews,independentTextFrames=checked_frames) if 'cachePreviews' in report else {})))
print(json.dumps(dict(pythonVersion=sys.version.split()[0],cases=cases)))
`;
const checked = spawnSync("python3", ["-c", python, directory], { encoding: "utf8", timeout: 120000, maxBuffer: 4 * 1024 * 1024 });
assert.equal(checked.status, 0, `${checked.error?.message ?? ""}\n${checked.stderr}`);
const independent = JSON.parse(checked.stdout), cases = independent.cases;
const evidence = { profile: "smartart-source-inventory-v1", repositoryCommit: provenance.commit, retainedDirectory: directory,
  runtime: { node: process.version, python: independent.pythonVersion, platform: process.platform, architecture: process.arch },
  command: `node scripts/inspect-smartart-inventory.mjs ${directory}${previewEnabled ? " --preview" : ""}${outputIndex >= 0 ? ` --output ${evidenceFile}` : ""}`, previewEnabled, cases,
  limitations: ["Source inspection, exact no-op and optional bounded preview probes only; no native application opening/editing/playback or visual-fidelity certification.", "Third-party source files remain local experiment inputs; redistribution/fixture licensing admission is pending.", "These pinned Transitional source cases do not establish complete producer/dialect/family support or cache freshness."] };
await writeFile(evidenceFile, JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify(cases.map(c => ({ name: c.name, layouts: c.layouts, caches: c.cacheStatuses, resources: c.resourceParts, images: c.imageParts, exactNoOp: c.exactNoOp, independentZipXml: c.independentZipXml })), null, 2));
