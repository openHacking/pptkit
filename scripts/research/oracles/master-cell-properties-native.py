"""Actual cell-format regression controls; broader first-save chart semantics remain separate."""
import copy
import importlib.util
import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as E
import fitz
HERE=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('cell_artwork_primitives',HERE/'master-artwork-native.py')
a=importlib.util.module_from_spec(spec);spec.loader.exec_module(a)
A,P,R,NS=a.A,a.P,a.R,a.NS
C='http://schemas.openxmlformats.org/drawingml/2006/chart'
sha=a.sha

def full(parts):
    result={}
    for name,body in parts.items():
        if not name.endswith(('.xml','.rels')):result[name]=sha(body);continue
        root=E.fromstring(body)
        if name=='docProps/core.xml':
            for child in root:
                if child.tag.rsplit('}',1)[-1] in ['modified','lastPrinted']:
                    datetime.fromisoformat(child.text.replace('Z','+00:00'));child.text='clock'
        elif name=='docProps/app.xml':
            for child in root:
                if child.tag.endswith('}TotalTime'):assert child.text.isdigit();child.text='time'
        result[name]=a.n.plain(root)
    return result

def tables(parts):
    result={}
    for name in parts:
        if name.startswith('ppt/slideMasters/') and name.endswith('.xml'):
            root=a.n.root(parts,name);owner=root.find('p:cSld',NS).get('name')
            if owner not in ['red','blue']:continue
            native=root.findall('.//a:tbl',NS);assert len(native)==1
            table=native[0];rows=table.findall('a:tr',NS)
            assert [r.get('h') for r in rows]==['508000','508000']
            assert [c.get('w') for c in table.findall('a:tblGrid/a:gridCol',NS)]==['1905000','1905000']
            cells=[cell for row in rows for cell in row.findall('a:tc',NS)];assert len(cells)==4
            assert cells[0].get('gridSpan')=='2' and cells[1].get('hMerge')=='1'
            result[owner]=[]
            for i,cell in enumerate(cells):
                pr=cell.find('a:tcPr',NS);attrs=dict(pr.attrib)
                if attrs.get('anchor')=='t':del attrs['anchor']
                children=[c.tag.rsplit('}',1)[-1] for c in pr]
                if i!=1:
                    assert {k:attrs.get(k) for k in ['marL','marR','marT','marB']}==dict.fromkeys(['marL','marR','marT','marB'],'50800'),'Lost authored cell margin'
                    assert children==['lnL','lnR','lnT','lnB','solidFill'],'Invalid border/fill order or lost border'
                    assert all(pr.find('a:'+s+'/a:noFill',NS) is not None for s in ['lnL','lnR','lnT','lnB']),'Unexpected visible table border'
                    assert cell.find('a:txBody/a:bodyPr',NS).attrib=={},'Cell frame misplaced in bodyPr'
                result[owner].append((cell.attrib,attrs,[a.n.plain(c) for c in pr]))
    assert set(result)=={'red','blue'}
    return result

def owners(parts,source):
    presentation=a.n.root(parts,'ppt/presentation.xml')
    size=presentation.find('p:sldSz',NS);assert size.get('cx')=='9144000' and size.get('cy')=='6858000'
    refs=presentation.findall('p:sldIdLst/p:sldId',NS);edges=a.n.relations(parts,'ppt/presentation.xml');assert len(refs)==4
    for i,ref in enumerate(refs):
        slide=edges[ref.get('{'+R+'}id')]['target'];root=a.n.root(parts,slide)
        assert root.get('showMasterSp','1')==('0' if i==1 else '1'),'Lost hidden artwork declaration'
        layout=a.n.target(parts,slide,'slideLayout');master=a.n.target(parts,layout,'slideMaster')
        assert a.n.root(parts,master).find('p:cSld',NS).get('name')==('blue' if i==3 else 'red')
        assert root.findall('.//a:tbl',NS)==[] and root.findall('.//{'+C+'}chart')==[],'Shared objects flattened onto slide'
    charts=set();workbooks=set()
    for part in source:
        if part.startswith('ppt/slideMasters/') and part.endswith('.xml'):
            before=a.n.root(source,part);after=a.n.root(parts,part)
            assert [a.n.plain(n) for n in before.findall('.//p:cxnSp',NS)]==[a.n.plain(n) for n in after.findall('.//p:cxnSp',NS)],'Changed native connector or part-local binding'
            for edge in a.n.relations(source,part).values():
                if not edge['type'].endswith('/chart'):continue
                chart=edge['target'];assert chart in parts;charts.add(chart)
                b=a.n.root(source,chart);n=a.n.root(parts,chart)
                for tag in ['numCache','strCache']:
                    assert [a.n.plain(x) for x in b.iter('{'+C+'}'+tag)]==[a.n.plain(x) for x in n.iter('{'+C+'}'+tag)],'Changed chart cache'
                old=a.n.target(source,chart,'package');new=a.n.target(parts,chart,'package');workbooks.add(new)
                assert parts[new]==source[old],'Changed embedded workbook payload'
    assert len(charts)==6 and len(workbooks)==6
    assert {p for p in parts if p.startswith('ppt/charts/') and p.endswith('.xml')}==charts
    assert {p for p in parts if p.startswith('ppt/embeddings/')}==workbooks
    for part in parts:
        if part.endswith('.rels'):
            root=E.fromstring(parts[part]);base='' if part=='_rels/.rels' else str(Path(part).parent.parent/Path(part).name.removesuffix('.rels'))
            for edge in a.n.relations(parts,base).values():
                if edge['mode']=='Internal':assert edge['target'] in parts,'Broken internal relationship'

def pages(file,baseline=None,white=0):
    doc=fitz.open(file);assert len(doc)==4
    old=fitz.open(baseline) if baseline else None;result=[]
    for i,page in enumerate(doc):
        text=page.get_text().strip().splitlines();assert text[-1]=='PAGE '+str(i)
        if i==1:assert text==['PAGE 1']
        else:
            owner='blue' if i==3 else 'red'
            for token in ['MASTER '+owner,owner+' shared table','Left '+owner,'Right '+owner,owner+' bar',owner+' line',owner+' pie']:assert text.count(token)==1
            assert not any(re.search(r'\b'+('blue' if owner=='red' else 'red')+r'\b',s) for s in text)
        drawings=page.get_drawings();lines=[d for d in drawings if d['type']=='s' and d['color']==(1,1,1)]
        assert len(lines)==(0 if i==1 else white),'Unexpected actual native table borders'
        record=a.pdf_page_record(doc,i)
        if old:a.same(record,a.pdf_page_record(old,i))
        result.append(record)
    return result

def main():
    manifest=json.loads(Path('docs/testing/master-advanced-evidence.json').read_text());prior=json.loads(Path('docs/testing/master-advanced-before-cell-fix-evidence.json').read_text())
    directory=Path(manifest['nativeDirectory']);old=Path(prior['nativeDirectory']);source=a.n.payload(directory/'input.pptx')
    assert sha((directory/'input.pptx').read_bytes())==manifest['sourceSha256']
    old_source=a.n.payload(old/'input.pptx')
    assert sha((old/'input.pptx').read_bytes())==prior['sourceSha256']
    assert source.keys()==old_source.keys()
    source_delta=[]
    for part in source:
        if source[part]==old_source[part]:continue
        source_delta.append(part)
        before_root=a.n.root(old_source,part)
        for cell in before_root.iter('{'+A+'}tc'):
            body=cell.find('a:txBody/a:bodyPr',NS);pr=cell.find('a:tcPr',NS)
            for old_key,new_key in [('lIns','marL'),('rIns','marR'),('tIns','marT'),('bIns','marB'),('anchor','anchor')]:
                if old_key in body.attrib:pr.set(new_key,body.attrib.pop(old_key))
            lines=[c for c in pr if c.tag.rsplit('}',1)[-1] in ['lnL','lnR','lnT','lnB']]
            pr[:]=lines+[c for c in pr if c not in lines]
        assert a.n.plain(before_root)==a.n.plain(a.n.root(source,part)),'Unrelated producer payload changed'
    assert source_delta==['ppt/slideMasters/slideMaster2.xml','ppt/slideMasters/slideMaster3.xml']
    expected=tables(source);owners(source,source)
    # Preserve the actual failure and isolate the sequence-only and frame-only effects.
    before=pages(old/'first.pdf',white=6)
    for i in [1,2,3]:pages(old/f'advanced-round{i}.pdf',old/'first.pdf',white=6)
    sequence=pages(old/'sequence-control.pdf');control=pages(old/'cell-frame-control.pdf')
    first=pages(directory/'first.pdf',old/'cell-frame-control.pdf')
    def origin(file):
        doc=fitz.open(file);return next(s['origin'] for b in doc[0].get_text('dict')['blocks'] for l in b.get('lines',[]) for s in l['spans'] if s['text']=='red shared table')
    printed=fitz.open(directory/'first.pdf')
    frame=[d['rect'] for d in printed[0].get_drawings() if d['fill']==(1,1,1)][-1]
    actual_origin=origin(directory/'first.pdf')[0]
    assert abs(actual_origin-(frame.x0+(20+4)*frame.width/720))<0.0001,'Actual inset does not match authored source frame'
    delta=origin(old/'sequence-control.pdf')[0]-actual_origin
    assert abs(delta-(7.2-4)*frame.width/720)<0.0001,'Declared 4pt inset did not replace actual default 7.2pt'
    rounds=[];native_base=None
    for i in [1,2,3]:
        file=directory/f'cell-fixed-round{i}.pptx';parts=a.n.payload(file)
        assert tables(parts)==expected;owners(parts,source)
        current=full(parts)
        if native_base is None:native_base=current
        else:a.same(current,native_base)
        for phase in ['saved','reopened']:
            ax=(directory/f'cell-fixed-round{i}-{phase}.ax.txt').read_text();assert 'URL: file://'+str(file) in ax and not re.search('发现.*内容有问题|修复',ax)
        assert str(file) not in (directory/f'cell-fixed-round{i}-closed.ax.txt').read_text()
        rounds.append({'round':i,'file':str(file),'sha256':sha(file.read_bytes()),'pages':pages(directory/f'cell-fixed-round{i}.pdf',directory/'first.pdf')})
    opened=(directory/'first-open.ax.txt').read_text();assert 'URL: file://'+str(directory/'input.pptx') in opened and not re.search('发现.*内容有问题|修复',opened)
    assert '16.104 (25121423)' in (directory/'powerpoint-version.ax.txt').read_text()
    # First native chart defaults are recorded in full rather than excluded by broad normalization.
    original=full(source);first_native=a.n.payload(directory/'cell-fixed-round1.pptx');changed={p:{'source':original.get(p),'native':native_base.get(p)} for p in sorted(original.keys()|native_base.keys()) if original.get(p)!=native_base.get(p)}
    negative=Path(tempfile.mkdtemp(prefix='negative-',dir=directory));bad=[]
    for name in ['margin','border','order','cache','workbook','hidden']:
        parts=copy.copy(first_native)
        if name in ['margin','border','order']:
            part='ppt/slideMasters/slideMaster2.xml';root=a.n.root(parts,part);pr=root.find('.//a:tcPr',NS)
            if name=='margin':pr.set('marL','91440')
            elif name=='border':pr.find('a:lnL/a:noFill',NS).tag='{'+A+'}solidFill'
            else:
                fill=pr[-1];pr.remove(fill);pr.insert(0,fill)
            parts[part]=E.tostring(root,encoding='utf-8')
        elif name=='cache':
            part='ppt/charts/chart1.xml';root=a.n.root(parts,part);next(root.iter('{'+C+'}numCache')).find('{'+C+'}pt/{'+C+'}v').text='99';parts[part]=E.tostring(root,encoding='utf-8')
        elif name=='workbook':
            part=a.n.target(parts,'ppt/charts/chart1.xml','package');parts[part]+=b'changed'
        else:
            part='ppt/slides/slide2.xml';root=a.n.root(parts,part);root.attrib.pop('showMasterSp');parts[part]=E.tostring(root,encoding='utf-8')
        file=negative/(name+'.pptx')
        with zipfile.ZipFile(file,'w',compression=zipfile.ZIP_DEFLATED) as z:
            for part,body in parts.items():z.writestr(part,body)
        try:assert tables(parts)==expected;owners(parts,source);a.same(full(parts),native_base)
        except AssertionError as error:bad.append({'file':str(file),'sha256':sha(file.read_bytes()),'reason':str(error)})
        else:raise AssertionError('Accepted corrupt '+name)
    deps=[Path(__file__),HERE/'master-artwork-native.py',HERE/'multi-master-native.py',HERE/'multi-master-operations.py']
    artifacts=[{'file':str(f),'bytes':f.stat().st_size,'sha256':sha(f.read_bytes())} for folder in [old,directory] for f in sorted(folder.rglob('*')) if f.is_file()]
    result={'schemaVersion':1,'verified':True,'recordedAt':datetime.now(timezone.utc).isoformat(),'directory':str(directory),'sourceSha256':manifest['sourceSha256'],'application':'PowerPoint 16.104 (25121423)','firstPages':first,'rounds':rounds,'negatives':bad,'sourceChangedParts':source_delta,'historicalSourceSha256':prior['sourceSha256'],'firstSaveFullPackageDelta':changed,'actualInsetCorrectionPdfPoints':delta,'artifacts':artifacts,'dependencies':[{'file':str(f),'sha256':sha(f.read_bytes())} for f in deps],'scope':'Cell margin/anchor placement and border-before-fill repair. Actual old failure/isolated controls/new native three rounds; full later package stability, chart caches/workbook closure/native connector/owner/visibility controls and all-page raster/trace/drawing/image/font stability. Full source-to-first-save chart style equivalence, native shared editing and general advanced master admission remain unaccepted.'}
    Path('docs/testing/master-cell-properties-native-evidence.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'verified':True,'firstPages':4,'reopenedPages':12,'zipNegatives':len(bad),'artifacts':len(artifacts),'firstSaveChangedParts':len(changed)}))
if __name__=='__main__':main()
