"""Actual master-heading edit: exact owner change, physical bystanders and three reopen rounds."""
import copy
import importlib.util
import json
import math
import re
import struct
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as E
import fitz
HERE=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('shared_cell_primitives',HERE/'master-cell-properties-native.py');c=importlib.util.module_from_spec(spec);spec.loader.exec_module(c)
a=c.a;A,P,NS=c.A,c.P,c.NS;sha=c.sha

def font_tables(data):
    assert data[:4] in [b'\x00\x01\x00\x00',b'true']
    result={}
    for i in range(struct.unpack_from('>H',data,4)[0]):
        tag,checksum,start,length=struct.unpack_from('>4sIII',data,12+16*i);assert start+length<=len(data) and tag not in result
        result[tag]=data[start:start+length]
    return result

def cmap_signature(body):
    # This captured Mac producer uses one format-6 subtable. No generic font rewrite is admitted.
    assert struct.unpack_from('>HHHHI',body)==(0,1,1,0,12)
    fmt,length,language,first,count=struct.unpack_from('>HHHHH',body,12)
    assert fmt==6 and length==10+2*count and len(body)==12+length+2 and body[-2:]==b'\0\0'
    glyphs=struct.unpack_from('>'+str(count)+'H',body,22)
    return (fmt,length,language,first,count,sorted(glyphs))

def fonts_equal(before,after,page):
    left=before[page].get_fonts();right=after[page].get_fonts();assert left==right,'Font resources changed'
    records=[]
    for f in left:
        x=before.extract_font(f[0])[3];y=after.extract_font(f[0])[3]
        tx,ty=font_tables(x),font_tables(y);assert tx.keys()==ty.keys()
        assert {k:v for k,v in tx.items() if k!=b'cmap'}=={k:v for k,v in ty.items() if k!=b'cmap'},'Font outline/metrics/non-cmap table changed'
        assert cmap_signature(tx[b'cmap'])==cmap_signature(ty[b'cmap']),'Unexpected cmap topology or glyph mapping set'
        offset,length=next((start,length) for tag,_,start,length in [struct.unpack_from('>4sIII',x,12+16*i) for i in range(struct.unpack_from('>H',x,4)[0])] if tag==b'cmap')
        assert len(x)==len(y) and x[:offset]==y[:offset] and x[offset+length:]==y[offset+length:],'Font stream bytes outside cmap changed'
        records.append({'resource':f[4],'name':f[3],'beforeSha256':sha(x),'afterSha256':sha(y),'changedTables':[k.decode() for k in tx if tx[k]!=ty[k]],'nonCmapTablesExact':True,'outsideCmapStreamBytesExact':True,'format6GlyphMappingPermutationOnly':True})
    return records

def physical(before_file,after_file):
    before=fitz.open(before_file);after=fitz.open(after_file);assert len(before)==len(after)==4;records=[]
    for i in range(4):
        x,y=before[i],after[i];assert x.rect==y.rect and x.get_drawings()==y.get_drawings() and x.get_image_info(hashes=True)==y.get_image_info(hashes=True)
        fonts=fonts_equal(before,after,i);tx,ty=x.get_texttrace(),y.get_texttrace()
        if i in [1,3]:
            assert x.get_text()==y.get_text() and tx==ty
            assert x.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples==y.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples,'Unchanged page pixels changed'
        else:
            assert x.get_text().replace('MASTER red\n','MASTER red 2\n',1)==y.get_text()
            assert tx[1:]==ty[1:] and len(tx)==len(ty)
            assert {k:v for k,v in tx[0].items() if k not in ['chars','bbox']}=={k:v for k,v in ty[0].items() if k not in ['chars','bbox']}
            assert tx[0]['chars']==ty[0]['chars'][:len(tx[0]['chars'])],'Original heading glyphs moved/changed'
            assert [chr(ch[0]) for ch in ty[0]['chars'][len(tx[0]['chars']):]]==[' ','2']
            # Compare every pixel outside the edited heading's old/new text bounds; all other artwork stays exact.
            old=x.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False);new=y.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
            box=fitz.Rect(tx[0]['bbox'])|fitz.Rect(ty[0]['bbox']);box=box+(-1,-1,1,1)
            x0=max(0,math.floor(box.x0*2));x1=min(old.width,math.ceil(box.x1*2));y0=max(0,math.floor(box.y0*2));y1=min(old.height,math.ceil(box.y1*2))
            aa,bb=old.samples,new.samples;stride=old.stride
            for row in range(old.height):
                start=row*stride;end=start+stride
                if y0<=row<y1:assert aa[start:start+x0*3]==bb[start:start+x0*3] and aa[start+x1*3:end]==bb[start+x1*3:end]
                else:assert aa[start:end]==bb[start:end],'Bystander artwork pixels changed'
        records.append({'page':i,'editedHeading':i in [0,2],'fonts':fonts,'allOtherTraceDrawingsImagesPixelsExact':True})
    return records

def expect_edit(parts,source):
    expected=copy.copy(source);part='ppt/slideMasters/slideMaster2.xml';root=a.n.root(source,part)
    shape=next(s for s in root.findall('p:cSld/p:spTree/p:sp',NS) if s.find('p:nvSpPr/p:cNvPr',NS).get('name')=='red-heading')
    paragraph=shape.find('p:txBody/a:p',NS);run=paragraph.find('a:r',NS);assert run.find('a:t',NS).text=='MASTER red'
    run.find('a:t',NS).text='MASTER red 2';run.find('a:rPr',NS).set('dirty','0');paragraph.find('a:endParaRPr',NS).set('dirty','0');expected[part]=E.tostring(root,encoding='utf-8')
    core=a.n.root(source,'docProps/core.xml');revision=core.find('{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}revision');assert revision.text=='1';revision.text='2';expected['docProps/core.xml']=E.tostring(core,encoding='utf-8')
    a.same(c.full(parts),c.full(expected));c.tables(parts);c.owners(parts,source)

def main():
    prepared=json.loads(Path('docs/testing/master-shared-evidence.json').read_text());directory=Path(prepared['nativeDirectory']);source=a.n.payload(directory/'input.pptx');binding=json.loads((directory/'input-binding.json').read_text())
    assert prepared['sourceSha256']==binding['sha256']
    assert sha((directory/'input.pptx').read_bytes())==binding['sha256']==sha(Path(binding['sourceFile']).read_bytes())
    prior=json.loads(Path('docs/testing/master-cell-properties-native-evidence.json').read_text());assert binding['sourceFile']==prior['rounds'][2]['file'] and binding['sha256']==prior['rounds'][2]['sha256']
    before=fitz.open(directory/'before.pdf');baseline=fitz.open(Path(binding['sourceFile']).with_suffix('.pdf'))
    for i in range(4):a.same(a.pdf_page_record(before,i),a.pdf_page_record(baseline,i))
    physical_records=physical(directory/'before.pdf',directory/'after-edit.pdf');rounds=[];native=None
    fresh_file=directory/'fresh-core-edit.pptx';assert sha(fresh_file.read_bytes())==prepared['freshEditedSha256']
    fresh_ax=(directory/'fresh-core-edit-open.ax.txt').read_text();assert 'URL: file://'+str(fresh_file) in fresh_ax and not re.search('发现.*内容有问题|修复',fresh_ax)
    fresh=fitz.open(directory/'fresh-core-edit.pdf');edited=fitz.open(directory/'after-edit.pdf');assert len(fresh)==len(edited)==4
    for i in range(4):a.same(a.pdf_page_record(fresh,i),a.pdf_page_record(edited,i))
    for i in [1,2,3]:
        file=directory/f'shared-round{i}.pptx';parts=a.n.payload(file);expect_edit(parts,source)
        if native is None:native=c.full(parts)
        else:a.same(c.full(parts),native)
        edited=fitz.open(directory/'after-edit.pdf');reopened=fitz.open(directory/f'shared-round{i}.pdf')
        pages=[]
        for p in range(4):
            record=a.pdf_page_record(reopened,p);a.same(record,a.pdf_page_record(edited,p));pages.append(record)
        for phase in ['saved','reopened']:
            ax=(directory/f'shared-round{i}-{phase}.ax.txt').read_text();assert 'URL: file://'+str(file) in ax and not re.search('发现.*内容有问题|修复',ax)
        assert str(file) not in (directory/f'shared-round{i}-closed.ax.txt').read_text()
        rounds.append({'round':i,'file':str(file),'sha256':sha(file.read_bytes()),'pages':pages})
    assert '16.104 (25121423)' in (directory/'powerpoint-version.ax.txt').read_text()
    out=Path(tempfile.mkdtemp(prefix='negative-',dir=directory));bad=[];positive=a.n.payload(directory/'shared-round1.pptx')
    for name in ['blue-master-changed','local-copy','hidden-flag','target-format']:
        parts=copy.copy(positive)
        if name=='blue-master-changed':part='ppt/slideMasters/slideMaster3.xml';root=a.n.root(parts,part);next(root.iter('{'+A+'}t')).text='MASTER blue 2'
        elif name=='local-copy':
            part='ppt/slides/slide1.xml';root=a.n.root(parts,part);master=a.n.root(parts,'ppt/slideMasters/slideMaster2.xml');shape=copy.deepcopy(master.find('p:cSld/p:spTree/p:sp',NS));root.find('p:cSld/p:spTree',NS).append(shape)
        elif name=='hidden-flag':part='ppt/slides/slide2.xml';root=a.n.root(parts,part);root.attrib.pop('showMasterSp')
        else:part='ppt/slideMasters/slideMaster2.xml';root=a.n.root(parts,part);root.find('p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr',NS).set('sz','1900')
        parts[part]=E.tostring(root,encoding='utf-8');file=out/(name+'.pptx')
        with zipfile.ZipFile(file,'w',compression=zipfile.ZIP_DEFLATED) as z:
            for part,body in parts.items():z.writestr(part,body)
        try:expect_edit(a.n.payload(file),source)
        except AssertionError as error:bad.append({'kind':'zip','file':str(file),'sha256':sha(file.read_bytes()),'reason':str(error)})
        else:raise AssertionError('Accepted corrupt '+name)
    for name in ['covered-table','swapped-pages','same-raster-font-advance']:
        doc=fitz.open(directory/'after-edit.pdf');file=out/(name+'.pdf')
        if name=='covered-table':doc[0].draw_rect(doc[0].search_for('red shared table')[0],color=(1,1,1),fill=(1,1,1),overlay=True)
        elif name=='swapped-pages':doc.select([1,0,2,3])
        else:
            f=doc[0].get_fonts()[0];program=doc.extract_font(f[0])[3];gid=doc[0].get_texttrace()[0]['chars'][0][1];altered=a.m.alter_font_advance(program,gid)
            streams=[x for x in range(1,doc.xref_length()) if doc.xref_is_stream(x) and doc.xref_stream(x)==program];assert len(streams)==1;doc.update_stream(streams[0],altered)
        doc.save(file)
        if name=='same-raster-font-advance':
            positive_pdf=fitz.open(directory/'after-edit.pdf');negative_pdf=fitz.open(file);assert all(x.get_pixmap().samples==y.get_pixmap().samples for x,y in zip(positive_pdf,negative_pdf))
        try:physical(directory/'before.pdf',file)
        except AssertionError as error:bad.append({'kind':'pdf','file':str(file),'sha256':sha(file.read_bytes()),'reason':str(error),'sameRaster':name=='same-raster-font-advance'})
        else:raise AssertionError('Accepted corrupt '+name)
    artifacts=[{'file':str(f),'bytes':f.stat().st_size,'sha256':sha(f.read_bytes())} for f in sorted(directory.rglob('*')) if f.is_file()]
    deps=[Path(__file__),HERE/'master-cell-properties-native.py',HERE/'master-artwork-native.py',HERE/'multi-master-native.py',HERE/'multi-master-operations.py']
    result={'schemaVersion':1,'recordedAt':datetime.now(timezone.utc).isoformat(),'verified':True,'application':'PowerPoint 16.104 (25121423)','directory':str(directory),'sourceSha256':binding['sha256'],'freshCoreEditedSha256':prepared['freshEditedSha256'],'freshCoreFirstPrintExactToNativeEdit':True,'baselineNativeEvidence':'docs/testing/master-cell-properties-native-evidence.json','baselineNativeEvidenceSha256':sha(Path('docs/testing/master-cell-properties-native-evidence.json').read_bytes()),'physicalEdit':physical_records,'rounds':rounds,'negatives':bad,'artifacts':artifacts,'dependencies':[{'file':str(f),'sha256':sha(f.read_bytes())} for f in deps],'scope':'One actual shared red master heading edit; two visible pages change, hidden and blue pages stay exact. All other package content exact except named edited paragraph dirty=0, edit revision 1→2 and valid clocks/TotalTime. PDF heading-font cmap format6 mapping permutation only; all non-cmap tables exact. Three actual save-close-exact-reopen rounds. Table/chart data or formatting editing, master images, general reset/rebinding, source import/writes remain separate.'}
    Path('docs/testing/master-shared-native-evidence.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n');print(json.dumps({'verified':True,'rounds':3,'reopenedPages':12,'zipNegatives':4,'pdfNegatives':3,'artifacts':len(artifacts)}))
if __name__=='__main__':main()
