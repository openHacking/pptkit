"""Native Ω-only styles. Neither the no-edit nor append acceptance rules change."""
import copy, hashlib, importlib.util, json, math, re, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image, ImageDraw

spec=importlib.util.spec_from_file_location('append',Path(__file__).with_name('native-notes-append.py'))
append=importlib.util.module_from_spec(spec)
spec.loader.exec_module(append)
NS=append.NS
STAGES={'bold':{'b':'1'},'italic':{'b':'0','i':'1'},'regular':{'b':'0','i':'0'}}
MATRIX=r'BT\s+(?:[-\d.]+\s+Tc\s+)?([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm\s+/([A-Za-z0-9]+)\s+1\s+Tf'

def inspect_package(file,directory,expected,stage):
    # First certify the immutable append source with its original strict oracle.
    append.inspect_package(directory/'append-round3.pptx',directory,expected)
    rows=[];ancillary={}
    with zipfile.ZipFile(directory/'append-round3.pptx') as source,zipfile.ZipFile(file) as z:
        assert len(z.namelist())==len(set(z.namelist()))
        assert set(z.namelist())==set(source.namelist())
        for part in source.namelist():
            before,after=source.read(part),z.read(part)
            if part in append.ANCILLARY:
                ancillary[part]={'sourceSha256':hashlib.sha256(before).hexdigest(),'nativeSha256':hashlib.sha256(after).hexdigest(),'sourceXml':before.decode(),'nativeXml':after.decode()}
            elif part.startswith('ppt/notesSlides/notesSlide') and part.endswith('.xml'):
                index=int(Path(part).stem.removeprefix('notesSlide'))
                target=ET.fromstring(before)
                if index in append.TARGETS:
                    p=append.paragraph(target)
                    r=p.findall('a:r',NS)[1].find('a:rPr',NS)
                    r.attrib.update(STAGES[stage]);r.set('dirty','0');r.set('err','1')
                    p.find('a:endParaRPr',NS).attrib.update(STAGES[stage])
                actual=ET.fromstring(after)
                assert append.notes.tree(actual)==append.notes.tree(target),'style tree changed outside named Ω/ending flags: '+part
                rows.append({'slide':index,'semanticSha256':append.notes.digest(actual),'omegaProperties':dict(append.paragraph(actual).findall('a:r',NS)[1].find('a:rPr',NS).attrib) if index in append.TARGETS else None})
            else:assert after==before,'unrelated style payload changed: '+part
    return {'sha256':append.notes.sha(file),'notes':sorted(rows,key=lambda r:r['slide']),'ancillaryRecords':ancillary}

def expected_face(variant,index,stage):
    family='Arial' if variant=='A' or 10<=index<=12 else 'Georgia'
    return {'Arial':{'bold':'Arial-BoldMT','italic':'Arial-ItalicMT','regular':'ArialMT'},
            'Georgia':{'bold':'Georgia-Bold','italic':'Georgia-Italic','regular':'Georgia'}}[family][stage]

def inspect_pdf(file,directory,expected,stage,reference=None):
    rows=[]
    with fitz.open(directory/'append-round3-notes.pdf') as source,fitz.open(file) as doc:
        assert len(source)==len(doc)==15
        prior=fitz.open(reference) if reference else None
        try:
            for i,(page,old) in enumerate(zip(doc,source),1):
                assert tuple(page.rect)==tuple(old.rect)==(0,0,595,842)
                traces,before=page.get_texttrace(),old.get_texttrace()
                current=[t for t in traces if t['bbox'][1]>410 and ''.join(chr(c[0])for c in t['chars'])=='Ω'] if i in append.TARGETS else []
                original=[t for t in before if t['bbox'][1]>410 and ''.join(chr(c[0])for c in t['chars'])=='Ω'] if i in append.TARGETS else []
                assert len(current)==len(original)==(1 if i in append.TARGETS else 0),'style append trace count/fill+stroke'
                assert [append.strip_sequence(t)for t in traces if not any(t is a for a in current)]==[append.strip_sequence(t)for t in before if not any(t is a for a in original)],'original or neighbor trace changed'
                pix,oldpix=page.get_pixmap(alpha=False),old.get_pixmap(alpha=False)
                image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
                oldimage=Image.frombytes('RGB',(oldpix.width,oldpix.height),oldpix.samples)
                matrices=[];ink=None
                if current:
                    t=current[0];face=expected_face(expected['variant'],i,stage)
                    assert t['font']==face and t['type']==0 and t['linewidth'] is None and abs(t['size']-41)<.001
                    assert len(t['chars'])==1 and t['chars'][0][1]!=0
                    assert t['chars'][0][2]==original[0]['chars'][0][2],'Ω origin moved'
                    resources={f[4]:f[3].split('+')[-1] for f in page.get_fonts()}
                    stream=b''.join(doc.xref_stream(x)for x in page.get_contents()).decode('latin1')
                    for match in re.findall(MATRIX,stream):
                        if resources[match[6]]!=face:continue
                        values=[float(v)for v in match[:6]]
                        # Header/labels use the same Arial resource at 17pt.
                        if abs(values[0]-170.8333)>.0001:continue
                        assert values[1]==values[2]==0 and values[0]==values[3]>0,'synthesized shear/scale'
                        matrices.append(values)
                    assert len(matrices)==1,'missing Ω text matrix'
                    b=t['bbox'];bounds=(math.floor(b[0])-1,math.floor(b[1])-1,math.ceil(b[2])+1,math.ceil(b[3])+1)
                    pixels=image.crop(bounds).tobytes();ink=sum(min(pixels[j:j+3])<200 for j in range(0,len(pixels),3));assert ink>20,'style extractable but invisible'
                    o=original[0]['bbox'];union=(min(bounds[0],math.floor(o[0])-1),min(bounds[1],math.floor(o[1])-1),max(bounds[2],math.ceil(o[2])+1),max(bounds[3],math.ceil(o[3])+1))
                    ImageDraw.Draw(image).rectangle(union,fill='white');ImageDraw.Draw(oldimage).rectangle(union,fill='white')
                assert image.tobytes()==oldimage.tobytes(),'outside-Ω style raster changed'
                if stage=='regular' or not current:assert pix.samples==oldpix.samples,'regular/control full raster not restored'
                append.notes.table.mixed.font_resources(doc,page)
                if prior:
                    assert traces==prior[i-1].get_texttrace(),'later style trace drift'
                    assert pix.samples==prior[i-1].get_pixmap(alpha=False).samples,'later style raster drift'
                rows.append({'slide':i,'actualFont':current[0]['font'] if current else None,'omegaTrace':current,'textMatrices':matrices,'syntheticStroke':False,'syntheticShear':False,'visibleOmegaInk':ink,'outsideOmegaRasterExact':True,'fullRasterRestored':stage=='regular' or not current,'rasterSha256':hashlib.sha256(pix.samples).hexdigest()})
        finally:
            if prior:prior.close()
    return {'sha256':append.notes.sha(file),'pages':rows}

def inspect(directory):
    expected=json.loads((directory/'expected.json').read_text());stages=[]
    for stage in STAGES:
        rounds=[]
        for number in [1,2,3]:
            name=f'{stage}-round{number}'
            package=inspect_package(directory/(name+'.pptx'),directory,expected,stage)
            if rounds:assert package['notes']==rounds[0]['package']['notes']
            pdf=inspect_pdf(directory/(name+'-notes.pdf'),directory,expected,stage,directory/(stage+'-round1-notes.pdf') if number>1 else None)
            rounds.append({'name':name,'package':package,'pdf':pdf})
        stages.append({'stage':stage,'rounds':rounds})
    result={'variant':expected['variant'],'stages':stages,'nativeNotesPages':135,'limits':'Ω-only native styles; no original body style operations, theme response, complete notes page or source style writer.'}
    (directory/'styles-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    return result

if __name__=='__main__':
    result=inspect(Path(sys.argv[1]));print(json.dumps({'variant':result['variant'],'nativeNotesPages':135,'stages':3,'rounds':9}))
