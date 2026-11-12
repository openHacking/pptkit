"""Ordinary slide PDF gate, separate from Notes printing and blank slide-image placeholders."""
import copy,hashlib,importlib.util,json,math,re,struct,sys,zipfile
from pathlib import Path
import fitz
from PIL import Image

spec=importlib.util.spec_from_file_location('themes',Path(__file__).with_name('native-notes-themes.py'))
themes=importlib.util.module_from_spec(spec);spec.loader.exec_module(themes)
SHA=themes.notes.sha
CM=re.compile(rb'q\s+0\.24\s+0\s+0\s+0\.24\s+([-\d.]+)\s+(-38\.31999|-138\.4)\s+cm\s+BT')

def true_type_bounds(data,glyph):
    """Read only the TrueType head/loca/glyf bounds needed by this named oracle."""
    assert data[:4]in [b'\x00\x01\x00\x00',b'true'] and len(data)>=12
    count=struct.unpack_from('>H',data,4)[0];assert 12+16*count<=len(data)
    tables={}
    for i in range(count):
        tag,_,start,length=struct.unpack_from('>4sIII',data,12+16*i);assert tag not in tables and start+length<=len(data);tables[tag]=data[start:start+length]
    head,loca,glyf,maxp=(tables[tag]for tag in [b'head',b'loca',b'glyf',b'maxp'])
    assert len(head)>=54 and len(maxp)>=6
    units=struct.unpack_from('>H',head,18)[0];mode=struct.unpack_from('>h',head,50)[0];total=struct.unpack_from('>H',maxp,4)[0]
    assert units>0 and mode in [0,1] and 0<=glyph<total
    width=2 if mode==0 else 4;assert len(loca)>=(total+1)*width
    start,end=(struct.unpack_from('>H'if mode==0 else'>I',loca,j*width)[0]*(2 if mode==0 else 1)for j in [glyph,glyph+1])
    assert 0<=start<end<=len(glyf) and end-start>=10
    bounds=struct.unpack_from('>hhhh',glyf,start+2);assert bounds[0]<bounds[2] and bounds[1]<bounds[3]
    return units,bounds

def fonts(doc,page):
    result=[]
    themes.notes.table.mixed.font_resources(doc,page)
    for font in page.get_fonts():
        _,extension,kind,data=doc.extract_font(font[0]);assert data and extension=='ttf'
        result.append({'resource':font[4],'baseFont':font[3],'kind':kind,'embeddedSha256':hashlib.sha256(data).hexdigest()})
    return sorted(result,key=lambda r:r['resource'])

def translated(trace):
    target=copy.deepcopy(trace);size=round(trace['size'])
    assert (trace['font'],size)in [('ArialMT',16),('ArialUnicodeMS',48)]
    delta=-1.20001220703125 if size==16 else -0.959991455078125
    def box(b):return(b[0],b[1]+delta,b[2],b[3]+delta)
    target['bbox']=box(trace['bbox']);target['chars']=tuple((c[0],c[1],(c[2][0],c[2][1]+delta),box(c[3]))for c in trace['chars'])
    return target

def inspect_pdf(file,baseline,controls,reference=None):
    rows=[];glyph_fonts={}
    with fitz.open(file)as doc,fitz.open(baseline)as source:
        assert len(doc)==len(source)==len(controls)
        prior=fitz.open(reference)if reference else None
        try:
            for index,(page,old,c)in enumerate(zip(doc,source,controls),1):
                assert tuple(page.rect)==tuple(old.rect)==(0,0,720,260)
                traces,before=page.get_texttrace(),old.get_texttrace()
                label=''.join(chr(ch[0])for t in before if abs(t['size']-16)<.001 for ch in t['chars'])
                assert label==f"notes-{c['variant']}-{c['id']}",'ordinary slide page identity'
                assert traces==[translated(t)for t in before],'ordinary slide glyph/face/geometry changed beyond named Y offsets'
                assert all(ch[1]!=0 for t in traces for ch in t['chars'])
                resource=fonts(doc,page);assert resource==fonts(source,old),'ordinary slide font resource or embedded program changed'
                pix=page.get_pixmap(alpha=False);image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples);ink=[]
                for trace in traces:
                    b=trace['bbox']
                    if b[2]==b[0]:
                        # Zero-advance Arabic marks have zero-width trace boxes.
                        # Measure their actual embedded glyph outlines, not a
                        # one-pixel column or an invented Unicode hyphen box.
                        matches=[f for f in page.get_fonts()if f[3].split('+')[-1]==trace['font']];assert len(matches)==1
                        data=doc.extract_font(matches[0][0])[3];digest=hashlib.sha256(data).hexdigest()
                        boxes=[]
                        assert trace['dir']==(1,0)
                        for char in trace['chars']:
                            key=(digest,char[1])
                            if key not in glyph_fonts:glyph_fonts[key]=true_type_bounds(data,char[1])
                            units,(x0,y0,x1,y1)=glyph_fonts[key];scale=trace['size']/units
                            x,y=char[2];boxes.append((x+x0*scale,y-y1*scale,x+x1*scale,y-y0*scale))
                        b=(min(b[0]for b in boxes),min(b[1]for b in boxes),max(b[2]for b in boxes),max(b[3]for b in boxes))
                    crop=image.crop((math.floor(b[0])-1,math.floor(b[1])-1,math.ceil(b[2])+1,math.ceil(b[3])+1)).tobytes()
                    count=sum(min(crop[j:j+3])<200 for j in range(0,len(crop),3));assert count>5,'extractable but invisible slide text';ink.append(count)
                if prior:
                    assert traces==prior[index-1].get_texttrace();assert pix.samples==prior[index-1].get_pixmap(alpha=False).samples,'later ordinary slide raster drift'
                # Reverse ONLY the two observed native Y translations in a diagnostic
                # in-memory PDF, then require the entire source raster to match.
                # The original native PDF on disk remains unchanged.
                changed=0
                for xref in page.get_contents():
                    raw=doc.xref_stream(xref)
                    def undo(match):
                        nonlocal changed
                        changed+=1
                        y=b'-39.51999'if match[2]==b'-38.31999'else b'-139.36'
                        return b'q 0.24 0 0 0.24 '+match[1]+b' '+y+b' cm BT'
                    normalized=CM.sub(undo,raw);doc.update_stream(xref,normalized)
                assert changed==sum(abs(t['size']-16)<.001 for t in traces)+1,'unrecognized slide transformation count'
                assert page.get_pixmap(alpha=False).samples==old.get_pixmap(alpha=False).samples,'full slide raster differs after reversing named Y offsets'
                rows.append({'page':index,'id':c['id'],'variant':c['variant'],'fonts':resource,'labelYShift':-1.20001220703125,'bodyYShift':-0.959991455078125,'glyphsAndXExact':True,'fullRasterExactAfterReversingNamedYOffsets':True,'visibleTraceInk':ink,'nativeRasterSha256':hashlib.sha256(pix.samples).hexdigest()})
        finally:
            if prior:prior.close()
    return{'sha256':SHA(file),'pages':rows}

def negatives(directory,controls):
    output=directory/'slide-negative-controls';output.mkdir(exist_ok=True);records=[]
    for mode in ['white-body','font-resource','font-program','wrong-geometry','swap-pages','later-ink']:
        file=output/(mode+'.pdf')
        with fitz.open(directory/'office-round1-slides.pdf')as doc:
            page=doc[0]
            if mode=='white-body':
                trace=next(t for t in page.get_texttrace()if abs(t['size']-48)<.001);page.draw_rect(fitz.Rect(trace['bbox'])+(-1,-1,1,1),color=None,fill=(1,1,1),overlay=True)
            elif mode=='font-resource':doc.xref_set_key(page.get_fonts()[0][0],'BaseFont','/Wrong')
            elif mode=='font-program':
                font=page.get_fonts()[0];kind,value=doc.xref_get_key(font[0],'FontDescriptor/FontFile2');assert kind=='xref';xref=int(value.split()[0]);data=bytearray(doc.xref_stream(xref));data[-1]^=1;doc.update_stream(xref,bytes(data))
            elif mode=='wrong-geometry':
                count=0
                for xref in page.get_contents():
                    raw=doc.xref_stream(xref)
                    def wrong(match):
                        nonlocal count
                        count+=1;return b'q 0.24 0 0 0.24 '+str(float(match[1])+10).encode()+b' '+match[2]+b' cm BT'
                    doc.update_stream(xref,CM.sub(wrong,raw))
                assert count>0
            elif mode=='swap-pages':doc.select([1,0]+list(range(2,len(doc))))
            else:page.draw_rect(fitz.Rect(650,210,680,240),color=None,fill=(0,0,0),overlay=True)
            doc.save(file)
        try:inspect_pdf(file,directory/'baseline-slides.pdf',controls,directory/'office-round1-slides.pdf'if mode=='later-ink'else None)
        except AssertionError as error:records.append({'name':mode,'file':str(file),'sha256':SHA(file),'reason':str(error)})
        else:raise AssertionError('ordinary slide corruption accepted: '+mode)
    return{'rejectedCount':len(records),'controls':records}

def inspect(root):
    manifest=themes.ROOT/'docs/testing/native-notes-themes-evidence.json';notes=json.loads(manifest.read_text());assert Path(notes['root'])==root.resolve();groups=[]
    for group in notes['groups']:
        name=group['group'];directory=root/name;expected=json.loads((directory/'expected.json').read_text())
        controls=json.loads((directory/'controls.json').read_text())if name=='override'else[{**c,'variant':name}for c in expected['controls']]
        baseline_input='native-baseline-round3'if name=='override'else'baseline';baseline_evidence=next(i for i in group['inputs']if i['name']==baseline_input)
        assert SHA(directory/(baseline_input+'.pptx'))==baseline_evidence['sha256']
        rounds=[]
        for number in [1,2,3]:
            stage=f'office-round{number}';input=next(i for i in group['inputs']if i['name']==stage);assert SHA(directory/(stage+'.pptx'))==input['sha256']
            rounds.append({'name':stage,'inputSha256':input['sha256'],'pdf':inspect_pdf(directory/(stage+'-slides.pdf'),directory/'baseline-slides.pdf',controls,directory/'office-round1-slides.pdf'if number>1 else None)})
        groups.append({'group':name,'baseline':{'inputSha256':baseline_evidence['sha256'],'pdfSha256':SHA(directory/'baseline-slides.pdf')},'rounds':rounds,'negatives':negatives(directory,controls)})
    result={'schemaVersion':1,'root':str(root.resolve()),'application':notes['application'],'notesEvidenceSha256':SHA(manifest),'groups':groups,'total':{'nativeSlidePDFPages':180,'preThemeSlidePDFPages':60,'actualSlidePDFMutationsRejected':18},'scope':'Ordinary slide PDFs exported later from the exact captured native inputs. New PDF exports do not constitute additional native save rounds. Named Y translations remain differences; reversing them is an in-memory diagnostic only. Font-program bytes/glyphs/X positions and normalized full raster verified. Notes slide-image remains a separate open gate.'}
    (themes.ROOT/'docs/testing/native-notes-theme-slides-evidence.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n');return result

if __name__=='__main__':print(json.dumps(inspect(Path(sys.argv[1]))['total']))
