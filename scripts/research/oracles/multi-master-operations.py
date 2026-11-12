"""Captured native layout, placeholder-text and selected-slide theme operations."""
import copy
import importlib.util
import json
import posixpath
import sys
import struct
import math
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as E
import fitz
from PIL import Image
spec=importlib.util.spec_from_file_location('native',Path(__file__).with_name('multi-master-native.py')); n=importlib.util.module_from_spec(spec);spec.loader.exec_module(n)
P,A,R,NS=n.P,n.A,n.R,n.NS

def xml(parts,part,node):parts[part]=E.tostring(node,encoding='utf-8',xml_declaration=True)
def shape(parts,part):
    root=n.root(parts,part); bodies=[s for s in root.findall('p:cSld/p:spTree/p:sp',NS) if (ph:=s.find('p:nvSpPr/p:nvPr/p:ph',NS)) is not None and ph.get('type')=='body' and ph.get('idx')=='1']
    assert len(bodies)==1
    return root,bodies[0]
def same(actual,expected):
    assert actual.keys()==expected.keys(), 'Unexpected semantic keys'
    diff=[k for k in actual if actual[k]!=expected[k]];assert not diff, ('Unapproved source change',diff)
def tree(node):return(node.tag,sorted(node.attrib.items()),'' if node.text is None or node.text.isspace() else node.text,tuple(tree(c)for c in node))
def fonts(document,page):return sorted((f[3].split('+')[-1],n.sha(document.extract_font(f[0])[3]))for f in page.get_fonts())
def glyph_program(data,gid):
    assert data[:4] in [b'\x00\x01\x00\x00',b'true']
    count=struct.unpack_from('>H',data,4)[0];tables={}
    for i in range(count):
        tag,_,start,length=struct.unpack_from('>4sIII',data,12+16*i);assert start+length<=len(data);tables[tag]=data[start:start+length]
    head,loca,glyf,maxp,hhea,hmtx=(tables[t]for t in [b'head',b'loca',b'glyf',b'maxp',b'hhea',b'hmtx'])
    units=struct.unpack_from('>H',head,18)[0];mode=struct.unpack_from('>h',head,50)[0];total=struct.unpack_from('>H',maxp,4)[0];assert 0<=gid<total and mode in [0,1]
    width=2 if mode==0 else 4;offsets=[struct.unpack_from('>H'if mode==0 else'>I',loca,g*width)[0]*(2 if mode==0 else 1)for g in [gid,gid+1]];start,end=offsets;assert 0<=start<=end<=len(glyf)
    program=glyf[start:end]
    # Named ASCII bystanders have simple TrueType glyphs; fail rather than
    # treating composite references as comparable raw programs across subsets.
    assert not program or struct.unpack_from('>h',program,0)[0]>=0, 'Composite glyph not admitted by this oracle'
    metrics=struct.unpack_from('>H',hhea,34)[0];assert 0<metrics<=total
    advance=struct.unpack_from('>H',hmtx,min(gid,metrics-1)*4)[0]
    lsb=struct.unpack_from('>h',hmtx,gid*4+2 if gid<metrics else metrics*4+(gid-metrics)*2)[0]
    return units,advance,lsb,n.sha(program)
def glyph_trace(document,page):
    result=copy.deepcopy(page.get_texttrace())
    for trace in result:
        candidates=[f for f in page.get_fonts()if f[3].split('+')[-1]==trace['font']];assert len(candidates)==1, 'Ambiguous actual font resource'
        data=document.extract_font(candidates[0][0])[3]
        trace['chars']=tuple((ch[0],glyph_program(data,ch[1]),ch[2],ch[3])for ch in trace['chars'])
    return result
def alter_font_advance(data,gid):
    data=bytearray(data);count=struct.unpack_from('>H',data,4)[0];tables={}
    for i in range(count):
        tag,_,start,length=struct.unpack_from('>4sIII',data,12+16*i);tables[tag]=(12+16*i,start,length)
    _,hhea,_=tables[b'hhea'];metrics=struct.unpack_from('>H',data,hhea+34)[0];assert gid<metrics
    record,start,length=tables[b'hmtx'];offset=start+gid*4;value=struct.unpack_from('>H',data,offset)[0];assert value<65535;struct.pack_into('>H',data,offset,value+1)
    def checksum(body):
        body=bytes(body)+b'\0'*((-len(body))%4);return sum(struct.unpack('>'+str(len(body)//4)+'I',body))&0xffffffff
    struct.pack_into('>I',data,record+4,checksum(data[start:start+length]))
    _,head,_=tables[b'head'];struct.pack_into('>I',data,head+8,0);struct.pack_into('>I',data,head+8,(0xB1B0AFBA-checksum(data))&0xffffffff)
    return bytes(data)
def edited_memory_trace(page):
    chars=[({k:v for k,v in t.items()if k not in ['chars','bbox','seqno']},ch)for t in page.get_texttrace()for ch in t['chars']]
    assert len(chars)==11
    props,ch=chars[6];assert ch[0]==48 and ch[2][0]==99.1875
    assert props['spacewidth']==6.999999046325684
    props['spacewidth']=3.499999523162842
    delta=-0.0047760009765625
    chars[6]=(props,(ch[0],ch[1],(ch[2][0]+delta,ch[2][1]),(ch[3][0]+delta,ch[3][1],ch[3][2]+delta,ch[3][3])))
    return chars
def flat_trace(page):return[({k:v for k,v in t.items()if k not in ['chars','bbox','seqno']},ch)for t in page.get_texttrace()for ch in t['chars']]
def pdf(file,stage,baseline=None,bystanders=None):
    d=fitz.open(file);assert len(d)==4; other=fitz.open(baseline) if baseline else None; old=fitz.open(bystanders) if bystanders else None
    names=['TimesNewRomanPSMT','TimesNewRomanPSMT','Georgia','ArialMT']; colors=['0044CC','0044CC','CC2200','112233']; bg=['E0E8F8','E0E8F8','F8E8E0','FFFFFF'];texts=['Owned '+str(i)for i in range(4)]
    if stage!='layout':texts[0]+=' EDIT'
    if stage=='theme':names[0]='DengXian-Regular';colors[0]='4472C4';bg[0]='FFFFFF'
    result=[]
    for i,p in enumerate(d):
        spans=[s for b in p.get_text('dict')['blocks'] if 'lines'in b for l in b['lines']for s in l['spans']]
        assert p.get_text().strip()==texts[i] and len(spans)==1, 'Wrong visible text/page mapping'
        assert spans[0]['font']==names[i] and spans[0]['color']==int(colors[i],16), 'Wrong actual font/color'
        pix=p.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False);im=Image.frombytes('RGB',[pix.width,pix.height],pix.samples)
        assert im.getpixel((pix.width//2,pix.height//2))==tuple(bytes.fromhex(bg[i])), 'Wrong inherited background'
        ink=tuple(int(c*255)for c in p.get_texttrace()[0]['color']);assert dict((color,count)for count,color in im.getcolors(pix.width*pix.height)).get(ink,0)>25, 'Text ink missing'
        if other:
            if stage=='text' and i==0: assert flat_trace(p)==edited_memory_trace(other[i]), 'Unexpected edited-memory to first-save glyph change'
            else: assert p.get_texttrace()==other[i].get_texttrace(), 'Changed post-operation glyph trace'
            assert pix.samples==other[i].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples, 'Changed post-operation raster'
            assert fonts(d,p)==fonts(other,other[i]), 'Changed embedded font program'
        if old and i:
            assert glyph_trace(d,p)==glyph_trace(old,old[i]), 'Bystander Unicode/glyph program/advance/position changed'
            assert pix.samples==old[i].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples, 'Bystander page changed'
        if old and i==0 and stage=='text':
            prior=old[i].get_texttrace();box=prior[0]['bbox'];crop=(math.floor(box[0]*2),math.floor(box[1]*2),math.ceil(box[2]*2),math.ceil(box[3]*2))
            beforepix=old[i].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False);beforeim=Image.frombytes('RGB',[beforepix.width,beforepix.height],beforepix.samples)
            assert im.crop(crop).tobytes()==beforeim.crop(crop).tobytes(), 'Original text ink changed during append'

        result.append({'page':i+1,'text':texts[i],'actualFace':names[i],'color':colors[i],'rasterSha256':n.sha(pix.samples),'fontPrograms':fonts(d,p),'rawTraceSha256':n.sha(json.dumps(p.get_texttrace(),sort_keys=True).encode()),'bystanderGlyphProgramTraceSha256':n.sha(json.dumps(glyph_trace(d,p),sort_keys=True).encode())if i else None})
    return result

def expected_layout(parts):
    parts=copy.copy(parts);_,rows=n.semantic(parts);source,body=shape(parts,rows[0]['slide']);_,target=shape(parts,rows[1]['layout'])
    local=body.find('p:spPr/a:xfrm',NS);inherited=target.find('p:spPr/a:xfrm',NS)
    assert local is not None and inherited is not None and tree(local)==tree(inherited), 'Native must inherit the same geometry'
    body.find('p:spPr',NS).remove(local);xml(parts,rows[0]['slide'],source)
    rel=n.relname(rows[0]['slide']);root=n.root(parts,rel);next(e for e in root if e.get('Type').endswith('/slideLayout')).set('Target',posixpath.relpath(rows[1]['layout'],posixpath.dirname(rows[0]['slide'])));xml(parts,rel,root)
    return n.semantic(parts)[0]
def expected_text(parts):
    parts=copy.copy(parts);_,rows=n.semantic(parts);root,body=shape(parts,rows[0]['slide']);p=body.find('p:txBody/a:p',NS);runs=p.findall('a:r',NS);assert len(runs)==1 and runs[0].find('a:t',NS).text=='Owned 0'
    props=runs[0].find('a:rPr',NS);props.set('dirty','0');end=p.find('a:endParaRPr',NS);end.set('dirty','0')
    for text,lang,alt in [(' ','zh-CN','en-US'),('EDIT','en-US','zh-CN')]:
        run=E.Element('{'+A+'}r');style=copy.deepcopy(props);style.set('lang',lang);style.set('altLang',alt);run.append(style);E.SubElement(run,'{'+A+'}t').text=text;p.insert(list(p).index(end),run)
    xml(parts,rows[0]['slide'],root);return n.semantic(parts)[0]

def office_profile(parts,row,directory):
    assets=directory/'office-assets';old=assets/'Office 2013 - 2022 Theme.thmx';modern=assets/'Office Theme.thmx'
    with zipfile.ZipFile(old)as z:legacy=E.fromstring(z.read('theme/theme/theme1.xml'))
    with zipfile.ZipFile(modern)as z:current=E.fromstring(z.read('theme/theme/theme1.xml'))
    actual=n.root(parts,row['theme'])
    for name in ['clrScheme','fmtScheme','fontScheme']:
        wanted=copy.deepcopy(legacy.find('a:themeElements/a:'+name,NS));wanted.set('name','Office')
        if name=='fontScheme':
            for kind in ['majorFont','minorFont']:
                base=wanted.find('a:'+kind,NS);recent=current.find('a:themeElements/a:fontScheme/a:'+kind,NS)
                base.find('a:latin',NS).set('typeface',next(f.get('typeface')for f in base if f.get('script')=='Hans'))
                scripts={f.get('script')for f in base if f.get('script')}
                for font in recent:
                    if font.get('script') and font.get('script')not in scripts:base.append(copy.deepcopy(font))
        assert tree(actual.find('a:themeElements/a:'+name,NS))==tree(wanted), 'Unexpected Office localized '+name
    return {'label':'Office 主题 (selected-slide command)','nativeName':actual.get('name'),'assets':[{'file':str(p),'sha256':n.sha(p.read_bytes())}for p in [old,modern]],'observedProfile':'Legacy palette/format; legacy font scheme with Hans Latin names and 17 modern supplemental entries per major/minor. Structural match, not a claim about the native internal algorithm.'}

def theme_preservation(parts,baseline,directory):
    before,oldrows=n.semantic(baseline);after,rows=n.semantic(parts,(5,17));assert after.keys()>=before.keys()
    assert rows[0]['masterId']=='2147483657' and rows[0]['layoutId']=='2147483669'
    assert rows[1:]==oldrows[1:], 'Theme scope leaked to unselected pages'
    newkeys=set(after)-set(before);expected={'master/2147483657','theme/master/2147483657'}|{'layout/'+str(i)for i in range(2147483658,2147483670)}
    assert newkeys==expected|{k+'/relationships'for k in expected}
    projected=copy.deepcopy({k:after[k]for k in before})
    projected['content-types']=[entry for entry in projected['content-types']if entry[1]not in expected]
    pres=projected['presentation'];lst=next(c for c in pres[3]if c[0]=='{'+P+'}sldMasterIdLst');refs=lst[3];newref=[c for c in refs if dict(c[1]).get('id')=='2147483657'];assert len(newref)==1;refs.remove(newref[0])
    edges=projected['presentation/relationships'];newedge=[e for e in edges if e[2]=='master/2147483657'];assert len(newedge)==1;edges.remove(newedge[0])
    edges=projected['slide/256/relationships'];newedge=next(e for e in edges if e[0].endswith('/slideLayout'));assert newedge[2]=='layout/2147483669';edges[edges.index(newedge)]=(newedge[0],newedge[1],'layout/2147483651');edges.sort()
    same(projected,before)
    profile=office_profile(parts,rows[0],directory)
    _,oldlayout=shape(baseline,oldrows[0]['layout']);_,newlayout=shape(parts,rows[0]['layout']);assert tree(oldlayout)==tree(newlayout), 'Copied layout placeholder changed'
    return after,rows,profile

def inspect(directory):
    baseline_file=Path('docs/testing/multi-master-native-evidence.json').resolve();baseline_bytes=baseline_file.read_bytes();baseline=json.loads(baseline_bytes)
    assert Path(baseline['directory'])==directory and n.sha((directory/'round3.pptx').read_bytes())==baseline['rounds'][-1]['sha256']
    source=n.payload(directory/'round3.pptx');results=[];negative=[];out=directory/'operation-negatives';out.mkdir(exist_ok=True)
    base_pdf=directory/'round3.pdf'
    for stage in ['layout','text','theme']:
        before=source
        input_file=directory/({'layout':'layout-input','text':'text-input','theme':'theme-input'}[stage]+'.pptx')
        predecessor=directory/({'layout':'round3','text':'layout-round3','theme':'text-round3'}[stage]+'.pptx')
        assert input_file.read_bytes()==predecessor.read_bytes(), 'Stage input was not an exact native predecessor copy'
        counts=(5,17)if stage=='theme'else(4,5)
        first_file=directory/(stage+'-round1.pptx');first=n.payload(first_file)
        profile=None
        if stage=='layout':expected=expected_layout(before)
        elif stage=='text':expected=expected_text(before)
        else:expected,_,profile=theme_preservation(first,before,directory)
        actual,rows=n.semantic(first,counts);same(actual,expected)
        oldaliases,_=n.graph(before);newaliases,_=n.graph(first,counts)
        oldmapped=set(oldaliases)|{n.relname(k)for k in oldaliases};newmapped=set(newaliases)|{n.relname(k)for k in newaliases}
        oldother=set(before)-oldmapped-{'[Content_Types].xml'};newother=set(first)-newmapped-{'[Content_Types].xml'}
        assert oldother==newother, 'Unexpected unowned payload'
        metadata={}
        for part in oldother:
            if part not in ['docProps/core.xml','docProps/app.xml']:assert before[part]==first[part], ('Unchanged ancillary payload changed',part)
            elif before[part]!=first[part]:metadata[part]={'beforeSha256':n.sha(before[part]),'afterSha256':n.sha(first[part]),'beforeXml':before[part].decode(),'afterXml':first[part].decode()}
        prior_pdf=directory/'layout-round3.pdf' if stage=='text' else base_pdf
        start=directory/(stage+'-first.pdf');first_pages=pdf(start,stage,bystanders=prior_pdf);rounds=[]
        for i in range(1,4):
            file=directory/f'{stage}-round{i}.pptx';data=n.payload(file);sem,domains=n.semantic(data,counts);same(sem,expected)
            changed=[p for p in data if p not in first or data[p]!=first[p]];assert set(data)==set(first) and set(changed)<= {'docProps/core.xml'}, ('Unexpected later native payload change',changed)
            pf=directory/f'{stage}-round{i}.pdf';pages=pdf(pf,stage,start,prior_pdf)
            if i>1:
                saved=fitz.open(directory/(stage+'-round1.pdf'));current=fitz.open(pf)
                for j in range(4):
                    assert current[j].get_texttrace()==saved[j].get_texttrace(), 'Raw saved-round trace changed'
                    assert fonts(current,current[j])==fonts(saved,saved[j]), 'Raw saved-round font changed'
            rounds.append({'file':str(file),'sha256':n.sha(file.read_bytes()),'pdf':str(pf),'pdfSha256':n.sha(pf.read_bytes()),'domains':domains,'pages':pages})
        capture=directory/({'layout':'layout-applied','text':'text-edited','theme':'theme-selected-slide'}[stage]+'.png');assert capture.exists()
        results.append({'changedMetadata':metadata,'stage':stage,'inputFile':str(directory/({'layout':'layout-input','text':'text-input','theme':'theme-input'}[stage]+'.pptx')),'inputSha256':n.sha((directory/({'layout':'layout-input','text':'text-input','theme':'theme-input'}[stage]+'.pptx')).read_bytes()),'firstPdf':str(start),'firstPdfSha256':n.sha(start.read_bytes()),'firstPages':first_pages,'operationScreenshot':str(capture),'screenshotSha256':n.sha(capture.read_bytes()),'officeProfile':profile,'rounds':rounds})
        def rejection(label,parts):
            file=out/(stage+'-'+label+'.pptx')
            with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED)as z:
                for p,b in parts.items():z.writestr(p,b)
            try:actual,_=n.semantic(n.payload(file),counts);same(actual,expected)
            except (AssertionError,KeyError,StopIteration)as e:negative.append({'stage':stage,'name':label,'file':str(file),'sha256':n.sha(file.read_bytes()),'reason':str(e)})
            else:raise AssertionError('Accepted corruption '+label)
        # Actual files independently break target text, bystander text, notes, owner theme, color and layout relation.
        for label,part,kind in [('target-text',rows[0]['slide'],'text'),('bystander',rows[1]['slide'],'text'),('notes',rows[0]['notes'],'text'),('color',rows[0]['theme'],'color'),('font',rows[0]['theme'],'font'),('layout-edge',n.relname(rows[0]['slide']),'edge')]:
            parts=copy.copy(first);r=n.root(parts,part)
            if kind=='text':next(r.iter('{'+A+'}t')).text='WRONG'
            elif kind=='color':r.find('a:themeElements/a:clrScheme/a:accent1/a:srgbClr',NS).set('val','00FF00')
            elif kind=='font':r.find('a:themeElements/a:fontScheme/a:minorFont/a:latin',NS).set('typeface','Wrong font')
            else:next(e for e in r if e.get('Type').endswith('/slideLayout')).set('Target',posixpath.relpath(rows[2]['layout'],posixpath.dirname(rows[0]['slide'])))
            xml(parts,part,r);rejection(label,parts)
        for name in ['blank-label','page-order']:
            file=out/(stage+'-'+name+'.pdf');d=fitz.open(directory/(stage+'-round1.pdf'))
            if name=='blank-label':d[0].draw_rect(d[0].search_for('Owned')[0],fill=(1,1,1),color=(1,1,1),overlay=True)
            else:d.select([1,0,2,3])
            d.save(file)
            try:pdf(file,stage,start,prior_pdf)
            except AssertionError as e:negative.append({'stage':stage,'name':name,'file':str(file),'sha256':n.sha(file.read_bytes()),'reason':str(e)})
            else:raise AssertionError('Accepted PDF corruption')
        # A real font-metric mutation can preserve pixels while violating the
        # source glyph program/advance invariant; pixel checks alone must not pass it.
        file=out/(stage+'-font-advance.pdf');d=fitz.open(directory/(stage+'-round1.pdf'));page=d[1]
        candidates=[f for f in page.get_fonts()if f[3].split('+')[-1]=='TimesNewRomanPSMT'];assert len(candidates)==1
        program=d.extract_font(candidates[0][0])[3];gid=page.get_texttrace()[0]['chars'][0][1]
        streams=[xref for xref in range(1,d.xref_length())if d.xref_is_stream(xref)and d.xref_stream(xref)==program];assert len(streams)==1
        d.update_stream(streams[0],alter_font_advance(program,gid));d.save(file)
        patched=fitz.open(file);original=fitz.open(directory/(stage+'-round1.pdf'))
        assert patched[1].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples==original[1].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples, 'Metric control unexpectedly changed raster'
        try:pdf(file,stage,bystanders=prior_pdf)
        except AssertionError as error:negative.append({'stage':stage,'name':'font-advance','file':str(file),'sha256':n.sha(file.read_bytes()),'sameRasterAsPositive':True,'reason':str(error)})
        else:raise AssertionError('Accepted raster-identical corrupted glyph metric')
        source=n.payload(directory/(stage+'-round3.pptx'))
    # Native first-page contact for the three actual operations.
    images=[]
    for stage in ['layout','text','theme']:
        d=fitz.open(directory/(stage+'-round3.pdf'));pix=d[0].get_pixmap(matrix=fitz.Matrix(.8,.8),alpha=False);images.append(Image.frombytes('RGB',[pix.width,pix.height],pix.samples))
    w,h=images[0].size;contact=Image.new('RGB',(w,h*3),'white')
    for i,im in enumerate(images):contact.paste(im,(0,i*h))
    contact_file=directory/'operations-contact.png';contact.save(contact_file)
    evidence={'baselineEvidence':str(baseline_file),'baselineEvidenceSha256':n.sha(baseline_bytes),'textPreSaveDifference':{'spanCountBefore':3,'spanCountAfter':2,'characterIndex':6,'unicode':48,'xBefore':99.1875,'xAfter':99.18272399902344,'xDeltaPt':-0.0047760009765625,'reportedSpacewidthBefore':6.999999046325684,'reportedSpacewidthAfter':3.499999523162842,'fontProgramAndRasterEqual':True,'savedRound1To2To3RawTraceEqual':True},'schemaVersion':1,'recordedAt':datetime.now(timezone.utc).isoformat(),'directory':str(directory),'application':{'name':'Microsoft PowerPoint for Mac','version':'16.104'},'stages':results,'negatives':negative,'contact':{'file':str(contact_file),'sha256':n.sha(contact_file.read_bytes())},'total':{'nativeOperationStages':3,'nativeSaveCloseReopenRounds':9,'firstOperationPages':12,'reopenedPages':36,'actualZipNegatives':18,'actualPdfNegatives':9},'scope':'Actual cross-master layout change, placeholder text append, and Office selected-slide theme operation only. Complete artwork/template semantics, Core shared-edit APIs and source style/layout writers remain open.'}
    Path('docs/testing/multi-master-operations-evidence.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2)+'\n');print(json.dumps(evidence['total']))
if __name__=='__main__':inspect(Path(sys.argv[1]).resolve())
