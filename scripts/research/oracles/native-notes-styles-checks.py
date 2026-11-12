"""Actual Ω style ZIP/PDF counterexamples, including pure-italic and matrix failures."""
import importlib.util, json, re, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz

spec=importlib.util.spec_from_file_location('styles',Path(__file__).with_name('native-notes-styles.py'))
styles=importlib.util.module_from_spec(spec);spec.loader.exec_module(styles)

def check(directory):
    expected=json.loads((directory/'expected.json').read_text());output=directory/'style-negative-controls';output.mkdir(exist_ok=True);records=[]
    def reject(stage,name,file,run):
        try:run()
        except AssertionError as error:records.append({'stage':stage,'name':name,'file':str(file),'sha256':styles.append.notes.sha(file),'reason':str(error)})
        else:raise AssertionError('style corruption accepted: '+name)
    for stage in styles.STAGES:
        for name in ['omega-flag','original-flag','language','local-slot','neighbor']:
            file=output/(stage+'-'+name+'.pptx')
            with zipfile.ZipFile(directory/(stage+'-round1.pptx')) as z:payload={p:z.read(p)for p in z.namelist()}
            part='ppt/notesSlides/notesSlide'+('10' if name=='local-slot' else '2' if name=='neighbor' else '3')+'.xml'
            n=ET.fromstring(payload[part]);p=styles.append.paragraph(n);runs=p.findall('a:r',styles.NS)
            if name=='omega-flag':runs[1].find('a:rPr',styles.NS).set('b','0' if stage=='bold' else '1')
            elif name=='original-flag':runs[0].find('a:rPr',styles.NS).set('i','1')
            elif name=='language':runs[1].find('a:rPr',styles.NS).set('lang','el')
            elif name=='local-slot':runs[1].find('a:rPr/a:ea',styles.NS).set('typeface','Georgia')
            else:runs[0].find('a:t',styles.NS).text='Changed'
            payload[part]=ET.tostring(n,encoding='utf-8',xml_declaration=True)
            with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED) as z:
                for owner,body in payload.items():z.writestr(owner,body)
            reject(stage,name,file,lambda:styles.inspect_package(file,directory,expected,stage))
        for name in ['white-omega','white-original','swap-pages','shear-matrix','later-ink']:
            file=output/(stage+'-'+name+'.pdf')
            with fitz.open(directory/(stage+'-round1-notes.pdf')) as doc:
                page=doc[2]
                if name=='swap-pages':doc.select([1,0]+list(range(2,15)))
                elif name=='later-ink':page.draw_rect(fitz.Rect(20,700,50,730),color=None,fill=(0,0,0),overlay=True)
                elif name=='shear-matrix':
                    face=styles.expected_face(expected['variant'],3,stage);resources={f[4]:f[3].split('+')[-1]for f in page.get_fonts()};count=0
                    for xref in page.get_contents():
                        stream=doc.xref_stream(xref).decode('latin1')
                        def mutate(match):
                            nonlocal count
                            if resources[match[7]]!=face or abs(float(match[1])-170.8333)>.0001:return match[0]
                            count+=1
                            values=list(match.groups());values[2]='50'
                            return 'BT '+' '.join(values[:6])+' Tm /'+values[6]+' 1 Tf'
                        changed=re.sub(styles.MATRIX,mutate,stream)
                        doc.update_stream(xref,changed.encode('latin1'))
                    assert count==1
                else:
                    body=[t for t in page.get_texttrace()if t['bbox'][1]>410];trace=body[-1] if name=='white-omega' else body[0]
                    page.draw_rect(fitz.Rect(trace['bbox'])+(-2,-2,2,2),color=None,fill=(1,1,1),overlay=True)
                doc.save(file)
            reject(stage,name,file,lambda:styles.inspect_pdf(file,directory,expected,stage,directory/(stage+'-round1-notes.pdf') if name=='later-ink' else None))
    result={'rejectedCount':len(records),'controls':records,'scope':'Repacked ZIP files and saved PDF mutations; source language/flags/slots and actual ink/order/matrix controls.'}
    (directory/'styles-negative-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    return result

if __name__=='__main__':print(json.dumps({'rejectedCount':check(Path(sys.argv[1]))['rejectedCount']}))
