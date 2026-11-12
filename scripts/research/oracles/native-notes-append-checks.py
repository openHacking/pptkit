"""Repacked ZIP and rendered PDF failures must be rejected by the append oracle."""
import hashlib, importlib.util, json, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz

spec=importlib.util.spec_from_file_location('append',Path(__file__).with_name('native-notes-append.py'))
append=importlib.util.module_from_spec(spec)
spec.loader.exec_module(append)

def check(directory):
    expected=json.loads((directory/'expected.json').read_text())
    output=directory/'append-negative-controls'
    output.mkdir(exist_ok=True)
    rejected=[]
    def reject(name,file,run):
        try:
            run()
        except AssertionError as error:
            rejected.append({'name':name,'file':str(file),'sha256':append.notes.sha(file),'reason':str(error)})
        else:
            raise AssertionError('corruption accepted: '+name)
    part='ppt/notesSlides/notesSlide3.xml'
    def xml_mutation(name,owner,mutate):
        file=output/(name+'.pptx')
        with zipfile.ZipFile(directory/'append-round1.pptx') as z:
            payload={p:z.read(p) for p in z.namelist()}
        n=ET.fromstring(payload[owner]); mutate(n)
        payload[owner]=ET.tostring(n,encoding='utf-8',xml_declaration=True)
        with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED) as z:
            for p,b in payload.items():z.writestr(p,b)
        reject(name,file,lambda:append.inspect_package(file,directory,expected))
    def run(n):return append.paragraph(n).findall('a:r',append.NS)[1]
    xml_mutation('append-text',part,lambda n:setattr(run(n).find('a:t',append.NS),'text','Ψ'))
    xml_mutation('append-language',part,lambda n:run(n).find('a:rPr',append.NS).set('lang','el'))
    xml_mutation('append-bold',part,lambda n:run(n).find('a:rPr',append.NS).set('b','1'))
    xml_mutation('original-language',part,lambda n:append.paragraph(n).find('a:r/a:rPr',append.NS).set('lang','zh-CN'))
    xml_mutation('local-slot','ppt/notesSlides/notesSlide10.xml',lambda n:run(n).find('a:rPr/a:ea',append.NS).set('typeface','Georgia'))
    xml_mutation('notes-master','ppt/notesMasters/notesMaster1.xml',lambda n:n.set('showMasterSp','0'))
    xml_mutation('notes-theme','ppt/theme/theme2.xml',lambda n:n.set('name','Wrong domain'))
    xml_mutation('slide-neighbor','ppt/slides/slide3.xml',lambda n:n.set('showMasterSp','0'))
    xml_mutation('notes-relation','ppt/notesSlides/_rels/notesSlide3.xml.rels',lambda n:list(n)[0].set('Target','../slides/slide4.xml'))
    xml_mutation('control-note','ppt/notesSlides/notesSlide2.xml',lambda n:setattr(append.paragraph(n).find('a:r/a:t',append.NS),'text','Changed'))
    for mode in ['white-append','white-original','swap-pages','later-round']:
        file=output/(mode+'.pdf')
        with fitz.open(directory/'append-round1-notes.pdf') as doc:
            if mode=='swap-pages':doc.select([1,0]+list(range(2,15)))
            else:
                page=doc[2]
                if mode=='later-round':
                    page.draw_rect(fitz.Rect(20,700,50,730),color=None,fill=(0,0,0),overlay=True)
                else:
                    traces=[t for t in page.get_texttrace() if t['bbox'][1]>410]
                    trace=traces[-1] if mode=='white-append' else traces[0]
                    page.draw_rect(fitz.Rect(trace['bbox'])+(-2,-2,2,2),color=None,fill=(1,1,1),overlay=True)
            doc.save(file)
        reject(mode,file,lambda:append.inspect_pdf(file,directory,expected,directory/'append-round1-notes.pdf' if mode=='later-round' else None))
    result={'rejectedCount':len(rejected),'controls':rejected,'scope':'Actual repacked ZIPs and PDF files; not in-memory comparison-only mutations.'}
    (directory/'append-negative-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    return result

if __name__=='__main__':
    print(json.dumps({'rejectedCount':check(Path(sys.argv[1]))['rejectedCount']}))
