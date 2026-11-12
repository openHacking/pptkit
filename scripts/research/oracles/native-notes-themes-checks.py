"""Repacked theme-domain and actual PDF corruptions, including the named kerning exception."""
import importlib.util,json,sys,zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec=importlib.util.spec_from_file_location('themes',Path(__file__).with_name('native-notes-themes.py'))
themes=importlib.util.module_from_spec(spec);spec.loader.exec_module(themes)

def check(directory):
    expected=json.loads((directory/'expected.json').read_text());group=expected['group']
    controls=json.loads((directory/'controls.json').read_text())if group=='override'else[{**c,'variant':group}for c in expected['controls']]
    positive=themes.inspect_theme_package(directory/'office-round1.pptx',directory,expected,controls);rows=positive['domains']
    output=directory/'theme-negative-controls';output.mkdir(exist_ok=True);records=[]
    def reject(name,file,run):
        try:run()
        except (AssertionError,KeyError)as error:records.append({'name':name,'file':str(file),'sha256':themes.notes.sha(file),'reason':str(error)})
        else:raise AssertionError('theme corruption accepted: '+name)
    def xml(name,part,change):
        file=output/(name+'.pptx')
        with zipfile.ZipFile(directory/'office-round1.pptx')as z:payload={p:z.read(p)for p in z.namelist()}
        n=ET.fromstring(payload[part]);change(n);payload[part]=ET.tostring(n,encoding='utf-8',xml_declaration=True)
        with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED)as z:
            for owner,body in payload.items():z.writestr(owner,body)
        reject(name,file,lambda:themes.inspect_theme_package(file,directory,expected,controls))
    def rpr(n):return themes.notes.body(n).find('p:txBody/a:p/a:r/a:rPr',themes.NS)
    first=rows[0]
    xml('notes-text',first['notesPart'],lambda n:setattr(themes.notes.body(n).find('p:txBody/a:p/a:r/a:t',themes.NS),'text','Wrong'))
    xml('notes-language',first['notesPart'],lambda n:rpr(n).set('lang','el'))
    xml('notes-style',first['notesPart'],lambda n:rpr(n).set('b','1'))
    local=next(r for r in rows if r['id']=='local-tc')
    xml('local-slot',local['notesPart'],lambda n:rpr(n).find('a:ea',themes.NS).set('typeface','Georgia'))
    xml('notes-theme',first['masterThemePart'],lambda n:n.find('a:themeElements/a:fontScheme/a:minorFont/a:font',themes.NS).set('typeface','Wrong'))
    xml('master-type',first['masterPart'],lambda n:themes.notes.body(n).find('p:nvSpPr/p:nvPr/p:ph',themes.NS).set('type','title'))
    xml('master-size',first['masterPart'],lambda n:themes.notes.body(n).find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',themes.NS).set('sz','1800'))
    xml('notes-borrow-slide-theme','ppt/notesMasters/_rels/notesMaster1.xml.rels',lambda n:next(c for c in n if c.get('Type').endswith('/theme')).set('Target','../theme/theme1.xml'))
    xml('slide-neighbor',first['slidePart'],lambda n:setattr(themes.notes.named(n,'SlideProbe').find('p:txBody/a:p/a:r/a:t',themes.NS),'text','Wrong'))
    xml('wrong-slide-layout','ppt/slides/_rels/slide1.xml.rels',lambda n:next(c for c in n if c.get('Type').endswith('/slideLayout')).set('Target','../slideLayouts/slideLayout1.xml'))
    xml('office-font-scheme','ppt/theme/theme1.xml',lambda n:n.find('a:themeElements/a:fontScheme/a:minorFont/a:latin',themes.NS).set('typeface','Georgia'))
    xml('presentation-default-kern','ppt/presentation.xml',lambda n:n.find('p:defaultTextStyle/a:lvl1pPr/a:defRPr',themes.NS).set('kern','0'))
    def swap(n):
        ids=n.find('p:sldIdLst',themes.NS);second=ids[1];ids.remove(second);ids.insert(0,second)
    xml('presentation-page-order','ppt/presentation.xml',swap)
    if group=='override':
        target=next(r for r in rows if r['variant']=='B')
        xml('override-font',target['overridePart'],lambda n:n.find('a:fontScheme/a:minorFont/a:latin',themes.NS).set('typeface','Arial'))
        parent,name=target['notesPart'].rsplit('/',1);rel=parent+'/_rels/'+name+'.rels'
        xml('override-borrows-A',rel,lambda n:n.remove(next(c for c in n if c.get('Type').endswith('/themeOverride'))))
    baseline=directory/('native-baseline-round3-notes.pdf'if group=='override'else'baseline-notes.pdf')
    modes=['white-body','font-resource','swap-pages','later-ink']+(['white-shifted-header-A']if group!='B'else[])
    for mode in modes:
        file=output/(mode+'.pdf')
        with fitz.open(directory/'office-round1-notes.pdf')as doc:
            page=doc[0];trace=next(t for t in page.get_texttrace()if t['bbox'][1]>410)
            if mode=='swap-pages':doc.select([1,0]+list(range(2,len(doc))))
            elif mode=='later-ink':page.draw_rect(fitz.Rect(450,650,480,680),color=None,fill=(0,0,0),overlay=True)
            elif mode=='font-resource':
                fonts=[f for f in page.get_fonts()if f[3].split('+')[-1]==trace['font']];assert fonts
                for font in fonts:doc.xref_set_key(font[0],'BaseFont','/Wrong')
            else:
                box=trace['bbox']if mode=='white-body'else next(t for t in page.get_texttrace()if 350<t['bbox'][1]<410)['chars'][-1][3]
                page.draw_rect(fitz.Rect(box)+(-1,-1,1,1),color=None,fill=(1,1,1),overlay=True)
            doc.save(file)
        reject(mode,file,lambda:themes.inspect_pdf(file,baseline,controls,True,directory/'office-round1-notes.pdf'if mode=='later-ink'else None))
    result={'rejectedCount':len(records),'controls':records,'scope':'Saved actual ZIP/PDF mutations. Shared-override repair prompt is a separate native observation, not counted as an oracle mutation.'}
    (directory/'theme-negative-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n');return result

if __name__=='__main__':print(json.dumps({'rejectedCount':check(Path(sys.argv[1]))['rejectedCount']}))
