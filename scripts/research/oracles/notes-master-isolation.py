"""Independent diagnostic-source oracle. Multi-notes-master input is NON-CONFORMING."""
import copy,hashlib,json,posixpath,re,sys,tempfile,zipfile
from pathlib import Path
import xml.etree.ElementTree as E
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
P='{'+NS['p']+'}';A='{'+NS['a']+'}';R='{'+NS['r']+'}'

def read(file):
    with zipfile.ZipFile(file)as z:
        assert len(z.namelist())==len(set(z.namelist()));return{n:z.read(n)for n in z.namelist()}
def related(parts,owner,role,id=None,optional=False):
    part=posixpath.join(posixpath.dirname(owner),'_rels',posixpath.basename(owner)+'.rels')if owner else'_rels/.rels'
    nodes=list(E.fromstring(parts[part]));assert len({n.attrib['Id']for n in nodes})==len(nodes)
    nodes=[n for n in nodes if n.attrib['Type'].endswith('/'+role)and(id is None or n.attrib['Id']==id)]
    if optional and not nodes:return None
    assert len(nodes)==1 and nodes[0].attrib.get('TargetMode')is None
    target=posixpath.normpath(posixpath.join(posixpath.dirname(owner),nodes[0].attrib['Target']));assert target in parts;return target

def location(parts,source):
    node=E.fromstring(parts[source['part'].lstrip('/')]);path=source['path'];steps=path.split('/@')[0].split('/')[1:]
    for i,step in enumerate(steps):
        m=re.fullmatch(r'([^:]+):([^\[]+)\[(\d+)\]',step);assert m,step
        prefix,name,index=m.groups();tag='{'+NS[prefix]+'}'+name
        if i==0:assert node.tag==tag and index=='1'
        else:node=node.findall(tag)[int(index)-1]
    if '/@'not in path:return node,None
    attr=path.split('/@')[1]
    if ':'in attr:prefix,local=attr.split(':');attr='{'+NS[prefix]+'}'+local
    assert attr in node.attrib;return node,node.attrib[attr]

def inspect(parts,controls,stage,report=None):
    single_master=any(c.get('overridePart')for c in controls)
    main=related(parts,'','officeDocument');presentation=E.fromstring(parts[main]);listed=presentation.findall('p:notesMasterIdLst/p:notesMasterId',NS);assert len(listed)==1,'multiple listed notes masters'
    listed_parts={related(parts,main,'notesMaster',n.attrib[R+'id'])for n in listed}
    types=E.fromstring(parts['[Content_Types].xml']);actual={n.attrib['PartName'].lstrip('/')for n in types if n.attrib.get('ContentType','').endswith('presentationml.notesMaster+xml')};assert len(actual)==(1 if single_master else 2) and listed_parts<=actual,'named notes master part inventory'
    ids=presentation.findall('p:sldIdLst/p:sldId',NS);assert len(ids)==len(controls)==30;rows=[];unlisted=[]
    for node,c in zip(ids,controls):
        slide=related(parts,main,'slide',node.attrib[R+'id']);assert slide==f"ppt/slides/slide{c['sourceSlide']}.xml",'actual presentation order'
        notes=related(parts,slide,'notesSlide');master=related(parts,notes,'notesMaster');master_theme=related(parts,master,'theme');theme=related(parts,notes,'themeOverride',optional=True)or master_theme;assert related(parts,notes,'slide')==slide
        assert notes==c['notesPart']and master==c['masterPart']and theme==c['themePart']and master_theme==c.get('masterThemePart',c['themePart']),'actual notes/master/theme chain'
        if master not in listed_parts:unlisted.append('/'+notes)
        root=E.fromstring(parts[notes]);shapes=root.findall('p:cSld/p:spTree/p:sp',NS)
        bodies=[s for s in shapes if s.find('p:nvSpPr/p:nvPr/p:ph',NS)is not None and s.find('p:nvSpPr/p:nvPr/p:ph',NS).attrib['type']=='body'];assert len(bodies)==1
        body=bodies[0];assert body.find('p:nvSpPr/p:nvPr/p:ph',NS).attrib['idx']=='7';runs=body.findall('p:txBody/a:p/a:r',NS);assert len(runs)==1
        run=runs[0];text=c['text']+'X'*((stage+1)//2 if c['variant']=='A'else stage//2);assert run.find('a:t',NS).text==text,'scalar notes text'
        props=run.find('a:rPr',NS);assert props.attrib=={'lang':c['language']};slots=[props.find('a:'+k,NS)for k in ['latin','ea','cs']];assert[n.attrib['typeface']if n is not None else None for n in slots]==c['localSlots'],'local source slots'
        label=next(s for s in shapes if s.find('p:nvSpPr/p:cNvPr',NS).attrib['name']=='Notes Control Label');assert label.find('p:txBody/a:p/a:r/a:t',NS).text==f"notes-{c['variant']}-{c['id']}"
        mr=E.fromstring(parts[master]);mb=[s for s in mr.findall('p:cSld/p:spTree/p:sp',NS)if s.find('p:nvSpPr/p:nvPr/p:ph',NS)is not None and s.find('p:nvSpPr/p:nvPr/p:ph',NS).attrib['type']=='body'];assert len(mb)==1 and mb[0].find('p:nvSpPr/p:nvPr/p:ph',NS).attrib['idx']=='99'
        assert mb[0].find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',NS).attrib=={'sz':'4000'}
        tr=E.fromstring(parts[theme]);assert tr.tag==A+('themeOverride'if c.get('overridePart')else'theme')
        if not c.get('overridePart'):assert tr.attrib['name']=='PPTKit Notes '+c['variant']
        for role in ['major','minor']:
            collection=tr.find('.//a:'+role+'Font',NS);assert{n.attrib['script']:n.attrib['typeface']for n in collection.findall('a:font',NS)}==c['map'];assert[collection.find('a:'+k,NS).attrib['typeface']for k in ['latin','ea','cs']]==[c['map']['Latn'],'','']
        rows.append({'variant':c['variant'],'id':c['id'],'slide':slide,'notes':notes,'master':master,'theme':theme,'text':text,'declaredRules':c['map']})
    assert len(unlisted)==(0 if single_master else 15)
    checked=0
    if report:
        reported=report['inspection']['findings'];assert sorted(n['part']for n in reported if n['code']=='unlisted-master')==sorted(unlisted),'missing/mislocated structural diagnosis'
        assert len(report['inspection']['notesMasters'])==1
        for prop in report['textStyleInheritance']['properties']:
            if '/notesSlides/'not in prop['target']['part']:continue
            source,value=location(parts,prop['source']);location(parts,prop['target'])
            for d in prop.get('resolvedFrom',[]):location(parts,d)
            if prop['status']!='mapped':continue
            row=next(r for r in rows if '/'+r['notes']==prop['target']['part']);key=prop['property']
            if key.startswith('run.scriptFontFamilies.'):
                assert prop['source']['part']=='/'+row['theme']and source.attrib['script']==key.split('.')[-1]and value==prop['value'];checked+=1
            if key in ['run.fontFamily','run.eastAsianFontFamily','run.complexScriptFontFamily']:
                token=source.attrib['typeface'];match=re.fullmatch(r'\+(mj|mn)-(lt|ea|cs)',token)
                if match:
                    role='majorFont'if match[1]=='mj'else'minorFont';slot='latin'if match[2]=='lt'else match[2]
                    actual=E.fromstring(parts[row['theme']]).find('.//a:'+role+'/a:'+slot,NS).attrib['typeface'];assert actual==prop['value']
                    assert any(d['part']=='/'+row['theme']and location(parts,d)[1]==prop['value']for d in prop.get('resolvedFrom',[])),'wrong font resolution dependency'
                else:assert token==prop['value']
                checked+=1
            if key=='run.fontSize':assert prop['source']['part']=='/'+row['master']and value=='4000'and prop['value']==40 and prop['level']=='notes-master-placeholder';checked+=1
    return{'classification':'single-master-with-notes-font-overrides'if single_master else'non-conforming-multi-notes-master-source-control','listedNotesMasters':1,'actualNotesMasterParts':len(actual),'unlistedTargets':unlisted,'controls':rows,'mappedFontAndSizePropertiesChecked':checked}

def mutate(parts,part,change):
    result=copy.deepcopy(parts);root=E.fromstring(result[part]);change(root);result[part]=E.tostring(root);return result

def repack(parts):
    with tempfile.TemporaryDirectory(prefix='pptkit-notes-domain-negative-')as temp:
        file=Path(temp)/'damaged.pptx'
        with zipfile.ZipFile(file,'w')as z:
            for name,data in parts.items():z.writestr(name,data)
        return read(file),hashlib.sha256(file.read_bytes()).hexdigest()

def run(directory):
    records=json.loads((directory/'snapshots.json').read_text());controls=json.loads((directory/'controls.json').read_text());rounds=[];before=None;rejected=[]
    for stage,r in enumerate(records):
        parts=read(directory/r['file']);result=inspect(parts,controls,stage,r['report'])
        if before is not None:
            variant='A'if stage%2 else'B';assert set(parts)==set(before)
            for name,data in before.items():
                c=next((c for c in controls if c['notesPart']==name and c['variant']==variant),None)
                if c:
                    count=stage//2 if variant=='A'else(stage-1)//2;old=('<a:t>'+c['text']+'X'*count+'</a:t>').encode();new=old.replace(b'</a:t>',b'X</a:t>');assert data.count(old)==1 and parts[name]==data.replace(old,new),'non-character or wrong-owner bytes'
                else:assert parts[name]==data,'unmodified domain/part changed'
        before=parts
        def body(root):return next(s for s in root.findall('p:cSld/p:spTree/p:sp',NS)if s.find('p:nvSpPr/p:nvPr/p:ph',NS)is not None and s.find('p:nvSpPr/p:nvPr/p:ph',NS).attrib['type']=='body')
        def swap_order(root):
            node=root.find('p:sldIdLst',NS);second=node[1];node.remove(second);node.insert(0,second)
        single_master=any(c.get('overridePart')for c in controls)
        master_part='ppt/notesMasters/notesMaster1.xml'if single_master else'ppt/notesMasters/notesMaster2.xml'
        theme_part='ppt/theme/notes-font-override-B.xml'if single_master else'ppt/theme/independent-notes-theme-B.xml'
        theme_owner='ppt/notesSlides/_rels/notesSlide16.xml.rels'if single_master else'ppt/notesMasters/_rels/notesMaster2.xml.rels'
        theme_role='themeOverride'if single_master else'theme'
        mutations=[
            ('wrong-actual-master','ppt/notesSlides/_rels/notesSlide16.xml.rels',lambda root:next(n for n in root if n.attrib['Type'].endswith('/notesMaster')).set('Target','../notesMasters/notesMaster2.xml'if single_master else'../notesMasters/notesMaster1.xml')),
            ('borrow-A-theme',theme_owner,lambda root:next(n for n in root if n.attrib['Type'].endswith('/'+theme_role)).set('Target','../theme/independent-notes-theme.xml')),
            ('wrong-B-rule',theme_part,lambda root:root.find('.//a:minorFont/a:font',NS).set('typeface','Wrong')),
            ('wrong-master-type',master_part,lambda root:body(root).find('p:nvSpPr/p:nvPr/p:ph',NS).set('type','title')),
            ('wrong-B-local-type','ppt/notesSlides/notesSlide16.xml',lambda root:body(root).find('p:nvSpPr/p:nvPr/p:ph',NS).set('type','title')),
            ('wrong-B-language','ppt/notesSlides/notesSlide16.xml',lambda root:body(root).find('p:txBody/a:p/a:r/a:rPr',NS).set('lang','xx-XX')),
            ('wrong-B-text','ppt/notesSlides/notesSlide16.xml',lambda root:body(root).find('p:txBody/a:p/a:r/a:t',NS).__setattr__('text','Wrong')),
            ('wrong-B-local-font','ppt/notesSlides/notesSlide25.xml',lambda root:body(root).find('p:txBody/a:p/a:r/a:rPr/a:ea',NS).set('typeface','Wrong')),
            ('wrong-B-backlink','ppt/notesSlides/_rels/notesSlide16.xml.rels',lambda root:next(n for n in root if n.attrib['Type'].endswith('/slide')).set('Target','../slides/slide1.xml')),
            ('page-order','ppt/presentation.xml',swap_order)]
        for name,part,change in mutations:
            damaged,damaged_hash=repack(mutate(parts,part,change))
            try:inspect(damaged,controls,stage,r['report'])
            except AssertionError:rejected.append({'stage':stage,'mutation':name,'part':part,'sha256':damaged_hash})
            else:raise AssertionError('Accepted actual source corruption: '+name)
        result.update({'file':r['file'],'sha256':hashlib.sha256((directory/r['file']).read_bytes()).hexdigest(),'characterOnlyTransition':stage>0});rounds.append(result)
    return{'oracle':'Python stdlib raw OPC/XML actual relationships, interleaved source order, type/slots/dictionaries, located diagnosis/provenance and owner-scoped bytes','rounds':rounds,'rejectedCount':len(rejected),'rejectedActualMutations':rejected,'nativeAdmission':False,'fullSchemaConformance':'not validated'}

if __name__=='__main__':print(json.dumps(run(Path(sys.argv[1])),ensure_ascii=False))
