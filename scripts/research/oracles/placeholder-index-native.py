"""Named explicit-index source/native controls; not a general template importer."""
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

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('index_physical_primitives', HERE/'layout-rebind-native.py')
p = importlib.util.module_from_spec(spec); spec.loader.exec_module(p)
a, n, c, NS, A, P = p.a, p.a.n, p.c, p.c.NS, p.c.A, p.c.P
MAPS = {
    'Source': {'old-body': ('body', 0), 'old-photo': ('pic', 4294967295), 'title': ('title', 73), 'legacy': ('subTitle', 4)},
    'Target': {'decoy': ('body', 17), 'legacy': ('subTitle', 1), 'reserved': ('body', 2), 'photo': ('pic', 4294967295), 'new-body': ('body', 0), 'title': ('title', 73)},
}


def index(ph):
    raw = ph.get('idx', '0')
    assert re.fullmatch(r'\d+', raw), 'Non-integer native placeholder identity'
    value = int(raw); assert 0 <= value <= 4294967295, 'Native placeholder identity outside uint32'
    return value


def graph(parts, media_sha, before=False):
    presentation = n.root(parts, 'ppt/presentation.xml')
    refs = presentation.findall('p:sldIdLst/p:sldId', NS)
    assert len(refs) == 3 and [node.get('id') for node in refs] == ['256', '257', '258']
    assert presentation.find('p:sldSz', NS).attrib['cx'] == '9144000'
    assert presentation.find('p:sldSz', NS).attrib['cy'] == '6858000'
    edges = n.relations(parts, 'ppt/presentation.xml')
    layouts = {}
    for part in parts:
        if not (part.startswith('ppt/slideLayouts/') and part.endswith('.xml')): continue
        root = n.root(parts, part); name = root.find('p:cSld', NS).get('name')
        actual = {}
        for shape in root.findall('p:cSld/p:spTree/p:sp', NS):
            key = shape.find('p:nvSpPr/p:cNvPr', NS).get('name')
            ph = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
            assert ph is not None and key not in actual
            actual[key] = (ph.get('type', 'obj'), index(ph))
        if name in MAPS:
            assert actual == MAPS[name], ('Changed layout numeric identities', name, actual)
            assert len({value[1] for value in actual.values()}) == len(actual), 'Duplicate prompt index'
            layouts[name] = part
    assert set(layouts) == {'Source', 'Target'}
    media = set()
    for i, ref in enumerate(refs):
        slide = edges[ref.get('{'+c.R+'}id')]['target']
        assert slide == f'ppt/slides/slide{i+1}.xml', 'Changed slide order'
        root = n.root(parts, slide); target = not before and i < 2
        assert root.get('show', '1') == ('0' if i == 1 else '1'), 'Changed hidden page'
        layout = n.target(parts, slide, 'slideLayout')
        assert layout == layouts['Target' if target else 'Source'], 'Changed slide/layout owner'
        master = n.target(parts, layout, 'slideMaster')
        assert n.root(parts, master).find('p:cSld', NS).get('name') == ('blue' if target else 'red')
        theme = n.root(parts, n.target(parts, master, 'theme'))
        assert theme.find('a:themeElements/a:clrScheme/a:accent1/a:srgbClr', NS).get('val') == ('0044CC' if target else 'CC2200'), 'Changed owner theme'
        for family in ['majorFont','minorFont']:
            for slot in ['latin','ea','cs']:
                font = theme.find('a:themeElements/a:fontScheme/a:'+family+'/a:'+slot, NS).get('typeface')
                assert font == ('Courier New' if target else 'Arial'), 'Changed owner font slots'
        tree = root.find('p:cSld/p:spTree', NS); objects = {}
        assert all(node.tag in ['{'+P+'}nvGrpSpPr', '{'+P+'}grpSpPr', '{'+P+'}sp', '{'+P+'}pic'] for node in tree)
        ids = [node.get('id') for node in tree.findall('.//p:cNvPr', NS)]; assert len(ids) == len(set(ids)), 'Drawing identity collision'
        for node in tree:
            if node.tag not in ['{'+P+'}sp', '{'+P+'}pic']: continue
            name = node.find('.//p:cNvPr', NS).get('name'); assert name not in objects; objects[name] = node
        assert set(objects) == {name+'-'+str(i) for name in ['title', 'body', 'image', 'legacy', 'ordinary']}, 'Lost/duplicated/flattened page objects'
        for name, kind, expected in [('title', 'title', 73), ('body', 'body', 0), ('image', 'pic', 4294967295), ('legacy', 'subTitle', 1 if target else 4)]:
            node = objects[name+'-'+str(i)]; ph = node.find('.//p:ph', NS)
            assert ph is not None and ph.get('type', 'obj') == kind and index(ph) == expected, ('Changed local numeric binding', name)
        frames = {'title': (40,40,500,60), 'body': (40,130,300,100), 'image': (430,200,200,100), 'legacy': (40,440,400,50)} if not before and i == 1 else {'title': (30,30,500,50), 'body': (30,110,280,80), 'image': (400,150,160,80), 'legacy': (30,440,400,50)}
        for name, expected in frames.items():
            node = objects[name+'-'+str(i)]; xfrm = node.find('p:spPr/a:xfrm', NS)
            off, ext = xfrm.find('a:off', NS), xfrm.find('a:ext', NS)
            assert tuple(int(v) for v in [off.get('x'),off.get('y'),ext.get('cx'),ext.get('cy')]) == tuple(v*12700 for v in expected), 'Wrong bound placeholder frame'
        for name, text in [('title','TITLE '),('body','BODY '),('legacy','LEGACY '),('ordinary','Ordinary ')]:
            assert ''.join(node.text or '' for node in objects[name+'-'+str(i)].findall('.//a:t', NS)) == text+str(i), 'Changed local content'
        note = n.target(parts, slide, 'notesSlide')
        assert n.target(parts, note, 'slide') == slide, 'Wrong notes backlink'
        assert ''.join(n.root(parts, note).itertext()) == 'Index notes '+str(i)
        image = n.target(parts, slide, 'image'); media.add(image)
        assert a.sha(parts[image]) == media_sha, 'Changed source image payload'
        embed = objects['image-'+str(i)].find('.//a:blip', NS).get('{'+c.R+'}embed')
        assert n.relations(parts, slide)[embed]['target'] == image, 'Picture references another resource'
    assert len(media) == 1 and {name for name in parts if name.startswith('ppt/media/')} == media
    for name in parts:
        if name.endswith('.rels'):
            owner = '' if name == '_rels/.rels' else str(Path(name).parent.parent/Path(name).name.removesuffix('.rels'))
            edges = n.relations(parts, owner); assert len(edges) == len(E.fromstring(parts[name])), 'Duplicate relationship IDs'
            for edge in edges.values():
                if edge['mode'] == 'Internal': assert edge['target'] in parts, 'Broken internal edge'


def physical(before_file, after_file):
    before, after = fitz.open(before_file), fitz.open(after_file)
    assert len(before) == len(after) == 3
    for i in range(3):
        assert after[i].get_text().splitlines() == ['TITLE '+str(i), 'BODY '+str(i), 'LEGACY '+str(i), 'Ordinary '+str(i)]
    for i in [0,2]:
        assert p.glyph_trace(before,before[i]) == p.glyph_trace(after,after[i]), 'Changed preserve/unselected glyph program/metric/position'
        x,y = a.pdf_page_record(before,i),a.pdf_page_record(after,i)
        for key in ['size','text','images']: assert x[key] == y[key], 'Changed preserved content '+key
        if i == 2:
            assert x['drawings'] == y['drawings'] and x['rasterSha256'] == y['rasterSha256'], 'Changed unselected appearance'
        else:
            diffs = [(old,new) for old,new in zip(x['drawings'],y['drawings']) if old != new]
            assert len(x['drawings']) == len(y['drawings']) and len(diffs) == 1
            old,new = diffs[0]; assert set(k for k in old if old[k] != new[k]) == {'fill'}
            assert old['fill'] == [0.800000011920929,0.13331808149814606,0.0]
            assert new['fill'] == [0.0,0.2666666805744171,0.800000011920929]
    assert all(span['font'] == 'CourierNewPSMT' for span in after[1].get_texttrace()), 'Target font not used on reset page'
    images = after[1].get_image_info(hashes=True); assert len(images) == 1
    assert images[0]['digest'] == before[1].get_image_info(hashes=True)[0]['digest'], 'Image content changed after reset'
    assert images[0]['bbox'] != before[1].get_image_info()[0]['bbox'], 'Picture did not follow target frame'
    for i in range(3):
        page = after[i]; image = page.get_image_info()[0]; rect = fitz.Rect(image['bbox']); pix = page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
        for u,v,color in [(0.2,0.2,(36,87,214)),(0.2,0.8,(34,197,94)),(0.8,0.2,(217,74,102)),(0.8,0.8,(250,204,21))]:
            assert pix.pixel(round((rect.x0+rect.width*u)*2),round((rect.y0+rect.height*v)*2)) == color, 'Picture occluded/cropped incorrectly'


def main():
    manifest = json.loads(Path('docs/testing/placeholder-index-evidence.json').read_text()); directory = Path(manifest['nativeDirectory'])
    assert a.sha((directory/'input.pptx').read_bytes()) == manifest['sourceSha256']
    assert a.sha((directory/'before.pptx').read_bytes()) == manifest['beforeSha256']
    source = n.payload(directory/'input.pptx'); media = n.target(source,'ppt/slides/slide1.xml','image'); media_sha = a.sha(source[media])
    graph(n.payload(directory/'before.pptx'),media_sha,True); graph(source,media_sha)
    physical(directory/'before.pdf',directory/'first.pdf')
    first = p.page_records(directory/'first.pdf'); rounds = []; canonical = None
    for i in range(1,4):
        file = directory/f'index-round{i}.pptx'; parts = n.payload(file); graph(parts,media_sha)
        current = c.full(parts)
        if canonical is None: canonical = current
        else: assert current == canonical, 'Native package changed after first save'
        assert p.page_records(directory/f'index-round{i}.pdf') == first, 'Native reopened PDF changed'
        opened = (directory/f'index-round{i}-reopened.ax.txt').read_text(); assert 'URL: file://'+str(file) in opened and not re.search('发现.*内容有问题|修复',opened)
        rounds.append({'round':i,'file':str(file),'sha256':a.sha(file.read_bytes())})
    for name in ['before','first','index-round1','index-round2','index-round3']:
        options = (directory/(name+'-print-options.ax.txt')).read_text()
        assert '复选框 打印隐藏幻灯片, Value: 1' in options and '页码：3/3' in options
    for name,file in [('before','before.pptx'),('first','input.pptx')]:
        opened = (directory/(name+'-open.ax.txt')).read_text()
        assert 'URL: file://'+str(directory/file) in opened and not re.search('发现.*内容有问题|修复',opened)
    assert '版本 16.104 (25121423)' in (directory/'powerpoint-version.ax.txt').read_text()
    out = Path(tempfile.mkdtemp(prefix='negative-',dir=directory)); positive = n.payload(directory/'index-round1.pptx')
    def object_named(root,name): return next(node for node in root.iter() if node.tag in ['{'+P+'}sp','{'+P+'}pic'] and node.find('.//p:cNvPr',NS).get('name') == name)
    def ph(root,name): return object_named(root,name).find('.//p:ph',NS)
    def xml_patch(part,callback):
        def mutate(parts):
            root=E.fromstring(parts[part]);callback(root);parts[part]=E.tostring(root)
        return mutate
    def rel(root,kind): return next(node for node in root if node.get('Type').endswith('/'+kind))
    target_layout = next(part for part in positive if part.startswith('ppt/slideLayouts/') and part.endswith('.xml') and n.root(positive,part).find('p:cSld',NS).get('name') == 'Target')
    def payload(parts):
        part=n.target(parts,'ppt/slides/slide1.xml','image');parts[part]+=b'changed'
    cases = [
        ('body-to-decoy',xml_patch('ppt/slides/slide1.xml',lambda root: ph(root,'body-0').set('idx','17'))),
        ('picture-to-implicit',xml_patch('ppt/slides/slide1.xml',lambda root: ph(root,'image-0').set('idx','1'))),
        ('missing-max-index',xml_patch('ppt/slides/slide1.xml',lambda root: ph(root,'image-0').attrib.pop('idx'))),
        ('duplicate-target-index',xml_patch(target_layout,lambda root: ph(root,'reserved').set('idx','0'))),
        ('implicit-reservation-collision',xml_patch(target_layout,lambda root: ph(root,'legacy').set('idx','2'))),
        ('uint32-overflow',xml_patch('ppt/slides/slide1.xml',lambda root: ph(root,'image-0').set('idx','4294967296'))),
        ('negative-index',xml_patch('ppt/slides/slide1.xml',lambda root: ph(root,'body-0').set('idx','-1'))),
        ('fractional-index',xml_patch('ppt/slides/slide1.xml',lambda root: ph(root,'body-0').set('idx','0.5'))),
        ('wrong-placeholder-role',xml_patch('ppt/slides/slide1.xml',lambda root: ph(root,'body-0').set('type','title'))),
        ('unselected-implicit-index',xml_patch('ppt/slides/slide3.xml',lambda root: ph(root,'legacy-2').set('idx','1'))),
        ('wrong-layout-owner',xml_patch('ppt/slides/_rels/slide3.xml.rels',lambda root: rel(root,'slideLayout').set('Target','../slideLayouts/slideLayout3.xml'))),
        ('wrong-theme-owner',xml_patch('ppt/slideMasters/_rels/slideMaster2.xml.rels',lambda root: rel(root,'theme').set('Target','../theme/theme4.xml'))),
        ('unhidden-page',xml_patch('ppt/slides/slide2.xml',lambda root: root.set('show','1'))),
        ('missing-picture',xml_patch('ppt/slides/slide1.xml',lambda root: root.find('p:cSld/p:spTree',NS).remove(object_named(root,'image-0')))),
        ('changed-local-text',xml_patch('ppt/slides/slide1.xml',lambda root: object_named(root,'body-0').find('.//a:t',NS).__setattr__('text','Other'))),
        ('wrong-notes-backlink',xml_patch('ppt/notesSlides/_rels/notesSlide1.xml.rels',lambda root: rel(root,'slide').set('Target','../slides/slide2.xml'))),
        ('one-emu-frame',xml_patch('ppt/slides/slide1.xml',lambda root: object_named(root,'body-0').find('p:spPr/a:xfrm/a:off',NS).set('x','381001'))),
        ('changed-image-payload',payload),
        ('duplicate-drawing-id',xml_patch('ppt/slides/slide1.xml',lambda root: object_named(root,'body-0').find('.//p:cNvPr',NS).set('id','2'))),
        ('wrong-picture-embed',xml_patch('ppt/slides/slide1.xml',lambda root: object_named(root,'image-0').find('.//a:blip',NS).set('{'+c.R+'}embed','rId1'))),
        ('changed-slide-size',xml_patch('ppt/presentation.xml',lambda root: root.find('p:sldSz',NS).set('cx','9144001'))),
        ('missing-target-prompt',xml_patch(target_layout,lambda root: root.find('p:cSld/p:spTree',NS).remove(object_named(root,'new-body')))),
        ('changed-major-font',xml_patch(n.target(positive,'ppt/slideMasters/slideMaster3.xml','theme'),lambda root: root.find('a:themeElements/a:fontScheme/a:majorFont/a:latin',NS).set('typeface','Arial'))),
    ]
    negatives=[]
    for name,mutate in cases:
        parts=copy.copy(positive);mutate(parts);file=out/(name+'.pptx')
        with zipfile.ZipFile(file,'w',compression=zipfile.ZIP_DEFLATED) as archive:
            for part,body in parts.items(): archive.writestr(part,body)
        try: graph(n.payload(file),media_sha)
        except (AssertionError,KeyError) as error: negatives.append({'name':name,'file':str(file),'bytes':file.stat().st_size,'sha256':a.sha(file.read_bytes()),'reason':str(error)})
        else: raise AssertionError('Accepted corrupt actual ZIP '+name)
    for name in ['covered-picture','swapped-pages','missing-hidden-page','same-raster-font-advance']:
        doc=fitz.open(directory/'first.pdf')
        if name=='covered-picture': doc[0].draw_rect(fitz.Rect(doc[0].get_image_info()[0]['bbox']),fill=(1,1,1),color=(1,1,1))
        elif name=='swapped-pages': doc.select([2,1,0])
        elif name=='missing-hidden-page': doc.select([0,2])
        else:
            span=doc[0].get_texttrace()[0];needed={(char[0],char[1])for char in span['chars']}
            programs=[p.font_pairs(doc,font)[1]for font in doc[0].get_fonts()if font[3].split('+')[-1]==span['font'] and needed<=p.font_pairs(doc,font)[0]]
            assert len(programs)==1;program=programs[0];gid=span['chars'][0][1];changed=p.m.alter_font_advance(program,gid)
            old,new=p.m.glyph_program(program,gid),p.m.glyph_program(changed,gid);assert new==(old[0],old[1]+1,old[2],old[3])
            streams=[xref for xref in range(1,doc.xref_length())if doc.xref_is_stream(xref) and doc.xref_stream(xref)==program];assert len(streams)==1;doc.update_stream(streams[0],changed)
        file=out/(name+'.pdf');doc.save(file)
        try: assert p.page_records(file)==first,'Changed actual page records'
        except AssertionError as error: negatives.append({'name':name,'file':str(file),'bytes':file.stat().st_size,'sha256':a.sha(file.read_bytes()),'reason':str(error)})
        else: raise AssertionError('Accepted corrupt actual PDF '+name)
        if name=='same-raster-font-advance':
            patched,original=fitz.open(file),fitz.open(directory/'first.pdf')
            assert all(patched[i].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples==original[i].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples for i in range(3))
            assert p.glyph_trace(patched,patched[0])!=p.glyph_trace(original,original[0]),'Missed font metric damage behind identical pixels'
    first_parts=c.full(positive);source_parts=c.full(source)
    first_save={'added':sorted(first_parts.keys()-source_parts.keys()),'removed':sorted(source_parts.keys()-first_parts.keys()),'changed':[{'part':part,'sourceSha256':a.sha(source[part]),'nativeSha256':a.sha(positive[part])}for part in sorted(source_parts.keys()&first_parts.keys())if source_parts[part]!=first_parts[part]]}
    artifacts=[{'file':str(file),'bytes':file.stat().st_size,'sha256':a.sha(file.read_bytes())}for file in sorted(directory.iterdir())if file.is_file()and not file.name.startswith('~$')]
    deps=[Path(__file__),HERE/'layout-rebind-native.py',HERE/'master-cell-properties-native.py',HERE/'master-artwork-native.py',HERE/'multi-master-native.py',HERE/'multi-master-operations.py']
    result={'schemaVersion':1,'recordedAt':datetime.now(timezone.utc).isoformat(),'verified':True,'nativeIndexGateAccepted':True,'completeTemplateRebindAccepted':False,'fullFirstSaveEquivalence':False,'directory':str(directory),'sourceSha256':manifest['sourceSha256'],'beforeSha256':manifest['beforeSha256'],'rounds':rounds,'firstPages':first,'firstSaveDifferences':first_save,'negatives':negatives,'artifacts':artifacts,'dependencies':[{'file':str(file),'sha256':a.sha(file.read_bytes())}for file in deps],'scope':'Named authored zero/max/index-disambiguation and mixed implicit allocation controls: exact source/native prompt/local identities, owners, frames, content, image and notes/hidden relationships. Native idx omission is accepted only as schema-default zero. Preserve/unselected used glyph programs/positions and picture probes stay exact; only the expected master accent changes on preserve. All PDF records and package parts remain exact through three native reopen rounds except validated clock/TotalTime values. First-save differences outside the named graph are recorded, not universally declared source-equivalent. No full native-template import, general reset, cross-document identity mapping or source write authority.'}
    Path('docs/testing/placeholder-index-native-evidence.json').write_text(json.dumps(result,indent=2)+'\n')
    print(json.dumps({'verified':True,'nativeIndexGateAccepted':True,'rounds':3,'zipNegatives':len(cases),'pdfNegatives':4}))


if __name__ == '__main__': main()
