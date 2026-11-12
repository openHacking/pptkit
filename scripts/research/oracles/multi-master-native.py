"""Independent ZIP/ET/PDF admission of the captured multi-master authoring sample."""
import collections
import copy
import hashlib
import json
import posixpath
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as E
import fitz
from PIL import Image

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'p': P, 'a': A}
sha = lambda b: hashlib.sha256(b).hexdigest()
def root(parts, part): return E.fromstring(parts[part])
def payload(file):
    with zipfile.ZipFile(file) as z: return {name: z.read(name) for name in z.namelist()}
def relname(part): return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels') if part else '_rels/.rels'
def relations(parts, part):
    return {n.get('Id'): {'type': n.get('Type'), 'mode': n.get('TargetMode', 'Internal'), 'target': n.get('Target') if n.get('TargetMode') == 'External' else posixpath.normpath(posixpath.join(posixpath.dirname(part), n.get('Target'))).lstrip('/')} for n in root(parts, relname(part))} if relname(part) in parts else {}
def target(parts, part, kind):
    found = [r for r in relations(parts, part).values() if r['type'].endswith('/' + kind)]
    assert len(found) == 1 and found[0]['mode'] == 'Internal', ('Ambiguous/missing/internal edge', part, kind)
    assert found[0]['target'] in parts
    return found[0]['target']
def graph(parts, expected_counts=(4, 5)):
    main = target(parts, '', 'officeDocument'); presentation = root(parts, main)
    aliases = collections.defaultdict(set)
    def add(label, part): assert part in parts; aliases[part].add(label)
    add('presentation', main); add('theme/presentation', target(parts, main, 'theme'))
    masters, layouts, rows = {}, {}, []
    edges = relations(parts, main)
    for ref in presentation.findall('p:sldMasterIdLst/p:sldMasterId', NS):
        id = ref.get('id'); part = edges[ref.get('{'+R+'}id')]['target']; assert id not in masters
        masters[id] = part; add('master/'+id, part); add('theme/master/'+id, target(parts, part, 'theme'))
        for layout in root(parts, part).findall('p:sldLayoutIdLst/p:sldLayoutId', NS):
            lid = layout.get('id'); lp = relations(parts, part)[layout.get('{'+R+'}id')]['target']; assert lid not in layouts
            assert target(parts, lp, 'slideMaster') == part, 'Layout backlink crossed master domain'
            layouts[lid] = lp; add('layout/'+lid, lp)
    assert set(masters).isdisjoint(layouts), 'PowerPoint master/layout numeric IDs collide'
    assert (len(masters), len(layouts)) == expected_counts, 'Lost/unexpected explicit/default/unused owner'
    note_refs = presentation.findall('p:notesMasterIdLst/p:notesMasterId', NS); assert len(note_refs) == 1
    nm = edges[note_refs[0].get('{'+R+'}id')]['target']; add('notes-master', nm); add('theme/notes', target(parts, nm, 'theme'))
    for i, ref in enumerate(presentation.findall('p:sldIdLst/p:sldId', NS)):
        id = ref.get('id'); part = edges[ref.get('{'+R+'}id')]['target']; add('slide/'+id, part)
        lp = target(parts, part, 'slideLayout'); mp = target(parts, lp, 'slideMaster'); tp = target(parts, mp, 'theme')
        note = target(parts, part, 'notesSlide'); add('notes/'+id, note)
        assert target(parts, note, 'slide') == part and target(parts, note, 'notesMaster') == nm
        rows.append({'page': i+1, 'id': id, 'slide': part, 'layout': lp, 'master': mp, 'theme': tp, 'notes': note,
                     'layoutId': next(k for k,v in layouts.items() if v == lp), 'masterId': next(k for k,v in masters.items() if v == mp)})
    assert len(rows) == 4
    return aliases, rows

def guide_extension(node):
    expected = E.fromstring('<p:extLst xmlns:p="'+P+'" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst/></p:ext></p:extLst>')
    return plain(node) == plain(expected)
def plain(node): return (node.tag, sorted(node.attrib.items()), node.text or '', tuple(plain(c) for c in node))
def semantic(parts, expected_counts=(4, 5)):
    aliases, rows = graph(parts, expected_counts)
    label = lambda part: '|'.join(sorted(aliases.get(part, {part})))
    result = {}
    types = root(parts, '[Content_Types].xml')
    assert types.tag == '{http://schemas.openxmlformats.org/package/2006/content-types}Types'
    result['content-types'] = sorted((n.tag, n.get('Extension') if n.get('Extension') is not None else label(n.get('PartName').lstrip('/')), n.get('ContentType')) for n in types)
    def canonical(node, part, parent=None):
        attrs = dict(node.attrib)
        if node.tag == '{'+A+'}pPr' and attrs.get('indent') == '-0': attrs['indent'] = '0'
        for key, value in list(attrs.items()):
            if key.startswith('{'+R+'}'):
                edge = relations(parts, part)[value]
                attrs[key] = (edge['type'], edge['mode'], label(edge['target']) if edge['mode'] == 'Internal' else edge['target'])
        children = []
        for child in node:
            if node.tag == '{'+P+'}cNvSpPr' and child.tag == '{'+A+'}spLocks' and len(child) == 0:
                ph = parent.find('p:nvPr/p:ph', NS) if parent is not None else None
                image_locks = ph is not None and ph.get('type') == 'sldImg' and child.attrib == {'noGrp':'1','noRot':'1','noChangeAspect':'1'}
                if child.attrib == {'noGrp':'1'} or image_locks: continue
            if node.tag == '{'+P+'}presentation' and child.tag == '{'+P+'}extLst' and guide_extension(child): continue
            children.append(canonical(child, part, node))
        return (node.tag, sorted(attrs.items()), node.text or '', children)
    for part in aliases:
        result[label(part)] = canonical(root(parts, part), part)
        result[label(part)+'/relationships'] = sorted((e['type'], e['mode'], label(e['target']) if e['mode'] == 'Internal' else e['target']) for e in relations(parts, part).values())
    return result, rows

def check_package(file, expected):
    parts = payload(file); actual, rows = semantic(parts)
    assert actual.keys() == expected.keys()
    different = [p for p in actual if actual[p] != expected[p]]
    assert not different, ('Semantic native content changed', different)
    return rows

def pdf_check(file, baseline=None):
    document = fitz.open(file); assert len(document) == 4
    other = fitz.open(baseline) if baseline else None
    expected_fonts = ['Georgia', 'TimesNewRomanPSMT', 'Georgia', 'ArialMT']
    foreground = ['CC2200', '0044CC', 'CC2200', '112233']; backgrounds = ['F8E8E0', 'E0E8F8', 'F8E8E0', 'FFFFFF']
    rows = []
    for i, page in enumerate(document):
        assert page.get_text().strip() == 'Owned ' + str(i), 'Wrong page/text'
        spans = [s for b in page.get_text('dict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans']]
        assert len(spans) == 1 and spans[0]['font'] == expected_fonts[i] and spans[0]['color'] == int(foreground[i], 16)
        pix = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False); im = Image.frombytes('RGB', [pix.width,pix.height], pix.samples)
        rgb = lambda value: tuple(bytes.fromhex(value))
        assert im.getpixel((pix.width//2,pix.height//2)) == rgb(backgrounds[i]), 'Wrong inherited background'
        ink_color = tuple(int(channel * 255) for channel in page.get_texttrace()[0]['color'])
        assert dict((color,count) for count,color in im.getcolors(pix.width*pix.height)).get(ink_color,0) > 25, 'Missing actual text ink'
        fonts = sorted((f[3].split('+')[-1], sha(document.extract_font(f[0])[3])) for f in page.get_fonts())
        if other:
            before = other[i]; assert page.rect == before.rect
            assert page.get_texttrace() == before.get_texttrace(), 'Native glyph trace changed'
            assert pix.samples == before.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples, 'Native full page raster changed'
            assert fonts == sorted((f[3].split('+')[-1],sha(other.extract_font(f[0])[3])) for f in before.get_fonts()), 'Actual embedded font program changed'
        rows.append({'page':i+1,'text':page.get_text().strip(),'font':spans[0]['font'],'color':foreground[i],'actualPdfColor':page.get_texttrace()[0]['color'],'rasterInkRgb':ink_color,'fontPrograms':fonts,'rasterSha256':sha(pix.samples)})
    return rows

def inspect(directory):
    original, fixed = payload(directory/'input.pptx'), payload(directory/'fixed.pptx')
    assert original.keys() == fixed.keys()
    assert [p for p in original if original[p] != fixed[p]] == ['ppt/presentation.xml']
    old = root(original,'ppt/presentation.xml'); new = root(fixed,'ppt/presentation.xml')
    for a,b in zip(old.findall('p:sldMasterIdLst/p:sldMasterId',NS),new.findall('p:sldMasterIdLst/p:sldMasterId',NS)): a.set('id',b.get('id'))
    assert plain(old) == plain(new), 'Fix changed more than master IDs'
    expected, source_rows = semantic(fixed); assert [r['masterId'] for r in source_rows] == ['2147483654','2147483655','2147483654','2147483648']
    first = pdf_check(directory/'first-open.pdf'); rounds = []
    first_parts = payload(directory/'round1.pptx')
    excluded = ['ppt/presProps.xml','ppt/viewProps.xml','docProps/core.xml','docProps/app.xml']
    for number in range(1,4):
        file = directory/f'round{number}.pptx'; rows = check_package(file, expected); pdf = directory/f'round{number}.pdf'; pages = pdf_check(pdf,directory/'first-open.pdf')
        parts = payload(file); assert set(parts) == set(fixed)
        if number > 1: assert [p for p in parts if parts[p] != first_parts[p]] == ['docProps/core.xml']
        rounds.append({'file':str(file),'sha256':sha(file.read_bytes()),'pdf':str(pdf),'pdfSha256':sha(pdf.read_bytes()),'domains':rows,'pages':pages})
    negatives = []; out = directory/'negative-controls'; out.mkdir(exist_ok=True)
    def reject(name, parts):
        file = out/(name+'.pptx')
        with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED) as z:
            for part,body in parts.items():z.writestr(part,body)
        try: check_package(file, expected)
        except (AssertionError,KeyError,StopIteration) as error: negatives.append({'name':name,'file':str(file),'sha256':sha(file.read_bytes()),'reason':str(error)})
        else: raise AssertionError('Accepted actual PPTX corruption '+name)
    aliases, rows = graph(first_parts); red = rows[0]
    mutations = [ ('id-collision','ppt/presentation.xml','collision'), ('wrong-layout-owner',relname(red['layout']),'owner'),('wrong-master-theme',relname(red['master']),'theme'),('wrong-color',red['theme'],'color'),('wrong-font',red['theme'],'font'),('wrong-background',red['master'],'background'),('wrong-slide-text',red['slide'],'text'),('wrong-notes',red['notes'],'text'),('lost-unused-master','ppt/presentation.xml','lost'),('wrong-content-type','[Content_Types].xml','content-type') ]
    for name,part,kind in mutations:
        parts = copy.copy(first_parts); node = root(parts,part)
        if kind == 'collision': node.findall('p:sldMasterIdLst/p:sldMasterId',NS)[1].set('id', rows[0]['layoutId'])
        elif kind == 'lost': node.find('p:sldMasterIdLst',NS).remove(node.findall('p:sldMasterIdLst/p:sldMasterId',NS)[-1])
        elif kind in ['owner','theme']:
            edge = next(n for n in node if n.get('Type').endswith('/'+('slideMaster' if kind=='owner' else 'theme')))
            edge.set('Target', '../slideMasters/'+posixpath.basename(rows[1]['master']) if kind=='owner' else '../theme/'+posixpath.basename(rows[1]['theme']))
        elif kind == 'content-type': next(n for n in node if n.get('PartName') == '/' + red['master']).set('ContentType','application/xml')
        elif kind == 'color': node.find('a:themeElements/a:clrScheme/a:accent1/a:srgbClr',NS).set('val','00FF00')
        elif kind == 'font': node.find('a:themeElements/a:fontScheme/a:minorFont/a:latin',NS).set('typeface','Arial')
        elif kind == 'background': node.find('p:cSld/p:bg/p:bgPr/a:solidFill/a:schemeClr',NS).set('val','accent1')
        else: next(node.iter('{'+A+'}t')).text = 'WRONG CONTENT'
        parts[part] = E.tostring(node,encoding='utf-8',xml_declaration=True); reject(name,parts)
    for name in ['hidden-text','wrong-page-order']:
        file = out/(name+'.pdf'); d = fitz.open(directory/'round1.pdf')
        if name == 'hidden-text':
            box = d[0].search_for('Owned 0')[0]; d[0].draw_rect(box,fill=(1,1,1),color=(1,1,1),overlay=True)
        else:d.select([1,0,2,3])
        d.save(file)
        try:pdf_check(file,directory/'first-open.pdf')
        except AssertionError as error:negatives.append({'name':name,'file':str(file),'sha256':sha(file.read_bytes()),'reason':str(error)})
        else:raise AssertionError('Accepted PDF corruption '+name)
    differential = []
    for name, capture, repair in [('input','first-open-warning',True),('no-unused','no-unused-first-open',True),('high-ids','high-ids-first-open',False),('high-ids-no-unused','high-ids-no-unused-first-open',False),('fixed','fixed-first-open',False)]:
        file = directory/(name+'.pptx'); ax = directory/(capture+'.ax.txt'); screenshot = directory/(capture+'.png'); text = ax.read_text()
        assert str(file) in text
        if repair: assert '修复' in text and '发现' in text
        else: assert '标准窗口 '+name in text and '修复' not in text
        differential.append({'file':str(file),'sha256':sha(file.read_bytes()),'repairPrompt':repair,'repairAccepted':False,'ax':str(ax),'axSha256':sha(ax.read_bytes()),'screenshot':str(screenshot),'screenshotSha256':sha(screenshot.read_bytes())})
    for original_name, high_name in [('input','high-ids'),('no-unused','high-ids-no-unused')]:
        before,after = payload(directory/(original_name+'.pptx')),payload(directory/(high_name+'.pptx'))
        assert before.keys() == after.keys()
        assert [p for p in before if before[p]!=after[p]] == ['ppt/presentation.xml']
        x,y = root(before,'ppt/presentation.xml'),root(after,'ppt/presentation.xml')
        for a,b in zip(x.findall('p:sldMasterIdLst/p:sldMasterId',NS),y.findall('p:sldMasterIdLst/p:sldMasterId',NS)):a.set('id',b.get('id'))
        assert plain(x)==plain(y), 'Differential changed beyond master IDs'
    document = fitz.open(directory/'first-open.pdf'); thumbs=[]
    for page in document:
        pix=page.get_pixmap(matrix=fitz.Matrix(.75,.75),alpha=False);thumbs.append(Image.frombytes('RGB',[pix.width,pix.height],pix.samples))
    w,h=thumbs[0].size;contact=Image.new('RGB',(w*2,h*2),(210,210,210))
    for i,im in enumerate(thumbs):contact.paste(im,((i%2)*w,(i//2)*h))
    contact_file=directory/'native-contact.png';contact.save(contact_file)
    result = {'differentialControls':differential,'nativeContact':{'file':str(contact_file),'sha256':sha(contact_file.read_bytes())},'schemaVersion':1,'recordedAt':datetime.now(timezone.utc).isoformat(),'directory':str(directory),'application':{'name':'Microsoft PowerPoint for Mac','version':'16.104'},'sourceSha256':sha((directory/'input.pptx').read_bytes()),'fixedSha256':sha((directory/'fixed.pptx').read_bytes()),'sourceDomains':source_rows,'firstOpenPdf':{'file':str(directory/'first-open.pdf'),'sha256':sha((directory/'first-open.pdf').read_bytes())},'firstOpenPages':first,'rounds':rounds,'noncontentNativeChanges':{p:{'sourceXml':fixed[p].decode(),'nativeXml':first_parts[p].decode()} for p in excluded},'negatives':negatives,'total':{'firstOpenPages':4,'nativeReopenPages':12,'saveCloseReopenRounds':3,'actualZipNegatives':10,'actualPdfNegatives':2},'scope':'Native no-edit admission for the fixed public authoring source. Full semantic master/layout/theme/slide/notes content is compared with explicit shape noGrp, image noGrp/noRot/noChangeAspect, -0 and empty-guides normalization; relationships are resolved by native identity, not filenames. Metadata/presentation-view properties are separately recorded. Native layout/theme edits and complete artwork/template capability remain separate.'}
    Path('docs/testing/multi-master-native-evidence.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n');print(json.dumps(result['total']))

if __name__=='__main__':inspect(Path(sys.argv[1]).resolve())
