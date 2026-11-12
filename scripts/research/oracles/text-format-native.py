"""Independent admission of the named text-format/reset capture; no broad template claim."""
import collections
import copy
import hashlib
import importlib.util
import json
import posixpath
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as E
import fitz
from PIL import Image, ImageDraw

HERE = Path(__file__).resolve().parent
def load(name, file):
    spec = importlib.util.spec_from_file_location(name, file)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module
n = load('master_primitives', HERE/'multi-master-native.py')
m = load('font_primitives', HERE/'multi-master-operations.py')
P, A, R, NS = n.P, n.A, n.R, n.NS
sha = n.sha

def graph(parts, pages=2, layout_count=3):
    main = n.target(parts, '', 'officeDocument'); presentation = n.root(parts, main)
    aliases = collections.defaultdict(set)
    def add(label, part): assert part in parts; aliases[part].add(label)
    add('presentation', main); add('theme/presentation', n.target(parts, main, 'theme'))
    masters, layouts = {}, {}
    edges = n.relations(parts, main)
    for ref in presentation.findall('p:sldMasterIdLst/p:sldMasterId', NS):
        id = ref.get('id'); part = edges[ref.get('{'+R+'}id')]['target']; assert id not in masters
        masters[id] = part; add('master/'+id, part); add('theme/master/'+id, n.target(parts, part, 'theme'))
        for ref in n.root(parts, part).findall('p:sldLayoutIdLst/p:sldLayoutId', NS):
            lid = ref.get('id'); lp = n.relations(parts, part)[ref.get('{'+R+'}id')]['target']; assert lid not in layouts
            assert n.target(parts, lp, 'slideMaster') == part, 'Wrong layout owner'
            layouts[lid] = lp; add('layout/'+lid, lp)
    assert len(masters) == 1 and len(layouts) == layout_count and set(masters).isdisjoint(layouts), 'Wrong owner graph/ID namespace'
    note_refs = presentation.findall('p:notesMasterIdLst/p:notesMasterId', NS); assert len(note_refs) == 1
    nm = edges[note_refs[0].get('{'+R+'}id')]['target']; add('notes-master', nm); add('theme/notes', n.target(parts, nm, 'theme'))
    rows = []
    for index, ref in enumerate(presentation.findall('p:sldIdLst/p:sldId', NS)):
        id = ref.get('id'); part = edges[ref.get('{'+R+'}id')]['target']; add('slide/'+id, part)
        lp = n.target(parts, part, 'slideLayout'); mp = n.target(parts, lp, 'slideMaster'); tp = n.target(parts, mp, 'theme')
        note = n.target(parts, part, 'notesSlide'); add('notes/'+id, note)
        assert n.target(parts, note, 'slide') == part and n.target(parts, note, 'notesMaster') == nm
        rows.append({'page': index+1, 'id': id, 'slide': part, 'layout': lp, 'master': mp, 'theme': tp, 'notes': note})
    assert len(rows) == pages, 'Wrong page correspondence'
    return aliases, rows

def semantic(parts):
    aliases, rows = graph(parts); label = lambda part: '|'.join(sorted(aliases.get(part, {part})))
    def canonical(node, part, parent=None):
        attrs = dict(node.attrib)
        if node.tag == '{'+A+'}pPr' and attrs.get('indent') == '-0': attrs['indent'] = '0'
        # These are schema defaults, not numeric tolerances or arbitrary attribute drops.
        if node.tag == '{'+A+'}normAutofit' and attrs.get('fontScale') == '100000': del attrs['fontScale']
        if node.tag == '{'+A+'}buAutoNum' and attrs.get('startAt') == '1': del attrs['startAt']
        for key, value in list(attrs.items()):
            if key.startswith('{'+R+'}'):
                edge = n.relations(parts, part)[value]
                attrs[key] = (edge['type'], edge['mode'], label(edge['target']) if edge['mode'] == 'Internal' else edge['target'])
        children = []
        for child in node:
            if node.tag == '{'+P+'}cNvSpPr' and child.tag == '{'+A+'}spLocks' and len(child) == 0:
                ph = parent.find('p:nvPr/p:ph', NS) if parent is not None else None
                image = ph is not None and ph.get('type') == 'sldImg' and child.attrib == {'noGrp':'1','noRot':'1','noChangeAspect':'1'}
                if child.attrib == {'noGrp':'1'} or image: continue
            if node.tag == '{'+P+'}presentation' and child.tag == '{'+P+'}extLst' and n.guide_extension(child): continue
            children.append(canonical(child, part, node))
        return node.tag, sorted(attrs.items()), node.text or '', children
    types = n.root(parts, '[Content_Types].xml')
    result = {'content-types': sorted((node.tag, node.get('Extension') if node.get('Extension') is not None else label(node.get('PartName').lstrip('/')), node.get('ContentType')) for node in types)}
    rels = lambda part: sorted((edge['type'], edge['mode'], label(edge['target']) if edge['mode'] == 'Internal' else edge['target']) for edge in n.relations(parts, part).values())
    result['root-relationships'] = rels('')
    covered = {'[Content_Types].xml', '_rels/.rels'}
    for part in aliases:
        result[label(part)] = canonical(n.root(parts, part), part); result[label(part)+'/relationships'] = rels(part)
        covered.update([part, n.relname(part)])
    for part, body in parts.items():
        if part in covered or part in METADATA: continue
        result[part] = n.plain(n.root(parts, part)) if part.endswith('.xml') else sha(body)
    return result, rows

METADATA = ['ppt/presProps.xml', 'ppt/viewProps.xml', 'docProps/core.xml', 'docProps/app.xml']
def shape(root, name):
    found = [node for node in root.iter('{'+P+'}sp') if node.find('p:nvSpPr/p:cNvPr', NS).get('name') == name]
    assert len(found) == 1, ('Missing/duplicate object', name); return found[0]

def approved_source(parts):
    parts = copy.copy(parts); _, rows = graph(parts); part = rows[0]['slide']; root = n.root(parts, part)
    body = shape(root, 'body'); paragraph = body.find('p:txBody/a:p', NS); runs = paragraph.findall('a:r', NS)
    assert [''.join(run.find('a:t', NS).itertext()) for run in runs] == ['', 'Next', ' Plain']
    assert runs[0].find('a:rPr/a:solidFill/a:srgbClr/a:alpha', NS).get('val') == '0'
    paragraph.remove(runs[0])
    child = shape(root, 'child'); empty = child.find('p:txBody/a:p', NS); runs = empty.findall('a:r', NS)
    assert len(runs) == 1 and not ''.join(runs[0].find('a:t', NS).itertext())
    props, end = runs[0].find('a:rPr', NS), empty.find('a:endParaRPr', NS)
    assert props.attrib == end.attrib and [n.plain(c) for c in props] == [n.plain(c) for c in end], 'Empty insertion style not retained'
    empty.remove(runs[0]); parts[part] = E.tostring(root, encoding='utf-8', xml_declaration=True)
    return semantic(parts)[0]

def package(file, expected, source_parts):
    parts = n.payload(file); assert set(parts) == set(source_parts), 'Unexpected/lost part'
    actual, rows = semantic(parts); assert actual.keys() == expected.keys()
    diff = [key for key in expected if actual[key] != expected[key]]; assert not diff, ('Unapproved semantic change', diff)
    first = n.root(parts, rows[0]['layout']); artwork = shape(first, 'shared-text')
    assert artwork.find('p:nvSpPr/p:nvPr/p:ph', NS) is None, 'Shared artwork became an invisible prompt'
    assert len(first.findall('.//p:ph', NS)) == 1, 'Duplicate native placeholder identity'
    return parts, rows

def decorations(page, span):
    box = span['bbox']; x = (box[0]+box[2])/2
    return [tuple(d['rect']) for d in page.get_drawings() if d['fill'] != (1,1,1) and d['rect'].x0 <= x <= d['rect'].x1 and box[1] <= d['rect'].y0 <= box[3]+2 and d['rect'].height < 3]

def pdf(file, baseline=None):
    document = fitz.open(file); assert len(document) == 2, 'Wrong page count'
    other = fitz.open(baseline) if baseline else None; result = []
    for index, page in enumerate(document):
        assert page.get_text() == ['Shared\n1. Next Plain\nShape\n', 'Other\n'][index], 'Missing/reordered visible text (including Shared)'
        spans = [s for block in page.get_text('dict')['blocks'] for line in block.get('lines', []) for s in line['spans']]
        expected = [('Shared','Arial-BoldItalicMT'),('1. Next','Arial-BoldMT'),(' Plain','ArialMT'),('Shape','Arial-BoldItalicMT')] if index == 0 else [('Other','ArialMT')]
        assert [(s['text'],s['font']) for s in spans] == expected, 'Wrong actual regular/bold/italic face'
        assert all(s['color'] == (0x123456 if index == 0 else 0xABCDEF) for s in spans), 'Wrong actual color'
        if index == 0:
            assert not decorations(page, spans[2]), 'Off underline/strike reactivated over Plain'
            for span in [spans[0], spans[1], spans[3]]: assert len(decorations(page, span)) == 2, ('Missing underline/strike', span['text'])
        pix = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
        if other:
            assert pix.samples == other[index].get_pixmap(matrix=fitz.Matrix(2,2), alpha=False).samples, 'Changed full raster'
            assert page.get_texttrace() == other[index].get_texttrace(), 'Changed raw glyph trace'
            assert page.get_drawings() == other[index].get_drawings(), 'Changed decoration/drawing geometry'
            assert m.fonts(document,page) == m.fonts(other,other[index]), 'Changed embedded font program/metric'
        result.append({'page':index+1,'text':page.get_text(),'spans':spans,'rasterSha256':sha(pix.samples),'traceSha256':sha(json.dumps(page.get_texttrace(),sort_keys=True).encode()),'drawingsSha256':sha(json.dumps(page.get_drawings(),default=list,sort_keys=True).encode()),'fontPrograms':m.fonts(document,page)})
    return result

def inspect(directory):
    assert '版本 16.104 (25121423)' in (directory/'powerpoint-version.ax.txt').read_text()
    source = directory/'fixed-input.pptx'; assert sha(source.read_bytes()) == 'ef3d886dcf09d80e23b58cbbafb67cbb26be84fbb2a5c71797c2eb0a3219cd2d'
    original = n.payload(source); expected = approved_source(original); first = pdf(directory/'fixed-first.pdf'); rounds = []
    first_parts = n.payload(directory/'fixed-round1.pptx')
    for number in [1,2,3]:
        file = directory/f'fixed-round{number}.pptx'; parts, rows = package(file,expected,original); pdf_file = directory/f'fixed-round{number}.pdf'
        pages = pdf(pdf_file,directory/'fixed-first.pdf')
        for suffix in ['saved','reopened']:
            text = (directory/f'fixed-round{number}-{suffix}.ax.txt').read_text(); assert str(file) in text and '修复' not in text
        closed = (directory/f'fixed-round{number}-closed.ax.txt').read_text(); assert str(file) not in closed.splitlines()[0], 'Document not closed'
        if number > 1:
            assert [name for name in parts if parts[name] != first_parts[name]] == ['docProps/core.xml','docProps/app.xml']
            for name, mutable in [('docProps/core.xml', ['{http://purl.org/dc/terms/}modified','{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastPrinted']), ('docProps/app.xml',['{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}TotalTime'])]:
                before, after = n.root(first_parts,name), n.root(parts,name)
                for tag in mutable:
                    a,b = before.find(tag),after.find(tag); assert a is not None and b is not None
                    if tag.endswith('TotalTime'): assert b.text.isdecimal() and int(b.text) >= int(a.text)
                    a.text = b.text
                assert n.plain(before) == n.plain(after), ('Unexpected metadata mutation',name)
        rounds.append({'file':str(file),'sha256':sha(file.read_bytes()),'domains':rows,'pdf':str(pdf_file),'pdfSha256':sha(pdf_file.read_bytes()),'pages':pages})
    # Actual pre-fix capture is a real visibility counterexample, not a fabricated expectation.
    old = fitz.open(directory/'first.pdf'); fixed = fitz.open(directory/'fixed-first.pdf')
    assert 'Shared' not in old[0].get_text()
    assert old[1].get_texttrace() == fixed[1].get_texttrace() and old[1].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples == fixed[1].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples
    out = Path(tempfile.mkdtemp(prefix='negative-',dir=directory)); negatives = []
    def reject_pdf(name, file):
        try: pdf(file,directory/'fixed-first.pdf')
        except AssertionError as error: negatives.append({'name':name,'file':str(file),'sha256':sha(file.read_bytes()),'reason':str(error)})
        else: raise AssertionError('Accepted actual PDF corruption '+name)
    reject_pdf('actual-missing-shared',directory/'first.pdf')
    _, rows = graph(first_parts)
    for name in ['bold','italic','underline','strike','duplicate-placeholder','lost-shared-default','wrong-text','wrong-color','bad-scale','notes-geometry','cross-notes-theme','id-collision','wrong-layout-owner']:
        parts = copy.copy(first_parts); part = rows[0]['slide']; root = n.root(parts,part)
        body = shape(root,'body')
        if name in ['bold','italic','underline','strike']:
            prop = {'bold':'b','italic':'i','underline':'u','strike':'strike'}[name]
            run = next(run for run in body.findall('p:txBody/a:p/a:r',NS) if run.find('a:t',NS).text == ' Plain'); del run.find('a:rPr',NS).attrib[prop]
        elif name in ['duplicate-placeholder','lost-shared-default','wrong-color']:
            part = rows[0]['layout']; root = n.root(parts,part)
            if name == 'duplicate-placeholder': E.SubElement(shape(root,'shared-text').find('p:nvSpPr/p:nvPr',NS),'{'+P+'}ph',{'type':'body','idx':'1'})
            elif name == 'lost-shared-default': shape(root,'same').find('p:txBody/a:p/a:pPr/a:defRPr',NS).set('b','0')
            else: shape(root,'shared-text').find('p:txBody/a:p/a:r/a:rPr/a:solidFill/a:srgbClr',NS).set('val','FF0000')
        elif name == 'wrong-text': body.find('p:txBody/a:p/a:r/a:t',NS).text = 'WRONG'
        elif name == 'bad-scale': body.find('p:txBody/a:bodyPr/a:normAutofit',NS).set('fontScale','99999')
        elif name == 'notes-geometry': part = n.target(parts,rows[0]['notes'],'notesMaster'); root=n.root(parts,part); node=root.find('.//a:xfrm/a:off',NS);node.set('x',str(int(node.get('x'))+1))
        elif name == 'cross-notes-theme': part=n.relname(n.target(parts,rows[0]['notes'],'notesMaster'));root=n.root(parts,part);next(node for node in root if node.get('Type').endswith('/theme')).set('Target','../theme/theme1.xml')
        elif name == 'id-collision': part='ppt/presentation.xml';root=n.root(parts,part);root.find('p:sldMasterIdLst/p:sldMasterId',NS).set('id',n.root(parts,rows[0]['master']).find('p:sldLayoutIdLst/p:sldLayoutId',NS).get('id'))
        else: part=n.relname(rows[0]['layout']);root=n.root(parts,part);next(node for node in root if node.get('Type').endswith('/slideMaster')).set('Target','../notesMasters/notesMaster1.xml')
        parts[part]=E.tostring(root,encoding='utf-8',xml_declaration=True); file=out/(name+'.pptx')
        with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED) as z:
            for path, data in parts.items(): z.writestr(path,data)
        try: package(file,expected,original)
        except (AssertionError,KeyError,StopIteration) as error: negatives.append({'name':name,'file':str(file),'sha256':sha(file.read_bytes()),'reason':str(error)})
        else: raise AssertionError('Accepted actual ZIP corruption '+name)
    for name in ['covered-plain','swapped-pages']:
        doc=fitz.open(directory/'fixed-round1.pdf'); file=out/(name+'.pdf')
        if name=='covered-plain': doc[0].draw_rect(doc[0].search_for('Plain')[0],fill=(1,1,1),color=(1,1,1),overlay=True)
        else: doc.select([1,0])
        doc.save(file);reject_pdf(name,file)
    doc=fitz.open(directory/'fixed-round1.pdf'); page=doc[1]; font=next(f for f in page.get_fonts() if f[3].split('+')[-1]=='ArialMT'); program=doc.extract_font(font[0])[3]; gid=page.get_texttrace()[0]['chars'][0][1]
    altered=m.alter_font_advance(program,gid);before=m.glyph_program(program,gid);after=m.glyph_program(altered,gid);assert after==(before[0],before[1]+1,before[2],before[3])
    streams=[xref for xref in range(1,doc.xref_length()) if doc.xref_is_stream(xref) and doc.xref_stream(xref)==program];assert len(streams)==1
    doc.update_stream(streams[0],altered);file=out/'font-advance.pdf';doc.save(file);patched=fitz.open(file)
    assert patched[1].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples==fixed[1].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples
    reject_pdf('same-raster-font-advance',file)
    negatives[-1].update({'sameRasterAsPositive':True,'metricBefore':before,'metricAfter':after})
    result={'schemaVersion':1,'recordedAt':datetime.now(timezone.utc).isoformat(),'directory':str(directory),'application':'PowerPoint 16.104 (25121423), captured About panel','sourceSha256':sha(source.read_bytes()),'firstOpenPdf':first,'rounds':rounds,'negatives':negatives,'firstSaveChanges':['Default fontScale=100000 omitted','Default startAt=1 omitted','Body leading zero-character alpha-zero run removed','Empty group child run removed, exact full insertion style retained at endParaRPr'],'noncontentFirstSaveChanges':{part:{'before':original[part].decode(),'after':first_parts[part].decode()}for part in METADATA},'laterMetadataChanges':['core modified/lastPrinted timestamps','app TotalTime 0→3→3 only'],'dependencies':[{'file':str(file),'sha256':sha(file.read_bytes())}for file in [Path(__file__),HERE/'multi-master-native.py',HERE/'multi-master-operations.py']],'total':{'rounds':3,'firstOpenPages':2,'reopenedPages':6,'zipNegatives':13,'pdfNegatives':4},'scope':'Named style/reset/layout-artwork native stability only. Source zero-character identity loss is explicit. SVG/native physical list indentation and font/shaping parity, live native template links, broad script-language behavior and source style writes remain open.'}
    # Raw PDF coordinates are retained; reported font size is not used as a
    # graphics-affine scale (the print output has distinct x/y transforms).
    control = fitz.open(directory/'list-control-first.pdf'); cases=json.loads((directory/'list-control.json').read_text())['cases']; assert len(control)==len(cases)==8
    graph(n.payload(directory/'list-control.pptx'),pages=8,layout_count=1)
    observations=[]
    for case,page in zip(cases,control):
        chars=[{'character':c['c'],'origin':c['origin'],'font':span['font'],'reportedFontSize':span['size']} for block in page.get_text('rawdict')['blocks'] for line in block.get('lines',[]) for span in line['spans'] for c in span['chars']]
        xs=[c for c in chars if c['character']=='X']; assert len(xs)==2
        observations.append({'case':case,'text':page.get_text(),'characters':chars,'deltaPdfPt':xs[0]['origin'][0]-xs[1]['origin'][0]})
    result['listControls']={'file':str(directory/'list-control-first.pdf'),'sha256':sha((directory/'list-control-first.pdf').read_bytes()),'observations':observations,'scope':'Eight first-open controls. Large first-line clipping and empty-run marker style are corrected; general marker advances, wraps and full native/font geometry are not admitted.'}
    last=fitz.open(directory/'fixed-round3.pdf');documents=[old,fixed,last];sample=old[0].get_pixmap(matrix=fitz.Matrix(.7,.7),alpha=False);width,height=sample.width,sample.height
    contact=Image.new('RGB',(width*3,(height+22)*2),(220,220,220));draw=ImageDraw.Draw(contact)
    for column,(label,doc)in enumerate(zip(['Failed first','Fixed first','Saved round 3'],documents)):
        for index,page in enumerate(doc):
            pix=page.get_pixmap(matrix=fitz.Matrix(.7,.7),alpha=False);im=Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
            contact.paste(im,(column*width,index*(height+22)+22));draw.text((column*width+5,index*(height+22)+4),label+' / page '+str(index+1),fill=(0,0,0))
    contact_file=directory/'native-contact.png';contact.save(contact_file);result['contact']={'file':str(contact_file),'sha256':sha(contact_file.read_bytes())}
    result['artifacts']=[{'file':str(file),'sha256':sha(file.read_bytes()),'bytes':file.stat().st_size}for file in sorted(directory.rglob('*'))if file.is_file()]
    Path('docs/testing/text-origins-native-evidence.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=list)+'\n');print(json.dumps(result['total']))

if __name__=='__main__': inspect(Path(sys.argv[1]).resolve())
