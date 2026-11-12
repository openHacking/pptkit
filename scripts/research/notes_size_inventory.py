"""Independent ZIP/XML inventory for named boundary-size notes specimens."""
import posixpath
import xml.etree.ElementTree as E
import zipfile

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'

def inventory(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)): raise ValueError('duplicate part name')
        def xml(name): return E.fromstring(archive.read(name))
        def relationships(owner):
            directory, name = posixpath.split(owner)
            rel_path = posixpath.join(directory, '_rels', name + '.rels')
            entries = {}
            for relation in xml(rel_path).findall(REL + 'Relationship'):
                rid = relation.get('Id')
                if rid in entries or relation.get('TargetMode') == 'External': raise ValueError('duplicate or external relation')
                target = relation.get('Target')
                resolved = posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join(directory, target))
                if resolved.startswith('../') or resolved not in names: raise ValueError('missing or unsafe relation')
                entries[rid] = (relation.get('Type').rsplit('/', 1)[-1], resolved)
            return entries
        def by_type(owner, kind):
            matches = [target for name, target in relationships(owner).values() if name == kind]
            if len(matches) != 1: raise ValueError('missing/ambiguous ' + kind)
            return matches[0]
        def text(root): return ''.join(n.text or '' for n in root.iter(A + 't'))
        def image_info(root):
            images = [sp for sp in root.iter(P + 'sp') if any(ph.get('type') == 'sldImg' for ph in sp.iter(P + 'ph'))]
            if len(images) != 1: raise ValueError('missing/ambiguous slide image')
            image = images[0]; placeholder = next(image.iter(P + 'ph')); properties = image.find(P + 'spPr')
            transform = properties.find(A + 'xfrm') if properties is not None else None
            geometry = properties.find(A + 'prstGeom') if properties is not None else None
            frame = None
            if transform is not None:
                offset, extent = transform.find(A + 'off'), transform.find(A + 'ext')
                frame = {k: int(node.get(attr)) for k, node, attr in [('x', offset, 'x'), ('y', offset, 'y'), ('width', extent, 'cx'), ('height', extent, 'cy')]}
            return {'idx': placeholder.get('idx'), 'preset': geometry.get('prst') if geometry is not None else None, 'frame': frame}
        presentation = xml('ppt/presentation.xml'); size = presentation.find(P + 'sldSz'); notes_size = presentation.find(P + 'notesSz')
        rels = relationships('ppt/presentation.xml'); slides = []
        for link in presentation.find(P + 'sldIdLst').findall(P + 'sldId'):
            kind, slide = rels[link.get(R + 'id')]
            if kind != 'slide': raise ValueError('invalid slide identity')
            note = by_type(slide, 'notesSlide'); master = by_type(note, 'notesMaster')
            if by_type(note, 'slide') != slide: raise ValueError('incorrect notes back-link')
            notes_xml = xml(note)
            bodies = [sp for sp in notes_xml.iter(P + 'sp') if any(ph.get('type') == 'body' for ph in sp.iter(P + 'ph'))]
            if len(bodies) != 1: raise ValueError('missing/ambiguous body')
            slides.append({'slide': slide, 'notes': note, 'master': master, 'slideText': text(xml(slide)), 'notesText': text(bodies[0]), 'localImage': image_info(notes_xml)})
        if len({item['notes'] for item in slides}) != len(slides): raise ValueError('shared notes')
        masters = {item['master'] for item in slides}
        if len(masters) != 1: raise ValueError('missing/shared notes master')
        master_xml = xml(next(iter(masters)))
        return {'slideSize': [int(size.get('cx')), int(size.get('cy'))], 'notesSize': [int(notes_size.get('cx')), int(notes_size.get('cy'))],
                'masterImage': image_info(master_xml), 'slides': slides}
