#!/usr/bin/env python3
"""Independent read-only inventory of retained drawing metadata and text."""
import json
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
DEC = '{http://schemas.microsoft.com/office/drawing/2017/decorative}'
URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}'


def inspect(path):
    with zipfile.ZipFile(path) as archive:
        names = sorted((n for n in archive.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)), key=lambda n: int(re.search(r'(\d+)\.xml$', n).group(1)))
        pages = []
        for part in names:
            root = ET.fromstring(archive.read(part))
            objects = []
            for shape in root.iter(f'{P}sp'):
                properties = shape.find(f'{P}nvSpPr/{P}cNvPr')
                if properties is None:
                    continue
                decorations = []
                for extension in properties.findall(f'{A}extLst/{A}ext'):
                    if extension.get('uri', '').upper() == URI:
                        decorations.extend(node.get('val') for node in extension.findall(f'{DEC}decorative'))
                objects.append({'name': properties.get('name'), 'description': properties.get('descr'), 'decorativeValues': decorations, 'text': ''.join(node.text or '' for node in shape.iter(f'{A}t'))})
            pages.append({'part': part, 'objects': objects})
        presentation = ET.fromstring(archive.read('ppt/presentation.xml'))
        return {'firstSlideNumber': presentation.get('firstSlideNum'), 'pages': pages}


if __name__ == '__main__':
    print(json.dumps(inspect(sys.argv[1]), ensure_ascii=False))
