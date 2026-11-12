"""Add paired explicit paragraph-direction controls to a generated deck."""
import re
import sys
import zipfile
from pathlib import Path

folder = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/bidi-direction-native')
names = {'latin-first-ltr', 'latin-first-rtl', 'rtl-first-ltr', 'rtl-first-rtl'}
seen = set()
with zipfile.ZipFile(folder / 'base.pptx') as source, zipfile.ZipFile(folder / 'direction.pptx', 'w') as target:
    for info in source.infolist():
        data = source.read(info.filename)
        if info.filename == 'ppt/slides/slide1.xml':
            xml = data.decode('utf-8')
            def replace(shape):
                fragment = shape.group()
                name = re.search(r'<p:cNvPr[^>]*name="([^"]+)', fragment)
                assert name and name.group(1) in names and name.group(1) not in seen
                name = name.group(1); seen.add(name)
                assert fragment.count('<a:pPr ') == 1
                return fragment.replace('<a:pPr ', f'<a:pPr rtl="{int(name.endswith("-rtl"))}" ', 1)
            xml, count = re.subn(r'<p:sp>.*?</p:sp>', replace, xml, flags=re.S)
            assert count == 4 and seen == names
            data = xml.encode('utf-8')
        target.writestr(info, data)
