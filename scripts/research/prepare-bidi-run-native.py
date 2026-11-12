"""Add a:rtl to three non-strong-LTR runs in each named control shape."""
import re
import sys
import zipfile
from pathlib import Path

folder = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/bidi-run-native')
names = {'ltr-control', 'ltr-run-rtl', 'rtl-control', 'rtl-run-rtl'}
seen = set()
with zipfile.ZipFile(folder / 'base.pptx') as source, zipfile.ZipFile(folder / 'run-rtl.pptx', 'w') as target:
    for info in source.infolist():
        data = source.read(info.filename)
        if info.filename == 'ppt/slides/slide1.xml':
            def patch_shape(match):
                shape = match.group(); found = re.search(r'<p:cNvPr[^>]*name="([^"]+)', shape)
                assert found and found.group(1) in names and found.group(1) not in seen
                name = found.group(1); seen.add(name)
                if not name.endswith('-run-rtl'): return shape
                count = 0
                def patch_run(run):
                    nonlocal count
                    value = run.group(); text = re.search(r'<a:t[^>]*>(.*?)</a:t>', value, re.S)
                    assert text
                    if text.group(1) in ('مرحبا', ' - ', '123'):
                        assert value.count('</a:rPr>') == 1
                        count += 1; return value.replace('</a:rPr>', '<a:rtl/></a:rPr>')
                    return value
                shape = re.sub(r'<a:r>.*?</a:r>', patch_run, shape, flags=re.S)
                assert count == 3
                return shape
            xml, count = re.subn(r'<p:sp>.*?</p:sp>', patch_shape, data.decode(), flags=re.S)
            assert count == 4 and seen == names
            data = xml.encode()
        target.writestr(info, data)
