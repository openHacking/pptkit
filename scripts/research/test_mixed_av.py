#!/usr/bin/env python3
"""Exercise the independent mixed-media verifier with damaged packages."""
import importlib.util
import io
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('verify_mixed_av', Path(__file__).with_name('verify-mixed-av.py'))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


def mutate(data, name, old, new):
    source, output = ZipFile(io.BytesIO(data)), io.BytesIO()
    with source, ZipFile(output, 'w') as result:
        for entry in source.infolist():
            content = source.read(entry.filename)
            if entry.filename == name:
                assert old in content
                content = content.replace(old, new, 1)
            result.writestr(entry, content)
    return output.getvalue()


native = (ROOT / 'tests/fixtures/powerpoint-mixed-av.pptx').read_bytes()
generated = (ROOT / 'tests/fixtures/pptkit-mixed-av.pptx').read_bytes()
module.verify(native)
module.verify(generated)
slide = 'ppt/slides/slide1.xml'
rels = 'ppt/slides/_rels/slide1.xml.rels'
cases = [
    (slide, b'cmd="playFrom(0.0)"', b'cmd="wrong"'),
    (slide, b'dur="2000"', b'dur="2001"'),
    (slide, b'<p:cTn id="2"', b'<p:cTn id="1"'),
    (rels, b'/relationships/video"', b'/relationships/audio"'),
    ('ppt/media/movie.mp4', b'ftyp', b'xxxx'),
]
for case in cases:
    try:
        module.verify(mutate(generated, *case))
    except (AssertionError, KeyError, ValueError):
        pass
    else:
        raise AssertionError(f'verifier accepted corruption: {case}')
print(f'Validated two fixtures and {len(cases)} negative controls')
