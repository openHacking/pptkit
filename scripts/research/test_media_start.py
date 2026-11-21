#!/usr/bin/env python3
"""Negative controls for the native media start timing comparator."""
import importlib.util
import io
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('verify_media_start', Path(__file__).with_name('verify-media-start.py'))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
files = ['powerpoint-mixed-av.pptx', 'powerpoint-audio-start-auto.pptx', 'powerpoint-audio-start-object-click.pptx']
samples = [(ROOT / 'tests/fixtures' / name).read_bytes() for name in files]
module.verify(*samples)
generated = (ROOT / 'tests/fixtures/pptkit-audio-start-object-click.pptx').read_bytes()
assert module.inspect(generated) == module.inspect(samples[2])


def mutate(data, part, old, new):
    output = io.BytesIO()
    with ZipFile(io.BytesIO(data)) as source, ZipFile(output, 'w') as result:
        for item in source.infolist():
            content = source.read(item.filename)
            if item.filename == part:
                assert old in content
                content = content.replace(old, new, 1)
            result.writestr(item, content)
    return output.getvalue()


slide = 'ppt/slides/slide1.xml'
cases = [
    (1, slide, b'nodeType="afterEffect"', b'nodeType="clickEffect"'),
    (1, slide, b'delay="1000"', b'delay="999"'),
    (2, slide, b'<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="5"', b'<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="4"'),
    (0, slide, b'<p:cTn id="2"', b'<p:cTn id="1"'),
    (1, 'ppt/media/media2.wav', b'RIFF', b'XXXX'),
]
for index, part, old, new in cases:
    altered = list(samples); altered[index] = mutate(altered[index], part, old, new)
    try:
        module.verify(*altered)
    except (AssertionError, KeyError, AttributeError, StopIteration, ValueError):
        continue
    raise AssertionError(f'accepted corrupted sample: {part}, {old!r}')
print(f'Validated {len(cases)} negative controls')


video_auto = (ROOT / 'tests/fixtures/powerpoint-video-start-auto.pptx').read_bytes()
module.verify_video_auto(video_auto)
video_cases = [
    (slide, b'nodeType="afterEffect"', b'nodeType="clickEffect"'),
    (slide, b'delay="2000"', b'delay="1999"'),
    (slide, b'cmd="togglePause"', b'cmd="stop"'),
    (slide, b'<p:cTn id="2"', b'<p:cTn id="1"'),
    ('ppt/media/media1.mp4', b'ftyp', b'xxxx'),
]
for part, old, new in video_cases:
    try:
        module.verify_video_auto(mutate(video_auto, part, old, new))
    except (AssertionError, KeyError, AttributeError, StopIteration, ValueError):
        continue
    raise AssertionError(f'accepted corrupted video sample: {part}, {old!r}')
print(f'Validated video-auto source and {len(video_cases)} negative controls')
