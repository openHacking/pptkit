#!/usr/bin/env python3
"""Independently check the local PowerPoint theme-change observation."""
import hashlib
import json
import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
DIR = ROOT / '.tmp/corpus-theme-change-20261031'
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def relationship(z, part, kind):
    directory, filename = posixpath.split(part)
    rels = posixpath.join(directory, '_rels', filename + '.rels')
    matches = [r for r in ET.fromstring(z.read(rels)) if r.get('Type', '').endswith('/' + kind)]
    assert len(matches) == 1 and matches[0].get('TargetMode') != 'External', (part, kind)
    return posixpath.normpath(posixpath.join(directory, matches[0].get('Target')))


def inspect(path, expected_theme):
    with ZipFile(path) as z:
        result = []
        for index in (1, 2):
            slide = f'ppt/slides/slide{index}.xml'
            root = ET.fromstring(z.read(slide))
            texts = [t.text or '' for t in root.findall('.//a:t', NS)]
            assert texts == (['Inherited title'] if index == 1 else [])
            assert root.get('show') == (None if index == 1 else '0')
            layout = relationship(z, slide, 'slideLayout')
            master = relationship(z, layout, 'slideMaster')
            theme = relationship(z, master, 'theme')
            theme_name = ET.fromstring(z.read(theme)).get('name')
            assert theme_name == expected_theme, (slide, theme_name)
            result.append({'slide': slide, 'layout': layout, 'master': master,
                           'theme': theme, 'themeName': theme_name,
                           'visible': root.get('show') != '0', 'texts': texts})
        return result


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


if __name__ == '__main__':
    source, saved = DIR / 'theme-layout-source.pptx', DIR / 'theme-layout.pptx'
    for name in ('after-theme.ax.txt', 'after-theme.png', 'reopen.ax.txt', 'reopen.png',
                 'hidden-reopen.ax.txt', 'hidden-reopen.png'):
        assert (DIR / name).is_file(), name
    reopen = (DIR / 'reopen.ax.txt').read_text()
    hidden = (DIR / 'hidden-reopen.ax.txt').read_text()
    assert 'Archway: 所有幻灯片都使用' in reopen and 'Inherited title' in reopen
    assert '第 2 张幻灯片，共 2 张, 隐藏' in reopen
    assert '幻灯片 2 / 2' in hidden
    print(json.dumps({'sourceSha256': sha(source), 'savedSha256': sha(saved),
                      'source': inspect(source, 'PPTKit'), 'saved': inspect(saved, 'Archway'),
                      'captureSha256': {name: sha(DIR / name) for name in
                          ('after-theme.ax.txt', 'after-theme.png', 'reopen.ax.txt', 'reopen.png',
                           'hidden-reopen.ax.txt', 'hidden-reopen.png')}}, indent=2, ensure_ascii=False))
