#!/usr/bin/env python3
"""Observe custom property types through LibreOfficeDev headless resaves."""
import hashlib
import json
import platform
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
BASE = ROOT / '.tmp/custom-properties-libreoffice'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
GUID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def inventory(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if len(names) != len(set(names)):
            raise ValueError('duplicate ZIP entries')
        rels = ET.fromstring(z.read('_rels/.rels'))
        matches = [r for r in rels if r.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties']
        if len(matches) != 1 or matches[0].get('TargetMode') == 'External':
            raise ValueError('expected one internal custom properties relationship')
        target = matches[0].get('Target', '').lstrip('/')
        if target != 'docProps/custom.xml':
            raise ValueError('unexpected custom property part')
        root = ET.fromstring(z.read(target))
        if root.tag != f'{{{CUSTOM}}}Properties':
            raise ValueError('unexpected root')
        result = []
        for child in root:
            if child.tag != f'{{{CUSTOM}}}property' or len(child) != 1:
                raise ValueError('unexpected property')
            value = child[0]
            result.append({'name': child.get('name'), 'pid': child.get('pid'), 'fmtid': child.get('fmtid'), 'type': value.tag, 'value': ''.join(value.itertext())})
        return result

def main():
    previous = json.loads((ROOT / 'docs/testing/core-metadata-libreoffice-evidence.json').read_text())
    seed = ROOT / previous['generated']['path']
    if digest(seed) != previous['generated']['sha256']:
        raise ValueError('seed hash mismatch')
    BASE.mkdir(parents=True, exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix='run-', dir=BASE))
    source = run / 'input.pptx'
    ET.register_namespace('', CUSTOM)
    ET.register_namespace('vt', VT)
    props = ET.Element(f'{{{CUSTOM}}}Properties')
    samples = [('Text', 'lpwstr', '中文 & text'), ('Boolean', 'bool', 'false'), ('Integer', 'i4', '-2147483648'), ('Real', 'r8', '1.25'), ('Date', 'filetime', '2020-01-02T03:04:05Z')]
    for pid, (name, kind, value) in enumerate(samples, 2):
        item = ET.SubElement(props, f'{{{CUSTOM}}}property', {'fmtid': GUID, 'pid': str(pid), 'name': name})
        ET.SubElement(item, f'{{{VT}}}{kind}').text = value
    with zipfile.ZipFile(seed) as before, zipfile.ZipFile(source, 'w') as after:
        if 'docProps/custom.xml' in before.namelist():
            raise ValueError('seed already has custom properties')
        for info in before.infolist():
            data = before.read(info.filename)
            if info.filename == '_rels/.rels':
                marker = b'</Relationships>'
                if data.count(marker) != 1 or b'Id="customProbe"' in data: raise ValueError('unexpected root relationships')
                data = data.replace(marker, b'<Relationship Id="customProbe" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>' + marker)
            elif info.filename == '[Content_Types].xml':
                marker = b'</Types>'
                if data.count(marker) != 1: raise ValueError('unexpected content types')
                data = data.replace(marker, b'<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' + marker)
            after.writestr(info, data)
        after.writestr('docProps/custom.xml', ET.tostring(props, encoding='utf-8', xml_declaration=True))
    records = [{'round': 0, 'path': str(source.relative_to(ROOT)), 'sha256': digest(source), 'properties': inventory(source)}]
    logs = []
    current = source
    for number in range(1, 4):
        out = run / f'round-{number}'
        out.mkdir()
        profile = (run / f'profile-{number}').as_uri()
        process = subprocess.run([str(SOFFICE), f'-env:UserInstallation={profile}', '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', str(out), str(current)], capture_output=True, text=True, check=True)
        current = out / current.name
        if not current.is_file(): raise ValueError('conversion did not create output')
        logs.append({'round': number, 'stdout': process.stdout, 'stderr': process.stderr})
        records.append({'round': number, 'path': str(current.relative_to(ROOT)), 'sha256': digest(current), 'properties': inventory(current)})
    result = {'scope': 'One independently authored fixture, three headless resaves; not interactive PowerPoint.', 'application': subprocess.run([str(SOFFICE), '--version'], capture_output=True, text=True, check=True).stdout.strip(), 'platform': platform.platform(), 'runnerSha256': digest(Path(__file__)), 'seed': previous['generated'], 'records': records, 'logs': logs}
    (ROOT / 'docs/testing/custom-properties-libreoffice-evidence.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps({'status': 'complete', 'records': records}, ensure_ascii=False))

if __name__ == '__main__': main()
