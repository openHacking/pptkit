"""Validate retained observations of disabled native chart-data controls."""
import hashlib
import json
import plistlib
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
BASE = ROOT / '.tmp/native-scatter-coordinate-save/powerpoint-bI3pt5'
RECORD = ROOT / 'docs/testing/chart-scatter-edit-data-availability.json'

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

cases = [
    ('pptkit-saved', 'interactive-edit.pptx', 'interactive-disabled', 'interactive-edit'),
    ('independent-control', '../../native-scatter-research/powerpoint-grywvvg1/scatter-research-round-1.pptx', 'control-disabled', 'scatter-research-round-1'),
]
captures = []
for name, filename, prefix, window in cases:
    file = (BASE / filename).resolve()
    ax_path = BASE / f'{prefix}.ax.txt'
    image_path = BASE / f'{prefix}.jpg'
    ax = ax_path.read_text()
    assert f'Window: "{window}"' in ax and file.as_uri() in ax
    assert 'Details: 图表设计' in ax and '按钮 (disabled) 在 Excel 中编辑数据' in ax
    assert image_path.read_bytes().startswith(b'\xff\xd8\xff')
    captures.append({'sample': name, 'file': str(file.relative_to(ROOT)), 'fileSha256': sha(file), 'axSha256': sha(ax_path), 'screenshotSha256': sha(image_path)})
result = {'scope': 'The PowerPoint Chart Design ribbon showed Edit Data in Excel disabled for both a PPTKit-preservation sample and an independently generated PowerPoint-saved control. Microsoft Excel was not found in /Applications or the Spotlight bundle-ID query at the time. These observations do not establish why the controls were disabled or prove that the files are or are not editable in another environment. No interactive workbook edit was performed. Accessibility and screenshot captures are retained observations, not independently authenticated provenance.',
    'applicationVersion': plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],
    'verifierSha256': sha(Path(__file__)), 'captures': captures}
if sys.argv[1:] == ['--record']:
    RECORD.write_text(json.dumps(result, indent=2) + '\n')
else:
    record = Path(sys.argv[1]) if len(sys.argv) == 2 else RECORD
    assert json.loads(record.read_text()) == result
print(json.dumps({'verified': True, 'disabledControls': 2, 'interactiveEditPerformed': False}))
