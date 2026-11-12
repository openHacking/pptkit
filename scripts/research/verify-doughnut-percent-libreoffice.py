"""Verify retained LibreOffice headless PDF comparison for hole lexical forms."""
import hashlib,json,subprocess,sys
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2]
control=root/'.tmp/native-doughnut-percent/libreoffice-control/doughnut.pdf'
percent=root/'.tmp/native-doughnut-percent/powerpoint-JYogRH/percent.pdf'
evidence=Path(next((x for x in sys.argv[1:] if x!='--record'),root/'docs/testing/chart-doughnut-percent-libreoffice-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
def boxes(file):
 with fitz.open(file) as doc:
  assert len(doc)==3
  return [{tuple(round(x,3) for x in d['fill']):d['rect'] for d in page.get_drawings() if d.get('fill') and max(d['fill'])>.5 and min(d['fill'])<.1} for page in doc]
a,b=boxes(control),boxes(percent)
for page1,page2 in zip(a,b):
 assert page1.keys()==page2.keys()
 for color in page1: assert all(abs(x-y)<.01 for x,y in zip(page1[color],page2[color]))
blue=(0.,0.,1.);green=(0.,.667,0.);ratio=a[0][blue].width/a[0][green].width
assert abs(ratio-2/3)<.01
native=json.loads((root/'docs/testing/chart-doughnut-powerpoint-evidence.json').read_text())['measurements']['innerToOuterRadius']; assert abs(native-.75)<.01 and native-ratio>.07
version=subprocess.run(['soffice','--version'],capture_output=True,text=True,check=True,timeout=20).stdout.strip()
result={'scope':'One generated control with integer 50 and its 050% single-attribute variant were converted headlessly by LibreOfficeDev to PDF. Colored shape bounding boxes agree across the two PDFs. The inner-to-outer ring boundary ratio is compared with an independently recorded PowerPoint print; LibreOffice resave, interactive editing, other values and other versions are not established.','applicationVersion':version,'verifierSha256':sha(Path(__file__)),'controlPdfSha256':sha(control),'percentPdfSha256':sha(percent),'percentSourceSha256':sha(root/'.tmp/native-doughnut-percent/powerpoint-JYogRH/percent.pptx'),'powerPointEvidenceSha256':sha(root/'docs/testing/chart-doughnut-powerpoint-evidence.json'),'libreOfficeInnerToOuterRadius':ratio,'powerPointInnerToOuterRadius':native}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'slides':3,'libreOfficeRatio':ratio,'powerPointRatio':native}))
