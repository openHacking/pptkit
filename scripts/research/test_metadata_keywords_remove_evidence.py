#!/usr/bin/env python3
import copy,hashlib,json,subprocess,tempfile,unittest,zipfile
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];EVIDENCE=ROOT/'docs/testing/metadata-keywords-remove-libreoffice-evidence.json';VERIFIER=ROOT/'scripts/research/verify-metadata-keywords-remove-libreoffice.py';BASE=ROOT/'.tmp/metadata-keywords-remove-libreoffice'
class EvidenceTests(unittest.TestCase):
 def test_actual(self):self.assertEqual(subprocess.run(['python3',str(VERIFIER)],capture_output=True).returncode,0)
 def test_bad_evidence(self):
  original=json.loads(EVIDENCE.read_text());cases=[]
  x=copy.deepcopy(original);x['records'][3]['metadata']['keywords']='';cases.append(x)
  x=copy.deepcopy(original);x['records'][1]['sha256']='0'*64;cases.append(x)
  x=copy.deepcopy(original);x['records'][3]=x['records'][1];cases.append(x)
  x=copy.deepcopy(original);x['logs'][0]['exitCode']=1;cases.append(x)
  x=copy.deepcopy(original);x['seed']['sha256']='0'*64;cases.append(x)
  with tempfile.TemporaryDirectory() as directory:
   for i,data in enumerate(cases):
    path=Path(directory)/f'{i}.json';path.write_text(json.dumps(data));self.assertNotEqual(subprocess.run(['python3',str(VERIFIER),str(path)],capture_output=True).returncode,0)
 def test_rehashed_corrupt_packages(self):
  original=json.loads(EVIDENCE.read_text());record=next(x for x in original['records'] if x['round']==1 and x['kind']=='control');source=ROOT/record['path'];cases=[('docProps/core.xml',b'alpha beta',b'alpha zeta'),('[Content_Types].xml',b'application/vnd.openxmlformats-package.core-properties+xml',b'application/vnd.openxmlformats-officedocument.custom-properties+xml')]
  with tempfile.TemporaryDirectory(dir=BASE) as directory:
   for i,(part,old,new) in enumerate(cases):
    package=Path(directory)/f'{i}.pptx'
    with zipfile.ZipFile(source) as before,zipfile.ZipFile(package,'w') as after:
     for info in before.infolist():
      payload=before.read(info.filename)
      if info.filename==part:self.assertEqual(payload.count(old),1);payload=payload.replace(old,new)
      after.writestr(info,payload)
    data=copy.deepcopy(original);target=next(x for x in data['records'] if x['round']==1 and x['kind']=='control');target['path']=str(package.relative_to(ROOT));target['sha256']=hashlib.sha256(package.read_bytes()).hexdigest();path=Path(directory)/f'{i}.json';path.write_text(json.dumps(data));self.assertNotEqual(subprocess.run(['python3',str(VERIFIER),str(path)],capture_output=True).returncode,0)
if __name__=='__main__':unittest.main()
