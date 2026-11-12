#!/usr/bin/env python3
"""Negative controls for the retained extended-creation evidence verifier."""
import copy,hashlib,json,subprocess,tempfile,unittest,zipfile
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
VERIFIER=ROOT/'scripts/research/verify-extended-create-libreoffice.py'
EVIDENCE=ROOT/'docs/testing/extended-create-libreoffice-evidence.json'
class EvidenceTests(unittest.TestCase):
    def test_actual_evidence(self):
        self.assertEqual(subprocess.run(['python3',str(VERIFIER)],capture_output=True,text=True).returncode,0)
    def test_altered_evidence_is_rejected(self):
        original=json.loads(EVIDENCE.read_text()); cases=[]
        x=copy.deepcopy(original);x['records'][1]['metadata']['extended']='wrong';cases.append(x)
        x=copy.deepcopy(original);x['records'][1]['sha256']='0'*64;cases.append(x)
        x=copy.deepcopy(original);x['records'][3]=x['records'][1];cases.append(x)
        x=copy.deepcopy(original);x['logs'][0]['status']='failed';cases.append(x)
        x=copy.deepcopy(original);x['sourceAddedParts'].pop();cases.append(x)
        with tempfile.TemporaryDirectory() as directory:
            for i,data in enumerate(cases):
                path=Path(directory)/f'{i}.json';path.write_text(json.dumps(data));result=subprocess.run(['python3',str(VERIFIER),str(path)],capture_output=True,text=True)
                self.assertNotEqual(result.returncode,0, f'accepted corruption {i}')
    def test_rehashed_corrupt_packages_are_rejected(self):
        original=json.loads(EVIDENCE.read_text()); record=next(x for x in original['records'] if x['round']==1 and x['kind']=='edited'); source=ROOT/record['path']
        cases=[
            ('docProps/app.xml','Company &amp; 中文'.encode(), 'Corrupt &amp; 中文'.encode()),
            ('[Content_Types].xml',b'application/vnd.openxmlformats-officedocument.extended-properties+xml',b'application/vnd.openxmlformats-officedocument.custom-properties+xml'),
            ('_rels/.rels',b'extended-properties" Target=',b'extended-properties" TargetMode="External" Target='),
        ]
        with tempfile.TemporaryDirectory(dir=ROOT/'.tmp/extended-create-libreoffice') as directory:
            for i,(part,old,new) in enumerate(cases):
                package=Path(directory)/f'{i}.pptx'
                with zipfile.ZipFile(source) as before,zipfile.ZipFile(package,'w') as after:
                    for info in before.infolist():
                        payload=before.read(info.filename)
                        if info.filename==part:
                            self.assertEqual(payload.count(old),1);payload=payload.replace(old,new)
                        after.writestr(info,payload)
                data=copy.deepcopy(original); target=next(x for x in data['records'] if x['round']==1 and x['kind']=='edited');target['path']=str(package.relative_to(ROOT));target['sha256']=hashlib.sha256(package.read_bytes()).hexdigest()
                manifest=Path(directory)/f'{i}.json';manifest.write_text(json.dumps(data));result=subprocess.run(['python3',str(VERIFIER),str(manifest)],capture_output=True,text=True)
                self.assertNotEqual(result.returncode,0,f'accepted corrupted package {i}')
if __name__=='__main__': unittest.main()
