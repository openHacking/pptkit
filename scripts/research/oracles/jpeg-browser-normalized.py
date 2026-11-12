#!/usr/bin/env python3
"""Verify browser-converted output against independent Pillow and retained native samples."""
import importlib.util
from hashlib import sha256
from pathlib import Path
import json
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/'.tmp/jpeg-orientation-native-01'
path=ROOT/'scripts/research/oracles/jpeg-orientation-normalized.py';spec=importlib.util.spec_from_file_location('reference',path);reference=importlib.util.module_from_spec(spec);spec.loader.exec_module(reference)
def sha(b):return sha256(b).hexdigest()
def main():
    names=['orientation6.jpg','browser-normalized.png','browser-normalized-candidate.pptx','browser-normalized-powerpoint.png']
    values=[(BASE/name).read_bytes() for name in names];source,output,pptx,screen=values
    browser=json.loads((BASE/'browser-normalized.json').read_text());producer=json.loads((BASE/'browser-normalized-producer.json').read_text());before=json.loads((ROOT/'docs/testing/jpeg-orientation-native-evidence.json').read_text())
    reference.check(sha(source)==before['files']['orientation6.jpg']==browser['sourceSha256']==producer['sourceJpegSha256'],'source digest')
    reference.check(sha(output)==browser['outputSha256']==producer['normalizedPngSha256'] and sha(pptx)==producer['pptxSha256'],'output digest')
    reference.check(browser['requests']==[] and (browser['width'],browser['height'],browser['sourceOrientation'],browser['mimeType'],browser['orientationNormalization'])==(200,300,6,'image/png','pixels'),'browser result')
    reference.check([w['code'] for w in producer['warnings']]==['asset-orientation-normalized'],'export warning')
    result=reference.inspect(source,output,pptx,screen);bad=0
    for sample in [(source,source,pptx,screen),(source,output,(BASE/'candidate.pptx').read_bytes(),screen),(source,output,pptx,(BASE/'powerpoint.png').read_bytes())]:
        try:reference.inspect(*sample)
        except Exception:bad+=1
    reference.check(bad==3,'negative controls')
    scripts=['scripts/prepare-browser-jpeg-normalization.mjs','scripts/prepare-jpeg-browser-normalized-pptx.mjs','scripts/jpeg-normalization-browser.test.mjs','packages/pptx-exporter/src/adapters/browser-jpeg-normalization.ts']
    result={'scope':'Chromium 149 browser JPEG normalization and PowerPoint 16.104 manual slideshow on one EXIF 6 sample; pixel checks do not independently witness UI actions or guarantee color/crop fidelity.',
      'directory':str(BASE.relative_to(ROOT)),'files':{name:sha(data) for name,data in zip(names,values)},'browserReportSha256':sha((BASE/'browser-normalized.json').read_bytes()),'producerSha256':sha((BASE/'browser-normalized-producer.json').read_bytes()),'originalEvidenceSha256':sha((ROOT/'docs/testing/jpeg-orientation-native-evidence.json').read_bytes()),'referenceOracleSha256':sha(path.read_bytes()),'oracleSha256':sha(Path(__file__).read_bytes()),'scripts':{p:sha((ROOT/p).read_bytes()) for p in scripts},'observation':result,'negativeCasesRejected':bad}
    (ROOT/'docs/testing/jpeg-browser-normalized-evidence.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps(result['observation']))
if __name__=='__main__':main()
