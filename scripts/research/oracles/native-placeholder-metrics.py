#!/usr/bin/env python3
"""Verify bounded font-default/paragraph-mark native counterexamples."""
import hashlib,json,sys
from pathlib import Path
import fitz
probe=Path(sys.argv[1]).resolve();baseline=Path(sys.argv[2]).resolve();families=Path(sys.argv[3]).resolve();courier=Path(sys.argv[4]).resolve()
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def pixels(a,b):
 x=fitz.open(a);y=fitz.open(b);assert len(x)==len(y)==3
 return [x[i].get_pixmap(alpha=False).samples==y[i].get_pixmap(alpha=False).samples for i in range(3)]
checks={}
expected={'only-font':[False,False,True],'only-color':[False,True,True],'only-bold':[False,True,True],'end-arial':[True,True,True],'end-theme':[False,False,True],'paragraph-default-size':[False,False,True],'paragraph-default-style':[False,False,True],'minimal-list':[False,False,True],'layout-size-2800':[False,False,True]}
files={str(baseline):sha(baseline)}
for name,want in expected.items():
 ppt=probe/(name+'.pptx');pdf=probe/(name+'.pdf');found=pixels(baseline,pdf);assert found==want;checks[name]=found;files[str(ppt)]=sha(ppt);files[str(pdf)]=sha(pdf)
for name in ('Georgia','Times-New-Roman','Courier-New'):
 dir=courier if name=='Courier-New' else families;materialized=dir/(name+'-materialize.pdf');fixed=dir/(name+'-fixed.pdf');found=pixels(materialized,fixed);assert found==([False,False,True] if name=='Courier-New' else [True,True,True]);checks[name]=found
 for mode in ('materialize','fixed'):
  for ext in ('pptx','pdf'):
   f=dir/(name+'-'+mode+'.'+ext);files[str(f)]=sha(f)
report={'scope':'manually constructed native paragraph mark probes, not production fix','files':files,'pixelEqualityByCase':checks,'rootThemeFontGeneralFixDisproved':True,'userProfile':{'app':'Microsoft PowerPoint 16.104 (25121423)','hiddenSlideIncluded':True},'productionFontMetricsResolved':False};Path('docs/testing/native-placeholder-metrics-probes.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
