import sys,json,hashlib,zipfile,xml.etree.ElementTree as E,re,posixpath
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/package/2006/relationships'; O='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def inspect(file):
  with zipfile.ZipFile(file) as z:
    names=z.namelist(); slides=sorted((n for n in names if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)),key=lambda n:int(re.search(r'\d+',n).group()))
    def rels(part):
      rel=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
      if rel not in names: return []
      return [{'id':x.get('Id'),'type':x.get('Type','').split('/')[-1], 'target':posixpath.normpath(posixpath.join(posixpath.dirname(part),x.get('Target'))) if x.get('TargetMode','Internal')=='Internal' else None} for x in E.fromstring(z.read(rel)) if x.tag=='{%s}Relationship'%R]
    def one(items,kind):
      found=[x['target'] for x in items if x['type']==kind and x['target'] is not None]
      return found[0] if len(found)==1 else None
    result=[]
    for part in slides:
      root=E.fromstring(z.read(part)); slideRels=rels(part); layout=one(slideRels,'slideLayout'); master=one(rels(layout),'slideMaster') if layout else None; theme=one(rels(master),'theme') if master else None
      themeRoot=E.fromstring(z.read(theme)) if theme else None; accent=themeRoot.find('.//{%s}clrScheme/{%s}accent2/*'%(A,A)) if themeRoot is not None else None
      clicks=[]
      for click in root.iter('{%s}hlinkClick'%A):
        rid=click.get('{%s}id'%O); found=[r for r in slideRels if r['id']==rid]; clicks.append(found[0] if len(found)==1 else None)
      masterRoot=E.fromstring(z.read(master)) if master else None; layoutRoot=E.fromstring(z.read(layout)) if layout else None
      result.append({'part':part,'text':''.join(n.text or '' for n in root.iter('{%s}t'%A)), 'runSizes':[n.get('sz') for n in root.iter('{%s}rPr'%A)], 'layout':layout,'master':master,'theme':theme,'accent2':accent.get('val') if accent is not None else None,'masterPictures':len(list(masterRoot.iter('{%s}pic'%P))) if masterRoot is not None else None,'layoutPictures':len(list(layoutRoot.iter('{%s}pic'%P))) if layoutRoot is not None else None,'clicks':clicks})
    notes=sorted((n for n in names if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml',n)),key=lambda n:int(re.search(r'\d+',n).group()))
    return {'pages':result,'notes':[{'text':''.join(n.text or '' for n in E.fromstring(z.read(part)).iter('{%s}t'%A)), 'slide':one(rels(part),'slide')} for part in notes], 'media':[{'part':n,'sha256':hashlib.sha256(z.read(n)).hexdigest()} for n in names if n.startswith('ppt/media/') and not n.endswith('/')]}

files=sys.argv[1:]
if len(files)!=4: raise SystemExit('expected input and three PowerPoint saves')
rounds=[inspect(f) for f in files]
first=rounds[0]
for item in rounds:
    pages=item['pages']; notes=item['notes']
    assert [p['text'] for p in pages]==['Page A','Page B']
    assert [p['runSizes'] for p in pages]==[['2000'],['2400']]
    assert [n['text'] for n in notes]==['Notes A','Notes B']
    assert [n['slide'] for n in notes]==['ppt/slides/slide1.xml','ppt/slides/slide2.xml']
    assert pages[0]['layout'] is not None and pages[0]['layout']==pages[1]['layout']
    assert pages[0]['master'] is not None and pages[0]['master']==pages[1]['master']
    assert all(p['accent2']=='245678' and p['masterPictures']==1 for p in pages)
    assert all(len(pages[i]['clicks'])==1 and pages[i]['clicks'][0] is not None and pages[i]['clicks'][0]['type']=='slide' and pages[i]['clicks'][0]['target']==target for i,target in ((0,'ppt/slides/slide2.xml'),(1,'ppt/slides/slide1.xml')))
    assert len(item['media'])==1 and item['media'][0]['sha256']==first['media'][0]['sha256']
print(json.dumps({'scope':'Manual PowerPoint for Mac three Save As and reopen rounds; package inspection only. No interaction, link activation, notes visual verification or broad fidelity claim.', 'oracle':'Python zipfile and ElementTree', 'files':[{'file':f,'sha256':hashlib.sha256(open(f,'rb').read()).hexdigest()} for f in files], 'rounds':rounds},indent=2))
