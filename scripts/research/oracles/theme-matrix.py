"""Independent fixture oracle: source slot/color programs, native XML, PDF samples."""
import sys,json,zipfile,xml.etree.ElementTree as E,copy,re,math,os,fitz
expected_file,directory=sys.argv[1:]
expected=json.load(open(expected_file));A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def package(file):
    with zipfile.ZipFile(file) as z:return {n:z.read(n) for n in z.namelist() if not n.endswith('/')}
def at(parts,location,root=None):
    node=E.fromstring(parts[location['part'].lstrip('/')]) if root is None else root
    path=location['path'].split('/@')[0]
    for i,piece in enumerate(path.split('/')[1:]):
        match=re.fullmatch(r'([^:]+):([^\[]+)\[(\d+)\]',piece);assert match,piece
        prefix,name,index=match.groups();tag=A+name if prefix=='a' else '{http://schemas.openxmlformats.org/presentationml/2006/main}'+name
        if i==0:assert node.tag==tag and index=='1'
        else:
            found=node.findall(tag);assert len(found)>=int(index),'missing source node';node=found[int(index)-1]
    return node
aliases={'bg1':'lt1','bg2':'lt2','tx1':'dk1','tx2':'dk2'}
def palette(parts):
    root=E.fromstring(parts['ppt/theme/theme1.xml']);return {n.tag.split('}')[-1]:list(n)[0] for n in root.find('.//'+A+'clrScheme')}
def program(node,colors,placeholder=None):
    if node.tag==A+'srgbClr':return node.get('val'),list(node)
    if node.tag==A+'sysClr':return node.get('lastClr'),list(node)
    assert node.tag==A+'schemeClr','unsupported fixture color'
    token=node.get('val')
    if token=='phClr':assert placeholder is not None;rgb,mods=placeholder
    else:rgb,mods=program(colors[aliases.get(token,token)],colors)
    return rgb,mods+list(node)
def evaluate(prog):
    rgb,mods=prog;channels=[int(rgb[i:i+2],16)/255 for i in (0,2,4)];alpha=1;space='rgb'
    for n in mods:
        name=n.tag.split('}')[-1];amount=int(n.get('val'))/100000
        if name in ('tint','shade'):
            if space!='linear':channels=[v/12.92 if v<=.04045 else ((v+.055)/1.055)**2.4 for v in channels]
            space='linear';channels=[max(0,min(1,1-(1-v)*amount if name=='tint' else v*amount)) for v in channels]
        elif name in ('alpha','alphaMod','alphaOff'):alpha=max(0,min(1,amount if name=='alpha' else alpha*amount if name=='alphaMod' else alpha+amount))
        else:raise AssertionError('fixture transform must have an independent oracle: '+name)
    if space=='linear':channels=[12.92*v if v<=.0031308 else 1.055*v**(1/2.4)-.055 for v in channels]
    return ''.join('%02X'%int(v*255+.5) for v in channels),alpha
def source_value(parts,p):
    ref=at(parts,p['source']);index=int(ref.get('idx'));kind='fill' if p['property']=='cell.fill' else 'line'
    if index==0 or kind=='fill' and index==1000:return {'paint':None,'alpha':0,**({'width':0,'dash':'solid'} if kind=='line' else {})}
    theme=E.fromstring(parts['ppt/theme/theme1.xml']);name='lnStyleLst' if kind=='line' else 'bgFillStyleLst' if index>1000 else 'fillStyleLst';offset=index-1001 if kind=='fill' and index>1000 else index-1
    entries=theme.find('.//'+A+name);assert entries is not None and 0<=offset<len(entries),'missing matrix slot';slot=list(entries)[offset]
    fill=slot if kind=='fill' else slot.find(A+'solidFill') or slot.find(A+'noFill')
    # ElementTree empty elements are falsey; select by explicit None checks.
    if kind=='line':
        fill=slot.find(A+'solidFill')
        if fill is None:fill=slot.find(A+'noFill')
    assert fill is not None
    if fill.tag==A+'noFill':result={'paint':None,'alpha':0}
    else:
        assert fill.tag==A+'solidFill';node=list(fill)[0];colors=palette(parts);ph=None
        if node.tag==A+'schemeClr' and node.get('val')=='phClr':ph=program(list(ref)[0],colors)
        value,alpha=evaluate(program(node,colors,ph));result={'paint':value,'alpha':alpha}
    if kind=='line':
        dash=slot.find(A+'prstDash');result.update(width=int(slot.get('w','12700'))/12700,dash=('dot' if dash.get('val')=='sysDot' else dash.get('val')) if dash is not None else 'solid')
    return result
def wanted(p):
    value=p['value'];line=p['property']!='cell.fill';paint=value['paint'] if line else value
    result={'paint':paint.get('color') if paint['type']=='solid' else None,'alpha':paint.get('opacity',1) if paint['type']=='solid' else 0}
    if line:result.update(width=value.get('width',1),dash=value.get('dash','solid'))
    return result
def verify(parts,entry):
    count=0
    for p in entry['provenance']['properties']:
        if not p['property'].startswith('cell.') or p['status']!='mapped':continue
        assert source_value(parts,p)==wanted(p),(p,source_value(parts,p));count+=1
    return count
def native_value(node,colors,line=False):
    if node is None:return {'missingDirectProperty':True}
    fill=node.find(A+'solidFill');none=node.find(A+'noFill')
    if fill is not None:
        value,alpha=evaluate(program(list(fill)[0],colors));result={'paint':value,'alpha':alpha}
    elif none is not None:result={'paint':None,'alpha':0}
    else:return {'missingDirectProperty':True}
    if line:
        dash=node.find(A+'prstDash');result.update(width=int(node.get('w','12700'))/12700,dash=('dot' if dash.get('val')=='sysDot' else dash.get('val')) if dash is not None else 'solid')
    return result
observations=[]
for entry in expected:
    name=entry['name'];source=package(os.path.join(directory,name+'.pptx'));checked=verify(source,entry)
    try:native=package(os.path.join(directory,'native',name+'.pptx'))
    except (FileNotFoundError,zipfile.BadZipFile):native=None
    chosen=next(p for p in entry['provenance']['properties'] if p['property']=='cell.fill');damaged=copy.deepcopy(source);part=chosen['source']['part'].lstrip('/');root=E.fromstring(damaged[part]);ref=at(damaged,chosen['source'],root);ref.set('idx','1' if int(ref.get('idx')) in (0,1000) else '4294967295');damaged[part]=E.tostring(root)
    try:verify(damaged,entry)
    except AssertionError:rejected=True
    else:rejected=False
    assert rejected
    if native is not None:
        rows=E.fromstring(native['ppt/slides/slide1.xml']).find('.//'+A+'tbl').findall(A+'tr');colors=palette(native);diff=[]
        for p in entry['provenance']['properties']:
            if not p['property'].startswith('cell.') or p['status']!='mapped':continue
            node=rows[p['rowIndex']].findall(A+'tc')[p['columnIndex']].find(A+'tcPr');line=p['property']!='cell.fill';side=p['property'].split('.')[-1]
            actual=native_value(node.find(A+{'left':'lnL','right':'lnR','top':'lnT','bottom':'lnB'}[side]) if node is not None and line else node,colors,line);want=wanted(p)
            if actual!=want:diff.append({'row':p['rowIndex'],'column':p['columnIndex'],'property':p['property'],'expected':want,'actual':actual})
    pdf=[]
    for folder,label in [('pdf','resaved'),('source-pdf','source-control')]:
        file=os.path.join(directory,folder,name+'.pdf')
        if not os.path.exists(file):pdf.append({'kind':label,'text':'unavailable','fillPixels':'unavailable'});continue
        book=fitz.open(file);assert len(book)==1;page=book[0];text=page.get_text();missing=[f'R{r}C{c}' for r in range(4) for c in range(4) if f'R{r}C{c}' not in text];assert not missing,missing;pix=page.get_pixmap(matrix=fitz.Matrix(4,4),alpha=False);samples=[]
        for r,row in enumerate(entry['table']['rows']):
            for c,cell in enumerate(row['cells']):
                paint=cell['style']['fill'];alpha=paint.get('opacity',1) if paint['type']=='solid' else 0;rgb=paint.get('color','FFFFFF');want=tuple(int(int(rgb[i:i+2],16)*alpha+255*(1-alpha)+.5) for i in (0,2,4));actual=pix.pixel(int((30+c*70+55)*4),int((30+r*30+15)*4))[:3]
                samples.append({'row':r,'column':c,'expected':want,'actual':actual,'pass':max(abs(a-b) for a,b in zip(want,actual))<=2})
        png=name+'-'+label+'.png';page.get_pixmap(matrix=fitz.Matrix(2,2)).save(os.path.join(directory,png));pdf.append({'kind':label,'text':'pass','fillPixelTolerance':2,'fillPixels':'pass' if all(p['pass'] for p in samples) else 'partial','samples':samples,'png':png})
    observations.append({'name':name,'independentSourceProperties':checked,'deliberateIndexMutationRejected':rejected,'nativeDirectFormatting':{'status':'unavailable' if native is None else 'pass' if not diff else 'partial','differences':diff if native is not None else []},'pdf':pdf})
print(json.dumps({'verified':True,'oracle':'Python ZIP/ElementTree source slot and unrounded color/alpha program + deliberate mutation + native direct XML/PDF text/fill pixel samples','scenes':observations}))
