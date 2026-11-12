"""Strict named first-save review for six explicit positive-value master charts; no generic normalizer."""
import copy
import importlib.util
import json
import math
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as E
HERE=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('chart_cell_primitives',HERE/'master-cell-properties-native.py');c=importlib.util.module_from_spec(spec);spec.loader.exec_module(c)
a=c.a;sha=c.sha
C='http://schemas.openxmlformats.org/drawingml/2006/chart';A=c.A
C14='http://schemas.microsoft.com/office/drawing/2007/8/2/chart';C16='http://schemas.microsoft.com/office/drawing/2014/chart'
NS={'c':C,'a':A}
GUID='{C3380CC4-5D6E-409C-BE32-E72D297353CC}';INVERT='{6F2FDCE9-48DA-4B69-8628-5D25D57E5C99}'

def canonical(body,native,guids):
    root=E.fromstring(body);assert root.tag=='{'+C+'}chartSpace'
    language=root.find('c:lang',NS);assert language.attrib=={'val':'zh-CN' if native else 'en-US'};language.set('val','en-US')
    family=next(n for n in root.find('c:chart/c:plotArea',NS) if n.tag in ['{'+C+'}barChart','{'+C+'}lineChart','{'+C+'}pieChart'])
    is_pie=family.tag=='{'+C+'}pieChart';is_bar=family.tag=='{'+C+'}barChart'
    vals=family.findall('.//c:numCache/c:pt/c:v',NS);assert len(vals)==2 and all(float(v.text)>0 for v in vals),'Only the named positive two-point controls are admitted'
    def remove(parent,child):parent.remove(child)
    for parent in list(root.iter()):
        for node in list(parent):
            tag=node.tag
            if native and parent is root and tag=='{'+C+'}style':assert node.attrib=={'val':'2'} and len(node)==0;remove(parent,node)
            elif native and tag=='{'+C+'}autoTitleDeleted':assert node.attrib=={'val':'0'} and len(node)==0;remove(parent,node)
            elif tag in ['{'+C+'}wMode','{'+C+'}hMode']:assert node.attrib=={'val':'factor'} and len(node)==0 and not native;remove(parent,node)
            elif parent.tag=='{'+C+'}manualLayout' and tag in ['{'+C+'}x','{'+C+'}y','{'+C+'}w','{'+C+'}h']:
                assert set(node.attrib)=={'val'} and len(node)==0;value=float(node.get('val'));assert math.isfinite(value);node.set('val',value.hex())
            elif tag=='{'+C+'}overlap':assert is_bar and node.attrib=={'val':'0'} and not native and len(node)==0;remove(parent,node)
            elif native and tag=='{'+C+'}numFmt' and parent.tag=='{'+C+'}catAx':assert node.attrib=={'formatCode':'General','sourceLinked':'0'} and len(node)==0;remove(parent,node)
            elif native and tag=='{'+C+'}invertIfNegative':assert is_bar and node.attrib=={'val':'1'} and len(node)==0;remove(parent,node)
            elif native and tag=='{'+C+'}dLbls':
                expected=['showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize']+(['showLeaderLines'] if is_pie else [])
                assert [n.tag for n in node]==['{'+C+'}'+s for s in expected] and not node.attrib
                for n in node:assert n.attrib=={'val':'1' if n.tag.endswith('}showLeaderLines') else '0'} and len(n)==0
                remove(parent,node)
            elif native and tag=='{'+C+'}bubble3D':assert is_pie and parent.tag=='{'+C+'}dPt' and node.attrib=={'val':'0'} and len(node)==0;remove(parent,node)
            elif native and tag=='{'+C+'}extLst':
                assert not node.attrib and parent.tag in ['{'+C+'}ser','{'+C+'}dPt']
                expected=[INVERT,GUID] if is_bar else [GUID];assert [n.get('uri') for n in node]==expected
                for ext in node:
                    assert ext.attrib=={'uri':ext.get('uri')} and len(ext)==1
                    if ext.get('uri')==GUID:
                        uid=ext[0];assert uid.tag=='{'+C16+'}uniqueId' and len(uid)==0 and set(uid.attrib)=={'val'}
                        value=uid.get('val');assert re.fullmatch(r'\{[0-9A-F]{8}(-[0-9A-F]{4}){3}-[0-9A-F]{12}\}',value) and value not in guids,'Malformed/duplicate chart identity';guids.add(value)
                    else:
                        expected_xml=E.fromstring('<q:invertSolidFillFmt xmlns:q="'+C14+'" xmlns:a="'+A+'"><q:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></q:spPr></q:invertSolidFillFmt>')
                        assert a.n.plain(ext[0])==a.n.plain(expected_xml),'Unexpected negative-value formatting extension'
                remove(parent,node)
    return a.n.plain(root)

def review(parts,source):
    guids=set();records=[]
    for i in range(1,7):
        part=f'ppt/charts/chart{i}.xml';expected=canonical(source[part],False,set());actual=canonical(parts[part],True,guids)
        assert actual==expected,('Unapproved chart change',part)
        records.append({'part':part,'sourceSha256':sha(source[part]),'nativeSha256':sha(parts[part]),'canonicalSha256':sha(json.dumps(actual,sort_keys=True).encode())})
    assert len(guids)==10
    c.owners(parts,source)
    return records

def main():
    prior_path=Path('docs/testing/master-cell-properties-native-evidence.json');prior_bytes=prior_path.read_bytes();prior=json.loads(prior_bytes);directory=Path(prior['directory']);source=a.n.payload(directory/'input.pptx');rounds=[]
    assert sha((directory/'input.pptx').read_bytes())==prior['sourceSha256']
    for record in prior['rounds']:
        file=Path(record['file']);assert sha(file.read_bytes())==record['sha256'];rounds.append({'round':record['round'],'file':str(file),'charts':review(a.n.payload(file),source)})
    out=Path(tempfile.mkdtemp(prefix='chart-review-',dir=directory.parent));positive=a.n.payload(Path(prior['rounds'][0]['file']));bad=[]
    for name in ['color','font-size','one-ulp-layout','layout-mode','visible-label','category-format','negative-fill','bad-guid','duplicate-guid','unknown-extension','cache']:
        parts=copy.copy(positive);part='ppt/charts/chart1.xml';root=a.n.root(parts,part)
        if name=='color':root.find('.//c:ser/c:spPr/a:solidFill/a:srgbClr',NS).set('val','000000')
        elif name=='font-size':next(root.iter('{'+A+'}defRPr')).set('sz','1500')
        elif name=='one-ulp-layout':n=root.find('.//c:manualLayout/c:w',NS);n.set('val',repr(math.nextafter(float(n.get('val')),math.inf)))
        elif name=='layout-mode':E.SubElement(root.find('.//c:manualLayout',NS),'{'+C+'}wMode',{'val':'edge'})
        elif name=='visible-label':root.find('.//c:dLbls/c:showVal',NS).set('val','1')
        elif name=='category-format':root.find('.//c:catAx/c:numFmt',NS).set('formatCode','0%')
        elif name=='negative-fill':root.find('.//{'+C14+'}spPr/a:solidFill/a:srgbClr',NS).set('val','000000')
        elif name=='bad-guid':next(root.iter('{'+C16+'}uniqueId')).set('val','bad')
        elif name=='duplicate-guid':root2=a.n.root(parts,'ppt/charts/chart2.xml');next(root.iter('{'+C16+'}uniqueId')).set('val',next(root2.iter('{'+C16+'}uniqueId')).get('val'))
        elif name=='unknown-extension':E.SubElement(root.find('.//c:extLst',NS),'{'+C+'}ext',{'uri':'unknown'})
        else:root.find('.//c:numCache/c:pt/c:v',NS).text='99'
        parts[part]=E.tostring(root,encoding='utf-8');file=out/(name+'.pptx')
        with zipfile.ZipFile(file,'w',compression=zipfile.ZIP_DEFLATED) as z:
            for p,body in parts.items():z.writestr(p,body)
        try:review(a.n.payload(file),source)
        except AssertionError as error:bad.append({'file':str(file),'bytes':file.stat().st_size,'sha256':sha(file.read_bytes()),'name':name,'reason':str(error)})
        else:raise AssertionError('Accepted actual corrupt '+name)
    deps=[Path(__file__),HERE/'master-cell-properties-native.py',HERE/'master-artwork-native.py',HERE/'multi-master-native.py',HERE/'multi-master-operations.py']
    result={'schemaVersion':1,'recordedAt':datetime.now(timezone.utc).isoformat(),'verified':True,'sourceSha256':prior['sourceSha256'],'nativeEvidence':str(prior_path),'nativeEvidenceSha256':sha(prior_bytes),'rounds':rounds,'negatives':bad,'dependencies':[{'file':str(f),'sha256':sha(f.read_bytes())} for f in deps],'scope':'Six named positive two-point bar/line/pie charts, explicit colors/text, absent data labels and A/B categories. All chart XML exact after listed captured first-save additions/default omissions/binary64-equivalent layout lexical changes. Known locale/style2/unused negative-fill additions are observations for these controls, not universal formatting equivalence or support for negative values, visible labels, locale-dependent categories or native shared chart-data/style editing.'}
    Path('docs/testing/master-chart-first-save-evidence.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n');print(json.dumps({'verified':True,'rounds':3,'chartComparisons':18,'actualZipNegatives':len(bad)}))
if __name__=='__main__':main()
