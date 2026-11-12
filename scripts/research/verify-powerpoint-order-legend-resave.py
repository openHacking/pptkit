"""Independent inspection of the retained PowerPoint resave, not native automation."""
import io
import sys
import xml.etree.ElementTree as ET
import zipfile

ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart', 's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(sys.argv[1]) as deck:
    chart = ET.fromstring(deck.read('ppt/charts/chart1.xml'))
    series = chart.findall('.//c:barChart/c:ser', ns)
    assert [(s.find('c:idx', ns).attrib['val'], s.find('c:order', ns).attrib['val']) for s in series] == [('1', '0'), ('0', '1')]
    names = [s.find('.//c:tx//c:pt/c:v', ns).text for s in series]
    values = [[v.text for v in s.findall('.//c:val//c:pt/c:v', ns)] for s in series]
    categories = [[v.text for v in s.findall('.//c:cat//c:pt/c:v', ns)] for s in series]
    assert names == ['Order round 2', 'Order round 3']
    assert values == [['2.5', '5', '6'], ['3.5', '2', '3']]
    assert categories == [[' Q1 ', 'Category 3', '中文 & two']] * 2
    legend = chart.find('.//c:legendEntry', ns)
    assert legend.find('c:idx', ns).attrib['val'] == '1' and legend.find('c:delete', ns).attrib['val'] == '1'
    books = [name for name in deck.namelist() if name.startswith('ppt/embeddings/') and name.endswith('.xlsx')]
    assert len(books) == 1
    with zipfile.ZipFile(io.BytesIO(deck.read(books[0]))) as workbook:
        sheet = ET.fromstring(workbook.read('xl/worksheets/sheet1.xml'))
        for ref, expected in {'B1': 'Order round 3', 'C1': 'Order round 2', 'A3': 'Category 3'}.items():
            assert sheet.find(f'.//s:c[@r="{ref}"]/s:is/s:t', ns).text == expected
        for ref, expected in {'B2': '3.5', 'C2': '2.5', 'B3': '2', 'C3': '5', 'B4': '3', 'C4': '6'}.items():
            assert sheet.find(f'.//s:c[@r="{ref}"]/s:v', ns).text == expected
