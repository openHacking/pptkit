"""Independent ZIP/XML check of the three mixed internal-candidate rounds."""
import io
import sys
import zipfile
import xml.etree.ElementTree as ET

before_path, after_path, round_text = sys.argv[1:4]
with_legend = len(sys.argv) == 5 and sys.argv[4] == 'legend'
round_number = int(round_text)
assert round_number in (1, 2, 3)
chart_name = 'ppt/charts/chart1.xml'
book_name = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'
sheet_name = 'xl/worksheets/sheet1.xml'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

with zipfile.ZipFile(before_path) as before, zipfile.ZipFile(after_path) as after:
    assert set(before.namelist()) == set(after.namelist())
    for name in before.namelist():
        if name not in (chart_name, book_name):
            assert before.read(name) == after.read(name), name
    old_chart, new_chart = ET.fromstring(before.read(chart_name)), ET.fromstring(after.read(chart_name))
    old_parent = old_chart.find('.//' + C + 'barChart')
    new_parent = new_chart.find('.//' + C + 'barChart')
    old_series, new_series = old_parent.findall(C + 'ser'), new_parent.findall(C + 'ser')
    assert len(old_series) == len(new_series) == 2
    sid = lambda series: int(series.find(C + 'idx').attrib['val'])
    assert [sid(item) for item in new_series] == [sid(item) for item in reversed(old_series)]
    assert [int(item.find(C + 'order').attrib['val']) for item in new_series] == [0, 1]
    edited_id = sid(old_series[0])
    for item in new_series:
        assert item.find('.//' + C + 'cat//' + C + 'pt[@idx="1"]/' + C + 'v').text == f'Category {round_number}'
        if sid(item) == edited_id:
            assert item.find('.//' + C + 'tx//' + C + 'pt[@idx="0"]/' + C + 'v').text == f'Order round {round_number}'
            assert item.find('.//' + C + 'val//' + C + 'pt[@idx="0"]/' + C + 'v').text == str(round_number + 0.5 if with_legend else 10 + round_number)
    if with_legend:
        old_entry = old_chart.find('.//' + C + 'legendEntry')
        new_entry = new_chart.find('.//' + C + 'legendEntry')
        assert old_entry is not None and new_entry is not None
        old_index = int(old_entry.find(C + 'idx').attrib['val'])
        new_index = int(new_entry.find(C + 'idx').attrib['val'])
        assert sid(old_series[old_index]) == sid(new_series[new_index]) == 0
        assert new_entry.find(C + 'delete').attrib['val'] == '1'
        new_entry.find(C + 'idx').attrib['val'] = old_entry.find(C + 'idx').attrib['val']
    # Restore the permitted fields and physical order, then compare complete XML trees.
    old_by_id = {sid(item): item for item in old_series}
    for item in new_series:
        original = old_by_id[sid(item)]
        item.find(C + 'order').attrib['val'] = original.find(C + 'order').attrib['val']
        item.find('.//' + C + 'cat//' + C + 'pt[@idx="1"]/' + C + 'v').text = original.find('.//' + C + 'cat//' + C + 'pt[@idx="1"]/' + C + 'v').text
        if sid(item) == edited_id:
            for field in ('tx', 'val'):
                path = './/' + C + field + '//' + C + 'pt[@idx="0"]/' + C + 'v'
                item.find(path).text = original.find(path).text
    slots = [i for i, item in enumerate(new_parent) if item.tag == C + 'ser']
    for index, item in zip(slots, sorted(new_series, key=lambda item: [sid(s) for s in old_series].index(sid(item)))):
        new_parent[index] = item
    assert ET.tostring(old_chart) == ET.tostring(new_chart)
    with zipfile.ZipFile(io.BytesIO(before.read(book_name))) as old_book, zipfile.ZipFile(io.BytesIO(after.read(book_name))) as new_book:
        assert set(old_book.namelist()) == set(new_book.namelist())
        for name in old_book.namelist():
            if name != sheet_name:
                assert old_book.read(name) == new_book.read(name), name
        old_sheet, new_sheet = ET.fromstring(old_book.read(sheet_name)), ET.fromstring(new_book.read(sheet_name))
        old_cells = {cell.attrib['r']: cell for cell in old_sheet.iter(S + 'c')}
        new_cells = {cell.attrib['r']: cell for cell in new_sheet.iter(S + 'c')}
        col = 'B' if edited_id == 0 else 'C'
        updates = {col + '1': f'Order round {round_number}', col + '2': str(round_number + 0.5 if with_legend else 10 + round_number), 'A3': f'Category {round_number}'}
        for ref, value in updates.items():
            new_cell, old_cell = new_cells[ref], old_cells[ref]
            new_value = new_cell.find('.//' + S + ('t' if ref.endswith('1') or ref == 'A3' else 'v'))
            old_value = old_cell.find('.//' + S + ('t' if ref.endswith('1') or ref == 'A3' else 'v'))
            assert new_value.text == value
            new_value.text = old_value.text
        assert ET.tostring(old_sheet) == ET.tostring(new_sheet)
