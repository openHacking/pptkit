"""Independent bounded ZIP/XML oracle for repository chart-save experiments.

No PPTKit reader is imported. This is an evidence verifier, not a public importer.
"""
import io
import json
import math
import posixpath
import re
import sys
import zipfile
from decimal import Decimal
from pathlib import Path
import xml.etree.ElementTree as E

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
STRICT = {P: 'http://purl.oclc.org/ooxml/presentationml/main', C: 'http://purl.oclc.org/ooxml/drawingml/chart', S: 'http://purl.oclc.org/ooxml/spreadsheetml/main', R: 'http://purl.oclc.org/ooxml/officeDocument/relationships'}


def check(condition, reason):
    if not condition:
        raise ValueError(reason)


def package(data):
    check(len(data) <= 64 * 1024 * 1024, 'Archive input exceeds observation budget')
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        items = [i for i in archive.infolist() if not i.is_dir()]
        check(len(items) <= 4096 and sum(i.file_size for i in items) <= 128 * 1024 * 1024, 'Archive expansion exceeds observation budget')
        names = [i.filename for i in items]
        check(len(set(names)) == len(names), 'Duplicate ZIP names are ambiguous')
        check(all(not n.startswith('/') and posixpath.normpath(n) == n and not n.startswith('../') for n in names), 'Unsafe package path')
        return {i.filename: archive.read(i) for i in items}


def xml(data):
    check(len(data) <= 8 * 1024 * 1024, 'XML exceeds observation byte budget')
    text = data.decode('utf-16' if data[:2] in [b'\xff\xfe', b'\xfe\xff'] else 'utf-8-sig')
    check(not re.search(r'<!\s*(DOCTYPE|ENTITY)\b', text, re.I), 'DTD/entity declarations are outside observation scope')
    root = E.fromstring(text)
    check(sum(1 for _ in root.iter()) <= 200000, 'XML exceeds observation node budget')
    pending = [(root, 1)]
    while pending:
        node, depth = pending.pop()
        check(depth <= 128, 'XML exceeds observation depth budget')
        pending.extend((child, depth + 1) for child in node)
    return root


def ns(root, conventional):
    return STRICT[conventional] if root.tag.startswith('{' + STRICT[conventional] + '}') else conventional


def one(node, path):
    found = node.findall(path)
    check(len(found) == 1, 'Expected one source element: ' + path)
    return found[0]


def rels(parts, part):
    name = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels') if part else '_rels/.rels'
    rows = list(xml(parts[name])) if name in parts else []
    check(len({r.get('Id') for r in rows}) == len(rows), 'Duplicate relationship identity')
    return {r.get('Id'): r for r in rows}


def target(part, relation):
    check(relation is not None and relation.get('TargetMode', 'Internal') == 'Internal', 'Expected internal relationship')
    raw = relation.get('Target', '')
    check(raw and '#' not in raw and '?' not in raw, 'Unsupported target identity')
    result = posixpath.normpath(raw.lstrip('/') if raw.startswith('/') else posixpath.join(posixpath.dirname(part), raw))
    check(not result.startswith('../'), 'Target escapes package')
    return result


def numeric(text):
    check(text is not None, 'Missing numeric value')
    number = Decimal(text.strip())
    check(number.is_finite() and math.isfinite(float(number)), 'Nonfinite numeric value')
    return number


def same_number(a, b):
    x, y = numeric(str(a)), numeric(str(b))
    return x == y and (x != 0 or x.is_signed() == y.is_signed())


def inventory(data):
    parts = package(data)
    main_rels = [r for r in rels(parts, '').values() if r.get('Type', '').endswith('/officeDocument')]
    check(len(main_rels) == 1, 'Expected unique presentation entry')
    main = target('', main_rels[0]); root = xml(parts[main]); p = ns(root, P); r = STRICT[R] if p == STRICT[P] else R
    chart_ns = STRICT[C] if p == STRICT[P] else C
    links = rels(parts, main); charts = []
    for slide_index, slide in enumerate(root.findall(f'{{{p}}}sldIdLst/{{{p}}}sldId')):
        slide_part = target(main, links.get(slide.get('{' + r + '}id')))
        slide_root = xml(parts[slide_part]); slide_rels = rels(parts, slide_part)
        for ref in slide_root.iter('{' + chart_ns + '}chart'):
            part = target(slide_part, slide_rels.get(ref.get('{' + r + '}id')))
            chart = xml(parts[part]); c = ns(chart, C)
            plot = one(chart, f'{{{c}}}chart/{{{c}}}plotArea')
            families = [n for n in plot if n.tag in [f'{{{c}}}{kind}Chart' for kind in ['bar', 'line', 'pie']]]
            check(len(families) == 1, 'Expected unique mapped chart family')
            row = {'slideIndex': slide_index, 'part': part, 'family': families[0].tag.split('}')[1][:-5], 'series': [], '_root': chart}
            external = chart.find('{' + c + '}externalData')
            book_relation = rels(parts, part).get(external.get('{' + r + '}id')) if external is not None else None
            book = None
            if book_relation is not None and book_relation.get('TargetMode', 'Internal') == 'Internal':
                book_part = target(part, book_relation); book = package(parts[book_part]); row['workbookPart'] = book_part; row['_book'] = book
                entry = [x for x in rels(book, '').values() if x.get('Type', '').endswith('/officeDocument')]
                check(len(entry) == 1, 'Expected workbook entry')
                book_main = target('', entry[0]); book_root = xml(book[book_main]); s = ns(book_root, S); br = STRICT[R] if s == STRICT[S] else R
                book_rels = rels(book, book_main); sheets = {}
                for sheet in book_root.findall(f'{{{s}}}sheets/{{{s}}}sheet'):
                    sheet_part = target(book_main, book_rels.get(sheet.get('{' + br + '}id')))
                    sheet_root = xml(book[sheet_part]); cells = {}
                    for cell in sheet_root.findall(f'{{{s}}}sheetData/{{{s}}}row/{{{s}}}c'):
                        key = cell.get('r'); check(key and key not in cells, 'Ambiguous worksheet cell')
                        cells[key] = cell
                    check(sheet.get('name') not in sheets, 'Duplicate sheet name')
                    sheets[sheet.get('name')] = (sheet_part, sheet_root, cells)
                row['_sheets'] = sheets; row['_sheet_ns'] = s
            for ser in sorted(families[0].findall('{' + c + '}ser'), key=lambda n: int(one(n, '{' + c + '}order').get('val'))):
                val = one(ser, '{' + c + '}val'); points = val.findall(f'.//{{{c}}}pt')
                ordered = sorted(points, key=lambda n: int(n.get('idx')))
                check([int(n.get('idx')) for n in ordered] == list(range(len(ordered))), 'Cache is not dense and unique')
                atoms = [one(n, '{' + c + '}v') for n in ordered]
                formula = val.find(f'{{{c}}}numRef/{{{c}}}f')
                cells = []
                if book is not None and formula is not None:
                    m = re.fullmatch(r"(?:'((?:[^']|'')+)'|([^'!\[\]]+))!\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?", (formula.text or '').strip())
                    check(m is not None, 'Unsupported local numeric formula')
                    sheet_name = m[1].replace("''", "'") if m[1] else m[2]
                    sheet_part, sheet_root, worksheet = row['_sheets'][sheet_name]
                    check(m[5] in [None, m[3]], 'Oracle numeric ranges require one column')
                    start, end = int(m[4]), int(m[6] or m[4])
                    check(end - start + 1 == len(atoms), 'Workbook range/cache size mismatch')
                    for index in range(start, end + 1):
                        address = f'{m[3]}{index}'; cell = worksheet[address]
                        check(cell.get('t', 'n') == 'n' and cell.find('{' + s + '}f') is None, 'Non-scalar numeric cell')
                        cells.append({'worksheetPart': sheet_part, 'reference': address, '_root': sheet_root, '_node': one(cell, '{' + s + '}v')})
                row['series'].append({'sourceId': int(one(ser, '{' + c + '}idx').get('val')), 'order': int(one(ser, '{' + c + '}order').get('val')), 'values': [float(numeric(n.text)) for n in atoms], 'formula': formula.text if formula is not None else None, 'cells': cells, '_nodes': atoms})
            charts.append(row)
    return {'charts': charts, 'partCount': len(parts), '_parts': parts}


def public(value):
    if isinstance(value, dict):
        return {k: public(v) for k, v in value.items() if not k.startswith('_')}
    if isinstance(value, list):
        return [public(v) for v in value]
    return value


def signature(node, allowed):
    return [node.tag, sorted(node.attrib.items()), '<approved-number>' if node in allowed else node.text, node.tail, [signature(n, allowed) for n in node]]


def verify(before_data, after_data, edits):
    check(isinstance(edits, list) and 0 < len(edits) <= 10000, 'Expected bounded nonempty edit proposals')
    for edit in edits:
        check(isinstance(edit, dict) and all(type(edit.get(k)) is int and edit[k] >= 0 for k in ['chartIndex', 'seriesIndex', 'pointIndex']), 'Invalid numeric edit identity')
        check(type(edit.get('value')) in [int, float, str], 'Invalid numeric edit value')
        numeric(str(edit['value']))
    before, after = inventory(before_data), inventory(after_data)
    check(len(before['charts']) == len(after['charts']), 'Chart count changed')
    allowed_parts = set(); chart_atoms = {}; sheet_atoms = {}; seen = set()
    for edit in edits:
        key = (edit['chartIndex'], edit['seriesIndex'], edit['pointIndex'])
        check(key not in seen, 'Duplicate edit proposal'); seen.add(key)
        ci, si, pi = key
        check(ci < len(before['charts']), 'Chart edit index is out of range')
        a, b = before['charts'][ci], after['charts'][ci]
        check(sum(c['part'] == a['part'] for c in before['charts']) == 1, 'Edited chart source is shared')
        check(a.get('workbookPart') is not None and sum(c.get('workbookPart') == a['workbookPart'] for c in before['charts']) == 1, 'Edited workbook source is shared or absent')
        check(si < len(a['series']) and si < len(b['series']), 'Series edit index is out of range')
        check(pi < len(a['series'][si]['values']) and pi < len(b['series'][si]['values']), 'Point edit index is out of range')
        check(a['part'] == b['part'] and a.get('workbookPart') == b.get('workbookPart') and a['family'] == b['family'], 'Source identity changed')
        x, y = a['series'][si], b['series'][si]
        check(x['formula'] == y['formula'] and len(x['cells']) == len(y['cells']) == len(x['values']), 'Unbound workbook values')
        u, v = x['cells'][pi], y['cells'][pi]
        check(u['worksheetPart'] == v['worksheetPart'] and u['reference'] == v['reference'], 'Cell identity changed')
        # Compare the original XML lexemes, not the rounded float values used
        # in the public inventory. Distinct decimals can map to the same float.
        check(same_number(x['_nodes'][pi].text, u['_node'].text), 'Source cache/workbook disagree')
        check(same_number(y['_nodes'][pi].text, edit['value']) and same_number(v['_node'].text, edit['value']), 'Saved cache/workbook disagree with proposal')
        check(not same_number(x['_nodes'][pi].text, edit['value']), 'Expected actual numeric mutation')
        allowed_parts.update([a['part'], a['workbookPart']])
        chart_atoms.setdefault(ci, [set(), set()])[0].add(x['_nodes'][pi]); chart_atoms[ci][1].add(y['_nodes'][pi])
        sk = (ci, u['worksheetPart']); sheet_atoms.setdefault(sk, [set(), set()])[0].add(u['_node']); sheet_atoms[sk][1].add(v['_node'])
    check(set(before['_parts']) == set(after['_parts']), 'Outer part inventory changed')
    changes = {p for p in before['_parts'] if before['_parts'][p] != after['_parts'][p]}
    check(changes == allowed_parts, 'Unexpected outer payload changes: ' + str(sorted(changes ^ allowed_parts)))
    for ci, atoms in chart_atoms.items():
        a, b = before['charts'][ci], after['charts'][ci]
        check(signature(a['_root'], atoms[0]) == signature(b['_root'], atoms[1]), 'Non-numeric chart XML changed')
        check(set(a['_book']) == set(b['_book']), 'Workbook part inventory changed')
        allowed_sheets = {part for index, part in sheet_atoms if index == ci}
        actual_sheets = {part for part in a['_book'] if a['_book'][part] != b['_book'][part]}
        check(actual_sheets == allowed_sheets, 'Unexpected inner payload changes')
        for part in allowed_sheets:
            left, right = sheet_atoms[(ci, part)]
            # Stored roots include all worksheet attributes and non-value text.
            aa = next(v[1] for v in a['_sheets'].values() if v[0] == part)
            bb = next(v[1] for v in b['_sheets'].values() if v[0] == part)
            check(signature(aa, left) == signature(bb, right), 'Non-numeric worksheet XML changed')
    return {'verified': True, 'changedOuterParts': sorted(changes), 'unchangedOuterPayloads': len(before['_parts']) - len(changes), 'charts': public(after['charts']), 'partCount': after['partCount']}


if __name__ == '__main__':
    if len(sys.argv) == 3 and sys.argv[1] == 'inspect':
        result = public(inventory(Path(sys.argv[2]).read_bytes()))
    elif len(sys.argv) == 5 and sys.argv[1] == 'verify':
        spec = json.loads(Path(sys.argv[4]).read_text())
        result = verify(Path(sys.argv[2]).read_bytes(), Path(sys.argv[3]).read_bytes(), spec)
    else:
        raise SystemExit('Usage: chart-value-save-oracle.py inspect file | verify before after edits.json')
    print(json.dumps(result, allow_nan=False))
