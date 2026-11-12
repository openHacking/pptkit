#!/usr/bin/env python3
"""Audit retained REL-02 files; hashes authenticate files, not historical UI claims."""
import hashlib
import json
import posixpath
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import fitz

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
ROOT = Path(__file__).resolve().parents[2]


def verified(path, expected):
    data = path.read_bytes()
    if len(data) != expected['bytes'] or hashlib.sha256(data).hexdigest() != expected['sha256']:
        raise ValueError(f'Artifact hash or length mismatch: {path}')
    return data


def slides(path):
    with zipfile.ZipFile(path) as archive:
        main = ET.fromstring(archive.read('ppt/presentation.xml'))
        relationships = ET.fromstring(archive.read('ppt/_rels/presentation.xml.rels'))
        targets = {item.attrib['Id']: item.attrib['Target'] for item in relationships.findall(f'{REL}Relationship') if item.attrib.get('TargetMode') != 'External'}
        result = []
        for item in main.findall(f'{P}sldIdLst/{P}sldId'):
            target = targets[item.attrib[f'{R}id']]
            part = posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('ppt', target))
            root = ET.fromstring(archive.read(part))
            paragraphs = [''.join(node.text or '' for node in paragraph.iter(f'{A}t')) for paragraph in root.iter(f'{A}p')]
            result.append({'part': part, 'hidden': root.attrib.get('show') in ('0', 'false'), 'paragraphs': [value for value in paragraphs if value.strip()]})
        return result


def normalize(text):
    return ''.join(text.split())


def audit(record):
    fixture = record['fixture']
    source = ROOT / '.tmp' / 'fixtures' / f'{fixture}.pptx'
    saved = ROOT / '.tmp' / 'compatibility' / 'roundtrip' / f'{fixture}.pptx'
    pdf = ROOT / '.tmp' / 'compatibility' / 'reopened' / f'{fixture}.pdf'
    for path, info in ((source, record['input']), (saved, record['saved']), (pdf, record['reopenedPdf'])):
        verified(path, info)
    source_slides, saved_slides = slides(source), slides(saved)
    document = fitz.open(pdf)
    pages = [page.get_text('text') for page in document]
    visible = [slide for slide in saved_slides if not slide['hidden']]
    compared = []
    if len(visible) == len(pages):
        for index, (slide, page) in enumerate(zip(visible, pages)):
            page_text = normalize(page)
            missing = [text for text in slide['paragraphs'] if normalize(text) not in page_text]
            compared.append({'pdfPage': index + 1, 'slidePart': slide['part'], 'paragraphs': len(slide['paragraphs']), 'missingExactWithoutWhitespace': missing})
    same_structure = len(source_slides) == len(saved_slides)
    slide_comparisons = [{'slideIndex': index + 1, 'hiddenEqual': before['hidden'] == after['hidden'], 'paragraphsExact': before['paragraphs'] == after['paragraphs']} for index, (before, after) in enumerate(zip(source_slides, saved_slides))] if same_structure else []
    return {'fixture': fixture, 'sourceSha256': record['input']['sha256'], 'savedSha256': record['saved']['sha256'], 'pdfSha256': record['reopenedPdf']['sha256'], 'sourceSlides': source_slides, 'savedSlides': saved_slides, 'slideCountEqual': same_structure, 'slideComparisons': slide_comparisons, 'pdfPages': len(pages), 'pageMapping': 'visible-slide-order' if len(visible) == len(pages) else 'unresolved-page-count', 'textObservations': compared}


def main():
    source = ROOT / 'docs/testing/libreoffice-evidence.json'
    record = json.loads(source.read_text())
    result = {'schemaVersion': 1, 'sourceEvidence': str(source.relative_to(ROOT)), 'sourceEvidenceSha256': hashlib.sha256(source.read_bytes()).hexdigest(), 'auditScriptSha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest(), 'scope': 'Hash-verified retained artifact XML paragraphs and reopened PDF extraction only; no interactive, appearance, native editing or repair-dialog claims.', 'pymupdfVersion': fitz.VersionBind, 'fixtures': [audit(item) for item in record['fixtures']]}
    output = Path(sys.argv[1]) if len(sys.argv) == 2 else ROOT / '.tmp/corpus-retained-text-audit.json'
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    for fixture in result['fixtures']:
        print(fixture['fixture'], fixture['pageMapping'], 'pages', fixture['pdfPages'], 'unmatched', sum(len(p['missingExactWithoutWhitespace']) for p in fixture['textObservations']), 'changed slides', sum(not p['paragraphsExact'] or not p['hiddenEqual'] for p in fixture['slideComparisons']))


if __name__ == '__main__':
    main()
