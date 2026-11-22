"""Independently inspect section identities and membership in PPTX samples."""
import json
import sys
import zipfile
from xml.etree import ElementTree

NAMESPACES = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "s": "http://schemas.microsoft.com/office/powerpoint/2010/main",
}


def inspect(filename):
    with zipfile.ZipFile(filename) as archive:
        if archive.testzip() is not None:
            raise ValueError("invalid ZIP entry")
        root = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
    extensions = root.findall("p:extLst/p:ext", NAMESPACES)
    native = [ext for ext in extensions if ext.get("uri") == "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"]
    if len(native) != 1:
        raise ValueError("expected exactly one native section extension")
    sections = native[0].findall("s:sectionLst/s:section", NAMESPACES)
    result = [
        {
            "id": section.get("id"),
            "name": section.get("name"),
            "slides": [int(slide.get("id")) for slide in section.findall("s:sldIdLst/s:sldId", NAMESPACES)],
        }
        for section in sections
    ]
    slide_ids = [int(slide.get("id")) for slide in root.findall("p:sldIdLst/p:sldId", NAMESPACES)]
    if len({section["id"] for section in result}) != len(result):
        raise ValueError("duplicate section identities")
    if [slide for section in result for slide in section["slides"]] != slide_ids:
        raise ValueError("section membership does not match slide order")
    return result


if __name__ == "__main__":
    print(json.dumps([inspect(filename) for filename in sys.argv[1:]]))
