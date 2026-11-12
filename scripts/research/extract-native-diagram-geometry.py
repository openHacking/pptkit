"""Independent optional producer-artifact oracle; not a public format importer."""
import json
import re
import sys
import zipfile
from decimal import Decimal
from pathlib import Path
import xml.etree.ElementTree as E

D = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
T = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
S = "{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
MAX_BYTES = 8 * 1024 * 1024


def xml_bytes(data):
    if len(data) > MAX_BYTES:
        raise ValueError("Native XML artifact exceeds the observation byte budget")
    text = data.decode("utf-8-sig")
    if re.search(r"<!\s*(DOCTYPE|ENTITY)\b", text, re.I):
        raise ValueError("DTD/entity declarations are outside the observation profile")
    root = E.fromstring(text)
    if sum(1 for _ in root.iter()) > 200000:
        raise ValueError("Native XML artifact exceeds the observation node budget")
    return root


def plain_paragraph(element):
    chunks = []
    characters = 0
    tasks = [("node", element)]
    work = 0
    while tasks:
        work += 1
        if work > 200000:
            raise ValueError("Paragraph exceeds the observation work budget")
        kind, value = tasks.pop()
        if kind == "text":
            chunks.append(value)
            characters += len(value)
            if characters > 1000000:
                raise ValueError("Paragraph exceeds the observation text budget")
            continue
        if value.tag == T + "s":
            count = int(value.get(T + "c", "1"))
            if count < 1 or count > 10000:
                raise ValueError("Unsupported native whitespace count")
            chunks.append(" " * count)
            characters += count
        elif value.tag in [T + "tab", T + "line-break"]:
            chunks.append("\t" if value.tag == T + "tab" else "\n")
            characters += 1
        else:
            if value.text:
                chunks.append(value.text)
                characters += len(value.text)
            for child in reversed(list(value)):
                if child.tail:
                    tasks.append(("text", child.tail))
                tasks.append(("node", child))
        if characters > 1000000:
            raise ValueError("Paragraph exceeds the observation text budget")
    return "".join(chunks)


def points(raw):
    match = re.fullmatch(r"([+-]?(?:\d+(?:\.\d*)?|\.\d+))(pt|pc|in|cm|mm)", raw.strip()) if isinstance(raw, str) else None
    if not match:
        return 0.0 if raw == "0" else None
    factors = {"pt": Decimal(1), "pc": Decimal(12), "in": Decimal(72), "cm": Decimal(72) / Decimal("2.54"), "mm": Decimal(72) / Decimal("25.4")}
    value = Decimal(match[1]) * factors[match[2]]
    return float(value) if abs(value) <= 1000000000 else None


def stored_paragraphs(element):
    # Lists are native text structure, not child drawings. Descending through
    # arbitrary drawing nodes would instead duplicate another object's text.
    result = []
    tasks = [element]
    containers = [T + "list", T + "list-item", T + "list-header", T + "section", D + "text-box"]
    work = 0
    while tasks:
        work += 1
        if work > 200000:
            raise ValueError("Stored text traversal exceeds observation budget")
        node = tasks.pop()
        if node.tag in [T + "p", T + "h"]:
            result.append(plain_paragraph(node))
        elif node is element or node.tag in containers:
            for child in reversed(list(node)):
                if child.tag in containers + [T + "p", T + "h"]:
                    tasks.append(child)
    return result


def extract(source, native):
    with zipfile.ZipFile(source) as archive:
        name = "ppt/slides/slide1.xml"
        if archive.getinfo(name).file_size > MAX_BYTES:
            raise ValueError("Source slide exceeds observation budget")
        slide = xml_bytes(archive.read(name))
    frames = [f for f in slide.iter(P + "graphicFrame") if f.find("./" + A + "graphic/" + A + "graphicData").get("uri") == "http://schemas.openxmlformats.org/drawingml/2006/diagram"]
    if len(frames) != 1:
        raise ValueError("Exactly one source diagram frame is required")
    frame = frames[0]
    name = frame.find("./" + P + "nvGraphicFramePr/" + P + "cNvPr").get("name")
    transform = frame.find(P + "xfrm")
    off, ext = transform.find(A + "off"), transform.find(A + "ext")
    frame_points = dict(x=int(off.get("x")) / 12700, y=int(off.get("y")) / 12700, width=int(ext.get("cx")) / 12700, height=int(ext.get("cy")) / 12700)
    with Path(native).open("rb") as file:
        root = xml_bytes(file.read(MAX_BYTES + 1))
    # Restrict to drawing pages; master/notes artwork must not become diagram data.
    pages = list(root.iter(D + "page"))
    if len(pages) != 1:
        raise ValueError("Exactly one native drawing page is required")
    groups = [g for g in pages[0].iter(D + "g") if g.get(D + "name") == name]
    records = []
    if len(groups) == 1:
        group = groups[0]
        ordinals = {element: i + 1 for i, element in enumerate(pages[0].iter())}
        parents = {child: parent for parent in pages[0].iter() for child in parent}
        inherited = []
        parent = parents.get(group)
        while parent is not None:
            if parent.get(D + "transform"):
                inherited.append(parent.get(D + "transform"))
            parent = parents.get(parent)
        tasks = [(group, list(reversed(inherited)))]
        while tasks:
            element, ancestors = tasks.pop()
            transform = element.get(D + "transform")
            if element.tag == D + "g":
                chain = ancestors + ([transform] if transform else [])
                for child in reversed(list(element)):
                    tasks.append((child, chain))
            elif element.tag.startswith(D) and element.tag[len(D):] in ["custom-shape", "frame", "rect", "ellipse", "circle", "line", "connector", "path", "polygon", "polyline"]:
                geometry = element.find(D + "enhanced-geometry")
                text_parent = element.find(D + "text-box") if element.tag == D + "frame" else element
                paragraphs = stored_paragraphs(text_parent) if text_parent is not None else []
                record = dict(ordinal=ordinals[element], kind=element.tag[len(D):], paragraphs=paragraphs, ancestorTransforms=ancestors)
                for key, attribute in [("name", D + "name"), ("x", S + "x"), ("y", S + "y"), ("width", S + "width"), ("height", S + "height"), ("transform", D + "transform")]:
                    if element.get(attribute) is not None:
                        record[key] = element.get(attribute)
                if geometry is not None:
                    for key, attribute in [("preset", "type"), ("textAreas", "text-areas"), ("mirrorHorizontal", "mirror-horizontal"), ("mirrorVertical", "mirror-vertical")]:
                        if geometry.get(D + attribute) is not None:
                            record[key] = geometry.get(D + attribute)
                if len(records) >= 10000:
                    raise ValueError("Native drawing count exceeds observation budget")
                records.append(record)
    oracle = []
    for record in records:
        box = {key: points(record.get(key)) for key in ["x", "y", "width", "height"]}
        transformed = bool(record.get("transform", "").strip() or any(t.strip() for t in record["ancestorTransforms"]))
        status = "transform-source-only" if transformed else "axis-aligned" if all(v is not None for v in box.values()) and box["width"] >= 0 and box["height"] >= 0 else "incomplete"
        oracle.append(dict(ordinal=record["ordinal"], status=status, **({"boundsPoints": box} if status == "axis-aligned" else {})))
    return dict(snapshot=dict(frame=frame_points, shapes=records), groupAssociation=dict(sourceName=name, matchingGroups=len(groups), status="located" if len(groups) == 1 else "missing" if len(groups) == 0 else "ambiguous"), oracle=oracle,
                fontFaces=[dict(e.attrib) for e in root.iter("{urn:oasis:names:tc:opendocument:xmlns:style:1.0}font-face")])


if __name__ == "__main__":
    print(json.dumps(extract(sys.argv[1], sys.argv[2])))
