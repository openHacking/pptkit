"""Independently check the retained straight-connector movement observation."""
from hashlib import sha256
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
HASHES = {
    "native-connector-move-before.pptx": "236e395583f6dad2a3665c1133d92de3392c187e5c61858aedaee79a62f2e1ba",
    "native-connector-move-after.pptx": "81901be601596dd5f4ee47e2282e1191d6df0edbee34eb7afbe8cf0f4c73fbf2",
}


def inspect(data):
    from io import BytesIO
    with ZipFile(BytesIO(data)) as archive:
        root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
    result = {}
    for tag in ("p:sp", "p:cxnSp"):
        for item in root.findall(".//" + tag, NS):
            ident = item.find(".//p:cNvPr", NS)
            transform = item.find("./p:spPr/a:xfrm", NS)
            if ident is None or transform is None or ident.get("name") in result:
                raise ValueError("missing or duplicate drawing identity/transform")
            off, ext = transform.find("a:off", NS), transform.find("a:ext", NS)
            if off is None or ext is None:
                raise ValueError("missing transform geometry")
            connections = [item.find(".//a:stCxn", NS), item.find(".//a:endCxn", NS)]
            result[ident.get("name")] = {"id": ident.get("id"), "tag": tag, "box": tuple(int(v) for v in (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy"))), "connections": [dict(x.attrib) if x is not None else None for x in connections]}
    return result


def check(before, after):
    for deck in (before, after):
        if set(deck) != {"left", "right", "link"} or deck["link"]["tag"] != "p:cxnSp":
            raise ValueError("unexpected drawing inventory")
        link = deck["link"]
        if link["connections"] != [{"id": deck["left"]["id"], "idx": "3"}, {"id": deck["right"]["id"], "idx": "1"}]:
            raise ValueError("attachment identity or site changed")
        lx, ly, lw, lh = deck["left"]["box"]
        rx, ry, _, rh = deck["right"]["box"]
        x, y, width, height = link["box"]
        if (x, y, x + width, y + height) != (lx + lw, ly + lh // 2, rx, ry + rh // 2):
            raise ValueError("connector endpoints differ from cardinal sites")
    if before["left"]["box"] != after["left"]["box"] or before["right"]["box"][2:] != after["right"]["box"][2:]:
        raise ValueError("unexpected control or target size change")
    if before["right"]["box"][:2] == after["right"]["box"][:2]:
        raise ValueError("target did not move")


def main():
    snapshots = []
    for name, expected in HASHES.items():
        data = (ROOT / "tests/fixtures" / name).read_bytes()
        if sha256(data).hexdigest() != expected:
            raise ValueError(f"hash mismatch: {name}")
        snapshots.append(inspect(data))
    check(*snapshots)
    print("Verified retained source and PowerPoint-saved attachment identities and endpoints")


if __name__ == "__main__":
    main()
