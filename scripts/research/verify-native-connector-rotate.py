"""Check an observed PowerPoint target rotation independently of PPTKit geometry."""
import importlib.util
import math
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

path = Path(__file__).with_name("verify-native-connector-move.py")
spec = importlib.util.spec_from_file_location("connector_move", path)
base = importlib.util.module_from_spec(spec)
spec.loader.exec_module(base)
HASHES = {
    "native-connector-move-before.pptx": "236e395583f6dad2a3665c1133d92de3392c187e5c61858aedaee79a62f2e1ba",
    "native-connector-rotate-after.pptx": "d38cc889c4348928ad998eadbfa09e63f2b76a5adf2cf60b47ffd80289c5f272",
}


def transforms(data):
    with ZipFile(BytesIO(data)) as archive:
        root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
    found = {}
    for node in root.findall(".//p:sp", base.NS) + root.findall(".//p:cxnSp", base.NS):
        ident = node.find(".//p:cNvPr", base.NS)
        transform = node.find("./p:spPr/a:xfrm", base.NS)
        if ident is None or transform is None or ident.get("name") in found:
            raise ValueError("missing or duplicate transform")
        found[ident.get("name")] = dict(transform.attrib)
    return found


def check(before, after):
    old, new = base.inspect(before), base.inspect(after)
    before_transform, after_transform = transforms(before), transforms(after)
    if set(old) != {"left", "right", "link"} or set(new) != set(old):
        raise ValueError("unexpected drawing inventory")
    for deck in (old, new):
        if deck["link"]["tag"] != "p:cxnSp" or deck["link"]["connections"] != [{"id": deck["left"]["id"], "idx": "3"}, {"id": deck["right"]["id"], "idx": "1"}]:
            raise ValueError("connection identity or site changed")
    if old["left"]["box"] != new["left"]["box"] or old["right"]["box"] != new["right"]["box"]:
        raise ValueError("target frame changed")
    if before_transform["right"] or after_transform["right"] != {"rot": "2727718"} or after_transform["link"] != {"flipV": "1"}:
        raise ValueError("rotation or connector direction differs")
    left_x, left_y, left_w, left_h = new["left"]["box"]
    right_x, right_y, right_w, right_h = new["right"]["box"]
    x, y, width, height = new["link"]["box"]
    angle = math.radians(2727718 / 60000)
    expected = (right_x + right_w / 2 - right_w / 2 * math.cos(angle), right_y + right_h / 2 - right_w / 2 * math.sin(angle))
    if (x, y + height) != (left_x + left_w, left_y + left_h // 2) or max(abs(x + width - expected[0]), abs(y - expected[1])) > 2:
        raise ValueError("endpoint does not match rotated cardinal site")


def main():
    files = []
    for name, expected in HASHES.items():
        data = (base.ROOT / "tests/fixtures" / name).read_bytes()
        if sha256(data).hexdigest() != expected:
            raise ValueError(f"hash mismatch: {name}")
        files.append(data)
    check(*files)
    print("Verified rotated target, native bindings and transformed endpoint")


if __name__ == "__main__":
    main()
