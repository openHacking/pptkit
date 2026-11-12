"""Check a retained PowerPoint elbow reroute using only ZIP/XML and hashes."""
import importlib.util
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
    "native-elbow-move-before.pptx": "76c7fb5ff261da2ec334417ebac7348a1c64049b86dedf8124b1b69459d71468",
    "native-elbow-move-after.pptx": "86bc9d6b844362e33d726256da17489d658eeeb713705099fc950e642fd1ed95",
}


def preset(data):
    with ZipFile(BytesIO(data)) as archive:
        root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
    items = root.findall(".//p:cxnSp/p:spPr/a:prstGeom", base.NS)
    if len(items) != 1:
        raise ValueError("unexpected connector geometry inventory")
    adjustment = items[0].find("a:avLst/a:gd", base.NS)
    return items[0].get("prst"), dict(adjustment.attrib) if adjustment is not None else None


def check(before, after):
    base.check(base.inspect(before), base.inspect(after))
    if preset(before) != ("bentConnector2", None):
        raise ValueError("unexpected original elbow preset")
    if preset(after) != ("bentConnector3", {"name": "adj1", "fmla": "val 50000"}):
        raise ValueError("unexpected rerouted elbow preset or adjustment")


def main():
    files = []
    for name, expected in HASHES.items():
        data = (base.ROOT / "tests/fixtures" / name).read_bytes()
        if sha256(data).hexdigest() != expected:
            raise ValueError(f"hash mismatch: {name}")
        files.append(data)
    check(*files)
    print("Verified PowerPoint elbow reroute, attachment sites and endpoints")


if __name__ == "__main__":
    main()
