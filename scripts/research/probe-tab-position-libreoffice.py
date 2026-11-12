"""Create a fresh PPTX control by removing only explicit tab-paragraph zero indents."""
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

root = Path(__file__).resolve().parents[2]
base = (root / ".tmp/native-tab-stops").resolve()
source = Path(sys.argv[1]).resolve()
if not source.is_relative_to(base) or source.name != "positions.pptx": raise ValueError("Expected a position-matrix PPTX under .tmp/native-tab-stops")
directory = Path(tempfile.mkdtemp(prefix="libreoffice-position-", dir=base))
shutil.copyfile(source, directory / "positions.pptx")
shutil.copyfile(source.parent / "manifest.json", directory / "manifest.json")
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
with zipfile.ZipFile(source) as original, zipfile.ZipFile(directory / "no-indent.pptx", "x", zipfile.ZIP_DEFLATED) as output:
    for info in original.infolist():
        data = original.read(info.filename)
        if info.filename.startswith("ppt/slides/slide") and info.filename.endswith(".xml") and "/_rels/" not in info.filename:
            document = ET.fromstring(data)
            for shape in document.iter(p + "sp"):
                if shape.find(".//" + a + "tabLst/" + a + "tab") is not None:
                    for prop in shape.iter(a + "pPr"):
                        if prop.attrib.pop("indent") != "-0": raise ValueError("Unexpected indentation")
            data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
        output.writestr(info, data)
subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(directory), str(directory / "no-indent.pptx")], check=True, capture_output=True, text=True)
print(directory)
