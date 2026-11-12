"""Create the small, renamed SIL-OFL fixture from the pinned Noto Sans KR source."""
import hashlib
from pathlib import Path
import subprocess
import sys

from fontTools.ttLib import TTFont

source, target = map(Path, sys.argv[1:3])
if hashlib.sha256(source.read_bytes()).hexdigest() != "69975a0ac8472717870aefeab0a4d52739308d90856b9955313b2ad5e0148d68":
    raise ValueError("Unexpected Noto Sans KR source")
if target.exists(): raise FileExistsError(target)
subprocess.run([sys.executable, "-m", "fontTools.subset", str(source), f"--output-file={target}",
                "--unicodes=U+0020-007E,U+200B,U+3042,U+304B,U+30A2,U+30AB,U+AC00,U+D55C,U+AE00,U+4E2D,U+6587",
                "--layout-features=*", "--name-IDs=*", "--name-legacy", "--name-languages=*", "--no-recalc-timestamp"], check=True)
font = TTFont(target, recalcTimestamp=False)
names = {1: "PPTKit Tab Measurement", 2: "Regular", 3: "PPTKit Tab Measurement Regular", 4: "PPTKit Tab Measurement Regular", 6: "PPTKitTabMeasurement-Regular", 16: "PPTKit Tab Measurement", 17: "Regular"}
for record in font["name"].names:
    if record.nameID in names: record.string = names[record.nameID].encode(record.getEncoding())
if "CFF " in font:
    cff = font["CFF "].cff; cff.fontNames = ["PPTKitTabMeasurement-Regular"]
    top = cff.topDictIndex[0]; top.FullName = "PPTKit Tab Measurement Regular"; top.FamilyName = "PPTKit Tab Measurement"
font.save(target, reorderTables=False)
print(hashlib.sha256(target.read_bytes()).hexdigest())
