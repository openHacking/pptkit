#!/usr/bin/env python3
"""Compare currency examples with ST_Cy from official ECMA schema archives."""
import hashlib
import io
import json
import sys
import zipfile
from pathlib import Path
from lxml import etree

XSD = "http://www.w3.org/2001/XMLSchema"
CASES = {"1.2500": True, ".1000": True, " 0001.2500\t": True, "0.0000": True,
         "-12.5000": False, "+1.0000": False, "1.000": False, "1.00000": False,
         "1e0": False, "1.0000\u00a0": False, "1.0000\u2028": False, "": False}


def check(archive, inner):
    with zipfile.ZipFile(archive) as outer:
        data = outer.read(inner)
    with zipfile.ZipFile(io.BytesIO(data)) as schemas:
        name = next(name for name in schemas.namelist() if name.endswith("shared-documentPropertiesVariantTypes.xsd"))
        content = schemas.read(name)
    root = etree.fromstring(content)
    node = root.find(f"{{{XSD}}}simpleType[@name='ST_Cy']")
    if node is None: raise ValueError("ST_Cy missing")
    schema_root = etree.Element(f"{{{XSD}}}schema", nsmap={"xs": XSD, "xsd": XSD})
    element = etree.SubElement(schema_root, f"{{{XSD}}}element", name="cy")
    element.append(etree.fromstring(etree.tostring(node)))
    element[0].attrib.pop("name")
    schema = etree.XMLSchema(schema_root)
    for text, expected in CASES.items():
        candidate = etree.Element("cy"); candidate.text = text
        if schema.validate(candidate) != expected: raise ValueError(f"unexpected ST_Cy result: {text!r}")
    return {"schema_sha256": hashlib.sha256(content).hexdigest(), "cases": len(CASES)}


if __name__ == "__main__":
    if len(sys.argv) != 3: raise SystemExit("usage: verify-currency-schema.py ECMA-PART1.zip ECMA-PART4.zip")
    result = {"strict": check(Path(sys.argv[1]), "OfficeOpenXML-XMLSchema-Strict.zip"), "transitional": check(Path(sys.argv[2]), "OfficeOpenXML-XMLSchema-Transitional.zip")}
    print(json.dumps(result, indent=2))
