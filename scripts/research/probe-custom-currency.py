#!/usr/bin/env python3
"""Inspect synthetic currency variants through PPTKit and LibreOffice resaves."""
import hashlib
import json
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "tests/fixtures/custom-currency-source.pptx"
CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def inventory(path):
    properties = []
    currency_elements = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            root = ET.fromstring(archive.read(name))
            currency_elements.extend({"part": name, "value": node.text} for node in root.iter(f"{{{VT}}}cy"))
            if root.tag != f"{{{CUSTOM}}}Properties":
                continue
            for item in root.findall(f"{{{CUSTOM}}}property"):
                children = list(item)
                properties.append({"name": item.get("name"), "type": children[0].tag.split("}")[-1] if len(children) == 1 else None, "value": children[0].text if len(children) == 1 else None})
    return {"properties": sorted(properties, key=lambda item: item["name"] or ""), "currency_elements": currency_elements}


def main():
    version = subprocess.run(["soffice", "--version"], capture_output=True, text=True, check=True).stdout.strip()
    js = """import {readFile,writeFile} from 'node:fs/promises';import {importPptx} from './packages/pptx-parser/dist/index.js';
const s=await importPptx(await readFile(process.argv[1]),{projection:'content',customPropertySaveProfile:'scalar-v1'});
try{const props=s.document.metadata.customProperties;if(props.length!==1||props[0].name!=='Control')throw Error('unexpected projection');
const no=await s.prepareSave({mode:'no-op'});if(!Buffer.from((await s.save(no.plan)).bytes).equals(await readFile(process.argv[1])))throw Error('no-op mismatch');
s.document.setMetadata({...s.document.metadata,customProperties:[{...props[0],value:'after'}]});const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));await writeFile(process.argv[2],(await s.save(p.plan)).bytes);}finally{s.dispose()}"""
    with tempfile.TemporaryDirectory(prefix="pptkit-currency-probe-") as tmp:
        root = Path(tmp); edited = root / "edited.pptx"
        subprocess.run(["node", "--input-type=module", "-e", js, str(SOURCE), str(edited)], cwd=ROOT, check=True, timeout=60)
        with zipfile.ZipFile(SOURCE) as before, zipfile.ZipFile(edited) as after:
            if set(before.namelist()) != set(after.namelist()): raise ValueError("package entries changed")
            changed = sorted(name for name in before.namelist() if before.read(name) != after.read(name))
        if changed != ["docProps/custom.xml"]: raise ValueError("unexpected changed parts")
        records = []
        for kind, source in [("control", SOURCE), ("edited", edited)]:
            records.append({"kind": kind, "round": 0, **inventory(source)})
            for index in range(1, 4):
                out = root / f"{kind}-{index}"; out.mkdir()
                profile = (root / f"profile-{kind}-{index}").as_uri()
                subprocess.run(["soffice", "-env:UserInstallation=" + profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", str(out), str(source)], check=True, capture_output=True, text=True, timeout=90)
                source = out / source.name
                records.append({"kind": kind, "round": index, **inventory(source)})
        print(json.dumps({"scope": "Synthetic PPTKit-created package with three manually substituted currency variants; separate unchanged and PPTKit-edited controls; no UI or native-authoring evidence.", "producer": version, "source_sha256": digest(SOURCE), "pptkit_edited_sha256": digest(edited), "pptkit_changed_parts": changed, "records": records}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
