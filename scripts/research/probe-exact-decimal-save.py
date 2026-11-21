#!/usr/bin/env python3
"""Compare exact PPTKit decimal edits with the untouched source across LibreOffice resaves."""
import hashlib
import json
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "tests/fixtures/custom-exact-decimal-source.pptx"
CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def inventory(path):
    properties = []
    decimal_elements = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            root = ET.fromstring(archive.read(name))
            decimal_elements.extend({"part": name, "value": node.text} for node in root.iter(f"{{{VT}}}decimal"))
            if root.tag != f"{{{CUSTOM}}}Properties":
                continue
            for item in root.findall(f"{{{CUSTOM}}}property"):
                children = list(item)
                properties.append({"name": item.get("name"), "type": children[0].tag.split("}")[-1] if len(children) == 1 else None, "value": children[0].text if len(children) == 1 else None})
    return {"properties": sorted(properties, key=lambda item: item["name"] or ""), "decimal_elements": decimal_elements}


def main():
    version = subprocess.run(["soffice", "--version"], capture_output=True, text=True, check=True).stdout.strip()
    js = """import {readFile,writeFile} from 'node:fs/promises';import {importPptx} from './packages/pptx-parser/dist/index.js';
const s=await importPptx(await readFile(process.argv[1]),{projection:'content',customDecimalProjection:'exact-save-v2',customPropertySaveProfile:'scalar-v1'});
try{const props=s.document.metadata.customProperties;if(props.length!==4||props[0].name!=='Tenth'||props[0].value!=='0.1')throw Error('unexpected projection');
const no=await s.prepareSave({mode:'no-op'});if(!Buffer.from((await s.save(no.plan)).bytes).equals(await readFile(process.argv[1])))throw Error('no-op mismatch');
s.document.setMetadata({...s.document.metadata,customProperties:props.map(p=>p.name==='Tenth'?{...p,value:'0.2'}:p.name==='Wide'?{...p,value:'9007199254740994.0001'}:p)});const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));await writeFile(process.argv[2],(await s.save(p.plan)).bytes);}finally{s.dispose()}"""
    base = ROOT / ".tmp/exact-decimal-save-libreoffice"
    base.mkdir(parents=True, exist_ok=True)
    root = Path(tempfile.mkdtemp(prefix="run-", dir=base)); control = root / "control.pptx"; edited = root / "edited.pptx"
    shutil.copyfile(SOURCE, control)
    subprocess.run(["node", "--input-type=module", "-e", js, str(SOURCE), str(edited)], cwd=ROOT, check=True, timeout=60)
    with zipfile.ZipFile(SOURCE) as before, zipfile.ZipFile(edited) as after:
        if set(before.namelist()) != set(after.namelist()): raise ValueError("package entries changed")
        changed = sorted(name for name in before.namelist() if before.read(name) != after.read(name))
    if changed != ["docProps/custom.xml"]: raise ValueError("unexpected changed parts")
    records = []; logs = []
    for kind, source in [("control", control), ("edited", edited)]:
        records.append({"kind": kind, "round": 0, "path": str(source.relative_to(ROOT)), "sha256": digest(source), **inventory(source)})
        for index in range(1, 4):
            out = root / f"{kind}-{index}"; out.mkdir()
            profile = (root / f"profile-{kind}-{index}").as_uri()
            result = subprocess.run(["soffice", "-env:UserInstallation=" + profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", str(out), str(source)], check=True, capture_output=True, text=True, timeout=90)
            logs.append({"kind": kind, "round": index, "exit_code": result.returncode, "stdout": result.stdout, "stderr": result.stderr})
            source = out / source.name
            records.append({"kind": kind, "round": index, "path": str(source.relative_to(ROOT)), "sha256": digest(source), **inventory(source)})
    print(json.dumps({"scope": "PPTKit-created package with exact decimal variants; compare original to PPTKit exact source-value edits; no UI or native-authoring evidence.", "producer": version, "source_sha256": digest(SOURCE), "pptkit_edited_sha256": digest(edited), "pptkit_changed_parts": changed, "records": records, "logs": logs}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
