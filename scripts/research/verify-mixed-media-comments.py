#!/usr/bin/env python3
"""Independent source-level checks for the synthetic media/comment fixture."""
import posixpath
import struct
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ROOT = Path(__file__).resolve().parents[2]
DEFAULT = ROOT / "tests/fixtures/mixed-media-comments-source.pptx"


def require(condition, message):
    if not condition:
        raise ValueError(message)


def relationships(archive, owner):
    directory, name = posixpath.split(owner)
    relpath = posixpath.join(directory, "_rels", name + ".rels")
    root = ET.fromstring(archive.read(relpath))
    result = {}
    for item in root:
        # MCE extension branches are not interpreted as relationship nodes.
        if item.tag == "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship":
            require(item.get("Id") not in result, "duplicate relationship")
            result[item.get("Id")] = item
    return result


def resolve(archive, owner, rel_id, kind):
    item = relationships(archive, owner).get(rel_id)
    require(item is not None and item.get("Type") == f"{R}/{kind}" and item.get("TargetMode") != "External", f"invalid {kind} relationship")
    target = posixpath.normpath(posixpath.join(posixpath.dirname(owner), item.get("Target", "")))
    require(target in archive.namelist(), f"missing {kind} target")
    return target


def inspect(path):
    with ZipFile(path) as archive:
        slide_name = "ppt/slides/slide1.xml"
        slide = ET.fromstring(archive.read(slide_name))
        audio = slide.findall(f".//{{{A}}}audioFile")
        require(len(audio) == 1, "audio declaration count")
        audio_part = resolve(archive, slide_name, audio[0].get(f"{{{R}}}link"), "audio")
        pictures = [pic for pic in slide.findall(f".//{{{P}}}pic") if pic.find(f".//{{{A}}}audioFile") is not None]
        require(len(pictures) == 1, "audio picture count")
        poster = pictures[0].find(f".//{{{A}}}blip")
        require(poster is not None, "missing audio poster")
        poster_part = resolve(archive, slide_name, poster.get(f"{{{R}}}embed"), "image")
        require(archive.read(poster_part).startswith(b"\x89PNG\r\n\x1a\n"), "invalid audio poster")
        data = archive.read(audio_part)
        require(len(data) == 124 and data[:4] == b"RIFF" and data[8:16] == b"WAVEfmt " and data[36:40] == b"data", "invalid waveform header")
        require(struct.unpack_from("<IHHIIHH", data, 16) == (16, 1, 1, 8000, 8000, 1, 8), "invalid waveform format")
        require(struct.unpack_from("<I", data, 4)[0] == len(data) - 8 and struct.unpack_from("<I", data, 40)[0] == len(data) - 44 and data[44:] == bytes([128]) * 80, "invalid waveform samples")
        comments = resolve(archive, slide_name, "smComments", "comments")
        authors = resolve(archive, "ppt/presentation.xml", "smAuthors", "commentAuthors")
        types = ET.fromstring(archive.read("[Content_Types].xml"))
        overrides = {item.get("PartName"): item.get("ContentType") for item in types if item.tag.endswith("}Override")}
        require(overrides.get("/" + audio_part) == "audio/wav", "audio content type mismatch")
        require(overrides.get("/" + comments) == "application/vnd.openxmlformats-officedocument.presentationml.comments+xml", "comments content type mismatch")
        require(overrides.get("/" + authors) == "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml", "authors content type mismatch")
        root = ET.fromstring(archive.read(comments)); author_root = ET.fromstring(archive.read(authors))
        require(root.tag == f"{{{P}}}cmLst" and author_root.tag == f"{{{P}}}cmAuthorLst", "comment root mismatch")
        rows, people = root.findall(f"{{{P}}}cm"), author_root.findall(f"{{{P}}}cmAuthor")
        require(len(rows) == 1 and len(people) == 1 and rows[0].get("authorId") == people[0].get("id") == "0", "comment author mismatch")
        require(rows[0].findtext(f"{{{P}}}text") == "Synthetic review note" and people[0].get("name") == "Fixture Author", "comment content mismatch")
        return {"audio": audio_part, "comments": comments, "authors": authors}


if __name__ == "__main__":
    if len(sys.argv) > 2:
        raise SystemExit("usage: verify-mixed-media-comments.py [INPUT]")
    print(inspect(Path(sys.argv[1]) if len(sys.argv) == 2 else DEFAULT))
