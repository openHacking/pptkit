#!/usr/bin/env python3
"""Independent inspection of the sanitized PowerPoint-created PCM sample."""
import hashlib
import io
import json
import wave
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "tests/fixtures/pcm-wave-powerpoint.pptx"
EXPECTED_WAV = "af6d351ebb501dff7a6b053de2b357ef5f873c4962738d7a7fc5c4cd19b98bdc"
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main", "a": "http://schemas.openxmlformats.org/drawingml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main", "rel": "http://schemas.openxmlformats.org/package/2006/relationships", "ct": "http://schemas.openxmlformats.org/package/2006/content-types"}


def inspect(data):
    with ZipFile(io.BytesIO(data)) as z:
        slide = ET.fromstring(z.read("ppt/slides/slide1.xml"))
        rels = ET.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels"))
        types = ET.fromstring(z.read("[Content_Types].xml"))
        pictures = slide.findall(".//p:pic", NS)
        assert len(pictures) == 1
        pic = pictures[0]
        obj = pic.find("p:nvPicPr/p:cNvPr", NS)
        audio = pic.find("p:nvPicPr/p:nvPr/a:audioFile", NS)
        media = pic.find("p:nvPicPr/p:nvPr/p:extLst/p:ext/p14:media", NS)
        poster = pic.find("p:blipFill/a:blip", NS)
        assert obj is not None and audio is not None and media is not None and poster is not None
        relationship = {x.get("Id"): x for x in rels.findall("rel:Relationship", NS)}
        link = relationship[audio.get(f"{{{NS['r']}}}link")]
        embed = relationship[media.get(f"{{{NS['r']}}}embed")]
        image = relationship[poster.get(f"{{{NS['r']}}}embed")]
        assert link.get("Type") == NS["r"] + "/audio"
        assert embed.get("Type") == "http://schemas.microsoft.com/office/2007/relationships/media"
        assert image.get("Type") == NS["r"] + "/image"
        assert link.get("Target") == embed.get("Target") == "../media/media1.wav"
        assert image.get("Target") == "../media/image1.png"
        assert all(r.get("TargetMode", "Internal") == "Internal" for r in (link, embed, image))
        wav = z.read("ppt/media/media1.wav")
        assert hashlib.sha256(wav).hexdigest() == EXPECTED_WAV
        with wave.open(io.BytesIO(wav), "rb") as reader:
            assert (reader.getnchannels(), reader.getsampwidth(), reader.getframerate(), reader.getnframes()) == (1, 2, 8000, 8000)
        defaults = {e.get("Extension"): e.get("ContentType") for e in types.findall("ct:Default", NS)}
        assert defaults.get("wav") == "audio/x-wav"
        targets = [x.get("spid") for x in slide.findall(".//p:timing//p:spTgt", NS)]
        assert targets and all(t == obj.get("id") for t in targets)
        commands = [x.get("cmd") for x in slide.findall(".//p:timing//p:cmd", NS)]
        assert "playFrom(0.0)" in commands and slide.find(".//p:timing//p:audio", NS) is not None
        return {"audioSha256": EXPECTED_WAV, "audioBytes": len(wav), "channels": 1, "bitsPerSample": 16, "sampleRate": 8000, "frames": 8000, "contentType": defaults["wav"], "audioTarget": link.get("Target"), "posterTarget": image.get("Target"), "timingTarget": obj.get("id"), "timingCommand": "playFrom(0.0)"}


if __name__ == "__main__":
    data = SOURCE.read_bytes()
    result = {"scope": "One PCM WAV inserted and saved with PowerPoint for Mac; core creator fields anonymized afterward. Package inspection only; no audible playback, native reopen or general compatibility claim.", "producer": "Microsoft PowerPoint for Mac 16.104", "source": "tests/fixtures/pcm-wave-powerpoint.pptx", "sha256": hashlib.sha256(data).hexdigest(), "observation": inspect(data)}
    print(json.dumps(result, ensure_ascii=False, indent=2))
