import { readGeneratedZip } from "../../tests/support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";

// Test-only, hand-authored sources. This is not a media or comment authoring API.
export function mixedMediaCommentsFixture(input) {
  const parts = readGeneratedZip(input), slideName = "ppt/slides/slide1.xml";
  const slide = parts.get(slideName).toString();
  const picture = '<p:pic><p:nvPicPr><p:cNvPr id="100" name="Synthetic audio"/><p:cNvPicPr/><p:nvPr><a:audioFile r:link="smAudio"/></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="635000" cy="635000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  if (!slide.includes('</p:spTree>')) throw new Error('Missing source shape tree');
  parts.set(slideName, Buffer.from(slide.replace('</p:spTree>', `${picture}</p:spTree>`)));
  const relName = 'ppt/slides/_rels/slide1.xml.rels', rel = parts.get(relName).toString();
  const base = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const links = `<Relationship Id="smAudio" Type="${base}audio" Target="../media/tone.wav"/><Relationship Id="smComments" Type="${base}comments" Target="../comments/comment1.xml"/>`;
  if (!rel.includes('<mc:AlternateContent>')) throw new Error('Missing inactive relationship branch');
  parts.set(relName, Buffer.from(rel.replace('<mc:AlternateContent>', `${links}<mc:AlternateContent>`)));
  const presRel = 'ppt/_rels/presentation.xml.rels', pres = parts.get(presRel).toString();
  if (!pres.includes('</Relationships>')) throw new Error('Missing presentation relationships');
  parts.set(presRel, Buffer.from(pres.replace('</Relationships>', `<Relationship Id="smAuthors" Type="${base}commentAuthors" Target="commentAuthors.xml"/></Relationships>`)));
  const ns = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  parts.set('ppt/comments/comment1.xml', Buffer.from(`<p:cmLst xmlns:p="${ns}"><p:cm authorId="0" dt="2026-01-01T00:00:00" idx="1"><p:pos x="10" y="20"/><p:text>Synthetic review note</p:text></p:cm></p:cmLst>`));
  parts.set('ppt/commentAuthors.xml', Buffer.from(`<p:cmAuthorLst xmlns:p="${ns}"><p:cmAuthor id="0" name="Fixture Author" initials="FA" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>`));
  // RIFF/WAVE, 8-bit mono PCM at 8 kHz, 80 silent samples; no decoder is invoked.
  const audio = Buffer.alloc(124, 128); audio.write('RIFF', 0); audio.writeUInt32LE(116, 4); audio.write('WAVEfmt ', 8);
  audio.writeUInt32LE(16, 16); audio.writeUInt16LE(1, 20); audio.writeUInt16LE(1, 22); audio.writeUInt32LE(8000, 24);
  audio.writeUInt32LE(8000, 28); audio.writeUInt16LE(1, 32); audio.writeUInt16LE(8, 34); audio.write('data', 36); audio.writeUInt32LE(80, 40);
  parts.set('ppt/media/tone.wav', audio);
  let types = parts.get('[Content_Types].xml').toString();
  const declarations = '<Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/><Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/><Override PartName="/ppt/media/tone.wav" ContentType="audio/wav"/>';
  if (!types.includes('</Types>')) throw new Error('Missing content types');
  types = types.replace('</Types>', `${declarations}</Types>`); parts.set('[Content_Types].xml', Buffer.from(types));
  return zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
}
