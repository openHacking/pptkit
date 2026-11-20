import { readGeneratedZip } from "../../tests/support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";
import { diagramTextFixture } from "../../packages/pptx-parser/test/diagram-text-fixture.mjs";
import { nativeTimingSpecimen } from "./native-timing-fixture.mjs";

// Test-only source construction: preserve the input package and add named raw specimens.
export function mixedAdvancedFixture(input) {
  const parts = readGeneratedZip(input), diagram = diagramTextFixture();
  const slidePart = "ppt/slides/slide1.xml", relPart = "ppt/slides/_rels/slide1.xml.rels";
  const slide = parts.get(slidePart).toString();
  const frame = diagram.previewFrame.replace('id="2"', 'id="99"').replace('<d:relIds ', `<d:relIds xmlns:d="${diagram.D}" `)
    .replace('r:dm="dm"', 'r:dm="smData"').replace('r:lo="lo"', 'r:lo="smLayout"')
    .replace('r:cs="cs"', 'r:cs="smColors"').replace('r:qs="qs"', 'r:qs="smStyle"');
  if (!slide.includes('</p:spTree>') || !slide.includes('<p:extLst>')) throw new Error('Unexpected mixed slide structure');
  parts.set(slidePart, Buffer.from(slide.replace('</p:spTree>', `${frame}</p:spTree>`).replace('<p:extLst>', `${nativeTimingSpecimen("fade-in", 2)}<p:extLst>`)));
  const rel = parts.get(relPart).toString(), base = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const extra = [['smData','diagramData','data'],['smLayout','diagramLayout','layout'],['smColors','diagramColors','colors'],['smStyle','diagramQuickStyle','style']]
    .map(([id,type,name]) => `<Relationship Id="${id}" Type="${base}${type}" Target="../diagrams/${name}.xml"/>`).join('')
    + '<Relationship Id="smCache" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/cache.xml"/>';
  if (!rel.includes('<mc:AlternateContent>')) throw new Error('Missing inactive relationship branch');
  parts.set(relPart, Buffer.from(rel.replace('<mc:AlternateContent>', `${extra}<mc:AlternateContent>`)));
  for (const [name, value] of Object.entries(diagram.parts)) if (name.startsWith('deck/diagrams/')) {
    const destination = name.replace('deck/', 'ppt/');
    let bytes = Buffer.from(value);
    if (name === 'deck/diagrams/data.xml') bytes = Buffer.from(bytes.toString().replace('relId="cache"', 'relId="smCache"'));
    parts.set(destination, bytes);
  }
  let types = parts.get('[Content_Types].xml').toString();
  for (const [name, type] of Object.entries(diagram.types)) if (name.startsWith('deck/diagrams/') && !name.endsWith('.rels')) {
    const contentType = type.startsWith('application/') ? type : `application/vnd.openxmlformats-officedocument.${type}+xml`;
    types = types.replace('</Types>', `<Override PartName="/${name.replace('deck/', 'ppt/')}" ContentType="${contentType}"/></Types>`);
  }
  parts.set('[Content_Types].xml', Buffer.from(types));
  return zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
}
