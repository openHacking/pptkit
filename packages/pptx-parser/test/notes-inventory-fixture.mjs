import { notesFixture, notesShape } from "./notes-fixture.mjs";
import { body } from "./content-fixture.mjs";

export function notesInventoryFixture(strict = false) {
  const text = value => body(`<a:p><a:r><a:t>${value}</a:t></a:r></a:p>`);
  const transform = '<a:xfrm rot="60000" flipH="true"><a:off x="-10" y="00020"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/><a:chOff x="1" y="2"/><a:chExt cx="30" cy="40"/></a:xfrm></p:grpSpPr>${notesShape({ id: "10", text: text("NESTED BODY") })}</p:grpSp>`;
  const field = type => body(`<a:p><a:fld id="{stored-id}" type="${type}"><a:rPr lang="fr-FR"/><a:t>2042-01-02 &amp; stored</a:t></a:fld></a:p>`);
  const content = notesShape({ id: "2", props: transform, text: text("LOCAL BODY") })
    + notesShape({ id: "3", role: "sldImg", text: "" })
    + notesShape({ id: "4", role: "hdr", text: text("HEADER") })
    + notesShape({ id: "5", role: "ftr", text: field("unrecognized-custom-type") })
    + notesShape({ id: "6", role: "dt", text: field("datetimeCustom") })
    + notesShape({ id: "7", role: "sldNum", text: field("slidenum") })
    + notesShape({ id: "8", placeholder: false, text: text("ARTWORK") }) + group;
  const f = notesFixture(content, strict);
  f.parts["deck/notes/z.xml"] = f.parts["deck/notes/z.xml"].replace('<p:notes ', '<p:notes showMasterSp="false" ').replace('</p:notes>', '<p:extLst><p:ext uri="future"><u:payload xmlns:u="urn:future"><a:fld type="hidden"><a:t>NOT A FIELD</a:t></a:fld></u:payload></p:ext></p:extLst></p:notes>');
  f.parts["deck/notes-master.xml"] = f.parts["deck/notes-master.xml"].replace('idx="1"', 'idx="99"').replace('</p:spTree>', notesShape({ id: "11", role: "sldImg", props: transform, text: "" }) + '</p:spTree>').replace('</p:notesMaster>', '<p:hf hdr="0" ftr="true" dt="false" sldNum="1"/></p:notesMaster>');
  return f;
}
