import { contentFixture, shape, body } from "./content-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
export const paragraphProfile = { projection: "content", editSaveProfile: "text-runs-v1", textTargetSaveProfile: "paragraph-stable-v1" };
export const paragraphTarget = (range = '<p:pRg st="0" end="2"/>', id = "2") => `<p:spTgt spid="${id}"><p:txEl>${range}</p:txEl></p:spTgt>`;
export function paragraphFixture({ strict = false, target = paragraphTarget(), txBox = true, id = "2", extra = "" } = {}) {
  const paragraphs = '<a:p><a:r><a:t>First</a:t></a:r><a:r><a:rPr b="1"/><a:t> tail</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:endParaRPr/></a:p><a:p><a:r><a:t>第三😀é</a:t></a:r></a:p>';
  const f = contentFixture(shape({ id, txBox, text: body(paragraphs) }) + extra, strict);
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace('</p:sld>', `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="500"><p:childTnLst><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="2" dur="500"/><p:tgtEl>${target}</p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing></p:sld>`);
  return f;
}
export const paragraphBytes = options => encodeStructure(paragraphFixture(options));
export function editParagraph(session, text, p = 0, r = 0, index = 0) {
  const slide = session.document.slides[0], element = structuredClone(slide.elements[index]);
  (element.type === "text" ? element : element.text).content[p].runs[r].text = text;
  slide.replaceElement(element.id, element);
}
