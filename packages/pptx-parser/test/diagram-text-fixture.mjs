import { diagramPreviewFixture } from "./diagram-preview-fixture.mjs";
export function diagramTextFixture(strict = false) {
  const f = diagramPreviewFixture(strict), xml = text => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll("\r", "&#13;");
  const paragraphs = values => values.map(text => `<a:p>${text.split("\n").map((line, i) => `${i ? '<a:br/>' : ''}<a:r><a:t xml:space="preserve">${xml(line)}</a:t></a:r>`).join("")}</a:p>`).join("");
  const setText = (source = ["Source & 😀"], cached = source, { association = '<d:prSet presAssocID="1"/>', extraPoints = "", sourceExtra = "", cacheExtra = "" } = {}) => {
    f.setData();
    const points = `<d:pt modelId="0" type="doc"/><d:pt modelId="1"><d:t><a:bodyPr/><a:lstStyle/>${paragraphs(source)}${sourceExtra}</d:t></d:pt><d:pt modelId="2" type="pres">${association}</d:pt>${extraPoints}`;
    const edges = '<d:cxn modelId="10" srcId="0" destId="1" srcOrd="0" destOrd="0"/><d:cxn modelId="11" type="presOf" srcId="1" destId="2" srcOrd="0" destOrd="0"/>';
    f.parts["deck/diagrams/data.xml"] = f.parts["deck/diagrams/data.xml"].replace('<d:extLst>', `<d:ptLst>${points}</d:ptLst><d:cxnLst>${edges}</d:cxnLst><d:extLst>`);
    const shape = f.cacheShape("2", { text: "temporary" }).replace('modelId="node-2"', 'modelId="2"').replace(/<dsp:txBody>[\s\S]*?<\/dsp:txBody>/, `<dsp:txBody><a:bodyPr/><a:lstStyle/>${paragraphs(cached)}${cacheExtra}</dsp:txBody>`);
    f.setCache(shape);
  };
  setText(); return Object.assign(f, { setText, paragraphs });
}
