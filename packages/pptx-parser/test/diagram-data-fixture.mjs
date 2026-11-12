import { diagramResourcesFixture } from "./diagram-resources-fixture.mjs";
export function diagramDataFixture(strict = false) {
  const f = diagramResourcesFixture(strict);
  const point = (id, type = "", content = "", attrs = "") => `<d:pt modelId="${id}"${type ? ` type="${type}"` : ""} ${attrs}>${content}</d:pt>`;
  const connection = (id, src, dst, attrs = "", type = "") => `<d:cxn modelId="${id}" srcId="${src}" destId="${dst}" srcOrd="0" destOrd="0"${type ? ` type="${type}"` : ""} ${attrs}/>`;
  const richText = '<d:t><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="02"/><a:r><a:rPr b="1"/><a:t>A &amp; B 😀</a:t></a:r><a:br/><a:fld id="field-1" type="slidenum"><a:rPr/><a:t>7</a:t></a:fld><a:endParaRPr/></a:p><a:p/></d:t>';
  const points = point("0", "doc") + point("1", "", richText) + point("2", "asst") + point("3", "parTrans", "", 'cxnId="10"') + point("4", "sibTrans", "", 'cxnId="10"')
    + point("5", "pres", '<d:prSet presAssocID="1" presStyleLbl="node0" custAng="0"/>') + point("6", "pres", '<d:prSet presAssocID="1"/>');
  const connections = connection("10", "0", "1", 'parTransId="3" sibTransId="4"').replace('srcOrd="0"', 'srcOrd="2"')
    + connection("11", "1", "5", 'presId="urn:test:layout"', "presOf").replace('srcOrd="0"', 'srcOrd="1"') + connection("12", "1", "6", "", "presOf");
  const setModel = (p = points, c = connections) => { f.setData(); f.parts["deck/diagrams/data.xml"] = f.parts["deck/diagrams/data.xml"].replace('<d:extLst>', `<d:ptLst>${p}</d:ptLst><d:cxnLst>${c}</d:cxnLst><d:extLst>`); };
  setModel();
  f.parts["deck/diagrams/cache.xml"] = f.cacheXml.replace('modelId="node-1"', 'modelId="5"');
  return Object.assign(f, { point, connection, richText, points, connections, setModel });
}
