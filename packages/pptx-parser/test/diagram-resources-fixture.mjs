import { smartArtFixture } from "./smartart-fixture.mjs";
import { relationships, pack, contentTypes } from "./opc-fixture.mjs";
export const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
export const CACHE_REL = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
export const CACHE_TYPE = "application/vnd.ms-office.drawingml.diagramDrawing+xml";

export function diagramResourcesFixture(strict = false) {
  const f = smartArtFixture(strict);
  const setData = (blocks = '<dsp:dataModelExt relId="cache" minVer="urn:test:version"/>', uri = DSP) => {
    f.parts["deck/diagrams/data.xml"] = `<d:dataModel xmlns:d="${f.D}" xmlns:a="${f.A}" xmlns:r="${f.R}" xmlns:dsp="${DSP}"><d:extLst><a:ext uri="${uri}">${blocks}</a:ext></d:extLst><a:blip r:embed="data-image"/></d:dataModel>`;
  };
  const cacheXml = `<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${f.A}" xmlns:r="${f.R}"><dsp:spTree><dsp:nvGrpSpPr/><dsp:grpSpPr/><dsp:sp modelId="node-1"><dsp:nvSpPr/><dsp:spPr><a:blipFill><a:blip r:embed="cache-image"/></a:blipFill></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`;
  f.parts["deck/diagrams/cache.xml"] = cacheXml;
  f.types["deck/diagrams/cache.xml"] = CACHE_TYPE;
  f.parts["deck/media/shared.png"] = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  f.types["deck/media/shared.png"] = "image/png";
  f.parts["deck/diagrams/_rels/data.xml.rels"] = relationships([f.rel("data-image", "image", "../media/shared.png")]);
  f.parts["deck/diagrams/_rels/cache.xml.rels"] = relationships([f.rel("cache-image", "image", "../media/shared.png"), f.rel("unreferenced-image", "image", "../media/shared.png")]);
  f.relations.push({ Id: "cache", Type: CACHE_REL, Target: "diagrams/cache.xml" }); f.setRelations(); setData();
  f.bytes = () => pack(f.parts, contentTypes(undefined, Object.entries(f.types).map(([name, type]) => ({ PartName: `/${name}`, ContentType: type.startsWith("application/") || type.startsWith("image/") ? type : `application/vnd.openxmlformats-officedocument.${type}+xml` }))));
  return Object.assign(f, { setData, cacheXml });
}
