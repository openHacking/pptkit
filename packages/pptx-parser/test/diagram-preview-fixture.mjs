import { diagramResourcesFixture, DSP } from "./diagram-resources-fixture.mjs";
import { transform, body, shape } from "./content-fixture.mjs";
import { png } from "./image-fixture.mjs";
export const cacheTextTransform = (attrs = "", { x = 381000, y = 317500, width = 1016000, height = 254000 } = {}) => `<dsp:txXfrm ${attrs}><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></dsp:txXfrm>`;
export function diagramPreviewFixture(strict = false) {
  const f = diagramResourcesFixture(strict);
  const cacheShape = (id = "2", { text = "Cached text", preset = "rect", fill = '<a:solidFill><a:srgbClr val="228844"/></a:solidFill>', xfrm = transform(), extra = "", hidden = false } = {}) => `<dsp:sp modelId="node-${id}"><dsp:nvSpPr><dsp:cNvPr id="${id}" name="Cache ${id}"${hidden ? ' hidden="1"' : ""}/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr>${xfrm}<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>${fill}</dsp:spPr>${text === undefined ? "" : body(`<a:p><a:r><a:rPr sz="1800"/><a:t>${text}</a:t></a:r></a:p>`).replaceAll('p:txBody', 'dsp:txBody')}${extra}</dsp:sp>`;
  const setCache = (content, rootProps = '<dsp:grpSpPr/>', attrs = "") => { f.parts["deck/diagrams/cache.xml"] = `<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${f.A}" xmlns:r="${f.R}" ${attrs}><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>${rootProps}${content}</dsp:spTree></dsp:drawing>`; };
  const frame = f.frame().replace('</p:nvGraphicFramePr>', '</p:nvGraphicFramePr><p:xfrm><a:off x="1270000" y="635000"/><a:ext cx="5080000" cy="2540000"/></p:xfrm>');
  f.setTree(frame + shape({ id: "10", txBox: true, text: body() }));
  f.parts["deck/media/shared.png"] = png;
  setCache(cacheShape());
  return Object.assign(f, { cacheShape, setCache, previewFrame: frame });
}
