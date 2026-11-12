import { contentFixture, transform } from "./content-fixture.mjs";
import { pack, contentTypes, relationships } from "./opc-fixture.mjs";

// Repository-authored 8 × 4 raster shared with the general fixture corpus.
export const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC", "base64");
export const gif = Buffer.from("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7", "base64");
// Complete 1×1 LZW stream, including the required end-of-information code.
export const validGif = Buffer.from('47494638396101000100800000000000ffffff21f90401000000002c00000000010001000002024401003b', 'hex');
// Independent synthetic SOF header, intentionally no pixel stream: metadata only.
export const jpegHeader = Buffer.from([255, 216, 255, 224, 0, 6, 255, 192, 0, 0, 255, 194, 0, 11, 8, 0, 4, 0, 8, 1, 1, 17, 0, 255, 217]);
export const picture = ({ id = 2, embed = "image", link, crop = "", placement = "<a:stretch><a:fillRect/></a:stretch>", fillAttrs = "", effects = "", props = "", preset = "rect", xfrm = transform(), nv = "", blipAttrs = "" } = {}) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}" descr="Raster" ${nv}/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill ${fillAttrs}><a:blip ${embed === undefined || embed === null ? "" : `r:embed="${embed}"`} ${link ? `r:link="${link}"` : ""} ${blipAttrs}>${effects}</a:blip>${crop}${placement}</p:blipFill><p:spPr>${xfrm}<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>${props}</p:spPr></p:pic>`;
export function imageFixture(content = picture(), strict = false) {
  const f = contentFixture(content, strict);
  f.parts["media/raster.bin"] = png;
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin")]);
  f.mediaTypes = { "media/raster.bin": "image/png" };
  return f;
}
export function encodeImages(f = imageFixture()) {
  return pack(f.parts, contentTypes(undefined, [
    ...Object.entries(f.types).map(([name, type]) => ({ PartName: `/${name}`, ContentType: `application/vnd.openxmlformats-officedocument.${type}+xml` })),
    ...Object.entries(f.mediaTypes).map(([name, type]) => ({ PartName: `/${name}`, ContentType: type })),
  ]));
}
