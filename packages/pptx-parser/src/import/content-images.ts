import type { ImageElementInput, PresentationDocument } from "@pptkit/core";
import type { structureContext } from "../presentation/structure-context.js";
import { children, invalidStructure } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";
import { PptxImportError, type ImportSourceIdentity, type PptxSessionLimits, type PptxSourceLocation } from "../types/import.js";
import type { ContentReader } from "./content-reader.js";
import { rasterMetadata } from "./image-metadata.js";

type Context = Awaited<ReturnType<typeof structureContext>>;
type ImageFields = Pick<ImageElementInput, "type" | "assetId" | "fit" | "crop" | "opacity">;

/** One source part owns one Core asset, including shared uses across slides/groups. */
export function contentImages(document: PresentationDocument, context: Context, limits: Required<PptxSessionLimits>, identities: ImportSourceIdentity[], checkpoint: () => Promise<void>) {
  const assets = new Map<string, string>();
  const headers = new Map<string, Awaited<ReturnType<typeof rasterMetadata>>>();
  let totalBytes = 0;
  return async (node: XmlElement, properties: XmlElement, r: ContentReader, at: PptxSourceLocation, fillOverride?: XmlElement): Promise<{ fields: ImageFields; source: { blip: XmlElement; mediaPart: string; relationshipId: string } } | { reason: string }> => {
    const omitted = (reason: string) => ({ reason });
    const fill = fillOverride ?? r.one(node, r.p, "blipFill");
    if (!fill) return omitted("Picture has no local blipFill; inherited image content is unresolved.");
    r.known(fill);
    const blip = r.one(fill, r.a, "blip");
    if (!blip) return omitted("Picture has no local image reference.");
    r.known(blip);
    const svgNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    const hasSvgExtension = children(blip).some(list => list.namespaceUri === r.a && list.local === "extLst" &&
      children(list).some(ext => ext.namespaceUri === r.a && ext.local === "ext" &&
        children(ext).some(child => child.namespaceUri === svgNamespace && child.local === "svgBlip")));
    if (hasSvgExtension) r.note(blip, "An SVG image extension is present; only a supported base image relationship can be projected. The SVG relationship, content and fidelity are not interpreted; rebuilding may discard the SVG representation.");
    const embed = r.attr(blip, "embed", context.r), link = r.attr(blip, "link", context.r);
    const embedded = embed ? context.relationship(r.part, embed, "image") : undefined;
    const linked = link ? context.relationship(r.part, link, "image") : undefined;
    if (embedded && embedded.targetMode !== "Internal") invalidStructure("Embedded image relationship must be internal.", r.part);
    if (linked && linked.targetMode !== "External") invalidStructure("Linked image relationship must be external.", r.part);
    if (linked) r.note(blip, `External image relationship ${linked.id} is inventoried only; its target is not fetched or registered as a Core URL asset. ${embedded ? "Only the embedded representation is projected; external updates are not mapped." : "The picture is omitted from Core."}`);
    if (!embedded) return omitted(linked ? "External-only picture is retained in source bytes; loading requires a separate explicit consumer operation." : "Picture has no embedded image relationship.");
    const part = context.target(embedded, "image");
    const preset = r.one(properties, r.a, "prstGeom");
    if (!preset || r.attr(preset, "prst") !== "rect" || r.one(properties, r.a, "custGeom")) return omitted("Only rectangular picture geometry is mapped; shaped/custom image clipping remains unsupported.");
    const adjustments = r.one(preset, r.a, "avLst");
    if (adjustments) r.known(adjustments);
    const stretch = r.one(fill, r.a, "stretch"), tile = r.one(fill, r.a, "tile");
    if (stretch && tile) invalidStructure("Conflicting image stretch and tile choices.", r.part);
    if (!stretch) return omitted("Picture requires explicit stretch placement; tiling and implicit placement are not mapped.");
    r.known(stretch);
    const rectangle = (rect: XmlElement) => {
      r.known(rect);
      return { left: percentage(rect, "l", 0, r), top: percentage(rect, "t", 0, r), right: percentage(rect, "r", 0, r), bottom: percentage(rect, "b", 0, r) };
    };
    const fillRect = r.one(stretch, r.a, "fillRect");
    if (fillRect && Object.values(rectangle(fillRect)).some(value => value !== 0)) return omitted("Nonzero destination fillRect offsets cannot be represented by the current Core image frame.");
    if (!r.bool(fill, "rotWithShape", true)) return omitted("Picture fill rotation independent of its shape is not mapped.");
    const sourceRect = r.one(fill, r.a, "srcRect"), crop = sourceRect ? rectangle(sourceRect) : undefined;
    if (crop && (Object.values(crop).some(value => value < 0 || value > 1) || crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1)) return omitted("Source crop outsets or an empty/inverted source rectangle cannot be represented by Core; crop is not clamped.");
    const alphas = children(blip).filter(n => n.namespaceUri === r.a && n.local === "alphaModFix");
    let opacity: number | undefined;
    if (alphas.length === 1) {
      const value = percentage(alphas[0]!, "amt", 1, r);
      if (value < 0) invalidStructure("Image alpha amount must be nonnegative.", r.part);
      if (value <= 1) opacity = value;
      else r.note(alphas[0]!, "Alpha amplification above 100% is not representable as Core image opacity.");
    }
    let assetId = assets.get(part);
    if (!assetId) {
      const payload = context.payload(part);
      const declared = payload.contentType.trim().toLowerCase();
      if (!["image/png", "image/apng", "image/jpeg", "image/gif"].includes(declared)) return omitted("Embedded image MIME/signature/dimensions are unsupported, malformed or mismatched; only matching PNG/JPEG/GIF headers with positive dimensions are projected.");
      if (payload.bytes.byteLength > limits.maxImportedAssetBytes || totalBytes + payload.bytes.byteLength > limits.maxTotalImportedAssetBytes) throw new PptxImportError("import-limit-exceeded", "Mapped images exceed maxImportedAssetBytes or maxTotalImportedAssetBytes.", { part });
      if (!headers.has(part)) headers.set(part, await rasterMetadata(payload.bytes, checkpoint));
      const metadata = headers.get(part);
      if (!metadata || (declared !== metadata.mimeType && !(declared === "image/apng" && metadata.mimeType === "image/png"))) return omitted("Embedded image MIME/signature/dimensions are unsupported, malformed or mismatched; only matching PNG/JPEG/GIF headers with positive dimensions are projected.");
      totalBytes += payload.bytes.byteLength;
      assetId = `source-image-${assets.size + 1}`;
      document.registerAsset({ id: assetId, kind: "image", ...metadata, ...(declared === "image/apng" ? { mimeType: declared } : {}), source: { type: "bytes", value: payload.bytes } });
      assets.set(part, assetId);
      identities.push({ kind: "asset", id: assetId, source: { part } });
    }
    identities.push({ kind: "asset", id: assetId, source: { ...at, relationshipId: embedded.id } });
    r.note(blip, "Image bytes and dimensions are header-classified only. Pixel decoding, EXIF orientation, color profiles and animated-image playback are not verified by import.");
    return { fields: { type: "image", assetId, fit: crop ? "crop" : "stretch", ...(crop ? { crop } : {}), ...(opacity !== undefined ? { opacity } : {}) }, source: { blip, mediaPart: part, relationshipId: embedded.id } };
  };
}

/** DrawingML percentages accept integer thousandths or a literal percent. */
function percentage(node: XmlElement, name: string, fallback: number, r: ContentReader): number {
  const text = r.attr(node, name)?.trim();
  if (text === undefined) return fallback;
  if (/^[+-]?\d+(?:\.\d+)?%$/.test(text)) {
    const value = Number(text.slice(0, -1)) / 100;
    if (Number.isFinite(value)) return value;
  } else if (/^[+-]?\d+$/.test(text)) {
    const value = Number(text);
    if (Number.isSafeInteger(value) && value >= -2147483648 && value <= 2147483647) return value / 100000;
  }
  return invalidStructure(`Invalid image percentage ${node.local}/@${name}.`, r.part);
}
