import { inspectIsoBmff } from "@pptkit/media-inspection";
import { attribute, children } from "../presentation/structure-xml.js";
import type { structureContext } from "../presentation/structure-context.js";
import { PptxImportError, type ImportedVideoSource, type PptxSessionLimits } from "../types/import.js";
import type { XmlElement } from "../xml/read-xml.js";

type Context = Awaited<ReturnType<typeof structureContext>>;
const MEDIA_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const MEDIA_REL = "http://schemas.microsoft.com/office/2007/relationships/media";
const MEDIA_EXT = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

/** Inventory effective source references without loading external data or projecting video. */
export function videoSources(context: Context, limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>) {
  const result: ImportedVideoSource[] = [];
  const byPicture = new WeakMap<XmlElement, readonly ImportedVideoSource[]>();
  const cache = new Map<string, Pick<ImportedVideoSource, "status" | "contentType" | "framing" | "reason">>();
  let total = 0;
  const inspect = async (part: string, slideId: string, objectId: string, picture: XmlElement, paths: WeakMap<XmlElement, string>) => {
    const named = (node: XmlElement, ns: string, name: string) => children(node).filter(child => child.namespaceUri === ns && child.local === name);
    const nv = named(picture, context.p, "nvPicPr")[0], nvPr = nv && named(nv, context.p, "nvPr")[0];
    if (!nvPr) return;
    const direct = named(nvPr, context.a, "videoFile");
    if (!direct.length) return;
    const declarations: { node: XmlElement; declaration: ImportedVideoSource["declaration"]; relationshipAttribute: "link" | "embed" }[] = direct.map(node => ({ node, declaration: "videoFile", relationshipAttribute: "link" }));
    for (const list of named(nvPr, context.p, "extLst")) for (const ext of named(list, context.p, "ext")) {
      await checkpoint();
      if (attribute(ext, "uri")?.toUpperCase() !== MEDIA_EXT) continue;
      for (const node of named(ext, MEDIA_NS, "media")) {
        const embed = attribute(node, "embed", context.r), link = attribute(node, "link", context.r);
        if (embed !== undefined || link === undefined) declarations.push({ node, declaration: "p14:media", relationshipAttribute: "embed" });
        if (link !== undefined) declarations.push({ node, declaration: "p14:media", relationshipAttribute: "link" });
      }
    }
    for (const fill of named(picture, context.p, "blipFill")) for (const node of named(fill, context.a, "blip")) {
      const embed = attribute(node, "embed", context.r), link = attribute(node, "link", context.r);
      if (embed !== undefined || link === undefined) declarations.push({ node, declaration: "poster", relationshipAttribute: "embed" });
      if (link !== undefined) declarations.push({ node, declaration: "poster", relationshipAttribute: "link" });
    }
    const first = result.length;
    for (const { node, declaration, relationshipAttribute } of declarations) {
      await checkpoint();
      if (result.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Video source inventory exceeds maxReportFindings.", { part });
      const raw = await context.sourceElement(part, node), relationshipId = attribute(node, relationshipAttribute, context.r);
      const source = { part, path: paths.get(node)!, sourceObjectId: objectId, elementIndex: raw.elementIndex, ...(relationshipId ? { relationshipId } : {}) };
      const base = { slideId, declaration, relationshipAttribute, source, sourceElementIndex: raw.elementIndex, ...(relationshipId ? { relationshipId } : {}) };
      const relation = context.findRelationship(part, relationshipId);
      const expectedType = declaration === "p14:media" ? MEDIA_REL : `${context.r}/${declaration === "poster" ? "image" : "video"}`;
      if (!relation || relation.type !== expectedType || relation.fragment !== undefined || (relationshipAttribute === "embed" && relation.targetMode !== "Internal") || (declaration !== "videoFile" && relationshipAttribute === "link" && relation.targetMode !== "External")) {
        result.push({ ...base, status: "unsupported", reason: "Missing, fragmented, wrong-type or wrong-mode media relationship." }); continue;
      }
      if (relation.targetMode === "External") { result.push({ ...base, status: "external", externalTarget: relation.target, reason: "External media is recorded but not fetched." }); continue; }
      if (!relation.targetPart) { result.push({ ...base, status: "unsupported", reason: "Media relationship has no internal target." }); continue; }
      const mediaPart = relation.targetPart, payload = context.payload(mediaPart), contentType = payload.contentType.trim().toLowerCase();
      if (declaration === "poster") { result.push({ ...base, mediaPart, status: "internal", contentType }); continue; }
      let info = cache.get(mediaPart);
      if (!info) {
        if (payload.bytes.length > limits.maxImportedAssetBytes || total + payload.bytes.length > limits.maxTotalImportedAssetBytes || payload.bytes.length > 64 * 1024 * 1024) throw new PptxImportError("import-limit-exceeded", "Video inventory exceeds imported asset byte limits.", { part: mediaPart });
        total += payload.bytes.length;
        if (contentType !== "video/mp4") info = { status: "unsupported", contentType, reason: "Video content type is outside the MP4 inventory profile." };
        else {
          try {
            const framing = inspectIsoBmff(payload.bytes, { maxBytes: Math.min(limits.maxImportedAssetBytes, 64 * 1024 * 1024) });
            info = { status: "iso-bmff-framing", contentType, framing: { majorBrand: framing.majorBrand, minorVersion: framing.minorVersion, compatibleBrands: framing.compatibleBrands, boxCount: framing.boxes.length } };
          } catch (error) {
            if (!(error instanceof TypeError || error instanceof RangeError)) throw error;
            info = { status: "unsupported", contentType, reason: error.message };
          }
        }
        cache.set(mediaPart, info);
      }
      result.push({ ...base, mediaPart, ...info });
    }
    const primaries = result.slice(first).filter(item => item.declaration === "videoFile");
    const conflicting = children(nvPr).some(node => (node.namespaceUri === context.a || node.namespaceUri === context.p) && ["audioFile", "wavAudioFile", "quickTimeFile", "audioCd"].includes(node.local));
    const primary = primaries.length === 1 && !conflicting ? context.findRelationship(part, primaries[0]!.relationshipId) : undefined;
    const primaryPart = primary?.type === `${context.r}/video` && primary.targetMode === "Internal" && primary.fragment === undefined ? primary.targetPart : undefined;
    for (let index = first; index < result.length; index++) {
      const item = result[index]!;
      if (item.declaration !== "p14:media") continue;
      const relation = context.findRelationship(part, item.relationshipId);
      const relatedPart = item.relationshipAttribute === "embed" && relation?.type === MEDIA_REL && relation.targetMode === "Internal" && relation.fragment === undefined ? relation.targetPart : undefined;
      result[index] = { ...item, companionStatus: primaryPart && relatedPart ? primaryPart === relatedPart ? "same-part" : "different-part" : "unverified" };
    }
    byPicture.set(picture, result.slice(first));
  };
  return { result, inspect, forPicture: (picture: XmlElement) => byPicture.get(picture) ?? [] };
}
