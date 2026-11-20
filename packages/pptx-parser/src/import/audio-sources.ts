import { inspectWavePcm } from "@pptkit/media-inspection";
import { attribute, children } from "../presentation/structure-xml.js";
import type { structureContext } from "../presentation/structure-context.js";
import { PptxImportError, type ImportedAudioSource, type PptxSessionLimits } from "../types/import.js";
import type { XmlElement } from "../xml/read-xml.js";

type Context = Awaited<ReturnType<typeof structureContext>>;
const MEDIA_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const MEDIA_REL = "http://schemas.microsoft.com/office/2007/relationships/media";
const MEDIA_EXT = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

/** Source inventory only: no audio assets, external loads, playback or edit bindings. */
export function audioSources(context: Context, limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>) {
  const result: ImportedAudioSource[] = [];
  const byOwner = new WeakMap<XmlElement, readonly ImportedAudioSource[]>();
  const cache = new Map<string, Pick<ImportedAudioSource, "status" | "contentType" | "pcm" | "reason">>();
  let total = 0;
  const inspect = async (part: string, slideId: string, objectId: string, nvPr: XmlElement | undefined, paths: WeakMap<XmlElement, string>) => {
    if (!nvPr) return;
    const firstResult = result.length;
    const nodes = children(nvPr);
    const declarations: { node: XmlElement; declaration: NonNullable<ImportedAudioSource["declaration"]>; relationshipAttribute: "link" | "embed" }[] = [];
    let hasAudio = false;
    for (const node of nodes) {
      await checkpoint();
      if (node.namespaceUri !== context.a) continue;
      if (node.local === "audioFile") { hasAudio = true; declarations.push({ node, declaration: "audioFile", relationshipAttribute: "link" }); }
      if (node.local === "wavAudioFile") { hasAudio = true; declarations.push({ node, declaration: "wavAudioFile", relationshipAttribute: "embed" }); }
    }
    if (hasAudio) for (const list of nodes) if (list.namespaceUri === context.p && list.local === "extLst") for (const ext of children(list)) {
      await checkpoint();
      if (ext.namespaceUri !== context.p || ext.local !== "ext" || attribute(ext, "uri")?.toUpperCase() !== MEDIA_EXT) continue;
      for (const node of children(ext)) {
        await checkpoint();
        if (node.namespaceUri !== MEDIA_NS || node.local !== "media") continue;
        const embed = attribute(node, "embed", context.r), link = attribute(node, "link", context.r);
        if (embed !== undefined || link === undefined) declarations.push({ node, declaration: "p14:media", relationshipAttribute: "embed" });
        if (link !== undefined) declarations.push({ node, declaration: "p14:media", relationshipAttribute: "link" });
      }
    }
    for (const { node, declaration, relationshipAttribute } of declarations) {
      await checkpoint();
      if (result.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Audio source inventory exceeds maxReportFindings.", { part });
      const raw = await context.sourceElement(part, node), relationshipId = attribute(node, relationshipAttribute, context.r);
      const source = { part, path: paths.get(node)!, sourceObjectId: objectId, elementIndex: raw.elementIndex, ...(relationshipId ? { relationshipId } : {}) };
      const base = { slideId, declaration, relationshipAttribute, source, sourceElementIndex: raw.elementIndex, ...(relationshipId ? { relationshipId } : {}) };
      const relation = context.findRelationship(part, relationshipId);
      const expectedType = declaration === "p14:media" ? MEDIA_REL : `${context.r}/audio`;
      if (!relation || relation.type !== expectedType || relation.fragment !== undefined || (relationshipAttribute === "embed" && relation.targetMode !== "Internal") || (declaration === "p14:media" && relationshipAttribute === "link" && relation.targetMode !== "External")) {
        result.push({ ...base, status: "unsupported", reason: "Missing, fragmented, wrong-type or wrong-mode audio relationship." }); continue;
      }
      if (relation.targetMode === "External") { result.push({ ...base, status: "external", reason: "External audio is recorded but not fetched." }); continue; }
      if (!relation.targetPart) { result.push({ ...base, status: "unsupported", reason: "Audio relationship has no internal target." }); continue; }
      const mediaPart = relation.targetPart;
      let info = cache.get(mediaPart);
      if (!info) {
        const payload = context.payload(mediaPart), contentType = payload.contentType.trim().toLowerCase();
        if (payload.bytes.length > limits.maxImportedAssetBytes || total + payload.bytes.length > limits.maxTotalImportedAssetBytes || payload.bytes.length > 64 * 1024 * 1024) throw new PptxImportError("import-limit-exceeded", "Audio inventory exceeds imported asset byte limits.", { part: mediaPart });
        total += payload.bytes.length;
        if (contentType !== "audio/wav" && contentType !== "audio/x-wav") info = { status: "unsupported", contentType, reason: "Audio content type is outside the WAVE profile." };
        else {
          try {
            const { channels, sampleRate, bitsPerSample, frameCount, durationSeconds } = inspectWavePcm(payload.bytes, { maxBytes: Math.min(limits.maxImportedAssetBytes, 64 * 1024 * 1024) });
            info = { status: "wave-pcm", contentType, pcm: { channels, sampleRate, bitsPerSample, frameCount, durationSeconds } };
          } catch (error) {
            if (!(error instanceof TypeError || error instanceof RangeError)) throw error;
            info = { status: "unsupported", contentType, reason: error.message };
          }
        }
        cache.set(mediaPart, info);
      }
      result.push({ ...base, mediaPart, ...info });
    }
    const primary = result.slice(firstResult).filter(item => item.declaration !== "p14:media");
    const direct = primary.length === 1 ? context.findRelationship(part, primary[0]!.relationshipId) : undefined;
    const directPart = direct?.type === `${context.r}/audio` && direct.targetMode === "Internal" && direct.fragment === undefined ? direct.targetPart : undefined;
    for (let index = firstResult; index < result.length; index++) {
      const item = result[index]!;
      if (item.declaration !== "p14:media") continue;
      const related = context.findRelationship(part, item.relationshipId);
      const relatedPart = item.relationshipAttribute === "embed" && related?.type === MEDIA_REL && related.targetMode === "Internal" && related.fragment === undefined ? related.targetPart : undefined;
      const companionStatus = directPart && relatedPart ? directPart === relatedPart ? "same-part" : "different-part" : "unverified";
      result[index] = { ...item, companionStatus };
    }
    byOwner.set(nvPr, result.slice(firstResult));
  };
  return { result, inspect, forOwner: (nvPr: XmlElement | undefined) => nvPr ? byOwner.get(nvPr) ?? [] : [] };
}
