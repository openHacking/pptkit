import { inspectWavePcm } from "@pptkit/media-inspection";
import { attribute, children } from "../presentation/structure-xml.js";
import type { structureContext } from "../presentation/structure-context.js";
import { PptxImportError, type ImportedAudioSource, type PptxSessionLimits } from "../types/import.js";
import type { XmlElement } from "../xml/read-xml.js";

type Context = Awaited<ReturnType<typeof structureContext>>;

/** Source inventory only: no audio assets, external loads, playback or edit bindings. */
export function audioSources(context: Context, limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>) {
  const result: ImportedAudioSource[] = [];
  const cache = new Map<string, Pick<ImportedAudioSource, "status" | "contentType" | "pcm" | "reason">>();
  let total = 0;
  const inspect = async (part: string, slideId: string, objectId: string, nvPr: XmlElement | undefined, paths: WeakMap<XmlElement, string>) => {
    if (!nvPr) return;
    for (const node of children(nvPr)) {
      await checkpoint();
      if (node.namespaceUri !== context.a || node.local !== "audioFile") continue;
      if (result.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Audio source inventory exceeds maxReportFindings.", { part });
      const raw = await context.sourceElement(part, node), relationshipId = attribute(node, "link", context.r);
      const source = { part, path: paths.get(node)!, sourceObjectId: objectId, elementIndex: raw.elementIndex, ...(relationshipId ? { relationshipId } : {}) };
      const base = { slideId, source, sourceElementIndex: raw.elementIndex, ...(relationshipId ? { relationshipId } : {}) };
      const relation = context.findRelationship(part, relationshipId);
      if (!relation || relation.type !== `${context.r}/audio` || relation.fragment !== undefined) {
        result.push({ ...base, status: "unsupported", reason: "Missing, fragmented or non-audio relationship." }); continue;
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
  };
  return { result, inspect };
}
