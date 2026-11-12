import { importPptx } from "@pptkit/pptx-parser";

export async function inspectEditableContent(bytes: Uint8Array) {
  const session = await importPptx(bytes, {
    projection: "content",
    sessionLimits: { maxImportedAssetBytes: 16 * 1024 * 1024, maxTotalImportedAssetBytes: 64 * 1024 * 1024 },
  });
  try {
    const objects = session.report.identities.filter(identity => identity.kind === "element");
    const imageSources = session.report.identities.filter(identity => identity.kind === "asset");
    const limitations = session.report.findings.filter(finding =>
      finding.operation === "import" && (finding.tier === "fallback" || finding.tier === "unsupported"));
    const prepared = await session.prepareSave({ mode: "no-op" });
    if (prepared.status !== "ready") throw new Error("Unchanged import unexpectedly cannot prepare a no-op save.");
    try {
      const original = await session.save(prepared.plan);
      return { objects, imageSources, limitations, originalBytes: original.bytes };
    } finally {
      prepared.plan.dispose();
    }
  } finally {
    session.dispose();
  }
}
