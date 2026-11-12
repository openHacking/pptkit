import { importPptx } from "@pptkit/pptx-parser";

export async function preserveOriginal(bytes: Uint8Array, signal?: AbortSignal) {
  const session = await importPptx(bytes, {
    projection: "structure",
    ...(signal ? { signal } : {}),
    sessionLimits: { maxLivePlans: 2 },
  });
  try {
    // This document contains the structural projection, not source drawing content.
    const sourceSlides = session.report.identities.filter((identity) => identity.kind === "slide");
    const prepared = await session.prepareSave({ mode: "no-op", ...(signal ? { signal } : {}) });
    if (prepared.status === "blocked") return prepared;
    try {
      const output = await session.save(prepared.plan);
      return { status: "saved" as const, bytes: output.bytes, sourceSlides, report: output.report };
    } finally {
      prepared.plan.dispose();
    }
  } finally {
    session.dispose();
  }
}
