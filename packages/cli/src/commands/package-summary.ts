import type { OpcPackageInspection } from '@pptkit/pptx-parser';

/** Summarize declared media parts without inspecting their payload or inferring reachability. */
export function packageSummary(pkg: OpcPackageInspection) {
  const entries = new Map(pkg.archive.entries.map(entry => [entry.name, entry]));
  const inbound = new Map<string, number>();
  for (const relationship of pkg.relationships) if (relationship.targetMode === 'Internal' && relationship.targetPart) inbound.set(relationship.targetPart, (inbound.get(relationship.targetPart) ?? 0) + 1);
  const media = pkg.parts.filter(part => /^(?:image|audio|video)\//i.test(part.contentType)).map(part => {
    const entry = entries.get(part.entryName);
    if (!entry) throw new Error('OPC part has no archive entry.');
    return { part: part.name, contentType: part.contentType, compressedBytes: entry.compressedBytes, expandedBytes: part.expandedBytes, incomingRelationships: inbound.get(part.name) ?? 0 };
  }).sort((a, b) => a.part < b.part ? -1 : a.part > b.part ? 1 : 0);
  return {
    schemaVersion: 'pptkit-cli-package-summary-v1',
    mediaClassification: 'declared-content-type',
    archive: { inputBytes: pkg.archive.inputBytes, expandedBytes: pkg.archive.expandedBytes, entryCount: pkg.archive.entries.length, partCount: pkg.parts.length },
    relationships: { total: pkg.relationships.length, internal: pkg.relationships.filter(r => r.targetMode === 'Internal').length, external: pkg.relationships.filter(r => r.targetMode === 'External').length },
    mediaTotals: { parts: media.length, compressedBytes: media.reduce((sum, part) => sum + part.compressedBytes, 0), expandedBytes: media.reduce((sum, part) => sum + part.expandedBytes, 0) },
    media,
  };
}
