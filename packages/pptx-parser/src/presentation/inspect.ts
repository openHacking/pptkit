import { readOpcPackage } from "../opc/read-package.js";
import { PackageInspectionError } from "../types/archive.js";
import type { OpcInspectionOptions, OpcPackageInspection } from "../types/opc.js";
import type { PresentationDialect, PresentationFileKind, PresentationPackageInspection } from "../types/presentation.js";

const ENTRY_RELATIONSHIPS = new Map<string, PresentationDialect>([
  ["http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "transitional"],
  ["http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument", "strict"],
]);
export const ROOT_NAMESPACES: Readonly<Record<PresentationDialect, string>> = {
  transitional: "http://schemas.openxmlformats.org/presentationml/2006/main",
  strict: "http://purl.oclc.org/ooxml/presentationml/main",
};
interface Format { kind: PresentationFileKind; macroEnabled: boolean; suggestedExtension: PresentationPackageInspection["suggestedExtension"] }
const MAIN_TYPES = new Map<string, Format>();
for (const [kind, extension] of [["presentation", "ppt"], ["template", "pot"], ["slideshow", "pps"]] as const) {
  MAIN_TYPES.set(`application/vnd.openxmlformats-officedocument.presentationml.${kind}.main+xml`, { kind, macroEnabled: false, suggestedExtension: `${extension}x` });
  MAIN_TYPES.set(`application/vnd.ms-powerpoint.${kind}.macroenabled.main+xml`, { kind, macroEnabled: true, suggestedExtension: `${extension}m` });
}

/** Resolve the package entry point without assuming any conventional filename. */
export async function readPresentationPackage(input: Uint8Array, options?: OpcInspectionOptions): Promise<PresentationPackageInspection> {
  const inspection = await readOpcPackage(input, options);
  return identifyPresentation(inspection);
}

export function identifyPresentation(inspection: OpcPackageInspection, checkRawRoot = true): PresentationPackageInspection {
  const entries = inspection.relationships.filter((r) => r.source === "/" && ENTRY_RELATIONSHIPS.has(r.type));
  if (entries.length !== 1) throw new PackageInspectionError("presentation-invalid", "Expected exactly one package officeDocument relationship.", "_rels/.rels");
  const entry = entries[0]!;
  if (entry.targetMode !== "Internal" || entry.fragment !== undefined) throw new PackageInspectionError("presentation-invalid", "Presentation entry must target an internal whole part.", entry.relationshipPart);
  const main = inspection.parts.find((part) => part.name === entry.targetPart)!;
  const format = MAIN_TYPES.get(main.contentType.toLowerCase());
  if (!format) throw new PackageInspectionError("presentation-unsupported", "Main part content type is not a supported PresentationML document type.", main.entryName);
  const dialect = ENTRY_RELATIONSHIPS.get(entry.type)!;
  if (checkRawRoot && (main.xml?.root.local !== "presentation" || main.xml.root.namespaceUri !== ROOT_NAMESPACES[dialect])) {
    throw new PackageInspectionError("presentation-invalid", "Presentation root and entry relationship dialect disagree.", main.entryName);
  }
  return { package: inspection, mainPart: main.name, mainRelationshipId: entry.id, dialect, ...format };
}
