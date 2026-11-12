import { DEFAULT_ZIP_LIMITS, readArchive } from "../archive/read-archive.js";
import { PackageInspectionError } from "../types/archive.js";
import type { OpcInspectionLimits, OpcInspectionOptions, OpcPackageInspection, OpcPartInfo, OpcRelationshipInfo } from "../types/opc.js";
import { DEFAULT_XML_LIMITS, readXml } from "../xml/read-xml.js";
import type { XmlDocument, XmlElement } from "../xml/read-xml.js";
import { captureConfiguration, MCE } from "../compatibility/configuration.js";
import { processCompatibility } from "../compatibility/process.js";
import { asciiKey, iriReference, partName, relationshipSource, resolveInternal } from "./names.js";
import { attributesOf, controlEncoding, CONTENT_TYPES_NS, elementsOf, mediaType, packageInvalid, RELATIONSHIPS_NS, RELATIONSHIPS_TYPE, XML_ID } from "./markup.js";

export const DEFAULT_OPC_LIMITS: Readonly<OpcInspectionLimits> = Object.freeze({ maxXmlParts: 10000, maxXmlCharacters: 32 * 1024 * 1024, maxXmlNodes: 1000000, maxXmlTextCharacters: 16 * 1024 * 1024, maxRelationships: 100000 });
const TYPES_ENTRY = "[Content_Types].xml";
interface Part extends OpcPartInfo { contentType: string; xml?: NonNullable<OpcPartInfo["xml"]>; markupCompatibility?: NonNullable<OpcPartInfo["markupCompatibility"]> }
export function snapshotOptions(options: OpcInspectionOptions) {
  if (!options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Invalid OPC inspection options");
  for (const value of [options.limits, options.xmlLimits, options.zipLimits]) if (value !== undefined && (!value || typeof value !== "object" || Array.isArray(value))) throw new TypeError("Invalid OPC limit options");
  const limits = { ...DEFAULT_OPC_LIMITS, ...options.limits }, xmlLimits = { ...DEFAULT_XML_LIMITS, ...options.xmlLimits }, zipLimits = { ...DEFAULT_ZIP_LIMITS, ...options.zipLimits };
  for (const group of [limits, xmlLimits, zipLimits]) for (const [name, value] of Object.entries(group)) if (!Number.isFinite(value) || value <= 0 || (name !== "maxCompressionRatio" && !Number.isSafeInteger(value))) throw new RangeError(`Invalid OPC limit ${name}`);
  return { limits, xmlLimits, zipLimits, signal: options.signal };
}

/** Validates an OPC envelope; does not establish a presentation's semantic validity. */
export async function readOpcPackage(input: Uint8Array, options: OpcInspectionOptions = {}): Promise<OpcPackageInspection> {
  return (await readOpcSource(input, options)).inspection;
}

/** Retained trees stay private to format readers and share the original aggregate parsing budget. */
export async function readOpcSource(input: Uint8Array, options: OpcInspectionOptions = {}, retainXml: boolean | "relationships" = false, retainPayloads = false): Promise<{ inspection: OpcPackageInspection; xml: ReadonlyMap<string, XmlElement>; xmlDocuments: ReadonlyMap<string, XmlDocument>; payloads?: ReadonlyMap<string, Uint8Array> }> {
  const { limits, xmlLimits, zipLimits, signal } = snapshotOptions(options);
  const xmlParts = new Map<string, XmlElement>();
  const xmlDocuments = new Map<string, XmlDocument>();
  const abort = (): void => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "Package inspection was aborted."); };
  const cooperate = async (): Promise<void> => { await new Promise<void>((resolve) => setTimeout(resolve, 0)); abort(); };
  const limit = (message: string, entry?: string): never => { throw new PackageInspectionError("package-limit-exceeded", message, entry); };
  // readArchive snapshots the caller's bytes synchronously before its first await.
  const { inspection: archive, parts: payloads } = await readArchive(input, { limits: zipLimits, ...(signal ? { signal } : {}) }, true);
  abort();
  for (const entry of archive.entries) if (!entry.directory && (/(?:^|\/)\[\d+\](?:\.last)?\.piece$/i.test(entry.name) || asciiKey(entry.name).startsWith("[trash]/"))) {
    throw new PackageInspectionError("package-unsupported", "Interleaved or trash physical items require unsupported package mapping.", entry.name);
  }
  const typesBytes = payloads.get(TYPES_ENTRY);
  if (!typesBytes) packageInvalid("Missing exact [Content_Types].xml stream.", TYPES_ENTRY);
  const parts: Part[] = [], byKey = new Map<string, Part>();
  for (const entry of archive.entries) {
    if (entry.directory || entry.name === TYPES_ENTRY) continue;
    const name = partName(`/${entry.name}`, entry.name, true), key = asciiKey(name);
    if (name.split("/").some((segment) => asciiKey(segment) === "[content_types].xml")) packageInvalid("Media Types stream cannot be a part.", entry.name);
    if (byKey.has(key)) packageInvalid("Equivalent OPC part names.", entry.name);
    const part: Part = { name, entryName: entry.name, contentType: "", relationshipPart: relationshipSource(name) !== undefined, expandedBytes: entry.expandedBytes };
    parts.push(part); byKey.set(key, part);
  }
  const totals = { parts: 0, characters: 0, nodes: 0, textCharacters: 0 };
  const metricLimits = { characters: limits.maxXmlCharacters, nodes: limits.maxXmlNodes, textCharacters: limits.maxXmlTextCharacters };
  const xml = async (entry: string, retain: boolean) => {
    abort();
    if (++totals.parts > limits.maxXmlParts) limit("Package XML part count exceeds maxXmlParts.", entry);
    if (totals.parts % 16 === 0) await cooperate();
    const bytes = payloads.get(entry)!;
    return readXml(bytes, { limits: xmlLimits, partName: entry, ...(signal ? { signal } : {}) }, retain, (metric, increment) => {
      totals[metric] += increment;
      if (totals[metric] > metricLimits[metric]) limit(`Package XML ${metric} budget exceeded.`, entry);
    });
  };
  const typesParsed = await xml(TYPES_ENTRY, true);
  if (retainXml === true && typesParsed.document) xmlDocuments.set(`/${TYPES_ENTRY}`, typesParsed.document);
  controlEncoding(typesParsed.declaredEncoding, TYPES_ENTRY);
  const types = typesParsed.root!;
  if (types.namespaceUri !== CONTENT_TYPES_NS || types.local !== "Types") packageInvalid("Invalid Media Types stream root.", TYPES_ENTRY);
  attributesOf(types, [], [], TYPES_ENTRY);
  const defaults = new Map<string, { extension: string; contentType: string }>(), overrides = new Map<string, { partName: string; contentType: string }>();
  const definitions = elementsOf(types, TYPES_ENTRY);
  for (const definition of definitions) {
    if (definition.namespaceUri !== CONTENT_TYPES_NS || !["Default", "Override"].includes(definition.local)) packageInvalid("Unknown Media Types declaration.", TYPES_ENTRY);
    if (elementsOf(definition, TYPES_ENTRY).length) packageInvalid("Media Types declarations must be empty elements.", TYPES_ENTRY);
    const keyName = definition.local === "Default" ? "Extension" : "PartName";
    const attrs = attributesOf(definition, [keyName, "ContentType"], [keyName, "ContentType"], TYPES_ENTRY), contentType = attrs.get("ContentType")!;
    mediaType(contentType, TYPES_ENTRY);
    if (definition.local === "Default") {
      const extension = attrs.get("Extension")!;
      if (!/^(?:[!$&'()*+,:=@A-Za-z0-9_~-]|%[\da-f]{2})+$/i.test(extension)) packageInvalid("Invalid default extension spelling.", TYPES_ENTRY);
      const key = asciiKey(extension);
      if (defaults.has(key)) packageInvalid("Duplicate default content type extension.", TYPES_ENTRY);
      defaults.set(key, { extension, contentType });
    } else {
      const name = partName(attrs.get("PartName")!, TYPES_ENTRY), key = asciiKey(name), target = byKey.get(key);
      if (!target) packageInvalid("Content type override references an absent part.", TYPES_ENTRY);
      if (overrides.has(key)) packageInvalid("Duplicate content type override.", TYPES_ENTRY);
      overrides.set(key, { partName: target.name, contentType });
    }
  }
  for (const part of parts) {
    const last = part.name.split("/").at(-1)!, dot = last.lastIndexOf("."), extension = dot < 0 ? "" : asciiKey(last.slice(dot + 1));
    const declared = overrides.get(asciiKey(part.name))?.contentType ?? defaults.get(extension)?.contentType;
    part.contentType = declared ?? (part.relationshipPart ? RELATIONSHIPS_TYPE : "");
    if (!part.contentType) packageInvalid("No content type for OPC part.", part.entryName);
    const type = mediaType(part.contentType, part.entryName);
    if (part.relationshipPart !== (type === RELATIONSHIPS_TYPE)) packageInvalid("Relationship part name and content type disagree.", part.entryName);
    if (part.relationshipPart && part.contentType.toLowerCase() !== RELATIONSHIPS_TYPE) packageInvalid("Relationship content type cannot have parameters.", part.entryName);
  }
  const relationships: OpcRelationshipInfo[] = [];
  let work = 0;
  for (const part of parts) {
    abort();
    const baseType = mediaType(part.contentType, part.entryName);
    if (baseType !== "application/xml" && baseType !== "text/xml" && !baseType.endsWith("+xml")) continue;
    const retainPart = retainXml === true || part.relationshipPart;
    const parsed = await xml(part.entryName, retainPart);
    if ((retainXml === true || (retainXml === "relationships" && part.relationshipPart)) && parsed.root) xmlParts.set(part.name, parsed.root);
    if ((retainXml === true || (retainXml === "relationships" && part.relationshipPart)) && parsed.document) xmlDocuments.set(part.name, parsed.document);
    if ([RELATIONSHIPS_TYPE, "application/vnd.openxmlformats-package.core-properties+xml", "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"].includes(baseType)) controlEncoding(parsed.declaredEncoding, part.entryName);
    part.xml = parsed.inspection;
    if (!part.relationshipPart) continue;
    let root = parsed.root!;
    if (parsed.inspection.namespaceUris.includes(MCE)) {
      const projected = await processCompatibility(root, captureConfiguration({ understoodNamespaces: [RELATIONSHIPS_NS] }), part.entryName, signal);
      part.markupCompatibility = projected.report;
      const roots = projected.nodes.filter((node): node is XmlElement => node.kind === "element");
      if (roots.length !== 1 || projected.nodes.some((node) => (node.kind === "text" || node.kind === "cdata") && !/^[\t\r\n ]*$/.test(node.value))) packageInvalid("Processed relationship markup must have one root element.", part.entryName);
      root = roots[0]!;
    }
    const sourceName = relationshipSource(part.name)!;
    const source = sourceName === "/" ? "/" : byKey.get(asciiKey(sourceName))?.name;
    if (!source || (source !== "/" && byKey.get(asciiKey(source))!.relationshipPart)) packageInvalid("Relationship part has a missing or forbidden source.", part.entryName);
    if (root.namespaceUri !== RELATIONSHIPS_NS || root.local !== "Relationships") packageInvalid("Invalid Relationships root.", part.entryName);
    attributesOf(root, [], [], part.entryName, "unsupported");
    const ids = new Set<string>();
    for (const element of elementsOf(root, part.entryName)) {
      if (++work % 512 === 0) await cooperate();
      if (element.namespaceUri !== RELATIONSHIPS_NS) throw new PackageInspectionError("package-unsupported", "Relationship markup extensions require an MCE processor.", part.entryName);
      if (element.local !== "Relationship" || elementsOf(element, part.entryName, true).length) packageInvalid("Invalid Relationship element.", part.entryName);
      const attrs = attributesOf(element, ["Id", "Type", "Target", "TargetMode"], ["Id", "Type", "Target"], part.entryName, "unsupported");
      const id = attrs.get("Id")!.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "").replace(/[\t\r\n ]+/g, " "), type = attrs.get("Type")!, target = attrs.get("Target")!, mode = attrs.get("TargetMode") ?? "Internal";
      if (!XML_ID.test(id) || ids.has(id)) packageInvalid("Invalid or duplicate relationship Id.", part.entryName);
      ids.add(id); iriReference(type, part.entryName); iriReference(target, part.entryName);
      if (mode !== "Internal" && mode !== "External") packageInvalid("Invalid relationship TargetMode.", part.entryName);
      if (relationships.length >= limits.maxRelationships) limit("Package relationship count exceeds maxRelationships.", part.entryName);
      const base = { id, type, source, relationshipPart: part.name, target, targetMode: mode } as const;
      if (mode === "External") relationships.push(base);
      else {
        const resolved = resolveInternal(source, target, part.entryName), destination = byKey.get(asciiKey(resolved.name));
        if (!destination) packageInvalid("Internal relationship target is absent.", part.entryName);
        if (destination.relationshipPart) packageInvalid("Relationships cannot target relationship parts.", part.entryName);
        relationships.push({ ...base, targetPart: destination.name, ...(resolved.fragment !== undefined ? { fragment: resolved.fragment } : {}) });
      }
    }
  }
  abort();
  return { inspection: { archive, parts, relationships, contentTypes: { defaults: [...defaults.values()], overrides: [...overrides.values()] }, xmlTotals: totals }, xml: xmlParts, xmlDocuments, ...(retainPayloads ? { payloads } : {}) };
}
