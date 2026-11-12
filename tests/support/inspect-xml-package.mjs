// Runs in a browser with a standards XML parser. No string-based XML parsing.
// This is structural/relationship verification, not an OOXML schema validator.
export function inspectXmlPackage(parts) {
  const issues = [];
  const documents = new Map();
  const names = new Set(parts.map(([name]) => name));
  for (const [name, text] of parts) {
    if (!name.endsWith(".xml") && !name.endsWith(".rels")) continue;
    const document = new DOMParser().parseFromString(text, "application/xml");
    if (document.getElementsByTagName("parsererror").length) issues.push(`${name}: invalid XML`);
    documents.set(name, document);
  }
  const relationships = new Map();
  const relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
  for (const [name, document] of documents) {
    if (!name.endsWith(".rels")) continue;
    const source = name === "_rels/.rels" ? "" : name.replace(/(^|\/)\_rels\//, "$1").slice(0, -5);
    if (source && !names.has(source)) issues.push(`${name}: missing source ${source}`);
    const ids = new Set();
    relationships.set(source, ids);
    for (const element of document.getElementsByTagNameNS(relationshipNamespace, "Relationship")) {
      const id = element.getAttribute("Id");
      if (!id || ids.has(id)) issues.push(`${name}: empty or duplicate relationship ${id}`);
      ids.add(id);
      if (!element.getAttribute("Type")) issues.push(`${name}: missing relationship type`);
      const target = element.getAttribute("Target");
      if (!target) issues.push(`${name}: missing target`);
      if (element.getAttribute("TargetMode") === "External") continue;
      const resolved = new URL(target ?? "", `https://package.invalid/${source}`).pathname.slice(1);
      if (!names.has(decodeURIComponent(resolved))) issues.push(`${name}: missing target ${resolved}`);
    }
  }
  const officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  for (const [name, document] of documents) {
    for (const element of document.getElementsByTagName("*")) {
      for (const attribute of element.attributes) {
        if (attribute.namespaceURI === officeRelationships && !relationships.get(name)?.has(attribute.value)) {
          issues.push(`${name}: unresolved ${attribute.name}=${attribute.value}`);
        }
      }
    }
  }
  const types = documents.get("[Content_Types].xml");
  if (!types) return [...issues, "Missing content types"];
  const typeNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
  const defaults = new Set([...types.getElementsByTagNameNS(typeNamespace, "Default")].map((node) => node.getAttribute("Extension")));
  const overrides = new Set();
  for (const node of types.getElementsByTagNameNS(typeNamespace, "Override")) {
    const name = node.getAttribute("PartName")?.slice(1);
    if (!names.has(name)) issues.push(`Content type references absent part ${name}`);
    if (overrides.has(name)) issues.push(`Duplicate content type ${name}`);
    overrides.add(name);
  }
  for (const name of names) {
    if (name === "[Content_Types].xml") continue;
    if (!overrides.has(name) && !defaults.has(name.split(".").at(-1))) issues.push(`No content type for ${name}`);
  }
  return issues;
}
