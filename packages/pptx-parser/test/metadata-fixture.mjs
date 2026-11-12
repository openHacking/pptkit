import { structureFixture } from "./structure-fixture.mjs";
import { contentTypes, pack, relationships } from "./opc-fixture.mjs";
export const CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
export const DC = "http://purl.org/dc/elements/1.1/", DT = "http://purl.org/dc/terms/", XSI = "http://www.w3.org/2001/XMLSchema-instance";
export const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const GUID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
export const custom = (pid, type, text = "", attrs = "") => `<property pid="${pid}" fmtid="${GUID}" ${attrs}><v:${type}>${text}</v:${type}></property>`;
export function metadataFixture(strict = false) {
  const f = structureFixture(strict), base = strict ? "http://purl.oclc.org/ooxml/officeDocument" : "http://schemas.openxmlformats.org/officeDocument/2006";
  f.metadataTypes = {};
  f.rootRelations = [f.rel("office", "officeDocument", "deck/main.data")];
  f.VT = `${base}/docPropsVTypes`;
  f.core = (body = "", attrs = "") => `<c:coreProperties xmlns:c="${CP}" xmlns:d="${DC}" xmlns:t="${DT}" xmlns:xsi="${XSI}" ${attrs}>${body}</c:coreProperties>`;
  f.properties = (key, body = "", attrs = "") => `<Properties xmlns="${base}/${strict ? `${key}Properties` : `${key}-properties`}" xmlns:v="${f.VT}" ${attrs}>${body}</Properties>`;
  f.add = (key, body) => {
    const part = `inventory/${key}.data`;
    f.parts[part] = key === "core" ? f.core(body) : f.properties(key, body);
    f.metadataTypes[part] = key === "core" ? "application/vnd.openxmlformats-package.core-properties+xml" : `application/vnd.openxmlformats-officedocument.${key}-properties+xml`;
    f.rootRelations.push({ Id: key, Target: part, Type: key === "core" ? "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" : `${base}/relationships/${strict ? `${key}Properties` : `${key}-properties`}` });
    return part;
  };
  return f;
}
export function encodeMetadata(f) {
  f.parts["_rels/.rels"] = relationships(f.rootRelations);
  const types = { ...Object.fromEntries(Object.entries(f.types).map(([name, type]) => [name, `application/vnd.openxmlformats-officedocument.${type}+xml`])), ...f.metadataTypes };
  return pack(f.parts, contentTypes(undefined, Object.entries(types).map(([name, type]) => ({ PartName: `/${name}`, ContentType: type }))));
}
