import { zip } from "./zip-fixture.mjs";
export const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
export const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
export const REL_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const escape = (s) => s.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
export const attrs = (object) => Object.entries(object).map(([name, value]) => `${name}="${escape(value)}"`).join(" ");
export const contentTypes = (defaults = [{ Extension: "xml", ContentType: "application/xml" }], overrides = []) => `<Types xmlns="${CT}">${defaults.map((d) => `<Default ${attrs(d)}/>`).join("")}${overrides.map((o) => `<Override ${attrs(o)}/>`).join("")}</Types>`;
export const relationships = (entries = []) => `<Relationships xmlns="${REL}">${entries.map((r) => `<Relationship ${attrs(r)}/>`).join("")}</Relationships>`;
export const pack = (parts = {}, types = contentTypes()) => zip([{ name: "[Content_Types].xml", body: types }, ...Object.entries(parts).map(([name, body]) => ({ name, body }))]).bytes;
export const basicParts = () => ({
  "doc/main.xml": '<doc xmlns="urn:document"><child/></doc>',
  "doc/data.bin": Buffer.from([0, 255, 1]),
  "_rels/.rels": relationships([{ Id: "root", Type: "urn:document", Target: "doc/main.xml" }]),
  "doc/_rels/main.xml.rels": relationships([{ Id: "data", Type: "urn:data", Target: "./data.bin" }]),
});
export const basicTypes = () => contentTypes([{ Extension: "XML", ContentType: "application/xml" }, { Extension: "bin", ContentType: "application/octet-stream" }]);
