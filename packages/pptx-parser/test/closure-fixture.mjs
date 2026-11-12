import { contentTypes, pack, REL, relationships } from "./opc-fixture.mjs";
export const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const closureTypes = () => contentTypes([{ Extension: "xml", ContentType: "application/xml" }, { Extension: "bin", ContentType: "application/octet-stream" }]);
export const sourceAlternates = (choice, fallback = '<Relationship Id="selected" Type="urn:asset" Target="active.bin"/>') =>
  `<Relationships xmlns="${REL}" xmlns:mc="${MC}" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent></Relationships>`;
export function closureParts() {
  return {
    "a.xml": "<a/>", "b.xml": "<b/>", "outside.xml": "<outside/>",
    "shared.bin": Buffer.from([1, 2]), "orphan.bin": Buffer.from([3, 4, 5]),
    "_rels/.rels": relationships([{ Id: "root", Type: "urn:root", Target: "a.xml" }]),
    "_rels/a.xml.rels": relationships([
      { Id: "b", Type: "urn:b", Target: "b.xml" }, { Id: "shared", Type: "urn:binary", Target: "shared.bin" },
      { Id: "self", Type: "urn:self", Target: "#fragment" },
      { Id: "external", Type: "urn:external", Target: "https://example.invalid/never", TargetMode: "External" },
    ]),
    "_rels/b.xml.rels": relationships([{ Id: "a", Type: "urn:cycle", Target: "a.xml" }, { Id: "shared", Type: "urn:binary", Target: "shared.bin" }]),
    "_rels/outside.xml.rels": relationships([{ Id: "shared", Type: "urn:binary", Target: "shared.bin" }]),
  };
}
export const closureBytes = (parts = closureParts()) => pack(parts, closureTypes());
export function alternateParts(choice = '<Relationship Id="hidden" Type="urn:asset" Target="shadow.bin"/>') {
  return { "a.xml": "<a/>", "active.bin": Buffer.from([1]), "shadow.bin": Buffer.from([2]), "_rels/a.xml.rels": sourceAlternates(choice) };
}
