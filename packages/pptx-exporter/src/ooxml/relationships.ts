import { REL } from "../constants/ooxml.js";
import type { Relationship } from "../types/internal.js";

export function relationshipsXml(relationships: Relationship[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map((relationship) => `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"/>`).join("")}</Relationships>`;
}

export function rootRelationshipsXml(): string {
  return relationshipsXml([
    { id: "rId1", type: REL.officeDocument, target: "ppt/presentation.xml" },
    { id: "rId2", type: REL.core, target: "docProps/core.xml" },
    { id: "rId3", type: REL.app, target: "docProps/app.xml" },
  ]);
}
