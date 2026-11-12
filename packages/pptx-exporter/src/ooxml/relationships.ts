import { REL } from "../constants/ooxml.js";
import type { Relationship } from "../types/internal.js";
import { escapeXml } from "./xml.js";

export function relationshipsXml(relationships: Relationship[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map((relationship) => `<Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(relationship.target)}"${relationship.targetMode === undefined ? "" : ` TargetMode="${escapeXml(relationship.targetMode)}"`}/>`).join("")}</Relationships>`;
}

export function rootRelationshipsXml(customProperties = false): string {
  return relationshipsXml([
    { id: "rId1", type: REL.officeDocument, target: "ppt/presentation.xml" },
    { id: "rId2", type: REL.core, target: "docProps/core.xml" },
    { id: "rId3", type: REL.app, target: "docProps/app.xml" },
    ...(customProperties ? [{ id: "rId4", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", target: "docProps/custom.xml" }] : []),
  ]);
}
