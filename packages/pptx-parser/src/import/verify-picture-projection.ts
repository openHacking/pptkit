import type { NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import { PptxImportError, type ImportedPictureSource } from "../types/import.js";
import { equalState } from "./state.js";

/** Compare full default projections using source part names instead of transient asset IDs. */
export function verifyPictureProjection(before: NormalizedPresentation, after: NormalizedPresentation, beforeSources: readonly ImportedPictureSource[], afterSources: readonly ImportedPictureSource[], target: ImportedPictureSource, newPart: string, newRelationshipId: string, maxValues: number): void {
  const fail = (): never => { throw new PptxImportError("save-candidate-invalid", "Picture candidate changed mapped content outside the selected reference."); };
  const key = (source: ImportedPictureSource) => JSON.stringify([source.slideId, source.elementId]);
  const index = new Map(afterSources.map(source => [key(source), source]));
  if (index.size !== afterSources.length || beforeSources.length !== afterSources.length || new Set(beforeSources.map(key)).size !== beforeSources.length) fail();
  const targetKey = key(target), targetBefore = beforeSources.find(source => key(source) === targetKey), targetAfter = index.get(targetKey);
  if (!targetBefore || !targetAfter || !equalState(targetBefore, target, maxValues) || targetAfter.mediaPart !== newPart || targetAfter.relationshipId !== newRelationshipId) fail();
  for (const source of beforeSources) {
    const next = index.get(key(source)) ?? fail();
    const common = (item: ImportedPictureSource) => ({ slideId: item.slideId, elementId: item.elementId, source: { ...item.source, relationshipId: undefined }, sourceElementIndex: item.sourceElementIndex, hasSvgExtension: item.hasSvgExtension, hasExternalLink: item.hasExternalLink });
    if (!equalState(common(source), common(next), maxValues) || (key(source) !== targetKey && (source.mediaPart !== next.mediaPart || source.relationshipId !== next.relationshipId))) fail();
  }
  const canonical = (snapshot: NormalizedPresentation, sources: readonly ImportedPictureSource[], changeTarget: boolean) => {
    const result = structuredClone(snapshot), assets = new Map<string, string>(), referenced = new Set<string>();
    for (const source of sources) {
      const existing = assets.get(source.assetId);
      if (existing !== undefined && existing !== source.mediaPart) fail();
      assets.set(source.assetId, source.mediaPart);
    }
    let changed = 0;
    for (const slide of result.slides) {
      const stack: NormalizedElement[] = [...slide.elements];
      while (stack.length) {
        const element = stack.pop()!;
        if (element.type === "group") stack.push(...element.children);
        if (element.type !== "image") continue;
        const part = assets.get(element.assetId) ?? fail();
        if (changeTarget && slide.id === target.slideId && element.id === target.elementId) { if (element.assetId !== target.assetId) fail(); element.assetId = newPart; changed++; }
        else element.assetId = part;
        referenced.add(element.assetId);
      }
    }
    if (changeTarget && changed !== 1) fail();
    result.assets = result.assets.map(asset => ({ ...asset, id: assets.get(asset.id) ?? fail() })).filter(asset => !changeTarget || referenced.has(asset.id)).sort((a, b) => a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
    return result;
  };
  const expected = canonical(before, beforeSources, true), actual = canonical(after, afterSources, false);
  const replacement = actual.assets.find(asset => asset.id === newPart) ?? fail();
  if (expected.assets.some(asset => asset.id === newPart)) fail();
  expected.assets.push(replacement); expected.assets.sort((a, b) => a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
  if (!equalState(expected, actual, maxValues)) fail();
}
