import type { NormalizedElement, NormalizedPresentation, NormalizedTextParagraph } from "@pptkit/core";
import type { LayoutOptions } from "../types/analysis.js";
import type { LayoutElement, LayoutResult } from "../types/layout.js";

type Element = NormalizedElement | LayoutElement;
type FieldValues = Map<string, string>;

function fields(element: Element, sourceRows?: readonly number[]): FieldValues {
  const values: FieldValues = new Map();
  function paragraphs(content: readonly NormalizedTextParagraph[], prefix: readonly number[]): void {
    content.forEach((paragraph, p) => paragraph.runs.forEach((run, r) => {
      if (run.field) values.set(JSON.stringify([...prefix, p, r, run.field.id, run.field.kind]), run.text);
    }));
  }
  if (element.type === "text") paragraphs(element.content, []);
  else if (element.type === "shape" && element.text) paragraphs(element.text.content, []);
  else if (element.type === "table") element.rows.forEach((row, r) => row.cells.forEach((cell, c) => paragraphs(cell.content, [sourceRows?.[r] ?? r, c])));
  return values;
}

function elements(document: NormalizedPresentation | LayoutResult): Map<string, Element> {
  const result = new Map<string, Element>();
  const walk = (items: readonly Element[]): void => { for (const item of items) { result.set(item.id, item); if (item.type === "group") walk(item.children); } };
  for (const master of document.masters ?? []) walk(master.elements ?? []);
  for (const layout of document.layouts) walk(layout.elements);
  for (const slide of document.slides) walk(slide.elements);
  return result;
}

/** Capture fields in sizing targets after the provisional refresh, before measurement. */
export function captureSizingFields(document: NormalizedPresentation, options: LayoutOptions): Map<string, FieldValues> {
  const found = elements(document), snapshots = new Map<string, FieldValues>();
  for (const rule of [...(options.tableSizing ?? []), ...(options.intrinsicSizing ?? [])]) {
    const element = found.get(rule.elementId);
    if (element) snapshots.set(rule.elementId, fields(element));
  }
  return snapshots;
}

/** Reject output whose measured fields changed when the final page order became known. */
export function assertSizingFieldsStable(layout: LayoutResult, snapshots: Map<string, FieldValues>): void {
  const found = elements(layout);
  const fragments = new Map<string, { elementId: string; sourceRowIndices: number[] }[]>();
  for (const fragment of layout.pagination?.fragments ?? []) {
    const list = fragments.get(fragment.sourceElementId) ?? []; list.push(fragment); fragments.set(fragment.sourceElementId, list);
  }
  for (const [id, before] of snapshots) {
    const instances = fragments.get(id) ?? [{ elementId: id, sourceRowIndices: [] }];
    const seen = new Set<string>();
    const fail = (): never => { throw new TypeError(`Field text in sizing target "${id}" changed after pagination; its measured geometry cannot be used.`); };
    const rowKeys = new Map<number, string[]>();
    if (fragments.has(id)) for (const key of before.keys()) {
      const row = JSON.parse(key)[0] as number, list = rowKeys.get(row) ?? []; list.push(key); rowKeys.set(row, list);
    }
    for (const instance of instances) {
      const element = found.get(instance.elementId);
      if (!element) { if (before.size) fail(); continue; }
      const after = fields(element, instance.sourceRowIndices.length ? instance.sourceRowIndices : undefined);
      for (const [key, value] of after) { if (before.get(key) !== value) fail(); seen.add(key); }
      if (instance.sourceRowIndices.length) {
        for (const row of instance.sourceRowIndices) for (const key of rowKeys.get(row) ?? []) if (!after.has(key)) fail();
      }
    }
    for (const key of before.keys()) if (!seen.has(key)) fail();
  }
}
