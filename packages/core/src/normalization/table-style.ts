import type { TableCellInput, TableElementInput, TextParagraphInput, TextRunInput } from "../types/element.js";
import type { StrokeStyleInput, TableCellStyleInput, TableRegionStyleInput, TableStyleInput, TextStylePresetInput } from "../types/style.js";
import { deepClone } from "../utils/clone.js";

import { tableStyleAnchors, type TableStyleAnchor as Anchor } from "../table/style-grid.js";
import type { PresentationInheritanceSource } from "../types/inheritance.js";
import type { StyleOriginMap } from "./style-origins.js";
import { StyleInputTrace, styleOriginsAt } from "./style-input-origins.js";
import { childSource, presetSource } from "./text-origins.js";
import { mergeTextParagraphStyleInput, mergeTextRunStyleInput } from "./normalize-styles.js";
const sides = ["left", "right", "top", "bottom"] as const;

/** Fieldwise strokes/insets; paint and discriminated unions replace atomically. */
function stroke(base: StrokeStyleInput | undefined, patch: StrokeStyleInput, trace?: StyleInputTrace): StrokeStyleInput {
  const fields = Object.entries(patch).filter(([, value]) => value !== undefined);
  for (const [key, value] of fields) if (key === "paint") trace?.replaceObject(key, value as object); else trace?.record(key);
  return { ...base, ...Object.fromEntries(fields) };
}
function margin(value: TableCellStyleInput["margin"]) {
  return typeof value === "number" ? { left: value, right: value, top: value, bottom: value } : value;
}
function layer(base: TableRegionStyleInput, patch: TableRegionStyleInput | undefined, trace?: StyleInputTrace): TableRegionStyleInput {
  if (!patch) return base;
  const next = { ...base };
  if (patch.fill !== undefined) { next.fill = patch.fill; trace?.replaceObject("fill", patch.fill); }
  if (patch.margin !== undefined) { next.margin = { ...margin(base.margin), ...Object.fromEntries(Object.entries(margin(patch.margin)!).filter(([, value]) => value !== undefined)) }; trace?.margin(patch.margin); }
  if (patch.verticalAlign !== undefined) { next.verticalAlign = patch.verticalAlign; trace?.record("verticalAlign"); }
  if (patch.textDirection !== undefined) { next.textDirection = patch.textDirection; trace?.record("textDirection"); }
  const borders = { ...base.borders };
  const seedBorder = (side: typeof sides[number]) => {
    if (borders[side] === undefined && trace) for (const [key, origin] of Object.entries(trace.origins)) {
      if (key.startsWith("stroke.")) trace.origins[`borders.${side}.${key.slice(7)}`] = origin;
    }
  };
  // A later uniform stroke applies to every side before its per-edge patch.
  if (patch.stroke !== undefined) {
    next.stroke = stroke(base.stroke, patch.stroke, trace?.scope("stroke"));
    for (const side of sides) { seedBorder(side); borders[side] = stroke(borders[side] ?? base.stroke, patch.stroke, trace && new StyleInputTrace(property => trace.source(property.replace(`borders.${side}.`, "stroke.")), trace.origins).scope(`borders.${side}`)); }
  }
  for (const side of sides) if (patch.borders?.[side] !== undefined) { seedBorder(side); borders[side] = stroke(borders[side] ?? next.stroke, patch.borders[side]!, trace?.scope(`borders.${side}`)); }
  next.borders = borders;
  if (patch.text !== undefined) next.text = {
    paragraph: mergeTextParagraphStyleInput(base.text?.paragraph, patch.text.paragraph, trace?.scope("text.paragraph")),
    run: mergeTextRunStyleInput(base.text?.run, patch.text.run, trace?.scope("text.run")),
  };
  return next;
}
function regions(style: TableStyleInput, a: Anchor, rows: number, columns: number): [string, TableRegionStyleInput | undefined][] {
  const top = a.row === 0, bottom = a.row + a.rowSpan === rows, left = a.column === 0, right = a.column + a.colSpan === columns;
  const band = (axis: "Rows" | "Columns"): [string, TableRegionStyleInput | undefined] => {
    const rule = style[`banded${axis}`]; if (!rule) return ["", undefined];
    const start = axis === "Rows" ? a.row : a.column, total = axis === "Rows" ? rows : columns;
    const first = axis === "Rows" ? style.firstRow : style.firstColumn, last = axis === "Rows" ? style.lastRow : style.lastColumn;
    if ((first !== undefined && start === 0) || (last !== undefined && start === total - 1)) return ["", undefined];
    const key = Math.floor((start - (first !== undefined ? 1 : 0)) / (rule.size ?? 1)) % 2 === 0 ? "first" : "second";
    return [`banded${axis}.${key}`, rule[key]];
  };
  return [band("Columns"), band("Rows"), ["firstColumn", left ? style.firstColumn : undefined], ["lastColumn", right ? style.lastColumn : undefined], ["firstRow", top ? style.firstRow : undefined], ["lastRow", bottom ? style.lastRow : undefined], ["topLeft", top && left ? style.topLeft : undefined], ["topRight", top && right ? style.topRight : undefined], ["bottomLeft", bottom && left ? style.bottomLeft : undefined], ["bottomRight", bottom && right ? style.bottomRight : undefined]];
}

export interface TableCellInputOrigins {
  style: StyleOriginMap;
  content: { style: StyleOriginMap; runs: StyleOriginMap[] }[];
}

/** Normalize one table's authoring cascade into detached effective cell inputs. */
export function tableStyleCells(table: TableElementInput, presets: Readonly<Record<string, TextStylePresetInput>>, owner?: PresentationInheritanceSource): { cells: TableCellInput[][]; origins?: TableCellInputOrigins[][] } {
  const style = table.style!, rows = table.rows.length, columns = table.columns.length;
  const origins: TableCellInputOrigins[][] | undefined = owner ? [] : undefined;
  const cells = tableStyleAnchors(table).map((anchors, rowIndex) => {
    const rowOrigins: TableCellInputOrigins[] = [];
    const row = anchors.map(a => {
      const declarations = owner ? new StyleInputTrace(property => childSource(owner, property)) : undefined;
      const trace = (source: PresentationInheritanceSource | undefined, rename: (property: string) => string = property => property) => source && declarations ? new StyleInputTrace(property => childSource(source, rename(property)), declarations.origins) : undefined;
      let resolved: TableRegionStyleInput = layer({}, style.wholeTable, trace(owner && childSource(owner, "style.wholeTable")));
      for (const side of sides) {
        const outer = side === "left" ? a.column === 0 : side === "right" ? a.column + a.colSpan === columns : side === "top" ? a.row === 0 : a.row + a.rowSpan === rows;
        const role = outer ? "outer" : side === "top" || side === "bottom" ? "innerHorizontal" : "innerVertical";
        const edge = style.borders?.[role];
        if (edge !== undefined) resolved = layer(resolved, { borders: { [side]: edge } }, trace(owner && childSource(owner, `style.borders.${role}`), property => property.slice(`borders.${side}.`.length)));
      }
      for (const [name, region] of regions(style, a, rows, columns)) resolved = layer(resolved, region, trace(owner && childSource(owner, `style.${name}`)));
      const preset = presets[a.cell.textStylePreset ?? ""];
      resolved = layer(resolved, {
        ...(preset?.frame?.margin !== undefined ? { margin: preset.frame.margin } : {}),
        ...(preset?.frame?.verticalAlign !== undefined ? { verticalAlign: preset.frame.verticalAlign } : {}),
        ...(preset !== undefined ? { text: { ...(preset.paragraph !== undefined ? { paragraph: preset.paragraph } : {}), ...(preset.run !== undefined ? { run: preset.run } : {}) } } : {}),
      }, trace(owner && presetSource(a.cell.textStylePreset ?? "", owner), property => property.startsWith("text.") ? property.slice(5) : `frame.${property}`));
      const inputIndex = a.inputIndex;
      const cellOwner = owner && (inputIndex === undefined ? owner : childSource(owner, `rows.${rowIndex}.cells.${inputIndex}`));
      resolved = layer(resolved, a.cell.style, trace(cellOwner && childSource(cellOwner, "style")));
      const { text, ...cellStyle } = resolved;
      const content: TextParagraphInput[] = typeof a.cell.content === "string" ? a.cell.content.replace(/\r\n?/g, "\n").split("\n").map(value => ({ runs: [{ text: value }] })) : a.cell.content;
      const paragraphs: TextParagraphInput[] = content.length ? content : [{ runs: [{ text: "" }] }];
      const contentOrigins: TableCellInputOrigins["content"] = [];
      const resolvedContent = paragraphs.map((p, pi) => {
        const runs: TextRunInput[] = p.runs.length ? p.runs : [{ text: "" }];
        const paragraphTrace = cellOwner && declarations ? new StyleInputTrace(property => childSource(cellOwner, `content.${pi}.style.${property.slice("text.paragraph.".length)}`), { ...declarations.origins }).scope("text.paragraph") : undefined;
        const style = mergeTextParagraphStyleInput(text?.paragraph, p.style, paragraphTrace);
        const runOrigins: StyleOriginMap[] = [];
        const resolvedRuns = runs.map((run, ri) => {
          const runTrace = cellOwner && declarations ? new StyleInputTrace(property => childSource(cellOwner, `content.${pi}.runs.${ri}.style.${property.slice("text.run.".length)}`), { ...declarations.origins }).scope("text.run") : undefined;
          const style = mergeTextRunStyleInput(text?.run, run.style, runTrace);
          if (runTrace) runOrigins.push(styleOriginsAt(runTrace.origins, "text.run"));
          return { ...run, style };
        });
        if (paragraphTrace) contentOrigins.push({ style: styleOriginsAt(paragraphTrace.origins, "text.paragraph"), runs: runOrigins });
        return { ...p, style, runs: resolvedRuns };
      });
      if (declarations) rowOrigins.push({ style: Object.fromEntries(Object.entries(declarations.origins).filter(([key]) => !key.startsWith("text."))), content: contentOrigins });
      return deepClone({ ...a.cell, content: resolvedContent, style: cellStyle });
    });
    origins?.push(rowOrigins); return row;
  });
  return { cells, ...(origins ? { origins } : {}) };
}
