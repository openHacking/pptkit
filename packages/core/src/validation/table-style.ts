import type { PresentationDiagnostic } from "../types/diagnostic.js";
import type { TableElementInput } from "../types/element.js";
import { tableStyleAnchors, TableStyleGridError } from "../table/style-grid.js";
import { isValidColor, isValidFontFamily, isValidPaint, scriptFontFamilyErrors } from "./style.js";
import { tabStopErrors } from "./tab-stops.js";

/** Table cascade has an intentionally closed grammar; unknown rules never disappear. */
export function validateTableStyle(table: TableElementInput, path: string, identity: Partial<PresentationDiagnostic>): PresentationDiagnostic[] {
  const errors: PresentationDiagnostic[] = [];
  const fail = (at: string, message: string) => errors.push({ ...identity, code: "invalid-table-style", severity: "error", path: at, message });
  const record = (value: unknown, keys: readonly string[], at: string): value is Record<string, unknown> => {
    if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) { fail(at, "Expected a plain table-style record."); return false; }
    for (const key of Object.keys(value)) if (!keys.includes(key)) fail(`${at}.${key}`, "Unknown table-style field.");
    return true;
  };
  const positive = (value: unknown) => typeof value === "number" && Number.isFinite(value) && value > 0;
  const nonnegative = (value: unknown) => typeof value === "number" && Number.isFinite(value) && value >= 0;
  const stroke = (value: unknown, at: string) => {
    if (!record(value, ["paint", "width", "dash", "beginArrow", "endArrow"], at)) return;
    if (!isValidPaint(value.paint as never)) fail(`${at}.paint`, "Invalid border paint.");
    if (value.width !== undefined && !nonnegative(value.width)) fail(`${at}.width`, "Border width must be finite and nonnegative.");
    if (value.dash !== undefined && !["solid", "dash", "dot", "dashDot"].includes(value.dash as string)) fail(`${at}.dash`, "Unsupported border dash.");
    for (const key of ["beginArrow", "endArrow"]) if (value[key] !== undefined && value[key] !== "none") fail(`${at}.${key}`, "Table borders do not support arrowheads.");
  };
  const text = (value: unknown, at: string) => {
    if (!record(value, ["paragraph", "run"], at)) return;
    if (value.paragraph !== undefined && record(value.paragraph, ["align", "direction", "indent", "hanging", "lineSpacing", "spaceBefore", "spaceAfter", "bullet", "tabStops"], `${at}.paragraph`)) {
      const p = value.paragraph;
      if (p.direction !== undefined && p.direction !== "ltr" && p.direction !== "rtl") fail(`${at}.paragraph.direction`, "Expected ltr or rtl.");
      if (p.tabStops !== undefined) for (const error of tabStopErrors(p.tabStops)) fail(`${at}.paragraph.tabStops${error.index === undefined ? "" : `.${error.index}`}`, error.message);
      if (p.align !== undefined && !["left", "right", "center", "justify"].includes(p.align as string)) fail(`${at}.paragraph.align`, "Unsupported paragraph alignment.");
      for (const key of ["indent", "hanging", "spaceBefore", "spaceAfter"]) if (p[key] !== undefined && !nonnegative(p[key])) fail(`${at}.paragraph.${key}`, "Expected a nonnegative finite paragraph metric.");
      if (p.lineSpacing !== undefined && !positive(p.lineSpacing)) fail(`${at}.paragraph.lineSpacing`, "Expected positive line spacing.");
      if (p.bullet !== undefined && record(p.bullet, ["type", "character", "style", "startAt"], `${at}.paragraph.bullet`)) {
        const b = p.bullet;
        if (!["none", "bullet", "number"].includes(b.type as string)) fail(`${at}.paragraph.bullet.type`, "Unsupported bullet type.");
        if (b.character !== undefined && (b.type !== "bullet" || typeof b.character !== "string" || !b.character.length)) fail(`${at}.paragraph.bullet.character`, "Bullet characters require a nonempty string.");
        if (b.style !== undefined && (b.type !== "number" || !["arabicPeriod", "arabicParen", "alphaLowerPeriod", "alphaUpperPeriod"].includes(b.style as string))) fail(`${at}.paragraph.bullet.style`, "Unsupported number style.");
        if (b.startAt !== undefined && (b.type !== "number" || !Number.isSafeInteger(b.startAt) || (b.startAt as number) < 1)) fail(`${at}.paragraph.bullet.startAt`, "Expected a positive safe integer.");
      }
    }
    if (value.run !== undefined && record(value.run, ["fontFamily", "eastAsianFontFamily", "complexScriptFontFamily", "scriptFontFamilies", "rightToLeft", "fontSize", "bold", "italic", "underline", "strike", "color", "colorOpacity", "language"], `${at}.run`)) {
      const r = value.run;
      for (const error of scriptFontFamilyErrors(r.scriptFontFamilies)) fail(`${at}.run.scriptFontFamilies${error.key === undefined ? "" : `.${error.key}`}`, error.message);
      for (const field of ["eastAsianFontFamily", "complexScriptFontFamily"]) if (r[field] !== undefined && r[field] !== null && !isValidFontFamily(r[field])) fail(`${at}.run.${field}`, "Expected a nonempty script font name or theme font role.");
      for (const key of ["bold", "italic", "underline", "strike", "rightToLeft"]) if (r[key] !== undefined && typeof r[key] !== "boolean") fail(`${at}.run.${key}`, "Expected boolean text flag.");
      if (r.fontSize !== undefined && !positive(r.fontSize)) fail(`${at}.run.fontSize`, "Expected positive font size.");
      if (r.color !== undefined && !isValidColor(r.color as never)) fail(`${at}.run.color`, "Invalid text color.");
      if (r.colorOpacity !== undefined && (!nonnegative(r.colorOpacity) || (r.colorOpacity as number) > 1)) fail(`${at}.run.colorOpacity`, "Invalid text alpha.");
      if (r.language !== undefined && (typeof r.language !== "string" || !r.language.length)) fail(`${at}.run.language`, "Expected nonempty language.");
      if (r.fontFamily !== undefined && !(typeof r.fontFamily === "string" && r.fontFamily.length)) {
        if (!record(r.fontFamily, ["theme"], `${at}.run.fontFamily`) || !["body", "heading"].includes((r.fontFamily as Record<string, unknown>)?.theme as string)) fail(`${at}.run.fontFamily`, "Expected a font name or theme font role.");
      }
    }
  };
  const region = (value: unknown, at: string) => {
    if (!record(value, ["fill", "stroke", "borders", "margin", "verticalAlign", "textDirection", "text"], at)) return;
    if (!isValidPaint(value.fill as never)) fail(`${at}.fill`, "Invalid cell fill.");
    if (value.stroke !== undefined) stroke(value.stroke, `${at}.stroke`);
    if (value.borders !== undefined && record(value.borders, ["left", "right", "top", "bottom"], `${at}.borders`)) for (const [side, edge] of Object.entries(value.borders)) if (edge !== undefined) stroke(edge, `${at}.borders.${side}`);
    if (value.margin !== undefined) {
      if (typeof value.margin === "number") { if (!nonnegative(value.margin)) fail(`${at}.margin`, "Expected nonnegative margin."); }
      else if (record(value.margin, ["left", "right", "top", "bottom"], `${at}.margin`)) for (const [side, metric] of Object.entries(value.margin)) if (metric !== undefined && !nonnegative(metric)) fail(`${at}.margin.${side}`, "Expected nonnegative margin.");
    }
    if (value.verticalAlign !== undefined && !["top", "middle", "bottom"].includes(value.verticalAlign as string)) fail(`${at}.verticalAlign`, "Unsupported vertical alignment.");
    if (value.textDirection !== undefined && !["horizontal", "rotate90", "rotate270", "eastAsianVertical", "mongolianVertical", "stackedVertical", "stackedVerticalRtl"].includes(value.textDirection as string)) fail(`${at}.textDirection`, "Unsupported text direction.");
    if (value.text !== undefined) text(value.text, `${at}.text`);
  };
  const at = `${path}.style`, value = table.style;
  if (record(value, ["wholeTable", "bandedRows", "bandedColumns", "firstColumn", "lastColumn", "firstRow", "lastRow", "topLeft", "topRight", "bottomLeft", "bottomRight", "borders"], at)) {
    for (const [key, rule] of Object.entries(value)) {
      if (rule === undefined) continue;
      if (key === "borders") { if (record(rule, ["outer", "innerHorizontal", "innerVertical"], `${at}.${key}`)) for (const [side, edge] of Object.entries(rule)) if (edge !== undefined) stroke(edge, `${at}.${key}.${side}`); }
      else if (key === "bandedRows" || key === "bandedColumns") {
        if (record(rule, ["size", "first", "second"], `${at}.${key}`)) {
          if (rule.size !== undefined && (!Number.isSafeInteger(rule.size) || (rule.size as number) < 1 || (rule.size as number) > 100000)) fail(`${at}.${key}.size`, "Band size must be 1..100000.");
          for (const band of ["first", "second"]) if (rule[band] !== undefined) region(rule[band], `${at}.${key}.${band}`);
        }
      } else if (["wholeTable", "firstColumn", "lastColumn", "firstRow", "lastRow", "topLeft", "topRight", "bottomLeft", "bottomRight"].includes(key)) region(rule, `${at}.${key}`);
    }
  }
  try { tableStyleAnchors(table); } catch (error) {
    if (!(error instanceof TableStyleGridError)) throw error;
    fail(`${path}.${error.path}`, error.message);
  }
  for (const [rowIndex, row] of table.rows.entries()) for (const [cellIndex, cell] of row.cells.entries()) if (cell.style !== undefined) {
    const cellPath = `${path}.rows.${rowIndex}.cells.${cellIndex}.style`;
    region(cell.style, cellPath);
    if (cell.style && Object.hasOwn(cell.style, "text")) fail(`${cellPath}.text`, "Local cell text belongs in content or textStylePreset.");
  }
  return errors;
}
