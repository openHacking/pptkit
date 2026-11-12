const stroke = (color, width = 1, dash = "solid", opacity = 1) => ({ paint: { type: "solid", color, opacity }, width, dash });
const cell = (text, borders, extra = {}) => ({ content: [{ runs: [{ text, style: { fontFamily: "Arial", fontSize: 9 } }] }], style: { margin: 0, stroke: stroke("666666"), borders }, ...extra });
export function tableBorderDeck(createPresentation, { theme = true } = {}) {
  const document = createPresentation({ size: { width: 420, height: 300 }, theme: { colors: { accent2: "3579BD" } } });
  const add = (id, columns, rows, extra = {}) => document.addSlide({ id }).addElement({ id: `${id}-table`, type: "table", box: { x: 40, y: 60, width: columns.reduce((a, b) => a + b, 0), height: rows.reduce((a, b) => a + b.height, 0) }, columns, rows, ...extra });
  add("direct", [100, 100], [
    { height: 50, cells: [cell("Direct A", { left: stroke("AA00AA", 3), top: stroke("FF0000", 4), right: stroke("00AA00", 2), bottom: stroke("0000FF", 2) }), cell("Direct B", { left: stroke("FF8800", 1), top: stroke("00AAAA", 2), right: stroke("333333", 2) })] },
    { height: 50, cells: [cell("Direct C", { top: stroke("777777", 5), bottom: stroke("FF8800", 3) }), cell("Direct D", { left: { paint: { type: "none" } }, bottom: stroke("0000FF", 2) })] },
  ]);
  add("merged", [60, 60, 100], [
    { height: 40, cells: [cell("Merged anchor", { left: stroke("AA00AA", 2), top: stroke("FF0000", 2), right: stroke("FF0000", 2), bottom: stroke("AAAA00", 2) }, { rowSpan: 2, colSpan: 2 }), cell("Upper neighbor", { left: stroke("00AA00", 4) })] },
    { height: 40, cells: [cell("Lower neighbor", { left: stroke("0000FF", 5) })] },
    { height: 40, cells: [cell("Bottom A", {}), cell("Bottom B", {}), cell("Bottom C", {})] },
  ]);
  add("wide", [60, 60, 60], [
    { height: 30, cells: [cell("Wide merged control", { left: stroke("AA00AA", 2), top: stroke("FF0000", 3), right: stroke("0000FF", 4), bottom: stroke("00AA00", 2) }, { rowSpan: 3, colSpan: 3 })] },
    { height: 30, cells: [] }, { height: 30, cells: [] },
  ]);
  if (theme) {
    const borders = { left: stroke({ theme: "accent2" }, 3, "dot", .5), top: stroke("FF0000", 2, "dash"), right: stroke("AA00AA", 2, "dashDot") };
    const alpha = cell("Theme alpha", borders); alpha.style.fill = { type: "solid", color: "112233", opacity: .25 };
    add("theme", [100, 100], [{ height: 60, cells: [alpha, cell("Theme neighbor", { bottom: stroke("00AA00", 3, "dot", .6) })] }], { opacity: .8, transform: { rotation: 10, flipH: true } });
  }
  return document;
}
