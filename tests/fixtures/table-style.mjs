export function tableStyleDeck(createPresentation, { theme = false } = {}) {
  const document = createPresentation({ size: { width: 400, height: 300 }, theme: { colors: { accent1: "2582B5", accent2: "B34F25" }, fonts: { body: "Arial", heading: "Arial" } } });
  const fill = color => ({ type: "solid", color }), stroke = (color, width = 1) => ({ paint: fill(color), width });
  const style = {
    wholeTable: { fill: fill("FAFAFA"), margin: 2, text: { run: { fontFamily: theme ? { theme: "body" } : "Arial", fontSize: 10, color: "111111" } } },
    borders: { outer: stroke("222222", 2), innerHorizontal: stroke("777777"), innerVertical: stroke("999999") },
    bandedColumns: { first: { text: { run: { italic: true } } } },
    bandedRows: { first: { fill: fill(theme ? { theme: "accent1" } : "DDEEFF") }, second: { fill: fill("EEEEEE") } },
    firstRow: { fill: fill(theme ? { theme: "accent2" } : "335577"), text: { run: { bold: true, color: "FFFFFF" } } },
    lastRow: { fill: fill("CCDDBB") },
    firstColumn: { text: { paragraph: { align: "center" } } },
    topLeft: { text: { run: { italic: false } } },
  };
  const slide = document.addSlide({ id: "styled" });
  slide.addElement({ id: "styled-table", name: "Region table", type: "table", box: { x: 30, y: 30, width: 320, height: 150 }, columns: [80,80,80,80], style,
    rows: Array.from({ length: 5 }, (_, r) => ({ height: 30, cells: Array.from({ length: 4 }, (_, c) => ({ content: `R${r}C${c}` })) })),
  });
  const source = slide.elements[0], merged = structuredClone(source);
  merged.rows[1].cells = [{ content: "Merged body", colSpan: 2, rowSpan: 2 }, { content: "Upper", style: { fill: fill("FFCC66"), borders: { left: stroke("AA0000", 3) } } }];
  merged.rows[2].cells = [{ content: "Lower" }];
  merged.rows[3].cells[0].content = [{ runs: [{ text: "Local", style: { bold: false, fontSize: 12 } }] }];
  slide.replaceElement(source.id, merged);
  return document;
}
