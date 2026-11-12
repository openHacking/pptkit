/** Repository-authored source-save/native fixture; fourth chart is an untouched control. */
export function createChartSaveFixture(createPresentation) {
  const deck = createPresentation({ size: { width: 960, height: 540 }, metadata: { title: "Existing chart value saves" } });
  for (const [index, family] of ["bar", "line", "pie", "bar"].entries()) {
    const slide = deck.addSlide({ id: `chart-slide-${index}` });
    const title = index === 3 ? "Unchanged control" : `${family[0].toUpperCase()}${family.slice(1)} values`;
    slide.addElement({ type: "text", id: `title-${index}`, box: { x: 40, y: 25, width: 860, height: 60 }, content: [{ runs: [{ text: title, style: { fontFamily: "Arial", fontSize: 28, color: "112233", bold: true } }] }] });
    slide.addElement({ type: "chart", id: `chart-${index}`, name: title, chartType: family,
      box: { x: 80, y: 105, width: 800, height: 360 }, categories: ["First", "Second", "Third"],
      legend: { visible: true, position: "bottom" },
      series: [{ name: "Primary", values: [2, 5, 8], color: "336699" },
        ...(family === "pie" ? [] : [{ name: "Control series", values: [3, 4, 6], color: "99BBCC" }])],
    });
    slide.setNotes(`Unchanged notes for ${title}.`);
  }
  return deck;
}
