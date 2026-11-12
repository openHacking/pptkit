export function createNativeConnectorFixture(createPresentation) {
  const p = createPresentation({ size: { width: 400, height: 250 }, metadata: { title: "Native connector profile" } });
  const stroke = { paint: { type: "solid", color: "225588" }, width: 2, endArrow: "triangle" };
  const shape = (id, x, y, kind = "rect") => ({ type: "shape", shape: kind, id, name: id, box: { x, y, width: 60, height: 40 }, style: { fill: { type: "solid", color: "DDEEFF" } } });
  const connection = (id, start, end) => ({ type: "connector", id, name: id, start, end, style: stroke });
  const anchor = (elementId, site) => ({ elementId, anchor: site });
  const h = p.addSlide({ id: "horizontal" });
  h.addElement(connection("horizontal-forward", anchor("rect-a", "right"), anchor("ellipse-b", "left")));
  h.addElement(connection("horizontal-reverse", anchor("ellipse-b", "left"), anchor("rect-a", "right")));
  h.addElement(shape("rect-a", 30, 40)); h.addElement(shape("ellipse-b", 260, 40, "ellipse"));
  const v = p.addSlide({ id: "vertical" });
  v.addElement(connection("vertical-forward", anchor("round-a", "bottom"), anchor("round-b", "top")));
  v.addElement(connection("vertical-reverse", anchor("round-b", "top"), anchor("round-a", "bottom")));
  v.addElement(shape("round-a", 80, 20, "roundRect")); v.addElement(shape("round-b", 80, 170, "roundRect"));
  p.addSlide({ id: "nested" }).addElement({ type: "group", id: "group", name: "group", box: { x: 50, y: 30, width: 300, height: 160 }, coordinateSize: { width: 300, height: 160 }, transform: { rotation: 15 }, children: [
    connection("group-edge", anchor("group-a", "right"), anchor("group-b", "left")), shape("group-a", 10, 30), shape("group-b", 200, 80, "ellipse"),
  ] });
  const f = p.addSlide({ id: "free" });
  for (const [index, end] of [{ x: 300, y: 200 }, { x: 40, y: 200 }, { x: 300, y: 40 }, { x: 40, y: 40 }, { x: 180, y: 100 }].entries()) f.addElement(connection(`free-${index}`, { x: 180, y: 100 }, end));
  return p;
}
