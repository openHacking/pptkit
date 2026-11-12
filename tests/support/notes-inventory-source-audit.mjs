/** DOM source oracle: no Parser/Core imports, raw element-preorder identity. */
export function notesInventorySourceAudit({ parts, report }) {
  const docs = {}, indices = {};
  for (const [name, xml] of Object.entries(parts)) {
    if (!name.endsWith(".xml") && name !== "deck/main.data" && !name.endsWith(".rels")) continue;
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw Error("Actual XML rejected by DOM");
    docs["/" + name] = doc;
    const nodes = [], pending = [doc.documentElement];
    while (pending.length) { const n = pending.pop(); nodes.push(n); for (let i = n.children.length - 1; i >= 0; i--) pending.push(n.children[i]); }
    indices["/" + name] = nodes;
  }
  let checked = 0;
  const select = source => {
    const node = indices[source.part]?.[source.elementIndex - 1];
    if (!node) throw Error("Invented raw source location"); checked++; return node;
  };
  const same = (condition, message) => { if (!condition) throw Error(message); };
  const visit = value => {
    if (!value || typeof value !== "object") return;
    if (Array.isArray(value)) { value.forEach(visit); return; }
    if (value.source?.elementIndex) {
      const node = select(value.source);
      if (value.name?.namespaceUri !== undefined) same(node.namespaceURI === value.name.namespaceUri && node.localName === value.name.local && node.nodeName === value.name.name, "Wrong expanded source name");
      if (value.attributes) {
        same(value.attributes.length === node.attributes.length, "Missing source attributes");
        for (const a of value.attributes) same(node.getAttributeNS(a.namespaceUri || null, a.local) === a.value, "Wrong lexical attribute");
      }
      if (typeof value.value === "string") {
        const actual = [...node.childNodes].filter(n => n.nodeType === 3 || n.nodeType === 4).map(n => n.nodeValue).join("");
        same(actual === value.value, "Wrong direct cached text");
      }
      if (value.paragraphElementIndex !== undefined) {
        const paragraph = select({ part: value.source.part, elementIndex: value.paragraphElementIndex });
        let ancestor = node.parentElement;
        while (ancestor && !(ancestor.localName === "p" && ancestor.namespaceURI === node.namespaceURI)) ancestor = ancestor.parentElement;
        same(paragraph.localName === "p" && paragraph.namespaceURI === node.namespaceURI && ancestor === paragraph, "Wrong enclosing paragraph source");
      }
      if (value.nonVisual) {
        const identity = select(value.nonVisual.source);
        let ancestor = identity.parentElement;
        while (ancestor && ancestor !== node && !["sp", "grpSp", "pic", "cxnSp", "graphicFrame", "spTree"].includes(ancestor.localName)) ancestor = ancestor.parentElement;
        same(ancestor === node, "Wrong drawing identity owner");
        same(identity.getAttribute("id") === (value.rawId ?? null), "Wrong raw drawing id");
        same(identity.getAttribute("name") === (value.drawingName ?? null), "Wrong raw drawing name");
      }
    }
    if (typeof value.part === "string" && Number.isInteger(value.elementIndex)) select(value);
    Object.values(value).forEach(visit);
  };
  visit(report);
  const P = report.dialect === "strict" ? "http://purl.oclc.org/ooxml/presentationml/main" : "http://schemas.openxmlformats.org/presentationml/2006/main";
  for (const part of report.parts) {
    for (const g of part.geometries) for (const component of g.components) same(select(component.source).parentElement === select(g.source), "Wrong geometry component source");
    for (const f of part.fields) for (const text of f.cachedTexts) { const field = select(f.source); let ancestor = select(text.source).parentElement; while (ancestor && !(ancestor.localName === "fld" && ancestor.namespaceURI === field.namespaceURI)) ancestor = ancestor.parentElement; same(ancestor === field, "Wrong field cache source"); }
  }
  for (const page of report.pages) for (const match of page.placeholderMatches) {
    const local = select(match.source).getAttribute("type");
    for (const candidate of match.candidates) {
      same(candidate.part === page.masterPart, "Borrowed another notes master");
      const drawing = report.parts.find(p => p.part === candidate.part).drawings.find(d => d.source.elementIndex === candidate.elementIndex);
      same(Boolean(drawing?.placeholder) && select(drawing.placeholder.source).getAttribute("type") === local, "Matched notes placeholder by something other than explicit type");
      same(select(candidate).namespaceURI === P, "Wrong candidate dialect");
    }
  }
  return { checkedSourceReferences: checked, parts: report.parts.length, fields: report.parts.reduce((n, p) => n + p.fields.length, 0) };
}
