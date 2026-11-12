/** Independent DOM/OPC oracle for fresh default notes-page geometry and owners. */
export function authoredNotesPageAudit(parts) {
  const P = "http://schemas.openxmlformats.org/presentationml/2006/main", A = "http://schemas.openxmlformats.org/drawingml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const docs = Object.fromEntries(Object.entries(parts).filter(([n]) => n.endsWith(".xml") || n.endsWith(".rels")).map(([n, xml]) => ["/" + n, new DOMParser().parseFromString(xml, "application/xml")]));
  const check = (condition, message) => { if (!condition) throw Error(message); };
  const one = nodes => { check(nodes.length === 1, "Missing/ambiguous notes component"); return nodes[0]; };
  const nodes = (node, ns, name) => [...node.getElementsByTagNameNS(ns, name)];
  for (const doc of Object.values(docs)) check(!nodes(doc, "*", "parsererror").length, "Invalid XML");
  const edge = (owner, kind, id) => {
    const slash = owner.lastIndexOf("/"), rels = docs[owner.slice(0, slash + 1) + "_rels/" + owner.slice(slash + 1) + ".rels"];
    const r = one([...rels.documentElement.children].filter(r => r.getAttribute("Type").endsWith("/" + kind) && (id === undefined || r.getAttribute("Id") === id)));
    check(r.getAttribute("TargetMode") !== "External", "External notes owner");
    const target = new URL(r.getAttribute("Target"), "https://fixture.invalid" + owner).pathname;
    check(Boolean(docs[target]), "Missing relationship target"); return target;
  };
  const role = (doc, type) => nodes(doc, P, "sp").filter(s => nodes(s, P, "ph").some(p => p.getAttribute("type") === type));
  const presentation = docs["/ppt/presentation.xml"], slideSize = one(nodes(presentation, P, "sldSz")), notesSize = one(nodes(presentation, P, "notesSz"));
  const nw = Number(notesSize.getAttribute("cx")), nh = Number(notesSize.getAttribute("cy"));
  check(nw === 6858000 && nh === 9144000, "Changed default notes size");
  const master = edge("/ppt/presentation.xml", "notesMaster"), masterDoc = docs[master];
  const frame = shape => {
    const rect = one(nodes(shape, A, "prstGeom")); check(rect.getAttribute("prst") === "rect", "Missing native rect");
    const off = one(nodes(shape, A, "off")), ext = one(nodes(shape, A, "ext"));
    const f = { x: Number(off.getAttribute("x")), y: Number(off.getAttribute("y")), width: Number(ext.getAttribute("cx")), height: Number(ext.getAttribute("cy")) };
    check(Object.values(f).every(Number.isSafeInteger) && f.x >= 0 && f.y >= 0 && f.width > 0 && f.height > 0 && f.x + f.width <= nw && f.y + f.height <= nh, "Invalid/off-page notes geometry"); return f;
  };
  const image = frame(one(role(masterDoc, "sldImg"))), body = frame(one(role(masterDoc, "body")));
  check(image.y + image.height < body.y, "Overlapping image and body");
  const sw = Number(slideSize.getAttribute("cx")), sh = Number(slideSize.getAttribute("cy"));
  check(Math.abs(image.width * sh - image.height * sw) <= Math.max(sw, sh), "Distorted slide image");
  check(Math.abs(image.x * 2 + image.width - nw) <= 1, "Image is not horizontally centered");
  check(Math.abs(image.y * 2 + image.height - 360 * 12700) <= 1, "Image is not vertically centered in its frame");
  check(image.width <= 468 * 12700 && image.height <= 288 * 12700 && (Math.abs(image.width - 468 * 12700) <= 1 || Math.abs(image.height - 288 * 12700) <= 1), "Image does not fill its fit frame");
  const rows = [];
  for (const id of nodes(presentation, P, "sldId")) {
    const slide = edge("/ppt/presentation.xml", "slide", id.getAttributeNS(R, "id")), note = edge(slide, "notesSlide"), local = docs[note];
    check(edge(note, "slide") === slide && edge(note, "notesMaster") === master, "Wrong notes owner/master");
    const ids = nodes(local, P, "cNvPr").map(n => n.getAttribute("id")); check(new Set(ids).size === ids.length, "Duplicate local drawing IDs");
    const localImage = one(role(local, "sldImg")); check(nodes(localImage, A, "xfrm").length === 0 && nodes(localImage, P, "txBody").length === 0, "Local image must inherit frame without fake text");
    const bodies = role(local, "body"); check(bodies.length <= 1, "Ambiguous local body");
    const label = nodes(docs[slide], A, "t").map(t => t.textContent).join("");
    const text = bodies.length ? nodes(bodies[0], A, "t").map(t => t.textContent).join("") : null;
    rows.push({ slide, note, label, text, bodyPresent: bodies.length === 1 });
  }
  check(new Set(rows.map(r => r.note)).size === rows.length, "Aliased copied notes");
  return { master, image, body, rows };
}
