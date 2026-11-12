/** Independent browser DOM/OPC oracle; no Parser or Core code is called here. */
export function nativeNotesSourceAudit({ parts, document, properties }) {
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main", P = "http://schemas.openxmlformats.org/presentationml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const docs = Object.fromEntries(Object.entries(parts).map(([name, text]) => ["/" + name, new DOMParser().parseFromString(text, "application/xml")]));
  for (const doc of Object.values(docs)) if (doc.getElementsByTagName("parsererror").length) throw Error("Invalid actual source XML");
  const children = (node, ns, name) => [...node.children].filter(n => n.namespaceURI === ns && n.localName === name);
  const one = nodes => { if (nodes.length !== 1) throw Error("Missing/ambiguous actual source node"); return nodes[0]; };
  const relation = (owner, kind, id, optional = false) => {
    const slash = owner.lastIndexOf("/"), rel = docs[owner.slice(0, slash + 1) + "_rels/" + owner.slice(slash + 1) + ".rels"];
    const nodes = [...rel.documentElement.children].filter(n => n.getAttribute("Type").endsWith("/" + kind) && (id === undefined || n.getAttribute("Id") === id));
    if (optional && nodes.length === 0) return undefined;
    const n = one(nodes); if (n.getAttribute("TargetMode") === "External") throw Error("External notes edge");
    return new URL(n.getAttribute("Target"), "https://fixture.invalid" + owner).pathname;
  };
  const body = part => {
    const tree = one([...docs[part].getElementsByTagNameNS(P, "spTree")]);
    return one(children(tree, P, "sp").filter(n => [...n.getElementsByTagNameNS(P, "ph")].some(p => p.getAttribute("type") === "body")));
  };
  const equal = (a, b, label) => { if (JSON.stringify(a) !== JSON.stringify(b)) throw Error(label); };
  const select = s => { const doc = docs[s.part]; if (!doc) throw Error("Invented source part"); return doc.evaluate(s.path, doc, p => doc.documentElement.lookupNamespaceURI(p), XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; };
  const rows = [], slides = [...docs["/ppt/presentation.xml"].getElementsByTagNameNS(P, "sldId")];
  if (slides.length !== document.slides.length) throw Error("Wrong notes projection order/count");
  const scriptSlots = { Latn: "latin", Grek: "latin", Hans: "ea", Hant: "ea", Jpan: "ea", Hang: "ea", Arab: "cs", Hebr: "cs" };
  for (const [index, id] of slides.entries()) {
    const slide = relation("/ppt/presentation.xml", "slide", id.getAttributeNS(R, "id")), part = relation(slide, "notesSlide"), master = relation(part, "notesMaster");
    if (relation(part, "slide") !== slide) throw Error("Wrong notes backlink");
    const theme = relation(part, "themeOverride", undefined, true) ?? relation(master, "theme"), root = docs[theme];
    const scheme = one([...root.getElementsByTagNameNS(A, "fontScheme")]);
    const local = body(part), inherited = body(master);
    const localPh = one([...local.getElementsByTagNameNS(P, "ph")]), masterPh = one([...inherited.getElementsByTagNameNS(P, "ph")]);
    if (localPh.getAttribute("idx") !== "7" || masterPh.getAttribute("idx") !== "99") throw Error("Type-vs-index control lost");
    const masterSize = one([...inherited.getElementsByTagNameNS(A, "defRPr")].filter(n => n.hasAttribute("sz"))).getAttribute("sz");
    const notesStyle = one([...docs[master].getElementsByTagNameNS(P, "notesStyle")]);
    const defaults = one(children(one(children(notesStyle, A, "lvl1pPr")), A, "defRPr"));
    const inheritedSlots = Object.fromEntries(["latin", "ea", "cs"].map(slot => [slot, one(children(defaults, A, slot)).getAttribute("typeface")]));
    const tx = one(children(local, P, "txBody")), paragraph = one(children(tx, A, "p")), sourceRuns = children(paragraph, A, "r");
    const projected = document.slides[index].notes;
    if (projected.length !== 1 || projected[0].runs.length !== sourceRuns.length) throw Error("Notes owner/run topology differs");
    const declarations = [];
    for (const [ri, run] of sourceRuns.entries()) {
      const rp = one(children(run, A, "rPr")), text = one(children(run, A, "t")).textContent, actual = projected[0].runs[ri];
      if (actual.text !== text || actual.style.language !== rp.getAttribute("lang") || actual.style.bold !== (rp.getAttribute("b") === "1") || actual.style.italic !== (rp.getAttribute("i") === "1") || actual.style.fontSize !== Number(masterSize) / 100) throw Error("Actual source run text/language/flags/size differs");
      const slots = Object.fromEntries(["latin", "ea", "cs"].map(name => { const nodes = children(rp, A, name); return [name, nodes.length ? one(nodes).getAttribute("typeface") : inheritedSlots[name]]; }));
      const collection = kind => one(children(scheme, A, kind + "Font"));
      const family = slot => {
        const value = slots[slot], match = /^\+(mn|mj)-(lt|ea|cs)$/.exec(value);
        if (!match) return value.startsWith("+") || !value ? undefined : value;
        return one(children(collection(match[1] === "mj" ? "major" : "minor"), A, slot)).getAttribute("typeface") || undefined;
      };
      if (actual.style.fontFamily !== family("latin") || actual.style.eastAsianFontFamily !== family("ea") || actual.style.complexScriptFontFamily !== family("cs")) throw Error("Actual local/generic source slot differs");
      const map = {};
      for (const [script, slot] of Object.entries(scriptSlots)) {
        const ref = /^\+(mn|mj)-(lt|ea|cs)$/.exec(slots[slot]); if (!ref) continue;
        const rule = children(collection(ref[1] === "mj" ? "major" : "minor"), A, "font").find(n => n.getAttribute("script") === script);
        if (rule) map[script] = rule.getAttribute("typeface");
      }
      equal(Object.entries(actual.style.scriptFontFamilies ?? {}).sort(), Object.entries(map).sort(), "Actual script dictionaries differ");
      declarations.push({ text, language: rp.getAttribute("lang"), slots, bold: rp.getAttribute("b"), italic: rp.getAttribute("i") });
    }
    rows.push({ slidePart: slide.slice(1), notesPart: part.slice(1), masterPart: master.slice(1), effectiveThemePart: theme.slice(1), runs: declarations });
  }
  let checked = 0;
  for (const p of properties.filter(p => p.target.part.includes("/notesSlides/"))) {
    if (!select(p.target) || !select(p.source)) throw Error("Missing actual notes source/target XPath");
    for (const d of p.resolvedFrom ?? []) if (!select(d)) throw Error("Invented notes source dependency");
    if (p.status !== "mapped") continue;
    const row = rows.find(r => "/" + r.notesPart === p.target.part);
    if (p.property.startsWith("run.scriptFontFamilies.")) {
      if (p.source.part !== "/" + row.effectiveThemePart || select(p.source).value !== p.value) throw Error("Borrowed notes font provenance");
      checked++;
    }
    if (p.property === "run.fontSize") {
      if (p.source.part !== "/" + row.masterPart || p.value !== 40 || select(p.source).value !== "4000" || p.level !== "notes-master-placeholder") throw Error("Wrong body type-match provenance");
      checked++;
    }
  }
  if (checked <= rows.length) throw Error("Missing font provenance coverage");
  return { rows, checkedMappedProperties: checked };
}
