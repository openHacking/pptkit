import assert from "node:assert/strict";
import test from "node:test";
import { chromium } from "@playwright/test";
import { readFile } from "node:fs/promises";
import { inspectPptx, importPptx, inspectXmlPart, compareOpcPackages } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { nativeDiagramSpecimen } from "./research/native-diagram-fixture.mjs";
import { buildNativeDiagramCandidate } from "./research/native-diagram-package.mjs";
import { projectDiagramContent } from "./research/diagram-content-model.mjs";
import { diagramResearchProfiles } from "./research/diagram-layout-profiles.mjs";
import { pdfTextEvidence } from "./research/pdf-text-evidence.mjs";
import { createDiagramIdentityState } from "./research/diagram-identity-state.mjs";
import { bindDiagramSourceContent } from "./research/diagram-source-bindings.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

const node = (id, paragraphs = [id], children = [], role = "content") => ({ id, paragraphs, children, role });
const candidate = family => family === "cycle" ? { family, closure: "layout-generated", roots: ["a", "b", "c"], nodes: [node("a"), node("b"), node("c")] }
  : { family, roots: ["a", "c"], nodes: [node("a", ["Lead & < 😀 \" ' \t\r", "Next\nline", ""], ["b"]), node("b", ["Detail"], [], family === "hierarchy" ? "assistant" : "content"), node("c")] };
const components = family => {
  const D = "http://schemas.openxmlformats.org/drawingml/2006/diagram", identifier = diagramResearchProfiles.find(p => p.family === family).layoutIdentifier;
  return Object.fromEntries([["layout", "layoutDef", identifier], ["colors", "colorsDef", "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"], ["style", "styleDef", "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"]].map(([key, root, id]) => [key, new TextEncoder().encode(`<d:${root} xmlns:d="${D}" uniqueId="${id}"/>`)]));
};

test("Parser source bindings survive XML reindex/order/lexical changes and feed stateful fresh generation", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage();
  await page.route("**/*", route => route.abort());
  for (const family of ["process", "hierarchy", "cycle"]) for (const cache of ["absent", "empty"]) for (const modelIds of ["int32", "guid"]) {
    const first = await buildNativeDiagramCandidate(candidate(family), components(family), { cache, modelIds, presentation: "layout" });
    const bound = bindDiagramSourceContent(first.inspection.diagramInventory.dataModels[0], first.specimen.layoutIdentifier, { scope: "owned-source/diagram", namespace: "12345678-1234-4ABC-8DEF" });
    const parts = readGeneratedZip(first.bytes), before = structuredClone(bound.state);
    const mutatedXml = await page.evaluate(xml => {
      const doc = new DOMParser().parseFromString(xml, "application/xml"), D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
      const list = doc.getElementsByTagNameNS(D, "ptLst")[0], points = [...list.children], root = points.find(p => p.getAttribute("type") === "doc").getAttribute("modelId");
      for (const p of points.slice().reverse()) list.appendChild(p);
      for (const p of points.filter(p => !p.hasAttribute("type") || p.getAttribute("type") === "asst")) {
        const raw = p.getAttribute("modelId"); p.setAttribute("modelId", raw.startsWith("{") ? ` ${raw} ` : ` +000${raw} `);
      }
      const roots = [...doc.getElementsByTagNameNS(D, "cxn")].filter(e => e.getAttribute("type") === "parOf" && e.getAttribute("srcId") === root);
      for (const e of roots) e.setAttribute("srcOrd", String(roots.length - 1 - Number(e.getAttribute("srcOrd"))));
      // XMLSerializer emits character-reference CRs as literal CRs; re-escape
      // those so this control changes topology/lexical IDs without changing text
      // through the next XML parser's required newline normalization.
      return new XMLSerializer().serializeToString(doc).replaceAll("\r", "&#13;");
    }, parts.get("ppt/diagrams/data1.xml").toString());
    parts.set("ppt/diagrams/data1.xml", Buffer.from(mutatedXml));
    const changed = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
    const changedModel = (await inspectPptx(changed)).diagramInventory.dataModels[0]; assert.deepEqual(changedModel.findings, []);
    const rebound = bindDiagramSourceContent(changedModel, first.specimen.layoutIdentifier, { scope: "owned-source/diagram", state: bound.state });
    assert.deepEqual(bound.state, before);
    assert.deepEqual(new Map(rebound.bindings.map(b => [b.canonicalModelId, b.id])), new Map(bound.bindings.map(b => [b.canonicalModelId, b.id])));
    assert.deepEqual(rebound.candidate.roots, bound.candidate.roots.slice().reverse());
    assert.ok(rebound.bindings.some(b => b.elementIndex !== bound.bindings.find(x => x.id === b.id).elementIndex));
    const oldById = new Map(bound.candidate.nodes.map(n => [n.id, n]));
    for (const n of rebound.candidate.nodes) assert.deepEqual(n, oldById.get(n.id));
    const outputState = createDiagramIdentityState(family, "87654321-4321-4CBA-9FED");
    const generated = await buildNativeDiagramCandidate(bound.candidate, components(family), { modelIds: "guid", identityState: outputState, presentation: "layout", cache });
    const regenerated = await buildNativeDiagramCandidate(rebound.candidate, components(family), { modelIds: "guid", identityState: generated.specimen.identityState, presentation: "layout", cache });
    assert.deepEqual(new Map(generated.specimen.identities.map(e => [e.id, e.modelId])), new Map(regenerated.specimen.identities.map(e => [e.id, e.modelId])));
    assert.notEqual(generated.specimen.dataXml, regenerated.specimen.dataXml);
    assert.ok(regenerated.specimen.identities.every(e => e.id !== e.modelId), "managed identities and emitted source identities have distinct domains");
  }
});

test("explicit identity snapshots retain content and typed graph identities through structural edits", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage();
  await page.route("**/*", route => route.abort());
  for (const family of ["process", "hierarchy", "cycle"]) for (const cache of ["absent", "empty"]) {
    let input = candidate(family), state = createDiagramIdentityState(family, "12345678-1234-4ABC-8DEF"), previous;
    const identities = new Map(), reserved = new Map(), graphIdentities = new Map();
    const retainGraphId = (key, value) => { if (graphIdentities.has(key)) assert.equal(value, graphIdentities.get(key), `stable graph role ${key}`); else graphIdentities.set(key, value); };
    for (let round = 0; round < 8; round++) {
      const before = structuredClone(state), built = await buildNativeDiagramCandidate(input, components(family), { presentation: "layout", cache, modelIds: "guid", identityState: state });
      assert.deepEqual(state, before, "caller state is never committed by package generation");
      for (const entry of built.specimen.identities) {
        if (identities.has(entry.id)) assert.equal(entry.modelId, identities.get(entry.id), `retained content ${entry.id} at round ${round}`);
        else identities.set(entry.id, entry.modelId);
      }
      state = JSON.parse(JSON.stringify(built.specimen.identityState));
      for (const entry of state.entries) { if (reserved.has(entry.key)) assert.equal(entry.modelId, reserved.get(entry.key)); else reserved.set(entry.key, entry.modelId); }
      const model = built.inspection.diagramInventory.dataModels[0]; assert.deepEqual(model.findings, []);
      const independent = await page.evaluate(xml => {
        const D = "http://schemas.openxmlformats.org/drawingml/2006/diagram", A = "http://schemas.openxmlformats.org/drawingml/2006/main", doc = new DOMParser().parseFromString(xml, "application/xml");
        if (doc.getElementsByTagName("parsererror").length) throw new Error("Malformed XML");
        const points = [...doc.getElementsByTagNameNS(D, "pt")], connections = [...doc.getElementsByTagNameNS(D, "cxn")];
        const byId = new Map(points.map(p => [p.getAttribute("modelId"), p])), edges = new Map(connections.map(p => [p.getAttribute("modelId"), p]));
        if (new Set([...byId.keys(), ...edges.keys()]).size !== points.length + connections.length) throw new Error("Duplicate ID");
        for (const edge of connections) {
          if (!byId.has(edge.getAttribute("srcId")) || !byId.has(edge.getAttribute("destId"))) throw new Error("Dangling endpoint");
          for (const [attr, kind] of [["parTransId", "parTrans"], ["sibTransId", "sibTrans"]]) if (edge.hasAttribute(attr)) {
            const transition = byId.get(edge.getAttribute(attr));
            if (transition?.getAttribute("type") !== kind || transition.getAttribute("cxnId") !== edge.getAttribute("modelId")) throw new Error("Transition mismatch");
          }
        }
        for (const point of points) {
          const properties = [...point.children].find(p => p.localName === "prSet");
          if (properties?.hasAttribute("presAssocID") && !byId.has(properties.getAttribute("presAssocID"))) throw new Error("Dangling presentation association");
        }
        return { documentId: points.find(p => p.getAttribute("type") === "doc").getAttribute("modelId"),
          content: points.filter(p => [null, "asst"].includes(p.getAttribute("type"))).map(p => ({ modelId: p.getAttribute("modelId"), role: p.getAttribute("type") === "asst" ? "assistant" : "content", paragraphs: [...p.getElementsByTagNameNS(A, "p")].map(para => [...para.children].map(run => run.localName === "br" ? "\n" : run.getElementsByTagNameNS(A, "t")[0]?.textContent ?? "").join("")) })),
          parents: connections.filter(e => e.getAttribute("type") === "parOf").map(e => [e.getAttribute("srcId"), e.getAttribute("destId"), Number(e.getAttribute("srcOrd"))]) };
      }, built.specimen.dataXml);
      const currentIds = new Map(built.specimen.identities.map(e => [e.id, e.modelId]));
      const expectedParents = input.roots.map((key, index) => [independent.documentId, currentIds.get(key), index]);
      for (const n of input.nodes) n.children.forEach((key, index) => expectedParents.push([currentIds.get(n.id), currentIds.get(key), index]));
      assert.deepEqual(independent.parents.slice().sort(), expectedParents.sort());
      const allIds = [...model.points, ...model.connections].map(p => p.modelId);
      assert.equal(new Set(allIds).size, allIds.length);
      for (const edge of model.connections) {
        assert.equal(edge.source.status, "resolved"); assert.equal(edge.destination.status, "resolved");
        const key = edge.kind === "parOf" ? [edge.kind, edge.destination.value] : [edge.kind, edge.source.value, edge.destination.value];
        retainGraphId(JSON.stringify(key), edge.modelId);
        if (edge.kind === "parOf") for (const kind of ["parentTransition", "siblingTransition"]) retainGraphId(JSON.stringify([kind, edge.destination.value]), edge[kind].value);
      }
      for (const point of model.points.filter(p => p.kind === "pres")) retainGraphId(JSON.stringify(["presentation", point.presentationAssociation.value, point.properties.find(a => a.local === "presName").value]), point.modelId);
      const projected = projectDiagramContent(model, built.specimen.layoutIdentifier).candidate;
      const expectedOrder = [], visit = key => { const node = input.nodes.find(n => n.id === key); expectedOrder.push(node); node.children.forEach(visit); }; input.roots.forEach(visit);
      assert.deepEqual(projected.nodes.map(n => [n.role, n.paragraphs]), expectedOrder.map(n => [n.role, n.paragraphs]));
      assert.deepEqual(independent.content, expectedOrder.map(n => ({ modelId: currentIds.get(n.id), role: n.role, paragraphs: n.paragraphs })));
      if (round === 1) assert.notEqual(built.specimen.dataXml, previous, "insertion is an actual structural change");
      previous = built.specimen.dataXml;
      if (round === 0) { input.roots.unshift("x"); input.nodes.push(node("x", ["Inserted 😀"])); }
      if (round === 1) { input.roots.reverse(); input.nodes.reverse(); }
      if (round === 2) input.nodes.find(n => n.id === "a").paragraphs = ["Edited\n正文", ""];
      if (round === 3 && family !== "cycle") { input.roots = input.roots.filter(k => k !== "a"); input.nodes.find(n => n.id === "c").children = ["a"]; }
      if (round === 4 && family !== "cycle") { input.nodes.find(n => n.id === "a").children = []; input.nodes.find(n => n.id === "c").children = ["b"]; input.roots.push("a"); }
      if (round === 5) { input.roots = input.roots.filter(k => k !== "a"); input.nodes = input.nodes.filter(n => n.id !== "a"); }
      if (round === 6) { input.roots.push("a"); input.nodes.push(node("a", ["Explicit undo restoration"])); }
    }
    assert.equal(state.entries.length, reserved.size);
  }
});

test("identity transactions reject wrong families, implicit policy and failed output without consuming reservations", async () => {
  const c = candidate("cycle"), state = createDiagramIdentityState("cycle", "12345678-1234-4ABC-8DEF"), before = structuredClone(state);
  assert.throws(() => nativeDiagramSpecimen(c, { identityState: state }), /explicit GUID/);
  assert.throws(() => nativeDiagramSpecimen(c, { modelIds: "guid", identityState: { ...state, family: "process" } }), /family/);
  for (const limits of [{ maxWork: 100 }, { maxOutputBytes: 100 }, { maxTextBytes: 1 }]) {
    assert.throws(() => nativeDiagramSpecimen(c, { modelIds: "guid", identityState: state, presentation: "layout", limits }), RangeError);
    assert.deepEqual(state, before);
  }
  await assert.rejects(() => buildNativeDiagramCandidate(c, {}, { modelIds: "guid", identityState: state, presentation: "layout" }), /layout XML/);
  assert.deepEqual(state, before);
  const a = nativeDiagramSpecimen(c, { modelIds: "guid", identityState: state, presentation: "layout" });
  const b = nativeDiagramSpecimen(c, { modelIds: "guid", identityState: state, presentation: "layout" });
  assert.deepEqual(a, b, "failed attempts leave a repeatable allocation baseline");
  assert.deepEqual(nativeDiagramSpecimen(c, { modelIds: "guid", identityState: a.identityState, presentation: "layout" }), a);
  // Find the exact shared-work boundary without mirroring internal checkpoint
  // counts; reconciliation and later XML output must both fit the same budget.
  const minWork = identityState => {
    let lo = 1, hi = 10000;
    while (lo < hi) {
      const middle = Math.floor((lo + hi) / 2);
      try { nativeDiagramSpecimen(c, { modelIds: "guid", identityState, presentation: "layout", limits: { maxWork: middle } }); hi = middle; }
      catch (error) { assert.ok(error instanceof RangeError); assert.match(error.message, /maxWork/); lo = middle + 1; }
    }
    return lo;
  };
  const threshold = minWork(a.identityState);
  assert.ok(threshold > minWork(undefined), "identity reconciliation consumes shared work");
  assert.deepEqual(nativeDiagramSpecimen(c, { modelIds: "guid", identityState: a.identityState, presentation: "layout", limits: { maxWork: threshold } }), a);
  assert.throws(() => nativeDiagramSpecimen(c, { modelIds: "guid", identityState: a.identityState, presentation: "layout", limits: { maxWork: threshold - 1 } }), /maxWork/);
  const capturedState = structuredClone(a.identityState), capturedCandidate = structuredClone(c);
  const preparing = buildNativeDiagramCandidate(capturedCandidate, components("cycle"), { modelIds: "guid", identityState: capturedState, presentation: "layout" });
  capturedState.namespace = "FFFFFFFF-FFFF-4FFF-8FFF"; capturedState.entries[0].key = "changed";
  capturedCandidate.nodes[0].paragraphs[0] = "changed while awaiting component validation";
  const captured = await preparing;
  assert.deepEqual(captured.specimen, a, "async preparation captures the input state and content before its first await");
});

test("explicit GUID specimens remap every typed identity/reference while preserving integer controls and semantic topology", async () => {
  const guid = /^\{00000000-0000-4000-8000-[0-9A-F]{12}\}$/;
  for (const family of ["process", "hierarchy", "cycle"]) for (const cache of ["empty", "absent"]) for (const presentation of ["content-only", "layout"]) {
    const c = candidate(family), integer = nativeDiagramSpecimen(c, { cache, presentation });
    const built = await buildNativeDiagramCandidate(c, components(family), { cache, presentation, modelIds: "guid" });
    assert.equal(integer.modelIdPolicy, "int32"); assert.equal(built.specimen.modelIdPolicy, "guid");
    const model = built.inspection.diagramInventory.dataModels[0]; assert.deepEqual(model.findings, []);
    assert.ok([...model.points, ...model.connections].every(p => guid.test(p.modelId)));
    assert.equal(new Set([...model.points, ...model.connections].map(p => p.modelId)).size, model.points.length + model.connections.length);
    for (const match of built.specimen.dataXml.matchAll(/(?:modelId|srcId|destId|cxnId|parTransId|sibTransId|presAssocID)="([^"]+)"/g)) assert.match(match[1], guid);
    for (const edge of model.connections) { assert.equal(edge.source.status, "resolved"); assert.equal(edge.destination.status, "resolved"); }
    const projected = projectDiagramContent(model, built.specimen.layoutIdentifier).candidate;
    assert.deepEqual(projected.nodes.map(n => [n.paragraphs, n.role]), c.nodes.map(n => [n.paragraphs, n.role]));
    const reordered = structuredClone(c); reordered.nodes.reverse();
    assert.equal(nativeDiagramSpecimen(reordered, { cache, presentation, modelIds: "guid" }).dataXml, built.specimen.dataXml);
    assert.equal(built.specimen.nativeAdmission, "untested");
  }
  const c = candidate("cycle"), int = nativeDiagramSpecimen(c, { presentation: "layout" });
  const intBytes = new TextEncoder().encode(int.dataXml + int.frameXml).byteLength;
  assert.equal(nativeDiagramSpecimen(c, { presentation: "layout", limits: { maxOutputBytes: intBytes } }).dataXml, int.dataXml);
  assert.throws(() => nativeDiagramSpecimen(c, { presentation: "layout", modelIds: "guid", limits: { maxOutputBytes: intBytes } }), /maxOutputBytes/);
  assert.throws(() => nativeDiagramSpecimen(c, { modelIds: "future" }), /identity/);
});

test("PDF lexical observations distinguish split words from unmatched text without claiming native admission", () => {
  assert.deepEqual(pdfTextEvidence("cycle\nlead\nCycle B", ["cycle lead", "Cycle B"]), {
    missingTextMarkers: [], allTextMarkersInPdf: true, missingWhitespaceInsensitiveMarkers: [],
    allWhitespaceInsensitiveMarkersInPdf: true, textMatchStatus: "word-spacing-match",
  });
  const fragmented = pdfTextEvidence("cyc\nle \nlea\nd\nCy\ncle B", ["cycle lead", "Cycle B"]);
  assert.deepEqual(fragmented.missingTextMarkers, ["cycle lead", "Cycle B"]);
  assert.equal(fragmented.allWhitespaceInsensitiveMarkersInPdf, true);
  assert.equal(fragmented.textMatchStatus, "whitespace-fragmented-match");
  const lost = pdfTextEvidence("PPTKit native diagram control", ["cycle lead", "Cycle B"]);
  assert.equal(lost.textMatchStatus, "unmatched-markers");
  assert.deepEqual(lost.missingWhitespaceInsensitiveMarkers, ["cycle lead", "Cycle B"]);
  assert.equal(pdfTextEvidence("中\n文 😀 < &", ["中文 😀 < &"]).textMatchStatus, "whitespace-fragmented-match");
  assert.equal(pdfTextEvidence("", ["missing"]).allTextMarkersInPdf, false);
  assert.throws(() => pdfTextEvidence(null, [])); assert.throws(() => pdfTextEvidence("", [""]));
  assert.throws(() => pdfTextEvidence("anything", [" \t\r\n"]));
});

test("layout presentation graphs have unique identities, resolved typed edges and one rooted presentation tree", async () => {
  for (const family of ["process", "hierarchy", "cycle"]) for (const cache of ["empty", "absent"]) {
    const input = candidate(family), built = await buildNativeDiagramCandidate(input, components(family), { cache, presentation: "layout" });
    const model = built.inspection.diagramInventory.dataModels[0];
    assert.deepEqual(model.findings, []);
    const points = new Map(model.points.map(p => [p.elementIndex, p]));
    const displays = model.points.filter(p => p.kind === "pres");
    const keys = displays.map(p => JSON.stringify([p.presentationAssociation.targetElementIndex, p.properties.find(a => a.local === "presName").value]));
    assert.equal(new Set(keys).size, keys.length);
    assert.equal(new Set([...model.points, ...model.connections].map(p => p.modelId)).size, model.points.length + model.connections.length);
    const parents = new Map();
    for (const edge of model.connections) {
      assert.equal(edge.source.status, "resolved"); assert.equal(edge.destination.status, "resolved");
      if (edge.kind === "presParOf") {
        assert.equal(points.get(edge.source.targetElementIndex).kind, "pres");
        assert.equal(points.get(edge.destination.targetElementIndex).kind, "pres");
        assert.equal(parents.has(edge.destination.targetElementIndex), false);
        parents.set(edge.destination.targetElementIndex, edge.source.targetElementIndex);
      } else if (edge.kind === "presOf") assert.equal(points.get(edge.destination.targetElementIndex).kind, "pres");
    }
    const roots = displays.filter(p => !parents.has(p.elementIndex)); assert.equal(roots.length, 1);
    assert.equal(roots[0].properties.find(a => a.local === "presName").value, { process: "linearFlow", hierarchy: "hierChild1", cycle: "cycle" }[family]);
    for (const p of displays) {
      const seen = new Set(); let current = p.elementIndex;
      while (parents.has(current)) { assert.equal(seen.has(current), false); seen.add(current); current = parents.get(current); }
      assert.equal(current, roots[0].elementIndex);
    }
    for (const p of model.points.filter(p => ["node", "asst"].includes(p.kind))) {
      assert.equal(p.properties.some(a => a.local === "custT"), false, "Plain text must remain eligible for native automatic fitting.");
    }
    const reversed = structuredClone(input); reversed.nodes.reverse();
    assert.equal(nativeDiagramSpecimen(reversed, { cache, presentation: "layout" }).dataXml, built.specimen.dataXml);
    assert.equal(built.specimen.nativeAdmission, "untested");
  }
});

test("presentation generation shares work/output limits and preserves empty graph semantics", () => {
  const c = candidate("hierarchy"), before = structuredClone(c);
  for (const limits of [{ maxWork: 35 }, { maxOutputBytes: 5000 }]) assert.throws(() => nativeDiagramSpecimen(c, { presentation: "layout", limits }), RangeError);
  assert.deepEqual(c, before);
  for (const family of ["process", "hierarchy", "cycle"]) {
    const empty = { family, roots: [], nodes: [], ...(family === "cycle" ? { closure: "layout-generated" } : {}) };
    const built = nativeDiagramSpecimen(empty, { presentation: "layout" });
    assert.equal((built.dataXml.match(/type="pres"/g) ?? []).length, 1);
    assert.equal((built.dataXml.match(/type="presOf"/g) ?? []).length, 1);
    assert.equal((built.dataXml.match(/type="parOf"/g) ?? []).length, 0);
  }
  assert.throws(() => nativeDiagramSpecimen(c, { presentation: "unknown" }));
});

test("exact family presentation associations preserve process detail, hierarchy assistants and generated Cycle closure", async () => {
  for (const family of ["process", "hierarchy", "cycle"]) {
    const built = await buildNativeDiagramCandidate(candidate(family), components(family), { presentation: "layout" });
    const model = built.inspection.diagramInventory.dataModels[0];
    const expected = { process: [21, 21], hierarchy: [30, 30], cycle: [20, 22] }[family];
    assert.deepEqual([model.points.length, model.connections.length], expected);
    const pointByIndex = new Map(model.points.map(p => [p.elementIndex, p]));
    const name = p => p.properties?.find(a => a.local === "presName")?.value;
    const association = p => pointByIndex.get(p.presentationAssociation.targetElementIndex).modelId;
    if (family === "process") {
      const detail = model.points.find(p => p.modelId === "2");
      const edge = model.connections.find(e => e.kind === "presOf" && e.source.targetElementIndex === detail.elementIndex);
      const display = pointByIndex.get(edge.destination.targetElementIndex);
      assert.equal(name(display), "desTx"); assert.equal(association(display), "1");
    } else if (family === "hierarchy") {
      const assistantText = model.points.find(p => name(p) === "rootText3");
      assert.equal(association(assistantText), "2");
      assert.equal(assistantText.properties.find(a => a.local === "presStyleLbl").value, "asst1");
      assert.equal(model.points.some(p => name(p) === "Name111"), true);
      assert.equal(model.points.some(p => name(p) === "Name37"), false);
    } else {
      const content = model.connections.filter(e => e.kind === "parOf");
      assert.ok(content.every(e => pointByIndex.get(e.source.targetElementIndex).kind === "doc"));
      assert.equal(model.points.filter(p => name(p) === "sibTrans").length, 3);
      assert.equal(model.points.filter(p => name(p) === "connectorText").length, 3);
    }
  }
});

test("native candidate XML has unique content/connection/transition IDs and exact typed endpoint bindings", async () => {
  for (const family of ["process", "hierarchy", "cycle"]) for (const cache of ["empty", "absent"]) {
    const built = await buildNativeDiagramCandidate(candidate(family), components(family), { cache });
    const model = built.inspection.diagramInventory.dataModels[0], diagram = built.inspection.slides[0].diagrams[0];
    assert.equal(model.status, "inspected"); assert.deepEqual(model.findings, []); assert.equal(model.points.length, 10); assert.equal(model.connections.length, 3);
    const ids = [...model.points, ...model.connections].map(p => p.modelId); assert.equal(new Set(ids).size, 13);
    for (const edge of model.connections) {
      assert.equal(edge.source.status, "resolved"); assert.equal(edge.destination.status, "resolved");
      const parent = model.points.find(p => p.elementIndex === edge.parentTransition.targetElementIndex), sibling = model.points.find(p => p.elementIndex === edge.siblingTransition.targetElementIndex);
      assert.equal(parent.kind, "parTrans"); assert.equal(sibling.kind, "sibTrans");
      assert.equal(parent.connection.targetElementIndex, edge.elementIndex); assert.equal(sibling.connection.targetElementIndex, edge.elementIndex);
    }
    assert.equal(diagram.cache.status, cache === "empty" ? "located" : "missing-reference");
    assert.equal(built.specimen.nativeAdmission, "untested");
    const projected = projectDiagramContent(model, built.specimen.layoutIdentifier).candidate;
    assert.deepEqual(projected.nodes.map(n => n.paragraphs), candidate(family).nodes.map(n => n.paragraphs));
    assert.equal(projected.nodes[1].role, family === "hierarchy" ? "assistant" : "content");
  }
});

test("serialization depends on authored topology/order, not node storage order or caller mutation", async () => {
  const input = candidate("process"), before = structuredClone(input), a = nativeDiagramSpecimen(input, { cache: "empty" });
  input.nodes.reverse(); assert.deepEqual(nativeDiagramSpecimen(input, { cache: "empty" }), a);
  assert.deepEqual(input.nodes.slice().reverse(), before.nodes);
  const parts = components("process"), pending = buildNativeDiagramCandidate(before, parts);
  parts.layout.fill(0); before.nodes[0].paragraphs[0] = "changed";
  const built = await pending;
  assert.equal(built.inspection.diagramInventory.dataModels[0].points[1].text[0].runs[0].text, candidate("process").nodes[0].paragraphs[0]);
});

test("malformed graphs, unknown fields, invalid XML scalars and unsupported assistant/closure cases reject", () => {
  const bad = [
    c => { c.family = "future"; }, c => { c.nodes[1].id = "a"; }, c => { c.nodes[0].children = ["missing"]; },
    c => { c.roots.push("a"); }, c => { c.roots = ["a"]; }, c => { c.nodes[1].children = ["a"]; },
    c => { c.nodes[1].role = "assistant"; }, c => { c.nodes[0].style = {}; }, c => { c.nodes[0].paragraphs = []; },
    c => { c.nodes[0].paragraphs = ["bad\u0000"]; }, c => { c.nodes[0].paragraphs = ["bad\uD800"]; },
    c => { c.nodes[0].paragraphs = ["bad\uFFFE"]; }, c => { c.nodes[0].paragraphs = [{ text: "rich" }]; },
  ];
  for (const change of bad) { const c = candidate("process"); change(c); assert.throws(() => nativeDiagramSpecimen(c)); }
  const c = candidate("hierarchy"); c.nodes[0].role = "assistant"; assert.throws(() => nativeDiagramSpecimen(c));
  const cyclic = candidate("cycle"); cyclic.closure = "source-cycle"; assert.throws(() => nativeDiagramSpecimen(cyclic));
  const getter = { get family() { throw new Error("must not execute"); } }; assert.throws(() => nativeDiagramSpecimen(getter), /accessor/);
});

test("node/depth/work/text/output budgets reject before returning a candidate and allow clean retries", () => {
  const c = candidate("process"), before = structuredClone(c);
  for (const limits of [{ maxNodes: 2 }, { maxDepth: 1 }, { maxWork: 2 }, { maxTextBytes: 1 }, { maxOutputBytes: 100 }]) assert.throws(() => nativeDiagramSpecimen(c, { limits }));
  for (const options of [{ cache: "stale" }, { drawingId: 0 }, { frame: { x: 0, y: 0, cx: 0, cy: 1 } }, { limits: { unknown: 1 } }, { limits: { maxWork: Infinity } }]) assert.throws(() => nativeDiagramSpecimen(c, options));
  assert.deepEqual(c, before); assert.ok(nativeDiagramSpecimen(c).dataXml.length > 0);
  for (const family of ["process", "hierarchy", "cycle"]) {
    const empty = { family, roots: [], nodes: [], ...(family === "cycle" ? { closure: "layout-generated" } : {}) };
    assert.equal(nativeDiagramSpecimen(empty).identities.length, 0);
  }
});

test("fresh package relationships/content types, frame units and explicit cache modes stay independent of old drawings", async () => {
  const built = await buildNativeDiagramCandidate(candidate("process"), components("process"), { cache: "empty", frame: { x: -1, y: 1, cx: 10, cy: 20 } });
  const parts = readGeneratedZip(built.bytes), xml = parts.get("ppt/slides/slide1.xml").toString();
  assert.match(xml, /<a:off x="-1" y="1"\/><a:ext cx="10" cy="20"\/>/);
  const roles = built.inspection.package.relationships.filter(r => r.source === "/ppt/slides/slide1.xml" && r.id.startsWith("pptkitDiagram"));
  assert.equal(roles.length, 5); assert.ok(roles.every(r => r.targetMode === "Internal"));
  assert.ok(roles.some(r => r.type.endsWith("/diagramQuickStyle") && r.targetPart.endsWith("/style1.xml")));
  assert.equal(built.inspection.diagramInventory.cacheMappings[0].shapes.length, 0);
  assert.equal((await inspectXmlPart(parts.get("ppt/diagrams/data1.xml"))).root.local, "dataModel");
  const invalid = components("process"); invalid.style = invalid.colors;
  await assert.rejects(buildNativeDiagramCandidate(candidate("process"), invalid));
  const wrongLayout = components("cycle"); await assert.rejects(buildNativeDiagramCandidate(candidate("process"), wrongLayout));
});

test("three unrelated text-save rounds preserve generated model/components and do not admit diagram edits", async () => {
  for (const cache of ["empty", "absent"]) for (const presentation of ["content-only", "layout"]) for (const identityPolicy of ["int32", "guid", "managed"]) {
    const modelIds = identityPolicy === "int32" ? "int32" : "guid", identityState = identityPolicy === "managed" ? createDiagramIdentityState("hierarchy", "12345678-1234-4ABC-8DEF") : undefined;
    let bytes = (await buildNativeDiagramCandidate(candidate("hierarchy"), components("hierarchy"), { cache, presentation, modelIds, identityState })).bytes;
    for (let round = 0; round < 3; round++) {
      const before = readGeneratedZip(bytes), session = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
      try {
        assert.equal(session.document.slides[0].elements.length, 1);
        const slide = session.document.slides[0], control = structuredClone(slide.elements[0]); control.content[0].runs[0].text = `Control round ${round}`;
        slide.removeElement(control.id); slide.insertElement(0, control);
        const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
        const saved = (await session.save(prepared.plan)).bytes;
        assert.deepEqual((await compareOpcPackages(bytes, saved)).differences.map(d => d.part), ["/ppt/slides/slide1.xml"]);
        for (const [name, payload] of before) if (name.startsWith("ppt/diagrams/")) assert.deepEqual(readGeneratedZip(saved).get(name), payload);
        bytes = saved;
      } finally { session.dispose(); }
    }
  }
});

test("independent browser DOM resolves generated identities, text and transitions across both cache policies", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage();
  await page.route("**/*", route => route.abort());
  const source = await readFile(new URL("../packages/pptx-parser/dist/browser/index.js", import.meta.url), "utf8");
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, source);
  for (const family of ["process", "hierarchy", "cycle"]) for (const cache of ["empty", "absent"]) for (const presentation of ["content-only", "layout"]) for (const identityPolicy of ["int32", "guid", "managed"]) {
    const modelIds = identityPolicy === "int32" ? "int32" : "guid", identityState = identityPolicy === "managed" ? createDiagramIdentityState(family, "12345678-1234-4ABC-8DEF") : undefined;
    const built = await buildNativeDiagramCandidate(candidate(family), components(family), { cache, presentation, modelIds, identityState }), parts = readGeneratedZip(built.bytes);
    const actual = await page.evaluate(async ({ bytes, xml }) => {
      const document = new DOMParser().parseFromString(xml, "application/xml"), D = "http://schemas.openxmlformats.org/drawingml/2006/diagram", A = "http://schemas.openxmlformats.org/drawingml/2006/main";
      if (document.getElementsByTagName("parsererror").length) throw new Error("Invalid generated XML");
      const raw = [...document.getElementsByTagName("*")], points = [...document.getElementsByTagNameNS(D, "pt")];
      const nodeText = points.filter(p => [null, "asst"].includes(p.getAttribute("type"))).map(p => [...p.getElementsByTagNameNS(A, "p")].map(para => [...para.children].map(run => run.localName === "br" ? "\n" : run.getElementsByTagNameNS(A, "t")[0]?.textContent ?? "").join("")));
      const edges = [...document.getElementsByTagNameNS(D, "cxn")].map(edge => ({ source: points.find(p => p.getAttribute("modelId") === edge.getAttribute("srcId")), destination: points.find(p => p.getAttribute("modelId") === edge.getAttribute("destId")) })).map(e => [raw.indexOf(e.source) + 1, raw.indexOf(e.destination) + 1]);
      const inspection = await globalThis.parser.inspectPptx(Uint8Array.from(bytes));
      return { nodeText, edges, inventory: inspection.diagramInventory };
    }, { bytes: [...built.bytes], xml: parts.get("ppt/diagrams/data1.xml").toString() });
    assert.deepEqual(actual.nodeText, candidate(family).nodes.map(n => n.paragraphs));
    assert.deepEqual(actual.edges, built.inspection.diagramInventory.dataModels[0].connections.map(c => [c.source.targetElementIndex, c.destination.targetElementIndex]));
    assert.deepEqual(actual.inventory, built.inspection.diagramInventory);
  }
});
