import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
import { DSP } from "../packages/pptx-parser/test/diagram-resources-fixture.mjs";
import { diagramDataFixture } from "../packages/pptx-parser/test/diagram-data-fixture.mjs";
import { shape, body } from "../packages/pptx-parser/test/content-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

test("independent browser DOM verifies raw frame/group identities and component links in both dialects", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage();
  await page.route("**/*", route => route.abort());
  const bundle = await readFile(new URL("../packages/pptx-parser/dist/browser/index.js", import.meta.url), "utf8");
  await page.evaluate(async bundle => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(bundle)))}`); }, bundle);
  for (const strict of [false, true]) {
    const f = diagramDataFixture(strict);
    f.setTree(f.frame("2") + `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3"/></p:nvGrpSpPr>${f.frame("4")}</p:grpSp>`);
    const bytes = f.bytes(), parts = readGeneratedZip(bytes), report = await inspectPptx(bytes);
    const actual = await page.evaluate(async ({ bytes, xml, dataXml, cacheXml, slideRels, dsp, p, a, d, r }) => {
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      if (doc.getElementsByTagName("parsererror").length) throw new Error("Invalid XML");
      const all = [...doc.getElementsByTagName("*")], diagrams = [...doc.getElementsByTagNameNS(p, "graphicFrame")].map(frame => {
        const ids = frame.getElementsByTagNameNS(d, "relIds")[0], nv = frame.getElementsByTagNameNS(p, "cNvPr")[0];
        const groups = []; let n = frame.parentElement;
        while (n) { if (n.namespaceURI === p && n.localName === "grpSp") groups.unshift(all.indexOf(n) + 1); n = n.parentElement; }
        assertGraphic(frame.getElementsByTagNameNS(a, "graphicData")[0].getAttribute("uri") === d);
        return { elementIndex: all.indexOf(frame) + 1, groupElementIndices: groups, drawingId: nv.getAttribute("id"), refs: ["dm", "lo", "cs", "qs"].map(attr => ids.getAttributeNS(r, attr)) };
      });
      function assertGraphic(ok) { if (!ok) throw new Error("Wrong diagram URI"); }
      const dataDoc = new DOMParser().parseFromString(dataXml, "application/xml"), raw = dataDoc.getElementsByTagNameNS(dsp, "dataModelExt")[0];
      const cacheRef = { elementIndex: [...dataDoc.getElementsByTagName("*")].indexOf(raw) + 1, relationshipId: raw.getAttribute("relId"), minimumVersion: raw.getAttribute("minVer") };
      const relDoc = new DOMParser().parseFromString(slideRels, "application/xml");
      const rel = [...relDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship")].find(n => n.getAttribute("Id") === cacheRef.relationshipId);
      const cacheTarget = rel.getAttribute("Target");
      const inspected = await globalThis.parser.inspectPptx(Uint8Array.from(bytes));
      const dataNodes = [...dataDoc.getElementsByTagName('*')];
      const rawPoints = [...dataDoc.getElementsByTagNameNS(d, 'pt')].map(n => ({ elementIndex: dataNodes.indexOf(n) + 1, modelId: n.getAttribute('modelId'), kind: n.getAttribute('type') ?? 'node' }));
      const rawConnections = [...dataDoc.getElementsByTagNameNS(d, 'cxn')].map(n => ({ elementIndex: dataNodes.indexOf(n) + 1, modelId: n.getAttribute('modelId'), kind: n.getAttribute('type') ?? 'parOf', source: n.getAttribute('srcId'), destination: n.getAttribute('destId'), sourceOrder: n.getAttribute('srcOrd') }));
      const rawPresentationMappings = [...dataDoc.getElementsByTagNameNS(d, 'pt')].filter(n => n.getAttribute('type') === 'pres').map(n => {
        const elementIndex = dataNodes.indexOf(n) + 1, associated = n.getElementsByTagNameNS(d, 'prSet')[0].getAttribute('presAssocID');
        const targetElementIndex = rawPoints.find(p => p.modelId === associated).elementIndex;
        const reference = value => ({ value, status: 'resolved', targetElementIndex: rawPoints.find(p => p.modelId === value).elementIndex });
        return { elementIndex, sources: [{ kind: 'presAssocID', elementIndex, reference: reference(associated) }, ...rawConnections.filter(c => c.kind === 'presOf' && c.destination === n.getAttribute('modelId')).map(c => ({ kind: 'presOf', elementIndex: c.elementIndex, reference: reference(c.source) }))], agreement: 'same-target', targetElementIndex };
      });
      const cacheDoc = new DOMParser().parseFromString(cacheXml, 'application/xml'), cacheNodes = [...cacheDoc.getElementsByTagName('*')];
      const rawCacheShapes = [...cacheDoc.getElementsByTagNameNS(dsp, 'sp')].map(n => ({ elementIndex: cacheNodes.indexOf(n) + 1, model: { value: n.getAttribute('modelId'), status: 'resolved', targetElementIndex: rawPoints.find(p => p.modelId === n.getAttribute('modelId')).elementIndex } }));
      let graphBudgetRejected = false;
      try { await globalThis.parser.inspectPptx(Uint8Array.from(bytes), { diagramLimits: { maxWorkUnits: 1 } }); } catch (error) { graphBudgetRejected = error.code === 'package-limit-exceeded'; }
      return { diagrams, inspected: inspected.slides[0].diagrams, inventory: inspected.diagramInventory, cacheRef, cacheTarget, rawPoints, rawConnections, rawPresentationMappings, rawCacheShapes, graphBudgetRejected };
    }, { bytes: [...bytes], xml: parts.get("deck/z.xml").toString(), dataXml: parts.get("deck/diagrams/data.xml").toString(), cacheXml: parts.get("deck/diagrams/cache.xml").toString(), slideRels: parts.get("deck/_rels/z.xml.rels").toString(), dsp: DSP, p: f.P, a: f.A, d: f.D, r: f.R });
    assert.deepEqual(actual.inspected, report.slides[0].diagrams);
    assert.deepEqual(actual.inventory, report.diagramInventory);
    assert.deepEqual(actual.inspected[0].cache.references, [actual.cacheRef]);
    assert.equal(actual.inspected[0].cache.target, actual.cacheTarget);
    assert.equal(actual.inventory.parts.find(p => p.part.endsWith("shared.png")).incomingRelationships.length, 3);
    const graph = actual.inventory.dataModels[0]; assert.equal(graph.status, "inspected"); assert.equal(actual.graphBudgetRejected, true);
    assert.deepEqual(graph.presentationMappings, actual.rawPresentationMappings);
    assert.equal(actual.inventory.cacheMappings.length, 1); assert.deepEqual(actual.inventory.cacheMappings[0].shapes, actual.rawCacheShapes);
    assert.deepEqual(graph.points.map(p => ({ elementIndex: p.elementIndex, modelId: p.modelId, kind: p.kind })), actual.rawPoints);
    assert.deepEqual(graph.connections.map(c => ({ elementIndex: c.elementIndex, modelId: c.modelId, kind: c.kind, source: c.source.value, destination: c.destination.value, sourceOrder: c.sourceOrder })), actual.rawConnections);
    for (const edge of graph.connections) { assert.equal(edge.source.targetElementIndex, actual.rawPoints.find(p => p.modelId === edge.source.value).elementIndex); assert.equal(edge.destination.targetElementIndex, actual.rawPoints.find(p => p.modelId === edge.destination.value).elementIndex); }
    assert.deepEqual(actual.diagrams, report.slides[0].diagrams.map(d => ({ elementIndex: d.elementIndex, groupElementIndices: d.groupElementIndices, drawingId: d.drawingId, refs: d.components.map(c => c.relationshipId) })));
  }
});

test("three unrelated text saves retain source-only diagram markup and every component payload", async () => {
  for (const strict of [false, true]) {
    const f = diagramDataFixture(strict); f.setTree(f.frame() + shape({ id: "10", txBox: true, text: body() }));
    let bytes = new Uint8Array(f.bytes());
    for (let round = 0; round < 3; round++) {
        const before = readGeneratedZip(bytes), session = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
        const sourceInventory = session.report.inspection.diagramInventory;
      try {
        assert.equal(session.document.slides[0].elements.length, 1);
        const element = structuredClone(session.document.slides[0].elements[0]); element.content[0].runs[0].text = `Round ${round}`;
        session.document.slides[0].removeElement(element.id); session.document.slides[0].insertElement(0, element);
        const prep = await session.prepareSave({ mode: "edit" }); assert.equal(prep.status, "ready"); bytes = (await session.save(prep.plan)).bytes;
        const after = readGeneratedZip(bytes), frame = xml => xml.toString().match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/)[0];
        assert.equal(frame(after.get("deck/z.xml")), frame(before.get("deck/z.xml")));
        for (const [name, value] of before) if (name !== "deck/z.xml") assert.deepEqual(after.get(name), value, name);
        const [diagram] = (await inspectPptx(bytes)).slides[0].diagrams; assert.ok(diagram.components.every(c => c.status === "located"));
        assert.equal(diagram.cache.status, "located"); assert.equal(diagram.cache.freshness, "unknown");
        assert.equal((await inspectPptx(bytes)).diagramInventory.dataModels[0].points[1].text[0].runs[0].text, "A & B 😀");
        const afterInventory = (await inspectPptx(bytes)).diagramInventory;
        assert.deepEqual(afterInventory.cacheMappings, sourceInventory.cacheMappings);
        assert.deepEqual(afterInventory.dataModels[0].presentationMappings, sourceInventory.dataModels[0].presentationMappings);
      } finally { session.dispose(); }
    }
  }
});
