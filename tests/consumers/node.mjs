import { createBrowserTextMeasurement } from "@pptkit/layout/dom";
import assert from "node:assert/strict";
import { readFile, realpath, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createPresentation, normalizePresentation, resolvePresentationInheritance, validatePresentation, prepareLayoutRebind, formatListMarker } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";
import { inspectPptx, inspectMarkupCompatibility, inspectPresentationPackage, inspectOpcPackage, inspectXmlPart, inspectZipArchive } from "@pptkit/pptx-parser";
import { importPptx, compareOpcPackages, inspectOpcClosure } from "@pptkit/pptx-parser";
import { importPptxFile, writePptxFile, PptxFileError } from "@pptkit/pptx-parser/node";
import { generatePptx, inspectGif, inspectPng, inspectApng } from "@pptkit/pptx-exporter";
import { writePptx } from "@pptkit/pptx-exporter/node";
import { renderLayoutToSvg, renderPresentationToSvg } from "@pptkit/svg-renderer";
import { runCli } from "@pptkit/cli";

assert.equal(inspectGif(Uint8Array.from(Buffer.from("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7", "base64"))).frames.length, 1);
assert.equal(formatListMarker({ type: "number", style: "alphaLowerPeriod", startAt: 53 }), "ba. ");
assert.equal(inspectPng(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAFUlEQVR4nGP4z8DwH4QZoIAJxoABAHXOA/+HXAe/AAAAAElFTkSuQmCC", "base64")).width, 4);
assert.equal(inspectApng(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAFUlEQVR4nGP4z8DwH4QZoIAJxoABAHXOA/+HXAe/AAAAAElFTkSuQmCC", "base64")).animation, undefined);
const directory = path.dirname(fileURLToPath(import.meta.url));
const diagramInventory = await inspectPptx(await readFile(path.join(directory, "smartart-inventory.pptx")));
assert.equal(diagramInventory.slides[0].diagrams[0].identityStatus, "unique");
assert.deepEqual(diagramInventory.slides[0].diagrams[0].components.map(c => c.status), ["located", "located", "located", "located"]);
assert.equal(diagramInventory.slides[0].diagrams[0].cache.status, "located");
assert.equal(diagramInventory.slides[0].diagrams[0].cache.freshness, "unknown");
assert.equal(diagramInventory.diagramInventory.parts.find(p => p.part.endsWith("shared.png")).incomingRelationships.length, 3);
const cachedPreview = await importPptx(await readFile(path.join(directory, "smartart-preview.pptx")), { projection: "content", diagramPreview: "cache" });
assert.equal(cachedPreview.report.diagramPreviews[0].status, "cache");
assert.equal(cachedPreview.document.slides[0].elements[0].children[0].children[0].children[1].content[0].runs[0].text, "Cached text");
assert.deepEqual(cachedPreview.document.slides[0].elements[0].children[0].children[0].children[1].box, { x: 30, y: 25, width: 80, height: 20 });
cachedPreview.dispose();
const diagramGraph = (await inspectPptx(await readFile(path.join(directory, "smartart-data.pptx")))).diagramInventory.dataModels[0];
assert.equal(diagramGraph.points[1].kind, "node"); assert.equal(diagramGraph.connections[0].source.status, "resolved");
assert.equal(diagramGraph.points[1].text[0].runs[0].text, "A & B 😀");
assert.equal(diagramGraph.presentationMappings[0].agreement, "same-target");
const cacheMapping = (await inspectPptx(await readFile(path.join(directory, "smartart-data.pptx")))).diagramInventory.cacheMappings[0];
assert.equal(cacheMapping.shapes[0].model.targetElementIndex, diagramGraph.points[5].elementIndex);
for (const name of ["core", "layout", "pptx-exporter", "svg-renderer", "cli", "pptx-parser"]) {
  const resolved = await realpath(fileURLToPath(import.meta.resolve(`@pptkit/${name}`)));
  assert.ok(resolved.startsWith(`${await realpath(directory)}${path.sep}node_modules${path.sep}`), `${name} must resolve within the installed consumer`);
  await assert.rejects(import(`@pptkit/${name}/dist/index.js`), { code: "ERR_PACKAGE_PATH_NOT_EXPORTED" });
  if (name !== "cli") assert.ok(Object.keys(await import(`@pptkit/${name}/browser`)).length > 0);
}
const document = createPresentation({ id: "installed-consumer" });
const installedPng = new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC", "base64"));
document.registerAsset({ id: "installed-bytes", kind: "image", source: { type: "bytes", value: installedPng } });
const slide = document.addSlide();
slide.setNotes([{ runs: [{ text: "Installed speaker notes", style: { bold: true, fontSize: 18 } }, { text: " Self", action: { type: "slide", slideId: slide.id } }] }]);
slide.addElement({ type: "image", assetId: "installed-bytes", box: { x: 500, y: 20, width: 80, height: 40 } });
slide.addElement({ type: "text", content: [{ runs: [{ text: "Installed package", style: { colorOpacity: .25 } }] }], box: { x: 40, y: 40, width: 300, height: 80 } });
slide.addElement({ type: "chart", chartType: "bar", categories: ["A", "B"], series: [{ name: "Data", values: [2, 3] }], box: { x: 40, y: 160, width: 400, height: 260 } });
slide.addElement({ type: "table", id: "installed-table", box: { x: 450, y: 160, width: 200, height: 80 }, columns: [100, 100], transform: { rotation: 15, flipH: true }, rows: [
  { height: 40, cells: [{ content: [{ runs: [{ text: "Installed merged cell", action: { type: "slide", slideId: slide.id } }] }], rowSpan: 2, colSpan: 2, style: { fill: { type: "solid", color: "AABBCC" }, margin: 3 } }] },
  { height: 40, cells: [] },
] });
assert.deepEqual(validatePresentation(document), []);
assert.equal(normalizePresentation(document).irVersion, 2);
const layout = resolveLayout(document);
const preview = await renderLayoutToSvg(layout);
assert.match(preview.slides[0].svg, /Installed package/);
assert.match(preview.slides[0].svg, /fill-opacity="0.25"/);
assert.deepEqual(preview.warnings.map(warning => warning.code), ["slide-action-not-linked"]);
const generated = await generatePptx(document);
assert.equal(generated.status, "generated-with-warnings");
assert.deepEqual(generated.warnings.map(w => w.code), ["table-transform-grouped"]);
assert.ok((await inspectZipArchive(generated.bytes)).entries.some((entry) => entry.name === "ppt/presentation.xml"));
assert.deepEqual([...generated.bytes.subarray(0, 2)], [80, 75]);
const written = await writePptx(document, { output: path.join(directory, "output/deck.pptx") });
assert.equal(written.status, "written-with-warnings");
assert.ok((await readFile(written.output)).equals(Buffer.from(generated.bytes)));
assert.equal(runCli(["--help"]).exitCode, 0);
console.log(JSON.stringify({ slideCount: generated.slideCount, bytes: generated.byteLength, previewCount: preview.slides.length }));

assert.equal(typeof createBrowserTextMeasurement, "function");

const xml = await inspectXmlPart(new TextEncoder().encode('<p:r xmlns:p="urn:installed"><p:c/></p:r>'));
assert.equal(xml.root.namespaceUri, "urn:installed");
assert.equal(xml.elements, 2);
await assert.rejects(inspectXmlPart(new TextEncoder().encode("<r>")), { code: "xml-invalid" });

const opc = await inspectOpcPackage(generated.bytes);
assert.ok(opc.relationships.some((relationship) => relationship.source === "/" && relationship.targetPart === "/ppt/presentation.xml"));
const comparison = await compareOpcPackages(generated.bytes, await readFile(written.output), { profile: "opc-xml-v1" });
assert.equal(comparison.equal, true);
assert.equal(comparison.differenceCount, 0);
assert.ok(Object.isFrozen(comparison.differences));
const dependencyClosure = await inspectOpcClosure(generated.bytes, { roots: ["/ppt/slides/slide1.xml"] });
assert.ok(dependencyClosure.parts.some((part) => part.name === "/ppt/slides/slide1.xml"));
assert.deepEqual(dependencyClosure.coverage, { outgoingComplete: true, incomingComplete: true });
assert.ok(dependencyClosure.expandedBytes <= opc.archive.expandedBytes);

assert.equal((await inspectPresentationPackage(generated.bytes)).suggestedExtension, "pptx");

assert.equal((await inspectMarkupCompatibility(new TextEncoder().encode('<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:Ignorable="x"><x:drop/><kept/></r>'), { understoodNamespaces: [] })).output.elements, 2);

assert.equal((await inspectPptx(generated.bytes)).slides.length, generated.slideCount);
const alphaImport = await importPptx(generated.bytes, { projection: "content" });
assert.equal(alphaImport.document.slides[0].elements.find(e => e.type === "text").content[0].runs[0].style.colorOpacity, .25);
alphaImport.dispose();
const imported = await importPptx(generated.bytes, { projection: "structure" });
const preservation = await imported.prepareSave({ mode: "no-op" });
assert.equal(preservation.status, "ready");
assert.equal(Buffer.compare(Buffer.from((await imported.save(preservation.plan)).bytes), Buffer.from(generated.bytes)), 0);
assert.equal(imported.document.slides.length, generated.slideCount);
imported.dispose();
const nativeSections = await importPptx(new Uint8Array(await readFile(path.join(directory, "native-sections.pptx"))), { projection: "structure" });
assert.deepEqual(nativeSections.document.slides.map(slide => slide.section), ["Opening", "Closing"]);
assert.equal(nativeSections.report.inspection.sectionList.sections[1].slides.length, 0);
assert.ok(nativeSections.report.findings.some(finding => finding.feature === "section.native-collection" && finding.operation === "export" && finding.tier === "unsupported"));
nativeSections.dispose();
assert.equal((await inspectPptx(generated.bytes)).metadata.extended.properties.find((p) => p.name === "Application").value.value, "PPTKit");

const resolvedBytes = await generatePptx(document, { maxAssetBytes: 1024, resolveAsset: () => ({ bytes: installedPng, mimeType: "image/png" }) });
assert.deepEqual(resolvedBytes.bytes, generated.bytes);

const lifecycle = createPresentation();
const sharedAsset = lifecycle.registerAsset({ kind: "image", source: { type: "bytes", value: installedPng } });
assert.equal(lifecycle.registerAsset({ kind: "image", source: { type: "bytes", value: installedPng } }, { deduplicate: "content" }), sharedAsset);
const lifecycleSlide = lifecycle.addSlide();
lifecycleSlide.addElement({ id: "shared-use", type: "image", assetId: sharedAsset.id, box: { x: 0, y: 0, width: 40, height: 20 } });
assert.equal(lifecycle.getAssetReferences(sharedAsset.id)[0].elementId, "shared-use");
assert.throws(() => lifecycle.removeAsset(sharedAsset.id), /referenced/);
lifecycle.replaceAsset(sharedAsset.id, { kind: "image", source: { type: "bytes", value: installedPng }, width: 8, height: 4 });
assert.equal((await generatePptx(lifecycle)).warnings.length, 0);
lifecycleSlide.removeElement("shared-use");
assert.equal(lifecycle.removeUnusedAssets({ retainAssetIds: [sharedAsset.id] }).length, 0);
assert.equal(lifecycle.removeUnusedAssets().length, 1);
assert.equal(lifecycle.assets.length, 0);

const copiedDocument = createPresentation(), sourceSlide = copiedDocument.addSlide({ id: "source" });
sourceSlide.addElement({ type: "shape", shape: "rect", id: "anchor", box: { x: 10, y: 10, width: 20, height: 20 }, action: { type: "slide", slideId: "source" } });
sourceSlide.addElement({ type: "connector", start: { elementId: "anchor" }, end: { x: 100, y: 100 } });
const copiedSlide = copiedDocument.duplicateSlide(sourceSlide.id); copiedDocument.removeSlide(sourceSlide.id);
assert.equal(copiedSlide.elements[1].start.elementId, copiedSlide.elements[0].id);
assert.equal(copiedSlide.elements[0].action.slideId, copiedSlide.id);
assert.deepEqual(validatePresentation(copiedDocument), []);
assert.deepEqual((await generatePptx(copiedDocument)).warnings, []);

const nativeDocument = createPresentation(), nativeSlide = nativeDocument.addSlide();
nativeSlide.addElement({ type: "connector", start: { elementId: "native-a", anchor: "right" }, end: { elementId: "native-b", anchor: "left" } });
nativeSlide.addElement({ type: "shape", id: "native-a", shape: "rect", box: { x: 0, y: 0, width: 40, height: 20 }, action: { type: "slide", slideId: nativeSlide.id } });
nativeSlide.addElement({ type: "shape", id: "native-b", shape: "ellipse", box: { x: 100, y: 0, width: 40, height: 20 }, action: { type: "url", url: "https://example.invalid/?x=1&y=2" } });
const nativeGenerated = await generatePptx(nativeDocument, { connectorMode: "native" });
assert.deepEqual(nativeGenerated.warnings, []);
assert.equal(resolveLayout(nativeDocument).slides[0].elements[0].connections.end.elementId, "native-b");
assert.equal((await inspectPptx(nativeGenerated.bytes)).slides.length, 1);
const nativeImported = await importPptx(nativeGenerated.bytes, { projection: "content" });
const nativeImportedLine = nativeImported.document.slides[0].elements[0];
assert.equal(nativeImportedLine.type, "connector");
assert.equal(nativeImportedLine.start.elementId, nativeImported.document.slides[0].elements[1].id);
assert.equal(nativeImportedLine.end.elementId, nativeImported.document.slides[0].elements[2].id);
assert.equal(nativeImported.document.slides[0].elements[1].action.slideId, nativeImported.document.slides[0].id);
assert.equal(nativeImported.document.slides[0].elements[2].action.url, "https://example.invalid/?x=1&y=2");
assert.deepEqual((await generatePptx(nativeImported.document, { connectorMode: "native" })).warnings, []);
nativeImported.dispose();
const contentSession = await importPptx(generated.bytes, { projection: "content" });
assert.equal(contentSession.document.slides[0].notes[0].runs[0].text, "Installed speaker notes");
assert.equal(contentSession.document.slides[0].notes[0].runs[1].action.slideId, contentSession.document.slides[0].id);
assert.ok(contentSession.report.identities.some(identity => identity.kind === "notes"));
assert.equal(contentSession.report.projection, "content");
assert.ok(contentSession.report.identities.some(identity => identity.kind === "element"));
assert.equal(contentSession.document.slides[0].elements[0].type, "image");
assert.equal(contentSession.document.assets[0].source.value, Buffer.from(installedPng).toString("base64"));
assert.ok(contentSession.report.identities.some(identity => identity.kind === "asset"));
assert.deepEqual((await generatePptx(contentSession.document)).warnings.map(warning => warning.code), ["native-table-preset-direct-formatting"]);
assert.equal((await contentSession.prepareSave({ mode: "edit" })).status, "blocked");
const contentPlan = await contentSession.prepareSave({ mode: "no-op" });
assert.equal(contentPlan.status, "ready");
assert.deepEqual((await contentSession.save(contentPlan.plan)).bytes, generated.bytes);
const installedTableGroup = normalizePresentation(contentSession.document).slides[0].elements.find(e => e.type === "group" && e.children[0]?.type === "table");
assert.deepEqual(installedTableGroup.transform, { rotation: 15, flipH: true, flipV: false });
const installedTable = installedTableGroup.children[0];
assert.deepEqual(installedTable.columns, [100, 100]);
assert.deepEqual(installedTable.transform, { rotation: 0, flipH: false, flipV: false });
assert.equal(installedTable.rows[1].cells.length, 0);
assert.equal(installedTable.rows[0].cells[0].content[0].runs[0].action.slideId, contentSession.document.slides[0].id);
const tableFresh = await generatePptx(contentSession.document);
const tableAgain = await importPptx(tableFresh.bytes, { projection: "content" });
assert.equal(normalizePresentation(tableAgain.document).slides[0].elements.find(e => e.type === "group" && e.children[0]?.type === "table").children[0].rows[0].cells[0].content[0].runs[0].text, "Installed merged cell");
tableAgain.dispose();
contentSession.dispose();

const installedCharts = createPresentation(), chartSlide = installedCharts.addSlide();
for (const chartType of ["bar", "line", "pie"]) chartSlide.addElement({ type: "chart", chartType, box: { x: 10, y: 20, width: 400, height: 240 }, transform: { rotation: 15, flipH: true }, categories: ["A", "B"], series: [{ name: "Installed chart", values: [0, 2.5] }] });
chartSlide.addElement({ type: "chart", chartType: "doughnut", holeSize: 1, box: { x: 10, y: 20, width: 400, height: 240 }, categories: ["A", "B"], series: [{ name: "Installed doughnut", values: [0, 2.5] }] });
const chartOutput = await generatePptx(installedCharts); assert.ok(chartOutput.warnings.some(w => w.code === "doughnut-small-hole-office-extension"));
const chartSession = await importPptx(chartOutput.bytes, { projection: "content" });
assert.deepEqual(chartSession.document.slides[0].elements.map(e => [e.chartType, e.series[0].values]), [["bar", [0, 2.5]], ["line", [0, 2.5]], ["pie", [0, 2.5]], ["doughnut", [0, 2.5]]]);
assert.equal(chartSession.document.slides[0].elements[0].transform.rotation, 15);
assert.equal(chartSession.document.slides[0].elements[3].holeSize, 1);
assert.ok(chartSession.report.findings.some(f => f.feature === "chart.data-source" && /cache agrees/.test(f.reason)));
assert.deepEqual(validatePresentation(chartSession.document), []);
assert.equal((await chartSession.prepareSave({ mode: "edit" })).status, "blocked");
const chartPlan = await chartSession.prepareSave({ mode: "no-op" }); assert.deepEqual((await chartSession.save(chartPlan.plan)).bytes, chartOutput.bytes);
chartSession.dispose();

const placeholderSession = await importPptx(new Uint8Array(await readFile(path.join(directory, "placeholder-inheritance.pptx"))), { projection: "content", sessionLimits: { maxPlaceholderEntries: 4 } });
const inherited = placeholderSession.report.placeholderInheritance.placeholders.find(p => p.source.part === "/deck/z.xml");
assert.equal(inherited.parent.status, "matched");
assert.equal(inherited.dimensions.width.source.part, "/deck/master.xml");
assert.equal(inherited.dimensions.x.source.part, "/deck/layout.xml");
assert.deepEqual(placeholderSession.document.slides[0].elements[0].box, { x: 10, y: 20, width: 200, height: 100 });
const placeholderFresh = await generatePptx(placeholderSession.document), placeholderAgain = await importPptx(placeholderFresh.bytes, { projection: "content" });
assert.equal(placeholderAgain.document.slides[0].elements.length, 1);
assert.deepEqual(placeholderAgain.document.slides[0].elements[0].box, { x: 10, y: 20, width: 200, height: 100 });
assert.equal((await placeholderSession.prepareSave({ mode: "edit" })).status, "blocked");
placeholderAgain.dispose(); placeholderSession.dispose();

const textInheritanceSession = await importPptx(new Uint8Array(await readFile(path.join(directory, "text-inheritance.pptx"))), { projection: "content" });
const inheritedText = normalizePresentation(textInheritanceSession.document).slides[0].elements[0].text;
assert.equal(inheritedText.frame.margin.left, 5);
assert.equal(inheritedText.content[0].runs[0].style.fontSize, 20);
assert.equal(inheritedText.content[0].runs[0].style.color, "13579B");
assert.equal(inheritedText.content[0].runs[0].style.fontFamily, "Installed Theme Font");
assert.ok(textInheritanceSession.report.textStyleInheritance.properties.some(property => property.property === "run.fontSize" && property.level === "layout-placeholder"));
assert.ok(textInheritanceSession.report.textStyleInheritance.properties.some(property => property.property === "run.color" && property.resolvedFrom?.some(source => source.part === "/deck/theme.xml")));
assert.ok(textInheritanceSession.report.textStyleInheritance.properties.some(property => property.property === "run.fontFamily" && property.resolvedFrom?.some(source => source.part === "/deck/font-override.xml")));
assert.ok(textInheritanceSession.report.textStyleInheritance.properties.every(property => Object.isFrozen(property) && Object.isFrozen(property.source)));
const textInheritanceOutput = await generatePptx(textInheritanceSession.document);
const textInheritanceAgain = await importPptx(textInheritanceOutput.bytes, { projection: "content" });
assert.equal(normalizePresentation(textInheritanceAgain.document).slides[0].elements[0].text.content[0].runs[0].style.fontSize, 20);
assert.equal((await textInheritanceSession.prepareSave({ mode: "edit" })).status, "blocked");
textInheritanceAgain.dispose(); textInheritanceSession.dispose();

const editingDocument = createPresentation();
editingDocument.addSlide().addElement({ type: "text", box: { x: 10, y: 10, width: 200, height: 50 }, content: "Installed source text" });
const editingInput = await generatePptx(editingDocument);
const editingSession = await importPptx(editingInput.bytes, { projection: "content", editSaveProfile: "text-runs-v1" });
assert.equal(editingSession.report.textRuns[0].status, "bound");
const editingSlide = editingSession.document.slides[0], editingElement = structuredClone(editingSlide.removeElement(editingSlide.elements[0].id));
editingElement.content[0].runs[0].text = "Installed preserved text <&🙂"; editingSlide.insertElement(0, editingElement);
const editingPlan = await editingSession.prepareSave({ mode: "edit" }); assert.equal(editingPlan.status, "ready");
assert.equal(editingPlan.plan.report.assurance, "scoped-preservation"); assert.equal(editingPlan.plan.report.changes[0].kind, "text");
const editingSaved = await editingSession.save(editingPlan.plan), editingReopened = await importPptx(editingSaved.bytes, { projection: "content" });
assert.deepEqual(normalizePresentation(editingReopened.document), normalizePresentation(editingSession.document));
editingReopened.dispose(); editingSession.dispose();

const ownerDocument = createPresentation(), ownerSlide = ownerDocument.addSlide();
ownerSlide.addElement({ type: "shape", shape: "ellipse", box: { x: 20, y: 20, width: 300, height: 80 }, text: { content: "Installed shape source" } });
ownerSlide.setNotes("Installed notes source");
const ownerSession = await importPptx((await generatePptx(ownerDocument)).bytes, { projection: "content", editSaveProfile: "text-runs-v2" });
assert.equal(ownerSession.report.textRuns[0].owner, "shape-text");
assert.equal(ownerSession.report.notesTextRuns[0].status, "bound");
assert.equal("elementId" in ownerSession.report.notesTextRuns[0], false);
assert.ok(Object.isFrozen(ownerSession.report.notesTextRuns[0].source));
const ownerImportedSlide = ownerSession.document.slides[0], ownerShape = structuredClone(ownerImportedSlide.removeElement(ownerImportedSlide.elements[0].id));
ownerShape.text.content[0].runs[0].text = "Installed shape edit <&🙂"; ownerImportedSlide.insertElement(0, ownerShape);
const ownerNotes = structuredClone(ownerImportedSlide.notes); ownerNotes[0].runs[0].text = "Installed notes edit <&🙂"; ownerImportedSlide.setNotes(ownerNotes);
const ownerPlan = await ownerSession.prepareSave({ mode: "edit" }); assert.equal(ownerPlan.status, "ready");
const ownerSaved = await ownerSession.save(ownerPlan.plan), ownerReopened = await importPptx(ownerSaved.bytes, { projection: "content", editSaveProfile: "text-runs-v2" });
assert.equal(ownerSaved.report.changes.length, 2);
assert.deepEqual(normalizePresentation(ownerReopened.document), normalizePresentation(ownerSession.document));
ownerReopened.dispose(); ownerSession.dispose();

ownerSlide.addElement({ type: "table", box: { x: 20, y: 150, width: 200, height: 60 }, columns: [100, 100], rows: [
  { height: 30, cells: [{ content: "Installed merged source", rowSpan: 2 }, { content: "Installed top control" }] },
  { height: 30, cells: [{ content: "Installed bottom source" }] },
] });
const tableSession = await importPptx((await generatePptx(ownerDocument)).bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
assert.deepEqual(tableSession.report.tableTextRuns.map(run => [run.rowIndex, run.cellIndex, run.columnIndex]), [[0, 0, 0], [0, 1, 1], [1, 0, 1]]);
assert.ok(Object.isFrozen(tableSession.report.tableTextRuns[0]));
const tableSlide = tableSession.document.slides[0], tableOriginal = tableSlide.elements.find(e => e.type === "table"), tableIndex = tableSlide.elements.indexOf(tableOriginal), tableCopy = structuredClone(tableOriginal);
tableCopy.rows[1].cells[0].content[0].runs[0].text = "Installed compact cell edit <&🙂";
tableSlide.removeElement(tableOriginal.id); tableSlide.insertElement(tableIndex, tableCopy);
const tablePlan = await tableSession.prepareSave({ mode: "edit" }); assert.equal(tablePlan.status, "ready");
const tableSaved = await tableSession.save(tablePlan.plan), tableReopened = await importPptx(tableSaved.bytes, { projection: "content" });
assert.deepEqual(normalizePresentation(tableReopened.document), normalizePresentation(tableSession.document));
tableReopened.dispose(); tableSession.dispose();

const fileInput = path.join(directory, "installed-file-input.pptx"), fileCopy = path.join(directory, "installed-file-copy.pptx");
await writeFile(fileInput, editingInput.bytes);
const importedFile = await importPptxFile(fileInput, { projection: "content", editSaveProfile: "text-runs-v1" });
const filePlan = await importedFile.session.prepareSave({ mode: "no-op" }); assert.equal(filePlan.status, "ready");
assert.equal(importedFile.session.assertSavePlanCurrent(filePlan.plan), undefined);
const fileWritten = await writePptxFile(importedFile.session, filePlan.plan, fileCopy, { mode: "create" });
assert.equal(fileWritten.status, "committed"); assert.deepEqual(await readFile(fileCopy), Buffer.from(editingInput.bytes));
await assert.rejects(writePptxFile(importedFile.session, filePlan.plan, fileCopy, { mode: "create" }), error => error instanceof PptxFileError && error.code === "file-target-exists");
const fileReplaced = await writePptxFile(importedFile.session, filePlan.plan, fileInput, { mode: "replace", expected: importedFile.file });
assert.equal(fileWritten.file.sha256, fileReplaced.file.sha256);
importedFile.session.dispose();

const transitionDeck = createPresentation();
transitionDeck.addSlide({ transition: { type: "fade", durationMs: 1234, afterMs: 0, onClick: false } });
const transitionResult = await generatePptx(transitionDeck);
const transitionSession = await importPptx(transitionResult.bytes, { projection: "content" });
assert.deepEqual(transitionSession.document.slides[0].transition, transitionDeck.slides[0].transition);
transitionSession.dispose();

for (const type of ["push", "wipe"]) for (const direction of ["left", "right", "up", "down"]) {
 const d = createPresentation(); d.addSlide({ transition: { type, direction, durationMs: 765 } });
 const s = await importPptx((await generatePptx(d)).bytes, { projection: "content" });
 assert.equal(s.document.slides[0].transition.direction, direction); assert.equal(s.document.slides[0].transition.type, type); s.dispose();
}

const comparedText = (await inspectPptx(await readFile(path.join(directory, "smartart-text.pptx")), { diagramTextComparison: "direct-paragraphs" })).diagramInventory.textComparisons[0];
assert.equal(comparedText.shapes[0].status, "equal"); assert.equal(comparedText.freshness, "unknown");

const { describeImportCapabilities } = await import("@pptkit/pptx-parser");
const capabilitySession = await importPptx(await readFile(path.join(directory, "smartart-text.pptx")), { projection: "content", diagramPreview: "cache" });
try {
  const capabilities = describeImportCapabilities(capabilitySession.report, { filter: { feature: "smartart.cache-preview" } });
  assert.equal(capabilities.schemaVersion, "pptkit-import-capabilities-v1"); assert.equal(capabilities.subjects[0].operations[6].tier, "unsupported");
} finally { capabilitySession.dispose(); }

const chartValueSession = await importPptx(await readFile(path.join(directory, "chart-value-source.pptx")), { projection: "content", chartEditSaveProfile: "existing-values-v1" });
try {
  assert.ok(chartValueSession.report.chartValues.every(v => v.status === "bound"));
  const slide = chartValueSession.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].values[1] = 27; slide.removeElement(chart.id); slide.insertElement(0, chart);
  const prepared = await chartValueSession.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
  const saved = await chartValueSession.save(prepared.plan), reopened = await importPptx(saved.bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
  try { assert.equal(reopened.document.slides[0].elements[0].series[0].values[1], 27); assert.ok(reopened.report.chartValues.every(v => v.status === "bound")); } finally { reopened.dispose(); }
} finally { chartValueSession.dispose(); }

// Installed public finite-time API is independent of presentation mutations.
const { compileFiniteTimeline, sampleFiniteTimeline, FiniteTimelineError } = await import("@pptkit/core");
const timeline = compileFiniteTimeline({ kind: "sequence", id: "root", children: [{ kind: "span", id: "a", durationMs: 10, repeat: 2, fill: "hold" }, { kind: "span", id: "zero", durationMs: 0 }] });
assert.equal(timeline.durationMs, 20);
assert.deepEqual(sampleFiniteTimeline(JSON.parse(JSON.stringify(timeline)), 20), [{ id: "a", phase: "ended", progress: 1, iteration: 1 }, { id: "zero", phase: "ended", progress: null, iteration: 0 }]);
assert.throws(() => compileFiniteTimeline({ kind: "span", id: "x", durationMs: 1, autoReverse: true }), error => error instanceof FiniteTimelineError && error.code === "unsupported-timeline-feature");
const browserFinite = await import("@pptkit/core/browser"); assert.deepEqual(browserFinite.sampleFiniteTimeline(timeline, 10), sampleFiniteTimeline(timeline, 10));

const { compileEventTimeline, replayEventTimeline, EventTimelineError } = await import("@pptkit/core");
const gated = compileEventTimeline({ pageId: "p", elementIds: ["s", "button"], stages: [{ id: "a", target: { pageId: "p", elementId: "s" }, trigger: { type: "object-click", elementId: "button" }, timing: { kind: "span", id: "t", durationMs: 10 } }] });
const eventRequest = { visitId: "v", timeMs: 7, elementIds: ["s", "button"], events: [{ id: "click", pageId: "p", visitId: "v", timeMs: 2, type: "object-click", elementId: "button" }] };
const gatedResult = replayEventTimeline(JSON.parse(JSON.stringify(gated)), eventRequest);
assert.equal(gatedResult.stages[0].samples[0].progress, .5); assert.equal(gatedResult.events[0].status, "consumed");
assert.deepEqual(browserFinite.replayEventTimeline(gated, eventRequest), gatedResult);
assert.throws(() => replayEventTimeline(gated, { ...eventRequest, elementIds: ["s"] }), error => error instanceof EventTimelineError && error.code === "invalid-event-timeline-input");

const boundDocument = createPresentation(), boundSlide = boundDocument.addSlide({ id: "bound-page" });
boundSlide.addElement({ id: "bound-text", type: "text", box: { x: 0, y: 0, width: 100, height: 40 }, content: "Bound" });
boundSlide.setEventTimeline({ stages: [{ id: "stage", target: { pageId: boundSlide.id, elementId: "bound-text" }, trigger: { type: "enter" }, timing: { kind: "span", id: "t", durationMs: 1 } }] });
const boundCopy = boundDocument.duplicateSlide(boundSlide.id);
assert.equal(boundCopy.eventTimeline.stages[0].target.elementId, boundCopy.elements[0].id);
assert.deepEqual(resolveLayout(boundDocument).slides[0].eventTimeline.plan, boundSlide.getEventTimelinePlan());
assert.throws(() => boundSlide.removeElement("bound-text"), EventTimelineError);
boundSlide.replaceElement("bound-text", { type: "text", box: { x: 0, y: 0, width: 100, height: 40 }, content: "Replaced" });
const { PptxExportError } = await import("@pptkit/pptx-exporter");
await assert.rejects(generatePptx(boundDocument), PptxExportError);
const boundStatic = await generatePptx(boundDocument, { eventTimelineMode: "omit" }); assert.equal(boundStatic.warnings.length, 2);

const { inspectPptxTiming } = await import("@pptkit/pptx-parser");
const nativeTiming = await inspectPptxTiming(await readFile(path.join(directory, "native-timing-source.pptx")));
assert.equal(nativeTiming.schemaVersion, "pptkit-native-timing-inventory-v1");
assert.equal(nativeTiming.slides[0].nodes.find(n => n.role === "drawing-target").drawingReference.status, "located");
assert.equal(nativeTiming.authoredMapping, "not-performed"); assert.equal(nativeTiming.nativePlayback, "not-performed");

const paragraphSession = await importPptx(await readFile(path.join(directory, "paragraph-target-source.pptx")), { projection: "content", editSaveProfile: "text-runs-v1", textTargetSaveProfile: "paragraph-stable-v1" });
try {
  const slide = paragraphSession.document.slides[0], element = structuredClone(slide.elements[0]); element.content[2].runs[0].text = "Installed paragraph edit👩‍💻"; slide.replaceElement(element.id, element);
  const prepared = await paragraphSession.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
  assert.deepEqual(await inspectPptxTiming((await paragraphSession.save(prepared.plan)).bytes), await inspectPptxTiming(await readFile(path.join(directory, "paragraph-target-source.pptx"))));
} finally { paragraphSession.dispose(); }

const { editTable, TableEditError } = await import("@pptkit/core");
const editableTable = { id: "editable", type: "table", box: { x: 0, y: 0, width: 80, height: 20 }, columns: [40, 40], rows: [{ height: 20, cells: [{ content: "Anchor" }, { content: "" }] }] };
const tableTransaction = editTable(editableTable, [{ type: "merge", region: { rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 2 } }, { type: "insert-row", index: 1, height: 20 }], { frame: "fit-grid" });
assert.equal(tableTransaction.table.rows[0].cells[0].colSpan, 2); assert.equal(tableTransaction.table.rows.length, 2);
assert.equal(tableTransaction.cellMapping[1].status, "merged"); assert.equal(editableTable.rows.length, 1);
assert.throws(() => editTable(editableTable, [{ type: "delete-row", index: 0 }], { frame: "preserve" }), TableEditError);

const { resolveTableBorders, TableBorderError } = await import("@pptkit/layout");
const borderDocument = createPresentation(); borderDocument.addSlide().addElement({ type: "table", box: { x: 0, y: 0, width: 80, height: 20 }, columns: [40, 40], rows: [{ height: 20, cells: [{ content: "A", style: { borders: { right: { paint: { type: "solid", color: "FF0000" }, width: 3 } } } }, { content: "B" }] }] });
const borderSource = normalizePresentation(borderDocument).slides[0].elements[0], borderLayout = resolveTableBorders(borderSource);
assert.equal(borderLayout.grid.vertical[1].paint.color, "FF0000"); assert.deepEqual(resolveLayout(borderDocument).slides[0].elements[0].borderGrid, borderLayout.grid);
assert.throws(() => resolveTableBorders(borderSource, { maxGridSlots: 1 }), TableBorderError);


const regionDocument = createPresentation();
regionDocument.addSlide().addElement({ type: "table", box: { x:0,y:0,width:80,height:40 }, columns:[40,40], style: { wholeTable: { text: { run: { bold:true } } }, bandedRows: { first: { fill: { type:"solid",color:"AA0000" } }, second: { fill: { type:"solid",color:"00AA00" } } } }, rows:[{height:20,cells:[{content:"First"}]},{height:20,cells:[]}] });
const regionTable = normalizePresentation(regionDocument).slides[0].elements[0];
assert.equal(regionTable.rows[0].cells.length,2); assert.equal(regionTable.rows[1].cells[0].style.fill.color,"00AA00"); assert.equal(regionTable.rows[0].cells[0].content[0].runs[0].style.bold,true);
assert.ok((await generatePptx(regionDocument)).warnings.some(w=>w.code === "table-style-materialized"));

const styledSource = await importPptx(await readFile(path.join(directory,"table-style-source.pptx")),{projection:"content",sessionLimits:{maxTableStyleWork:2000000}});
assert.equal(styledSource.report.tableStyleInheritance.tables[0].status,"resolved");
assert.equal(normalizePresentation(styledSource.document).slides[0].elements[0].rows[0].cells[0].style.fill.color,"3579BD");
assert.ok(Object.isFrozen(styledSource.report.tableStyleInheritance.properties));
styledSource.dispose();

const matrixSource=await importPptx(await readFile(path.join(directory,"theme-matrix-source.pptx")),{projection:"content",editSaveProfile:"text-runs-v3"});
assert.equal(normalizePresentation(matrixSource.document).slides[0].elements[0].rows[0].cells[0].style.fill.color,"3579BD");
assert.ok(matrixSource.report.tableStyleInheritance.properties.some(p=>p.resolvedFrom?.some(d=>d.part==="/deck/theme.xml"&&d.path.includes("lnStyleLst"))));
matrixSource.dispose();

const appearanceSource = await importPptx(await readFile(path.join(directory, "table-text-appearance.pptx")), { projection: "content" });
const appearanceRun = normalizePresentation(appearanceSource.document).slides[0].elements[0].rows[0].cells[0].content[0].runs[0];
assert.equal(appearanceRun.style.fontFamily, "Arial"); assert.equal(appearanceRun.style.color, "102030");
assert.ok(appearanceSource.report.tableStyleInheritance.properties.find(p => p.property === "run.fontFamily").resolvedFrom.some(d => d.path.includes("ea[1]")));
appearanceSource.dispose();

const scriptDocument = createPresentation();
scriptDocument.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 250, height: 60 }, content: [{ runs: [{ text: "A中ع", style: { fontFamily: "Latin", eastAsianFontFamily: "Eastern", complexScriptFontFamily: "Complex" } }] }] });
const scriptCore = await import("@pptkit/core");
assert.ok(Object.isFrozen(scriptCore.TEXT_FONT_SCRIPT_CODES));
assert.deepEqual(scriptCore.resolveTextFontSegments("中عא", { fontFamily: "Default", language: "zh-TW", scriptFontFamilies: { Hant: "Traditional", Arab: "Arabic", Hebr: "Hebrew" } }, normalizePresentation(scriptDocument).theme).map(s => [s.family, s.scriptFont]), [["Traditional", "Hant"], ["Arabic", "Arab"], ["Hebrew", "Hebr"]]);
const scriptNormalized = normalizePresentation(scriptDocument);
assert.deepEqual(scriptCore.resolveTextFontSegments("A中ع", scriptNormalized.slides[0].elements[0].content[0].runs[0].style, scriptNormalized.theme).map(s => s.family), ["Latin", "Eastern", "Complex"]);
const namespacedPreview = await renderPresentationToSvg(scriptDocument, { idNamespace: "installed-consumer" });
assert.match(namespacedPreview.slides[0].svg, /id="pptkit-ns-/);
const scriptPreview = (await renderPresentationToSvg(scriptDocument)).slides[0].svg;
assert.ok(scriptPreview.includes("<foreignObject"));
assert.ok(scriptPreview.includes('font-family:&quot;Eastern&quot;'));
assert.ok(scriptPreview.includes('font-family:&quot;Complex&quot;'));
assert.ok((await generatePptx(scriptDocument)).bytes.length > 100);
const scriptImported = await importPptx((await generatePptx(scriptDocument)).bytes, { projection: "content", textFontProfile: "script-families-v1", editSaveProfile: "text-runs-v1" });
assert.equal(scriptImported.report.textFontProfile, "script-families-v1");
assert.equal(scriptImported.document.slides[0].elements[0].content[0].runs[0].style.eastAsianFontFamily, "Eastern");
assert.ok(scriptImported.report.textStyleInheritance.properties.some(p => p.property === "run.complexScriptFontFamily" && p.status === "mapped"));
const scriptPage = scriptImported.document.slides[0], scriptNext = structuredClone(scriptPage.elements[0]); scriptNext.content[0].runs[0].text = "Saved 中文ع"; scriptPage.replaceElement(scriptNext.id, scriptNext);
const scriptReady = await scriptImported.prepareSave({ mode: "edit" }); assert.equal(scriptReady.status, "ready"); assert.ok((await scriptImported.save(scriptReady.plan)).bytes.length > 100); scriptImported.dispose();

const supplementalImported = await importPptx(new Uint8Array(await readFile("supplemental-font-source.pptx")), { projection: "content", textFontProfile: "script-families-v2", editSaveProfile: "text-runs-v3" });
try {
  assert.equal(supplementalImported.report.textFontProfile, "script-families-v2");
  const page = supplementalImported.document.slides[0], next = structuredClone(page.elements[0]);
  assert.equal(next.text.content[0].runs[0].style.scriptFontFamilies.Hant, "Traditional Source");
  assert.ok(supplementalImported.report.textStyleInheritance.properties.some(p => p.property === "run.scriptFontFamilies.Arab" && p.source.part === "/deck/theme.xml"));
  next.text.content[0].runs[0].text = "Installed 中عא"; page.replaceElement(next.id, next);
  const prepared = await supplementalImported.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
  assert.ok((await supplementalImported.save(prepared.plan)).bytes.length > 100);
} finally { supplementalImported.dispose(); }

const declaredSlotImported = await importPptxFile("declared-slot-font-source.pptx", { projection: "content", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" });
try {
  const session = declaredSlotImported.session, slide = session.document.slides[0], next = structuredClone(slide.elements[0]);
  assert.equal(session.report.textFontProfile, "script-families-v3");
  assert.deepEqual(next.text.content[0].runs[0].style.scriptFontFamilies, { Grek: "Greek Source" });
  assert.ok(session.report.textStyleInheritance.properties.some(p => p.property === "run.scriptFontFamilies.Hant" && p.status === "unsupported" && p.value === undefined));
  next.text.content[0].runs[0].text = "Installed v3 中عאΩ"; slide.replaceElement(next.id, next);
  const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
  assert.ok((await session.save(prepared.plan)).bytes.length > 100);
} finally { declaredSlotImported.session.dispose(); }

const inheritedNotes = await importPptxFile("notes-style-source.pptx", { projection: "content", notesTextProfile: "master-style-v1", textFontProfile: "script-families-v2", editSaveProfile: "text-runs-v3" });
try {
  const s = inheritedNotes.session; assert.equal(s.report.notesTextProfile, "master-style-v1");
  const slide = s.document.slides[0], notes = structuredClone(slide.notes); assert.equal(notes[0].runs[0].style.fontSize, 22); assert.equal(notes[0].runs[0].style.scriptFontFamilies.Hant, "Notes Traditional");
  notes[0].runs[0].text = "Installed notes 中عא"; slide.setNotes(notes);
  const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); assert.ok((await s.save(plan.plan)).bytes.length > 100);
} finally { inheritedNotes.session.dispose(); }

const { inspectPptxNotes } = await import("@pptkit/pptx-parser");
const notesInventory = await inspectPptxNotes(await readFile(path.join(directory, "notes-style-source.pptx")));
assert.equal(notesInventory.schemaVersion, "pptkit-notes-source-inventory-v1");
assert.equal(notesInventory.pages[0].body.status, "unique-direct");
assert.equal(notesInventory.authoredMapping, "not-performed");
assert.ok(Object.isFrozen(notesInventory.parts[0].drawings));

const multiMasterConsumer = createPresentation({ masters: [{ id: "installed-master", name: "Installed master", theme: { colors: { accent1: "ABCDEF" } } }] });
multiMasterConsumer.defineSlideLayout({ id: "installed-layout", name: "Installed layout", masterId: "installed-master" });
multiMasterConsumer.addSlide({ layoutId: "installed-layout" });
assert.equal((await inspectPptx((await generatePptx(multiMasterConsumer)).bytes)).slideMasters.length, 2);
multiMasterConsumer.defineSlideMaster({ id: "new-master", name: "New master", theme: { colors: { accent1: "118855" } } });
multiMasterConsumer.replaceSlideMaster("installed-master", { name: "Resume defaults" });
multiMasterConsumer.replaceSlideLayout("installed-layout", { name: "Rebound", masterId: "new-master", background: { type: "none" } });
assert.equal(normalizePresentation(multiMasterConsumer).layouts.find(l => l.id === "installed-layout").theme.colors.accent1, "118855");
assert.equal(normalizePresentation(multiMasterConsumer).slides[0].background.type, "none");
assert.equal((await inspectPptx((await generatePptx(multiMasterConsumer)).bytes)).slideMasters.length, 3);
const inheritance = resolvePresentationInheritance(multiMasterConsumer);
assert.deepEqual(inheritance.presentation, normalizePresentation(multiMasterConsumer));
assert.ok(Object.isFrozen(inheritance.origins));
assert.deepEqual(inheritance.origins.find(r => r.target.kind === "layout" && r.target.id === "installed-layout" && r.property === "theme.colors.accent1").source,
  { kind: "master", id: "new-master", path: "masters.1.theme.colors.accent1" });
multiMasterConsumer.slides[0].addElement({ id: "format-probe", type: "text", box: { x: 0, y: 0, width: 100, height: 30 }, frame: { wrap: false }, content: [{ runs: [{ text: "", style: { bold: false, eastAsianFontFamily: null } }] }] });
const formats = resolvePresentationInheritance(multiMasterConsumer).origins.filter(r => r.target.kind === "element" && r.target.id === "format-probe");
assert.equal(formats.find(r => r.property === "content.0.runs.0.style.eastAsianFontFamily").state, "cleared");
assert.deepEqual(formats.find(r => r.property === "frame.wrap").source, { kind: "element", id: "format-probe", path: "slides.0.elements.0.frame.wrap" });

multiMasterConsumer.slides[0].setNotes([{ runs: [{ text: "", style: { complexScriptFontFamily: null } }] }]);
multiMasterConsumer.slides[0].addElement({ id: "chart-format-probe", type: "chart", chartType: "pie", box: { x: 10, y: 10, width: 300, height: 200 }, categories: ["One"], series: [{ name: "Only", values: [1] }], titleStyle: { bold: false } });
const chartNotesOrigins = resolvePresentationInheritance(multiMasterConsumer).origins;
assert.equal(chartNotesOrigins.find(r => r.target.kind === "notes" && r.property === "notes.0.runs.0.style.complexScriptFontFamily").state, "cleared");
assert.deepEqual(chartNotesOrigins.find(r => r.target.kind === "element" && r.target.id === "chart-format-probe" && r.property === "titleStyle.bold").source, { kind: "element", id: "chart-format-probe", path: "slides.0.elements.1.titleStyle.bold" });

const regionalOrigins = resolvePresentationInheritance(regionDocument).origins;
assert.deepEqual(regionalOrigins.find(r => r.property === "rows.1.cells.0.style.fill.color").source, { kind: "element", id: regionDocument.slides[0].elements[0].id, path: "slides.0.elements.0.style.bandedRows.second.fill.color" });
assert.deepEqual(regionalOrigins.find(r => r.property === "rows.1.cells.1.content.0.runs.0.style.bold").source, { kind: "element", id: regionDocument.slides[0].elements[0].id, path: "slides.0.elements.0.style.wholeTable.text.run.bold" });

const ownedArtConsumer = createPresentation();
ownedArtConsumer.defineSlideMaster({ id: "owned-art", name: "Artwork", elements: [{ id: "owned-art-shape", type: "shape", shape: "rect", box: { x: 0, y: 0, width: 30, height: 30 }, style: { fill: { type: "solid", color: "112233" } } }] });
ownedArtConsumer.defineSlideLayout({ id: "owned-layout", name: "Art", masterId: "owned-art", showMasterShapes: true });
const ownedArtSlide = ownedArtConsumer.addSlide({ layoutId: "owned-layout", showInheritedShapes: false });
ownedArtSlide.setShowInheritedShapes(true); assert.equal(ownedArtSlide.showInheritedShapes, true);
assert.equal(normalizePresentation(ownedArtConsumer).masters[0].elements[0].id, "owned-art-shape");
assert.match((await renderPresentationToSvg(ownedArtConsumer)).slides[0].svg, /#112233/);
assert.equal((await inspectPptx((await generatePptx(ownedArtConsumer)).bytes)).slideMasters.length, 2);


const rebindConsumer = createPresentation();
for (const [id, x] of [["before", 0], ["after", 100]]) rebindConsumer.defineSlideLayout({ id, name: id, placeholders: [{ key: id, kind: "body", box: { x, y: 10, width: 200, height: 50 } }] });
const rebindSlide = rebindConsumer.addSlide({ layoutId: "before", elements: [{ id: "rebind-content", type: "text", placeholderKey: "before", content: "Retained" }] });
const rebindPlan = prepareLayoutRebind(rebindConsumer, [{ slideId: rebindSlide.id, layoutId: "after" }], { localAppearance: "adopt" });
assert.equal(rebindPlan.status, "ready"); assert.equal(rebindPlan.matches[0].method, "kind");
assert.equal(rebindConsumer.applyLayoutRebind(rebindPlan).status, "applied"); assert.equal(rebindSlide.layoutId, "after");
assert.equal(normalizePresentation(rebindConsumer).slides[0].elements[0].box.x, 100);
assert.equal((await inspectPptx((await generatePptx(rebindConsumer)).bytes)).slides.length, 1);
const indexConsumer=createPresentation(),indexBox={x:0,y:0,width:100,height:50};
indexConsumer.defineSlideLayout({id:'indexed-source',name:'Source',placeholders:[{key:'old',kind:'body',index:0,box:indexBox}]});
indexConsumer.defineSlideLayout({id:'indexed-target',name:'Target',placeholders:[{key:'decoy',kind:'body',index:9,box:indexBox},{key:'new',kind:'body',index:0,box:indexBox}]});
const indexSlide=indexConsumer.addSlide({layoutId:'indexed-source',elements:[{type:'text',placeholderKey:'old',content:'Index'}]});
const indexPlan=prepareLayoutRebind(indexConsumer,[{slideId:indexSlide.id,layoutId:'indexed-target'}],{localAppearance:'adopt'});assert.equal(indexPlan.matches[0].method,'index');assert.equal(indexPlan.matches[0].sourceIndex,0);assert.equal(indexPlan.matches[0].targetIndex,0);indexConsumer.applyLayoutRebind(indexPlan);assert.equal(indexSlide.elements[0].placeholderKey,'new');

const installedArea = createPresentation();
installedArea.addSlide().addElement({type:'chart',chartType:'area',grouping:'stacked',box:{x:0,y:0,width:300,height:200},categories:['A','B'],series:[{name:'S',values:[1,null]}]});
const installedAreaSession = await importPptx((await generatePptx(installedArea)).bytes,{projection:'content'});
try { assert.equal(installedAreaSession.document.slides[0].elements[0].chartType,'area'); assert.deepEqual(installedAreaSession.document.slides[0].elements[0].series[0].values,[1,null]); } finally { installedAreaSession.dispose(); }

const installedScatterDeck=createPresentation();
installedScatterDeck.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:300,height:200},series:[{name:'S',xValues:[1,null],yValues:[2,3]}]});
assert.equal(normalizePresentation(installedScatterDeck).slides[0].elements[0].chartType,'scatter');
assert.match((await renderPresentationToSvg(installedScatterDeck)).slides[0].svg,/data-chart-role="scatter-marker"/);
assert.ok((await generatePptx(installedScatterDeck)).bytes.length>100);
