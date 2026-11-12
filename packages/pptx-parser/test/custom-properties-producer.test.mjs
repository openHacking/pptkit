import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { importPptx, inspectPptx } from "../dist/index.js";
import { readOpcSource } from "../dist/opc/read-package.js";

const directory = new URL("./fixtures/custom-properties-libreoffice/", import.meta.url);
const expectedAfter = [
  ["Boolean", 2, "bool", "boolean", "0"],
  ["Date", 3, "filetime", "date", "2020-01-02T03:04:05Z"],
  ["Integer", 4, "r8", "real", "-2147483648"],
  ["PresentationFormat", 5, "lpwstr", "string", "On-screen Show (16:9)"],
  ["Real", 6, "r8", "real", "1.25"],
  ["Slides", 7, "r8", "real", "1"],
  ["Text", 8, "lpwstr", "string", "中文 & text"],
];
const expectedSource = [
  ["Text", 2, "lpwstr", "string", "中文 & text"],
  ["Boolean", 3, "bool", "boolean", "false"],
  ["Integer", 4, "i4", "integer", "-2147483648"],
  ["Real", 5, "r8", "real", "1.25"],
  ["Date", 6, "filetime", "date", "2020-01-02T03:04:05Z"],
];

test("independent LibreOfficeDev custom-property samples preserve actual types and identities without source rewriting", async () => {
  const manifest = JSON.parse(await readFile(new URL("provenance.json", directory), "utf8"));
  assert.match(manifest.application, /^LibreOfficeDev 26\.8\.0\.0\.alpha0/);
  assert.equal(manifest.records.length, 4);
  for (const record of manifest.records) {
    const source = await readFile(new URL(`round-${record.round}.pptx`, directory));
    assert.equal(createHash("sha256").update(source).digest("hex"), record.sha256);
    const inspected = await inspectPptx(source);
    const properties = inspected.metadata.custom.properties;
    assert.deepEqual(properties.map(p => [p.name, p.id, p.valueType, p.value.kind, p.value.value]), record.round === 0 ? expectedSource : expectedAfter);
    for (const projection of ["structure", "content"]) {
      const session = await importPptx(source, { projection });
      try {
        const findings = session.report.findings.filter(f => f.feature === "metadata.custom-property");
        assert.equal(findings.length, properties.length);
        assert.deepEqual(session.document.metadata.customProperties.find(p => p.name === "Date"), { name: "Date", type: "date", value: "2020-01-02T03:04:05.000Z" });
        for (const property of properties) assert.ok(findings.some(f => f.source.sourceObjectId === `${property.formatId}:${property.id}` && f.tier === "editable"));
        const prepared = await session.prepareSave({ mode: "no-op" });
        assert.equal(prepared.status, "ready");
        assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), source);
      } finally { session.dispose(); }
    }
  }
});


test("producer-resaved values remain bound to the actual reordered source identities", async () => {
  for (const round of [1, 2, 3]) {
    const source = await readFile(new URL(`round-${round}.pptx`, directory));
    const session = await importPptx(source, { projection: "content", customPropertySaveProfile: "scalar-v1" });
    try {
      const original = session.document.metadata.customProperties;
      session.document.setMetadata({ ...session.document.metadata, customProperties: original.map(p => p.name === "Text" ? { ...p, value: `Changed ${round}` } : p.name === "Boolean" ? { ...p, value: true } : p) });
      const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
      const changed = (await session.save(prepared.plan)).bytes, before = await readOpcSource(source, {}, false, true), after = await readOpcSource(changed, {}, false, true);
      assert.deepEqual(before.inspection.parts.map(p => p.name).sort(), after.inspection.parts.map(p => p.name).sort());
      for (const part of before.inspection.parts) if (part.name !== "/docProps/custom.xml") assert.deepEqual(after.payloads.get(part.entryName), before.payloads.get(part.entryName), part.name);
      const reopened = await importPptx(changed, { projection: "content" });
      try {
        assert.equal(reopened.document.metadata.customProperties.find(p => p.name === "Text").value, `Changed ${round}`);
        assert.equal(reopened.document.metadata.customProperties.find(p => p.name === "Boolean").value, true);
        assert.equal(reopened.document.metadata.customProperties.find(p => p.name === "Integer").type, "real");
      } finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});


test("producer-resaved custom names and values remain bound to their reordered identities", async () => {
  for (const round of [1, 2, 3]) {
    const source = await readFile(new URL(`round-${round}.pptx`, directory));
    const session = await importPptx(source, { projection: "content", customPropertySaveProfile: "scalar-and-name-v2" });
    try {
      const original = session.document.metadata.customProperties;
      session.document.setMetadata({ ...session.document.metadata, customProperties: original.map(p => p.name === "Text" ? { ...p, name: `Renamed ${round}`, value: `Changed ${round}` } : p) });
      const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
      const changed = (await session.save(prepared.plan)).bytes, before = await readOpcSource(source, {}, false, true), after = await readOpcSource(changed, {}, false, true);
      for (const part of before.inspection.parts) if (part.name !== "/docProps/custom.xml") assert.deepEqual(after.payloads.get(part.entryName), before.payloads.get(part.entryName), part.name);
      const reopened = await importPptx(changed, { projection: "content" });
      try { const entry = reopened.document.metadata.customProperties.find(p => p.name === `Renamed ${round}`); assert.equal(entry.value, `Changed ${round}`); assert.equal(reopened.document.metadata.customProperties.some(p => p.name === "Text"), false); }
      finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});
