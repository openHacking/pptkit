import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { chromium } from "@playwright/test";
import { resolveCycleLayout } from "./research/cycle-layout.mjs";
import { createDiagramCacheSession } from "./research/diagram-cache-session.mjs";
const digest = data => createHash("sha256").update(data).digest("hex");
const encode = text => new TextEncoder().encode(text);
const escape = text => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;");

test("owned Cycle geometry SVG cache validates actual payload/dependencies, rejects stale work and never trusts imported cache bytes", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage();
  await page.route("**/*", route => route.abort());
  const versions = new Map(), colors = new Map();
  const candidate = suffix => ({ family: "cycle", closure: "layout-generated", roots: ["a", "b", "c", "d"], nodes: ["a", "b", "c", "d"].map(id => ({ id, role: "content", children: [], paragraphs: [id + suffix] })) });
  const snapshot = (suffix, color = "#4266B1") => {
    const data = candidate(suffix), dataDigest = digest(JSON.stringify(data)), colorDigest = digest(color); versions.set(dataDigest, structuredClone(data)); colors.set(colorDigest, color);
    return { ownerId: "owned-cycle/geometry-preview", layoutIdentifier: "urn:microsoft.com/office/officeart/2005/8/layout/cycle2", coverage: "complete", generatorVersion: "owned-geometry-svg-v1", environmentKey: "geometry-only/no-text/no-native-cache", frame: { width: 300, height: 300 }, dependencies: [
      { role: "data", id: "owned-content", sha256: dataDigest }, { role: "layout", id: "owned-layout", sha256: digest("equal-circle-layout-v1") }, { role: "colors", id: "owned-colors", sha256: colorDigest }, { role: "style", id: "owned-style", sha256: digest("fill-only") }, { role: "theme", id: "owned-theme", sha256: digest("explicit-color") },
    ] };
  };
  let validations = 0;
  const generate = ({ snapshot }) => {
    const data = versions.get(snapshot.dependencies.find(p => p.role === "data").sha256), color = colors.get(snapshot.dependencies.find(p => p.role === "colors").sha256);
    const result = resolveCycleLayout(data, { frame: { x: 0, y: 0, width: snapshot.frame.width, height: snapshot.frame.height } });
    assert.equal(result.findings.length, 4, "geometry-only cache never promotes text fitting");
    const metadata = escape(JSON.stringify({ scope: "owned-geometry-only", snapshot }));
    return encode(`<svg xmlns="http://www.w3.org/2000/svg" width="300" height="300"><metadata>${metadata}</metadata>${result.edges.map(e => `<polygon data-source="${e.sourceId}" data-target="${e.targetId}" points="${e.polygon.map(p => `${p.x},${p.y}`).join(" ")}" fill="#9CB0D8"/>`).join("")}${result.nodes.map(n => `<circle data-node="${n.id}" cx="${n.center.x}" cy="${n.center.y}" r="${n.diameter / 2}" fill="${color}"/>`).join("")}</svg>`);
  };
  const validate = async ({ bytes, snapshot }) => {
    validations++;
    return page.evaluate(({ bytes, snapshot, color }) => {
      const document = new DOMParser().parseFromString(new TextDecoder().decode(Uint8Array.from(bytes)), "image/svg+xml");
      if (document.getElementsByTagName("parsererror").length || document.documentElement.namespaceURI !== "http://www.w3.org/2000/svg") return false;
      const metadata = document.querySelector("metadata"); if (!metadata) return false;
      const info = JSON.parse(metadata.textContent); if (info.scope !== "owned-geometry-only" || JSON.stringify(info.snapshot) !== JSON.stringify(snapshot)) return false;
      const circles = [...document.querySelectorAll("circle")], edges = [...document.querySelectorAll("polygon")]; if (circles.length !== 4 || edges.length !== 4) return false;
      // Independent square packing expectation; no call to the JS solver here.
      const diameter = 300 / (1 + 3 / Math.SQRT2), radius = 3 * diameter / (2 * Math.SQRT2), centers = [[150, 150 - radius], [150 + radius, 150], [150, 150 + radius], [150 - radius, 150]], ids = ["a", "b", "c", "d"];
      return circles.every((c, i) => c.getAttribute("data-node") === ids[i] && c.getAttribute("fill") === color && Math.abs(Number(c.getAttribute("cx")) - centers[i][0]) < 1e-8 && Math.abs(Number(c.getAttribute("cy")) - centers[i][1]) < 1e-8 && Math.abs(Number(c.getAttribute("r")) - diameter / 2) < 1e-8)
        && edges.every((e, i) => e.getAttribute("data-source") === ids[i] && e.getAttribute("data-target") === ids[(i + 1) % 4] && e.getAttribute("points").split(" ").length === 7);
    }, { bytes: [...bytes], snapshot, color: colors.get(snapshot.dependencies.find(p => p.role === "colors").sha256) });
  };
  const initial = snapshot("-old"), session = createDiagramCacheSession(initial, { sourceCache: generate({ snapshot: initial }) });
  assert.equal(session.inspect().state, "unknown"); assert.equal((await session.prepare({ policy: "block" })).status, "blocked");
  const first = await session.prepare({ policy: "rebuild", generate, validate }); const oldBytes = session.read(first); assert.equal(session.inspect().state, "unknown"); session.commit(first);
  assert.equal(session.inspect().state, "managed-current");
  const changed = snapshot("-new", "#087F8C"); session.update(changed); assert.equal(session.inspect().state, "invalidated");
  await assert.rejects(() => session.prepare({ policy: "rebuild", generate: () => oldBytes, validate }), { code: "validation-failed" }); assert.equal(session.inspect().state, "invalidated");
  let resolve, child; const gate = new Promise(r => { resolve = r; });
  const pending = session.prepare({ policy: "rebuild", generate: ({ snapshot, signal }) => { child = signal; return gate.then(() => generate({ snapshot })); }, validate });
  const beforeValidation = validations; const latest = snapshot("-latest", "#A06030"); session.update(latest); assert.equal(child.aborted, true); resolve(); await assert.rejects(pending, { code: "stale-plan" }); assert.equal(validations, beforeValidation);
  const replacement = await session.prepare({ policy: "rebuild", generate, validate }); const current = session.read(replacement); assert.notDeepEqual(current, oldBytes); session.commit(replacement);
  const reused = await session.prepare({ policy: "block" }); assert.deepEqual(session.read(reused), current); session.release(reused);
  const sourceAgain = createDiagramCacheSession(latest, { sourceCache: current }); assert.equal(sourceAgain.inspect().state, "unknown", "serialized bytes do not serialize the session witness"); sourceAgain.dispose(); session.dispose();
});
