import { nativeTableStyleSourceBytes } from "./native-table-font-styles.mjs";
import { readGeneratedZip } from "../support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";
import { nativeFontFamilies as F } from "../../scripts/research/font-native-cases.mjs";

const xml = value => value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
const themeMaps = {
  minor: { Latn: F.greek, Hans: F.ja, Hant: F.sc, Jpan: F.tc, Hang: F.generic, Arab: F.generic, Hebr: F.generic, Grek: F.latin },
  major: { Latn: F.latin, Hans: F.tc, Hant: F.ja, Jpan: F.sc, Hang: F.generic, Arab: F.latin, Hebr: F.latin, Grek: F.greek },
};
const collection = map => `<a:latin typeface="${xml(map.Latn)}"/><a:ea typeface=""/><a:cs typeface=""/>${Object.entries(map).map(([script, family]) => `<a:font script="${script}" typeface="${xml(family)}"/>`).join("")}`;

/** Source-only theme relationship rebinding and local-override control, not native theme application. */
export async function nativeTableThemeSourceBytes({ filled = false, swapped = false, localOverrides = false } = {}) {
  const fixture = await nativeTableStyleSourceBytes("regular", filled), parts = readGeneratedZip(fixture.bytes), maps = structuredClone(fixture.maps);
  let themePart = "ppt/theme/theme1.xml";
  if (swapped) {
    themePart = "ppt/theme/table-rebound-theme.xml";
    let theme = parts.get("ppt/theme/theme1.xml").toString();
    for (const kind of ["major", "minor"]) {
      let count = 0;
      theme = theme.replace(new RegExp(`<a:${kind}Font>[\\s\\S]*?</a:${kind}Font>`), () => { count++; return `<a:${kind}Font>${collection(themeMaps[kind])}</a:${kind}Font>`; });
      if (count !== 1) throw Error("Expected exactly one theme font collection"); maps[kind] = structuredClone(themeMaps[kind]);
    }
    parts.set(themePart, Buffer.from(theme));
    const relPart = "ppt/slideMasters/_rels/slideMaster1.xml.rels"; let count = 0;
    const relationships = parts.get(relPart).toString().replace(/<Relationship\b[^>]*\/>/g, node => {
      if (!/Type="[^"]*\/theme"/.test(node)) return node;
      count++; return node.replace(/Target="[^"]*"/, 'Target="../theme/table-rebound-theme.xml"');
    });
    if (count !== 1) throw Error("Expected one actual master theme relationship");
    parts.set(relPart, Buffer.from(relationships));
    const types = parts.get("[Content_Types].xml").toString();
    if (!types.endsWith("</Types>")) throw Error("Expected complete content type source");
    parts.set("[Content_Types].xml", Buffer.from(types.replace("</Types>", `<Override PartName="/${themePart}" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`)));
  }
  if (localOverrides) for (const index of fixture.targets) {
    const name = `ppt/slides/slide${index}.xml`; let count = 0;
    const text = parts.get(name).toString().replace(/<a:tbl>[\s\S]*?<\/a:tbl>/, table => table.replace(/<a:rPr\b([^>]*)\/>/g, (_, attributes) => {
      count++; return `<a:rPr${attributes}><a:latin typeface="${xml(F.latin)}"/><a:ea typeface="${xml(F.generic)}"/><a:cs typeface="${xml(F.generic)}"/></a:rPr>`;
    }));
    if (count !== 2) throw Error("Expected two local override runs"); parts.set(name, Buffer.from(text));
  }
  const controls = fixture.controls.map((c, i) => ({ ...c, localOverride: localOverrides && fixture.targets.includes(i+1), targetFamily: localOverrides && fixture.targets.includes(i+1) ? F.generic : maps[c.kind][c.script] }));
  return { ...fixture, bytes: zip([...parts].map(([name, body]) => ({ name, body }))).bytes, maps, controls, themePart, localFonts: { latin: F.latin, ea: F.generic, cs: F.generic }, sourceOnly: true };
}
