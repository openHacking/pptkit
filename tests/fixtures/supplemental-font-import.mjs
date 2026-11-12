import { createPresentation } from "../../packages/core/dist/index.js";
import { generatePptx } from "../../packages/pptx-exporter/dist/index.js";
import { createScriptFontFixture } from "./script-fonts.mjs";
import { readGeneratedZip } from "../support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";
import { supplementalFonts } from "../../packages/pptx-parser/test/supplemental-font-fixture.mjs";

/** Source fixture, not an author-map round trip: replace native references before import. */
export async function mixedSupplementalBytes() {
  const parts = readGeneratedZip((await generatePptx(createScriptFontFixture(createPresentation))).bytes);
  parts.set("ppt/theme/theme1.xml", Buffer.from(parts.get("ppt/theme/theme1.xml").toString().replace('</a:minorFont>', `${supplementalFonts}</a:minorFont>`)));
  parts.set("ppt/slides/slide1.xml", Buffer.from(parts.get("ppt/slides/slide1.xml").toString().replaceAll('lang="en-US"', 'lang="zh-TW"').replaceAll('typeface="Latin Face"', 'typeface="+mn-lt"').replaceAll('typeface="Eastern Face"', 'typeface="+mn-ea"').replaceAll('typeface="Complex Face"', 'typeface="+mn-cs"')));
  return zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
}
