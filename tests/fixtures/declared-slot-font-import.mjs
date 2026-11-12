import { mixedSupplementalBytes } from "./supplemental-font-import.mjs";
import { supplementalFonts } from "../../packages/pptx-parser/test/supplemental-font-fixture.mjs";
import { readGeneratedZip } from "../support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";

/** Six scalar owners with native Latin references and absent local EA/CS slots. */
export async function mixedDeclaredSlotBytes() {
  const parts = readGeneratedZip(await mixedSupplementalBytes());
  for (const name of ["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"]) {
    parts.set(name, Buffer.from(parts.get(name).toString().replace(/<a:latin\b[^>]*\/>/g, '<a:latin typeface="+mn-lt"/>').replace(/<a:(?:ea|cs)\b[^>]*\/>/g, "").replaceAll('lang="en-US"', 'lang="zh-TW"')));
  }
  const notesRules = supplementalFonts.replaceAll('typeface="', 'typeface="Notes ');
  parts.set("ppt/theme/theme2.xml", Buffer.from(parts.get("ppt/theme/theme2.xml").toString().replaceAll('typeface="Body Face"', 'typeface="Notes Body Face"').replace('</a:minorFont>', `${notesRules}</a:minorFont>`)));
  return zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
}
