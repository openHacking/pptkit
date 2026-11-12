import { contentFixture, shape, body } from "./content-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { nativeTimingSpecimen } from "../../../scripts/research/native-timing-fixture.mjs";
export function timingFixture({ strict = false, name = "fade-in", timing = nativeTimingSpecimen(name, 2), drawings = shape({ id: "2", text: body() }), rootAttributes = "" } = {}) {
  const f = contentFixture(drawings, strict);
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace('<p:sld ', `<p:sld ${rootAttributes} `).replace('</p:sld>', timing + '</p:sld>');
  return f;
}
export const timingBytes = options => encodeStructure(timingFixture(options));
