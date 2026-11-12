export const SECTION_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";
export const SECTION_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
export const SECTION_IDS = ["{01F07B81-39E6-4BBB-9B89-66EA253FBD29}", "{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274}", "{CA1E145A-94F4-4C2D-9BC0-76C4A01D48ED}"];
export const section = (ids = [], attrs = "", suffix = "") => `<s:section ${attrs}><s:sldIdLst>${ids.map((id) => `<s:sldId id="${id}"/>`).join("")}</s:sldIdLst>${suffix}</s:section>`;
export const sectionExtension = (body, attrs = "") => `<p:ext uri="${SECTION_URI}" xmlns:s="${SECTION_NS}" ${attrs}><s:sectionLst>${body}</s:sectionLst></p:ext>`;
export function withSections(f, body = section([300], `id="${SECTION_IDS[0]}" name="Opening"`) + section([], `id="${SECTION_IDS[1]}" name="Empty"`) + section([256], `id="${SECTION_IDS[2]}" name="Closing"`), extra = "") {
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace('</p:presentation>', `<p:extLst>${sectionExtension(body)}${extra}</p:extLst></p:presentation>`);
  return f;
}
