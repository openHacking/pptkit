import { attribute, children } from '../presentation/structure-xml.js';
import type { XmlElement } from '../xml/read-xml.js';
import type { ContentReader } from './content-reader.js';

const uri = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';
const namespace = 'http://schemas.microsoft.com/office/drawing/2017/decorative';
const xmlns = 'http://www.w3.org/2000/xmlns/';

/** Only a unique, exact extension with an explicit valid value is projected. */
export function readDecorative(cnv: XmlElement, reader: ContentReader): { value: boolean; node: XmlElement } | undefined {
  const lists = children(cnv).filter(node => node.namespaceUri === reader.a && node.local === 'extLst');
  if (lists.length !== 1) return undefined;
  const list = lists[0]!;
  const matches = children(list).filter(node => node.namespaceUri === reader.a && node.local === 'ext' && attribute(node, 'uri')?.toUpperCase() === uri);
  if (matches.length !== 1) return undefined;
  const ext = matches[0]!, nodes = children(ext);
  if (nodes.length !== 1 || nodes[0]!.namespaceUri !== namespace || nodes[0]!.local !== 'decorative') return undefined;
  const node = nodes[0]!, value = attribute(node, 'val')?.trim();
  if (value !== '1' && value !== '0' && value !== 'true' && value !== 'false') return undefined;
  const nonWhitespace = (element: XmlElement) => element.children.some(child => (child.kind === 'text' || child.kind === 'cdata') && !/^[\t\r\n ]*$/.test(child.value));
  if (ext.attributes.some(a => a.namespaceUri !== xmlns && (a.namespaceUri !== '' || a.local !== 'uri')) || node.attributes.some(a => a.namespaceUri !== xmlns && (a.namespaceUri !== '' || a.local !== 'val')) || children(node).length || nonWhitespace(ext) || nonWhitespace(node)) return undefined;
  reader.known(list); reader.attr(ext, 'uri'); reader.attr(node, 'val');
  return { value: value === '1' || value === 'true', node };
}
