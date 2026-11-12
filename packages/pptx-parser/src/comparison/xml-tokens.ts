import type { XmlDocument, XmlNode } from "../xml/read-xml.js";

export interface XmlComparisonToken { readonly value: string; readonly path: string }
const order = (a: string, b: string): number => a < b ? -1 : a > b ? 1 : 0;

/** Conservative source-tree tokens. Prefixes and namespace declarations remain significant. */
export function* xmlTokens(document: XmlDocument): Generator<XmlComparisonToken> {
  yield { value: JSON.stringify(["declaration", document.standalone ?? null]), path: "$declaration/standalone" };
  const stack = [{ children: document.children as readonly XmlNode[], index: 0, parent: 0 }];
  let elements = 0;
  while (stack.length) {
    const frame = stack[stack.length - 1]!;
    if (frame.index === frame.children.length) {
      stack.pop();
      if (frame.parent) yield { value: '["close"]', path: `$element[${frame.parent}]` };
      continue;
    }
    const node = frame.children[frame.index++]!, path = `$element[${frame.parent}]/child[${frame.index}]`;
    if (node.kind === "text" || node.kind === "cdata") {
      const chunks = [node.value];
      while (frame.index < frame.children.length) {
        const next = frame.children[frame.index]!;
        if (next.kind !== "text" && next.kind !== "cdata") break;
        frame.index++; chunks.push(next.value);
      }
      const text = chunks.join("");
      if (text) yield { value: JSON.stringify(["text", text]), path };
    } else if (node.kind === "element") {
      const index = ++elements;
      yield { value: JSON.stringify(["open", node.name, node.namespaceUri]), path: `$element[${index}]` };
      for (const attribute of [...node.attributes].sort((a, b) => order(a.name, b.name))) {
        yield { value: JSON.stringify(["attribute", attribute.name, attribute.namespaceUri, attribute.value]), path: `$element[${index}]/@${attribute.name}` };
      }
      stack.push({ children: node.children, index: 0, parent: index });
    } else if (node.kind === "comment") {
      yield { value: JSON.stringify(["comment", node.value]), path };
    } else if (node.kind === "processing-instruction") {
      yield { value: JSON.stringify(["processing-instruction", node.target, node.value]), path };
    }
  }
}
