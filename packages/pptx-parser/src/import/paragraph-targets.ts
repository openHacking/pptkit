import { attribute, children } from "../presentation/structure-xml.js";
import type { structureContext } from "../presentation/structure-context.js";
import type { XmlElement } from "../xml/read-xml.js";
import { PptxImportError } from "../types/import.js";
import type { TextRunBinding } from "./text-run-bindings.js";

/** Matches the existing content projection's numeric drawing identity domain. */
export function textTargetId(value: string | undefined): string | undefined {
  const text = value?.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, ""), number = Number(text);
  return text && /^[+]?\d+$/.test(text) && Number.isSafeInteger(number) && number <= 4294967295 ? String(number) : undefined;
}
type Context = Awaited<ReturnType<typeof structureContext>>;
interface Owner {
  readonly id: string;
  readonly paragraphs: ReadonlyMap<XmlElement, number>;
  readonly count: number;
  matches: number;
  invalid: boolean;
}

/** Private proof of unchanged direct paragraph identity, never animation execution. */
export async function paragraphStableOwners(
  context: Context, bindings: readonly TextRunBinding[], maxRecords: number, checkpoint: () => Promise<void>,
): Promise<ReadonlyMap<string, ReadonlyMap<string, number>>> {
  const byPart = new Map<string, Map<string, TextRunBinding[]>>();
  let records = 0;
  const charge = () => { if (++records > maxRecords) throw new PptxImportError("import-limit-exceeded", "Paragraph target proofs exceed maxReportFindings."); };
  for (const binding of bindings) {
    await checkpoint();
    if (binding.owner === "notes" || binding.owner === "table-cell" || binding.source.sourceObjectId === undefined) continue;
    charge();
    let part = byPart.get(binding.source.part); if (!part) { part = new Map(); byPart.set(binding.source.part, part); }
    const id = binding.source.sourceObjectId, list = part.get(id) ?? []; list.push(binding); part.set(id, list);
  }
  const result = new Map<string, ReadonlyMap<string, number>>();
  for (const [part, owners] of byPart) {
    const atoms = new Map<number, TextRunBinding[]>();
    for (const list of owners.values()) for (const binding of list) {
      await checkpoint();
      if (binding.sourceElementIndex !== undefined) {
        const at = atoms.get(binding.sourceElementIndex) ?? []; at.push(binding); atoms.set(binding.sourceElementIndex, at);
      }
    }
    const identities = new Map<string, number>(), candidates = new Map<string, Owner[]>();
    const stack: { node: XmlElement; owner?: Owner; paragraph?: number }[] = [{ node: context.rawXml.get(part)! }];
    let index = 0;
    while (stack.length) {
      await checkpoint();
      const item = stack.pop()!, node = item.node; index++;
      let owner = item.owner, paragraph = item.paragraph;
      if (node.namespaceUri === context.p && node.local === "cNvPr") {
        const raw = attribute(node, "id");
        // Content's integer reader also accepts lexical negative zero. It must
        // participate in collision checks even though targets do not admit it.
        const id = textTargetId(raw) ?? (raw !== undefined && /^[\t\r\n ]*-0+[\t\r\n ]*$/.test(raw) ? "0" : undefined);
        if (id !== undefined && owners.has(id)) identities.set(id, (identities.get(id) ?? 0) + 1);
      }
      if (node.namespaceUri === context.p && node.local === "sp") {
        // Nested/malformed owners never inherit the containing shape's proof.
        owner = undefined; paragraph = undefined;
        const direct = children(node), nv = direct.filter(n => n.namespaceUri === context.p && n.local === "nvSpPr");
        const ids = nv.length === 1 ? children(nv[0]!).filter(n => n.namespaceUri === context.p && n.local === "cNvPr") : [];
        const id = ids.length === 1 ? textTargetId(attribute(ids[0]!, "id")) : undefined;
        const bodies = direct.filter(n => n.namespaceUri === context.p && n.local === "txBody");
        if (id !== undefined && owners.has(id) && bodies.length === 1) {
          const bodyChildren = children(bodies[0]!);
          const paragraphs = new Map<XmlElement, number>();
          for (const child of bodyChildren) {
            await checkpoint();
            if (child.namespaceUri === context.a && child.local === "p") { charge(); paragraphs.set(child, paragraphs.size); }
          }
          const valid = paragraphs.size > 0 && bodyChildren.every(n => n.namespaceUri === context.a && ["bodyPr", "lstStyle", "p"].includes(n.local));
          charge(); owner = { id, paragraphs, count: paragraphs.size, matches: 0, invalid: !valid };
          const list = candidates.get(id) ?? []; list.push(owner); candidates.set(id, list);
        }
      }
      if (owner?.paragraphs.has(node)) paragraph = owner.paragraphs.get(node);
      for (const binding of atoms.get(index) ?? []) {
        await checkpoint();
        const id = binding.source.sourceObjectId!;
        if (!owner || owner.id !== id || paragraph !== binding.paragraphIndex || node.namespaceUri !== context.a || node.local !== "t") {
          // No matching source owner means this source identity cannot be proven.
          identities.set(id, Number.POSITIVE_INFINITY);
        } else owner.matches++;
      }
      for (let i = node.children.length - 1; i >= 0; i--) {
        const child = node.children[i]!;
        if (child.kind === "element") { await checkpoint(); stack.push({ node: child, ...(owner ? { owner } : {}), ...(paragraph !== undefined ? { paragraph } : {}) }); }
      }
    }
    const proven = new Map<string, number>();
    for (const [id, list] of owners) {
      await checkpoint();
      const choices = candidates.get(id), owner = choices?.length === 1 ? choices[0] : undefined;
      if (!owner || owner.invalid || identities.get(id) !== 1 || !list.some(b => b.status === "bound") || owner.matches !== list.filter(b => b.sourceElementIndex !== undefined).length) continue;
      const paragraphs = new Set(list.map(b => b.paragraphIndex));
      if (paragraphs.size !== owner.count || [...paragraphs].some(i => i < 0 || i >= owner.count)) continue;
      proven.set(id, owner.count);
    }
    result.set(part, proven);
  }
  return result;
}

/** Only the exact, transparent spTgt/txEl/pRg form is admitted. */
export function stableParagraphTarget(node: XmlElement, p: string, count: number): boolean {
  const transparent = (n: XmlElement, allowed: readonly string[]) => n.attributes.every(a => a.namespaceUri === "http://www.w3.org/2000/xmlns/" || (!a.namespaceUri && allowed.includes(a.local)))
    && n.children.every(c => c.kind === "element" || ((c.kind === "text" || c.kind === "cdata") && /^[\t\r\n ]*$/.test(c.value)));
  if (!transparent(node, ["spid"])) return false;
  const targets = children(node), text = targets[0];
  if (targets.length !== 1 || text?.namespaceUri !== p || text.local !== "txEl" || !transparent(text, [])) return false;
  const ranges = children(text), range = ranges[0];
  if (ranges.length !== 1 || range?.namespaceUri !== p || range.local !== "pRg" || !transparent(range, ["st", "end"]) || children(range).length) return false;
  const start = textTargetId(attribute(range, "st")), end = textTargetId(attribute(range, "end"));
  return start !== undefined && end !== undefined && Number(start) <= Number(end) && Number(end) < count;
}
