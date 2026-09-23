/**
 * Small helpers on top of @xmldom/xmldom for walking Word's OOXML body.
 * Not namespace-aware -- Word always writes elements with their literal "w:"/"a:"/etc.
 * prefix, so matching tagName strings directly is simpler and reliable here.
 */

export function directChildren(el: Element, tagName: string): Element[] {
  const out: Element[] = [];
  const children = el.childNodes;
  for (let i = 0; i < children.length; i++) {
    const node = children[i] as unknown as Element;
    if (node.nodeType === 1 && node.tagName === tagName) out.push(node);
  }
  return out;
}

export function firstChildTag(el: Element, tagName: string): Element | null {
  const children = el.childNodes;
  for (let i = 0; i < children.length; i++) {
    const node = children[i] as unknown as Element;
    if (node.nodeType === 1 && node.tagName === tagName) return node;
  }
  return null;
}

/** Concatenates every descendant <w:t> run's text, trimmed. */
export function textOf(el: Element | null): string {
  if (!el) return "";
  const runs = el.getElementsByTagName("w:t");
  let out = "";
  for (let i = 0; i < runs.length; i++) out += runs[i].textContent ?? "";
  return out.trim();
}
