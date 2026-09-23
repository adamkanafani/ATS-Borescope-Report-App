/**
 * Shared string-level OOXML text helpers -- operate directly on the raw document.xml text
 * rather than a parsed DOM, so edits can be applied as targeted splices without re-serializing
 * (and risking subtly reformatting) the rest of a multi-megabyte document. Safe for elements
 * that don't nest within the regions these are used on (w:p, w:tr, w:tc, w:tbl in the
 * Observations tables all hold plain text only here -- no nested tables/drawings).
 */

/** <w:p> never nests in body content, so the nearest preceding "<w:p" is always the true enclosing paragraph. */
export function findParagraphStart(xml: string, beforeIndex: number): number {
  const openTag = /<w:p(?=[\s>/])/g;
  let lastIndex = -1;
  let match: RegExpExecArray | null;
  while ((match = openTag.exec(xml)) && match.index < beforeIndex) {
    lastIndex = match.index;
  }
  if (lastIndex === -1) throw new Error(`Could not find an enclosing <w:p> before position ${beforeIndex}.`);
  return lastIndex;
}

/** Concatenates every <w:t> run's text between two offsets -- immune to Word splitting a
 *  heading's text across multiple runs (spell-check/revision boundaries), unlike a plain
 *  substring search for ">text<". */
export function extractText(xml: string, start: number, end: number): string {
  const runPattern = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  runPattern.lastIndex = start;
  let out = "";
  let match: RegExpExecArray | null;
  while ((match = runPattern.exec(xml)) && match.index < end) {
    out += match[1];
  }
  return out.trim();
}

/** Finds a paragraph styled `styleId` whose full text (across runs) equals `text`, searching
 *  forward from `fromIndex`. Anchored on the <w:pStyle> tag itself, not the heading's own
 *  text, since that text isn't reliably a single contiguous substring in the raw XML. */
export function findHeadingParagraph(
  xml: string,
  styleId: "Heading1" | "Heading2",
  text: string,
  fromIndex: number,
): { paragraphStart: number; paragraphEnd: number } | null {
  const styleTag = new RegExp(`<w:pStyle w:val="${styleId}"`, "g");
  styleTag.lastIndex = fromIndex;
  let match: RegExpExecArray | null;
  while ((match = styleTag.exec(xml))) {
    const paragraphStart = findParagraphStart(xml, match.index);
    const paragraphEnd = xml.indexOf("</w:p>", match.index);
    if (paragraphEnd === -1) continue;
    if (extractText(xml, paragraphStart, paragraphEnd) === text) {
      return { paragraphStart, paragraphEnd };
    }
  }
  return null;
}

/** Finds the next `<tagName ...>...</tagName>` pair at/after `fromIndex`. Only valid where the
 *  tag is known not to nest within itself in this region (true for w:tr/w:tc/w:tbl in the
 *  Observations tables, which hold plain text, not nested tables). */
export function findElementBounds(xml: string, tagName: string, fromIndex: number): { start: number; end: number } | null {
  const openTag = new RegExp(`<${tagName}(?=[\\s>/])`, "g");
  openTag.lastIndex = fromIndex;
  const openMatch = openTag.exec(xml);
  if (!openMatch) return null;
  const closeTag = `</${tagName}>`;
  const closeIdx = xml.indexOf(closeTag, openMatch.index);
  if (closeIdx === -1) return null;
  return { start: openMatch.index, end: closeIdx + closeTag.length };
}

/** All direct `<w:tr>` elements between [start, end), in document order. */
export function findRowsInRange(xml: string, start: number, end: number): { start: number; end: number }[] {
  const rows: { start: number; end: number }[] = [];
  let cursor = start;
  while (true) {
    const bounds = findElementBounds(xml, "w:tr", cursor);
    if (!bounds || bounds.start >= end) break;
    rows.push(bounds);
    cursor = bounds.end;
  }
  return rows;
}

/** All direct `<w:tc>` elements between [start, end), in document order. */
export function findCellsInRange(xml: string, start: number, end: number): { start: number; end: number }[] {
  const cells: { start: number; end: number }[] = [];
  let cursor = start;
  while (true) {
    const bounds = findElementBounds(xml, "w:tc", cursor);
    if (!bounds || bounds.start >= end) break;
    cells.push(bounds);
    cursor = bounds.end;
  }
  return cells;
}
