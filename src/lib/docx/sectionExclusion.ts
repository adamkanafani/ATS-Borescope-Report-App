import { findHeadingParagraph, findPrecedingPageBreakStart } from "./xmlTextUtils";
import type { PhotosSection } from "./photosSectionMap";

/**
 * Some ATS customers only pay for (and only want in the final report) specific inspection
 * areas -- confirmed from real completed sample reports back in Phase 1: a compressor-only
 * report's title/TOC/Combustion/Turbine/Exhaust sections are entirely dropped, not just left
 * blank. This removes BOTH halves of each excluded section -- its Observations H2 block
 * (heading + table, when `observationsHeadingBySection` has one for that section) and its
 * Photos H2 block (heading + photos) -- so the generated report reads as if that area was never
 * inspected at all, rather than showing an empty or half-filled section.
 *
 * `resolvedPhotosSubsections` must be the subsections THIS specific template actually has (see
 * templateFamilies.ts's resolvePhotosSubsections) -- not a family's full/optional list, or a
 * subsection this template doesn't have would look like a heading to search for and remove.
 */
export function removeExcludedSections(
  documentXml: string,
  excludedSections: Set<PhotosSection>,
  resolvedPhotosSubsections: PhotosSection[],
  observationsHeadingBySection: Partial<Record<PhotosSection, string>>,
): string {
  if (excludedSections.size === 0) return documentXml;

  const xml = documentXml;
  const removals: { start: number; end: number }[] = [];

  const photosHeading = findHeadingParagraph(xml, "Heading1", "Photos", 0);
  if (!photosHeading) throw new Error('Could not find a "Photos" (Heading1) heading in the template.');

  const observationsHeading = findHeadingParagraph(xml, "Heading1", "Observations", 0);
  if (observationsHeading) {
    collectH2Removals(
      xml,
      observationsHeading.paragraphStart,
      photosHeading.paragraphStart,
      excludedSections,
      observationsHeadingBySection,
      resolvedPhotosSubsections,
      removals,
    );
  }

  const trailingSectPr = xml.indexOf("<w:sectPr", photosHeading.paragraphStart);
  const photosSectionEnd = trailingSectPr === -1 ? xml.length : trailingSectPr;
  // Every resolved Photos subsection's own H2 text is exactly its PhotosSection name, so the
  // "heading text by section" map is just each name mapped to itself.
  const photosHeadingTextBySection = Object.fromEntries(resolvedPhotosSubsections.map((name) => [name, name])) as Partial<
    Record<PhotosSection, string>
  >;
  collectH2Removals(
    xml,
    photosHeading.paragraphStart,
    photosSectionEnd,
    excludedSections,
    photosHeadingTextBySection,
    resolvedPhotosSubsections,
    removals,
  );

  if (removals.length === 0) return xml;
  removals.sort((a, b) => b.start - a.start);
  let result = xml;
  for (const { start, end } of removals) {
    result = result.slice(0, start) + result.slice(end);
  }
  return result;
}

/** Locates every H2 heading named in `headingTextBySection` within [h1Start, h1End), in
 *  document order, and pushes a removal range for each excluded section: its own heading
 *  paragraph through the start of the next H2 found (or h1End for the last one) -- adjusted at
 *  both ends so the removal neither deletes nor orphans a page break:
 *   - the END stops short of whatever dedicated page-break paragraph precedes that boundary, so
 *     removing this section doesn't also delete the page break the *next* section (or whatever
 *     follows h1End) relies on;
 *   - the START is pulled back to swallow this section's *own* preceding page-break paragraph
 *     too, so removing it doesn't leave that paragraph behind as an orphaned extra blank page
 *     between whatever precedes it and the next (preserved) page break.
 *  See findPrecedingPageBreakStart. Two sections that happen to share the same heading text (not
 *  currently the case for any family) would collapse to one entry here, which is intentional --
 *  there's only one heading to remove. */
function collectH2Removals(
  xml: string,
  h1Start: number,
  h1End: number,
  excludedSections: Set<PhotosSection>,
  headingTextBySection: Partial<Record<PhotosSection, string>>,
  sectionOrder: PhotosSection[],
  removals: { start: number; end: number }[],
): void {
  const byHeadingText = new Map<string, { paragraphStart: number; excluded: boolean }>();
  for (const section of sectionOrder) {
    const headingText = headingTextBySection[section];
    if (!headingText || byHeadingText.has(headingText)) continue;
    const heading = findHeadingParagraph(xml, "Heading2", headingText, h1Start);
    if (heading && heading.paragraphStart < h1End) {
      byHeadingText.set(headingText, { paragraphStart: heading.paragraphStart, excluded: excludedSections.has(section) });
    }
  }

  const found = [...byHeadingText.values()].sort((a, b) => a.paragraphStart - b.paragraphStart);
  for (let i = 0; i < found.length; i++) {
    if (!found[i].excluded) continue;
    const boundary = i + 1 < found.length ? found[i + 1].paragraphStart : h1End;
    const start = findPrecedingPageBreakStart(xml, found[i].paragraphStart) ?? found[i].paragraphStart;
    const end = findPrecedingPageBreakStart(xml, boundary) ?? boundary;
    removals.push({ start, end });
  }
}
