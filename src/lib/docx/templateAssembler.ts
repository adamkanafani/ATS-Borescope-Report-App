import type { LoadedDocx } from "./zip";
import { addImageRelationship } from "./zip";
import type { RawPhotoUnit } from "./rawMdiParser";
import type { PhotosSection } from "./photosSectionMap";
import { detectTemplateFamily, resolveClassify, resolvePhotosSubsections, type TemplateFamily } from "./templateFamilies";
import { buildPhotoBlockXml } from "./photoBlockBuilder";
import { findHeadingParagraph } from "./xmlTextUtils";
import { aggregateObservations } from "./observationsAggregator";
import { fillObservations } from "./observationsFiller";
import { insertFrontMatterPhotos } from "./frontMatterPhotos";
import { removeExcludedSections } from "./sectionExclusion";

export interface AssembleResult {
  documentXml: string;
  inserted: number;
  unassigned: RawPhotoUnit[];
  bySection: Record<PhotosSection, number>;
  observationsFilled: string[];
  observationsSkipped: string[];
  operationalDataInserted: boolean;
  dataPlateInsertedCount: number;
}

export interface AssembleOptions {
  /** Per-unit (by index) section overrides -- checked before the keyword classifier. An
   *  explicit `null` forces exclusion even if the classifier would've placed it. Lets the
   *  review UI rescue (or veto) units without touching the deterministic classifier itself. */
  overrides?: Map<number, PhotosSection | null>;
  /** Manually-supplied photos for the Photos-section placeholders the raw MDI file never
   *  contains (the borescope only photographs turbine components) -- see frontMatterPhotos.ts.
   *  dataPlateImages maps to the template's two Data Plate slots in order. */
  operationalDataImage?: Buffer;
  dataPlateImages?: Buffer[];
  /** Photos subsections to leave out of the report entirely (not just left empty) -- both
   *  their Photos content and their Observations table, when one exists. Some ATS customers
   *  only pay for part of the inspection; see sectionExclusion.ts. */
  excludedSections?: Set<PhotosSection>;
  /** Which template family's subsection layout/classifier to use (see templateFamilies.ts).
   *  Auto-detected from the template's own Photos headings when omitted -- callers only need
   *  to pass this explicitly if they want to skip that detection (e.g. already resolved it at
   *  scan time and want generate time to fail the same way rather than re-guess). */
  family?: TemplateFamily;
}

/**
 * Inserts every classifiable raw photo unit into the template's "Photos" section, replacing
 * Brett's instructional example blocks under each subsection heading with the real units
 * (see reference/phase1-source-analysis.md and the plan file for how this was derived).
 * Mutates `template` in place (new media + relationships via addImageRelationship) and
 * returns the new documentXml -- caller still owns calling saveDocx() to write it out.
 */
export function assembleReport(template: LoadedDocx, units: RawPhotoUnit[], options: AssembleOptions = {}): AssembleResult {
  const overrides = options.overrides;
  const excludedSections = options.excludedSections ?? new Set<PhotosSection>();

  const family = options.family ?? detectTemplateFamily(template.documentXml);
  if (!family) {
    throw new Error("Could not determine which report layout this template uses (unrecognized Photos subsection headings).");
  }

  const observationsResult = fillObservations(template.documentXml, aggregateObservations(units, family.classify));
  const xml = observationsResult.documentXml;

  const photosHeading = findHeadingParagraph(xml, "Heading1", "Photos", 0);
  if (!photosHeading) {
    throw new Error('Could not find a "Photos" (Heading1) heading in the template.');
  }
  const photosHeadingPos = photosHeading.paragraphStart;

  // The subsections THIS template actually has -- not every subsection the family knows about
  // (see templateFamilies.ts: some, like 7FA's Variable Inlet Guide Vanes bucket, are optional
  // and only some templates in a family have them). classify() is wrapped the same way, so a
  // unit that would've landed in a subsection this template doesn't have falls back correctly.
  const photosSubsections = resolvePhotosSubsections(family, xml, photosHeadingPos);
  const classify = resolveClassify(family, new Set(photosSubsections));

  const grouped = new Map<PhotosSection, RawPhotoUnit[]>();
  const unassigned: RawPhotoUnit[] = [];
  for (const unit of units) {
    const section = overrides?.has(unit.index) ? overrides.get(unit.index) ?? null : unit.component ? classify(unit.component) : null;
    if (!section || !unit.imageBytes) {
      unassigned.push(unit);
      continue;
    }
    if (!grouped.has(section)) grouped.set(section, []);
    grouped.get(section)!.push(unit);
  }

  const headings = photosSubsections.map((name) => {
    const found = findHeadingParagraph(xml, "Heading2", name, photosHeadingPos);
    if (!found) throw new Error(`Could not find the "${name}" heading under Photos in the template.`);
    return { name, paragraphStart: found.paragraphStart };
  });

  const trailingSectPr = xml.indexOf("<w:sectPr", headings[headings.length - 1].paragraphStart);
  const bodyEnd = trailingSectPr === -1 ? xml.length : trailingSectPr;

  let uniqueIdCounter = 0;
  const bySection = Object.fromEntries(photosSubsections.map((name) => [name, 0])) as Record<PhotosSection, number>;

  const splices: { start: number; end: number; replacement: string }[] = [];

  for (let i = 0; i < headings.length; i++) {
    const { name, paragraphStart } = headings[i];
    const paragraphEnd = xml.indexOf("</w:p>", paragraphStart);
    if (paragraphEnd === -1) throw new Error(`Malformed heading paragraph for "${name}".`);
    const contentStart = paragraphEnd + "</w:p>".length;
    const contentEnd = i + 1 < headings.length ? headings[i + 1].paragraphStart : bodyEnd;

    const unitsForSection = excludedSections.has(name) ? [] : grouped.get(name) ?? [];
    const tables = unitsForSection.map((unit) => {
      const photoRelId = addImageRelationship(template, unit.imageBytes as Buffer, "jpg");
      return buildPhotoBlockXml(unit, { photoRelId, uniqueId: uniqueIdCounter++ });
    });
    bySection[name] = tables.length;

    // A blank paragraph between the heading and the first photo table (matching the spacing in
    // ATS's own sample reports -- without it the photo sits flush against the heading text),
    // and again between each photo table so consecutive photos aren't stacked flush against
    // each other with no visual break. w:lineRule="exact" pins this to an exact height instead
    // of the ~240twip default single line, matching the wider gap in Adam's reference example.
    const spacer =
      tables.length > 0 ? '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="700" w:lineRule="exact"/></w:pPr></w:p>' : "";
    splices.push({ start: contentStart, end: contentEnd, replacement: spacer + tables.join(spacer) });
  }

  // Apply from the last splice backwards so earlier positions (computed against the
  // original xml) stay valid as we go.
  splices.sort((a, b) => b.start - a.start);
  let result = xml;
  for (const { start, end, replacement } of splices) {
    result = result.slice(0, start) + replacement + result.slice(end);
  }

  template.documentXml = result;

  if (excludedSections.size > 0) {
    template.documentXml = removeExcludedSections(template.documentXml, excludedSections, photosSubsections, family.observationsHeadingBySection);
  }

  const frontMatterResult = insertFrontMatterPhotos(template, {
    operationalDataImage: options.operationalDataImage,
    dataPlateImages: options.dataPlateImages,
  });

  return {
    documentXml: template.documentXml,
    inserted: Object.values(bySection).reduce((sum, n) => sum + n, 0),
    unassigned,
    bySection,
    observationsFilled: observationsResult.filled,
    observationsSkipped: observationsResult.skipped,
    operationalDataInserted: frontMatterResult.operationalDataInserted,
    dataPlateInsertedCount: frontMatterResult.dataPlateInsertedCount,
  };
}
