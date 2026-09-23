import type { LoadedDocx } from "./zip";
import { findHeadingParagraph } from "./xmlTextUtils";

export interface FrontMatterPhotosInput {
  operationalDataImage?: Buffer;
  /** Up to two images -- the template itself has exactly two "Data Plate" photo placeholders
   *  (a plain one, and a second paired with a nameplate spec-details table), matching how
   *  inspectors sometimes take a second data plate photo. Extra entries beyond 2 are ignored. */
  dataPlateImages?: Buffer[];
}

export interface FrontMatterPhotosResult {
  operationalDataInserted: boolean;
  /** How many of the template's two Data Plate slots got a real photo. */
  dataPlateInsertedCount: number;
}

/**
 * The "Operational Data" (control-room HMI screenshot) and "Data Plate" (turbine nameplate
 * photo) placeholders sit at the very start of the Photos section, before the Inlet/
 * Compressor/Combustion/Turbine/Exhaust subsections -- and aren't in the raw MDI file at all,
 * since the borescope only photographs turbine components, never the HMI screen or the
 * physical nameplate (confirmed: neither string appears anywhere in a real raw export). Both
 * have to come from the inspector directly.
 *
 * When a photo IS supplied for a slot, this swaps its bytes into the exact relationship/media
 * slot the template's own placeholder picture already uses -- Word just displays whatever
 * bytes are there, so the template's own frame size/position is left completely untouched (the
 * same trick Word's own "Change Picture" does). When a slot ISN'T filled, its whole placeholder
 * block (photo + caption/label, and for the second Data Plate slot the nameplate spec-details
 * fields alongside it) is deleted outright, rather than leaving Brett's own "Insert Photo Here"
 * reminder sitting in a report that's otherwise meant to be finished.
 */
export function insertFrontMatterPhotos(template: LoadedDocx, input: FrontMatterPhotosInput): FrontMatterPhotosResult {
  const xml = template.documentXml;
  const photosHeading = findHeadingParagraph(xml, "Heading1", "Photos", 0);
  if (!photosHeading) throw new Error('Could not find a "Photos" (Heading1) heading in the template.');

  const firstSubsection = findHeadingParagraph(xml, "Heading2", "Inlet Section", photosHeading.paragraphStart);
  const searchEnd = firstSubsection ? firstSubsection.paragraphStart : xml.length;

  const opLabelIdx = findLabel(xml, "Operational Data", photosHeading.paragraphStart, searchEnd);
  const dpLabelIndexes = findAllLabels(xml, "Data Plate", opLabelIdx + 1, searchEnd);
  if (dpLabelIndexes.length === 0) {
    throw new Error('Could not find any "Data Plate" placeholder at the start of the template\'s Photos section.');
  }

  const dataPlateImages = input.dataPlateImages ?? [];
  const result: FrontMatterPhotosResult = { operationalDataInserted: false, dataPlateInsertedCount: 0 };
  const removals: { start: number; end: number }[] = [];

  if (input.operationalDataImage) {
    replacePlaceholderImage(template, opLabelIdx, input.operationalDataImage);
    result.operationalDataInserted = true;
  } else {
    removals.push(findUnitBounds(xml, opLabelIdx));
  }

  dpLabelIndexes.forEach((labelIdx, i) => {
    const image = dataPlateImages[i];
    if (image) {
      replacePlaceholderImage(template, labelIdx, image);
      result.dataPlateInsertedCount++;
    } else {
      removals.push(findUnitBounds(xml, labelIdx));
    }
  });

  if (removals.length > 0) {
    removals.sort((a, b) => b.start - a.start);
    let updated = template.documentXml;
    for (const { start, end } of removals) {
      updated = updated.slice(0, start) + updated.slice(end);
    }
    template.documentXml = updated;
  }

  return result;
}

function findLabel(xml: string, label: string, from: number, before: number): number {
  const idx = xml.indexOf(label, from);
  if (idx === -1 || idx >= before) {
    throw new Error(`Could not find the "${label}" placeholder at the start of the template's Photos section.`);
  }
  return idx;
}

function findAllLabels(xml: string, label: string, from: number, before: number): number[] {
  const indexes: number[] = [];
  let cursor = from;
  while (true) {
    const idx = xml.indexOf(label, cursor);
    if (idx === -1 || idx >= before) break;
    indexes.push(idx);
    cursor = idx + label.length;
  }
  return indexes;
}

/** Finds the placeholder photo immediately before `labelTextIndex` (the nearest preceding
 *  <wp:inline> picture -- the small watermark logo next to it is a separate <wp:anchor>
 *  drawing, so searching for <wp:inline> specifically skips right past it) and overwrites its
 *  underlying media file with `imageBytes`. */
function replacePlaceholderImage(template: LoadedDocx, labelTextIndex: number, imageBytes: Buffer): void {
  const relId = findPictureRelId(template.documentXml, labelTextIndex);
  const mediaPath = template.relationships.get(relId);
  if (!mediaPath) {
    throw new Error(`Relationship ${relId} not found in the template's relationships.`);
  }

  // mediaFiles is read-only bookkeeping from load time -- saveDocx() only ever serializes
  // whatever's registered directly on the zip, so that has to be updated too.
  template.mediaFiles.set(mediaPath, imageBytes);
  template.zip.file(mediaPath, imageBytes);
}

function findPictureRelId(xml: string, labelTextIndex: number): string {
  const inlineIdx = xml.lastIndexOf("<wp:inline", labelTextIndex);
  if (inlineIdx === -1) {
    throw new Error("Could not find the placeholder photo's <wp:inline> drawing in the template.");
  }
  const blipMatch = /<a:blip r:embed="([^"]+)"/.exec(xml.slice(inlineIdx, labelTextIndex));
  if (!blipMatch) {
    throw new Error("Could not find the placeholder photo's image relationship in the template.");
  }
  return blipMatch[1];
}

/**
 * Finds the whole removable unit for a placeholder: the enclosing 2-column <w:tbl> (photo
 * cell + a NESTED <w:tbl> holding the caption label, and for the second Data Plate slot a
 * few extra spec-detail rows too -- nesting-aware since a naive indexOf("</w:tbl>") would stop
 * at the inner table's own close first), plus the single spacer <w:p> that always immediately
 * follows it (Word requires a paragraph between two adjacent tables, the same rule behind the
 * nested-table-needs-a-trailing-paragraph fix elsewhere in this codebase) -- deleting the table
 * without it would leave two tables touching with nothing between them.
 */
function findUnitBounds(xml: string, labelTextIndex: number): { start: number; end: number } {
  const inlineIdx = xml.lastIndexOf("<wp:inline", labelTextIndex);
  if (inlineIdx === -1) {
    throw new Error("Could not find the placeholder photo's <wp:inline> drawing in the template.");
  }
  const tableStart = xml.lastIndexOf("<w:tbl>", inlineIdx);
  if (tableStart === -1) {
    throw new Error("Could not find the placeholder's enclosing table in the template.");
  }
  const tableEnd = findMatchingTableClose(xml, tableStart);

  let end = tableEnd;
  const nextP = xml.indexOf("<w:p", tableEnd);
  const nextTbl = xml.indexOf("<w:tbl>", tableEnd);
  if (nextP !== -1 && (nextTbl === -1 || nextP < nextTbl)) {
    const pClose = xml.indexOf("</w:p>", nextP);
    if (pClose !== -1) end = pClose + "</w:p>".length;
  }

  return { start: tableStart, end };
}

/** Depth-counts <w:tbl>/</w:tbl> tokens (exact matches only -- <w:tblPr>, <w:tblGrid>, etc.
 *  all have extra characters before their closing ">" so they never collide) to find the
 *  close tag that actually matches `tableStart`, tolerating any tables nested inside it. */
function findMatchingTableClose(xml: string, tableStart: number): number {
  const OPEN = "<w:tbl>";
  const CLOSE = "</w:tbl>";
  let depth = 0;
  let i = tableStart;
  while (i < xml.length) {
    const nextOpen = xml.indexOf(OPEN, i);
    const nextClose = xml.indexOf(CLOSE, i);
    if (nextClose === -1) throw new Error("Unbalanced <w:tbl> in the template.");
    if (nextOpen !== -1 && nextOpen < nextClose) {
      depth++;
      i = nextOpen + OPEN.length;
    } else {
      depth--;
      i = nextClose + CLOSE.length;
      if (depth === 0) return i;
    }
  }
  throw new Error("Could not find the placeholder's matching </w:tbl>.");
}
