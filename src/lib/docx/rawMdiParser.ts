import { DOMParser } from "@xmldom/xmldom";
import type { LoadedDocx } from "./zip";
import { directChildren, firstChildTag, textOf } from "./domHelpers";

export interface Measurement {
  /** The row label, e.g. "Point-to-Line" or "Measurement Plane". */
  label: string;
  value: string;
}

export interface RawPhotoUnit {
  index: number;
  filenameCaption: string | null;
  component: string | null;
  location: string | null;
  /** Extra rows between Location and the observation/classification row -- present on
   *  measurement photos (tip clashing, R-17 movement, etc.), e.g. [{label: "Point-to-Line",
   *  value: "0.091 in"}]. */
  measurements: Measurement[];
  /** The row immediately before Comments -- labeled "Observation" in one MDI software
   *  version, "Classification" in another; see parseCaptionTable(). */
  observation: string | null;
  comments: string | null;
  /** The nested caption table's own serialized XML -- reused as-is when building an
   *  inserted photo block, since its fonts/sizes already match Brett's rules exactly. */
  captionTableXml: string;
  imageRelId: string | null;
  imagePath: string | null;
  imageBytes: Buffer | null;
  warnings: string[];
}

const FILENAME_PATTERN = /\.(jpg|jpeg|png|bmp)\b/i;

/**
 * Every photo in a raw borescope-generated .docx is one top-level table, directly under
 * <w:body>, with two rows (verified by walking the real document -- see
 * reference/phase1-source-analysis.md, "Raw MDI output"):
 *   Row 0: cell 0 = the photo (<w:drawing>), cell 1 = a nested caption table whose rows
 *          are always Component, then Location, then (for measurement photos only) one or
 *          more extra rows like "Point-to-Line"/"Measurement Plane", then an observation/
 *          classification row (label text varies by MDI software version -- see
 *          parseCaptionTable()), then Comments last.
 *   Row 1: cell 0 = the auto-generated filename caption text (e.g.
 *          "Stage_1_Rotor_Blade_Leading_Edge_Pressure_Side_001.JPG") -- what inspectors
 *          delete before pasting into the ATS template. Cell 1 is empty.
 * Other top-level tables in the document (the header metadata table, the "Inspection
 * Points" summary table) don't match this 2-row/nested-caption-table shape and are skipped.
 */
export function parseRawMdiPhotoUnits(loaded: LoadedDocx): RawPhotoUnit[] {
  const doc = new DOMParser().parseFromString(loaded.documentXml, "text/xml");
  const body = doc.getElementsByTagName("w:body")[0] as unknown as Element;
  const topLevelTables = directChildren(body, "w:tbl");

  const units: RawPhotoUnit[] = [];
  let index = 0;

  for (const tbl of topLevelTables) {
    const rows = directChildren(tbl, "w:tr");
    if (rows.length !== 2) continue;

    const row0Cells = directChildren(rows[0], "w:tc");
    if (row0Cells.length < 2) continue;
    const [photoCell, captionCellWrapper] = row0Cells;

    const captionTable = directChildren(captionCellWrapper, "w:tbl")[0];
    if (!captionTable) continue;
    const parsedCaption = parseCaptionTable(captionTable);
    if (!parsedCaption) continue;
    normalizeTblPrBeforeTblGrid(captionTable);

    const warnings: string[] = [];
    const { component, location, measurements, observation, comments } = parsedCaption;

    const drawings = photoCell.getElementsByTagName("w:drawing");
    let imageRelId: string | null = null;
    if (drawings.length === 0) {
      warnings.push("No <w:drawing> found in this photo unit's image cell.");
    } else {
      if (drawings.length > 1) {
        warnings.push(`Found ${drawings.length} drawings in this unit's image cell; using the first.`);
      }
      imageRelId = firstBlipEmbed(drawings[0] as unknown as Element);
      if (!imageRelId) warnings.push("Found a <w:drawing> but no <a:blip r:embed> inside it.");
    }

    const row1Cells = directChildren(rows[1], "w:tc");
    let filenameCaption: string | null = row1Cells[0] ? textOf(row1Cells[0]) : null;
    if (filenameCaption && !FILENAME_PATTERN.test(filenameCaption)) {
      warnings.push(`Row 1's caption cell didn't look like a filename: ${JSON.stringify(filenameCaption)}.`);
      filenameCaption = null;
    }

    let imagePath: string | null = null;
    let imageBytes: Buffer | null = null;
    if (imageRelId) {
      imagePath = loaded.relationships.get(imageRelId) ?? null;
      if (!imagePath) {
        warnings.push(`Relationship id ${imageRelId} not found in word/_rels/document.xml.rels.`);
      } else {
        imageBytes = loaded.mediaFiles.get(imagePath) ?? null;
        if (!imageBytes) {
          warnings.push(`Relationship ${imageRelId} resolved to "${imagePath}", but that file wasn't among the loaded media.`);
        }
      }
    }

    units.push({
      index: index++,
      filenameCaption,
      component,
      location,
      measurements,
      observation,
      comments,
      captionTableXml: captionTable.toString(),
      imageRelId,
      imagePath,
      imageBytes,
      warnings,
    });
  }

  return units;
}

interface ParsedCaption {
  component: string | null;
  location: string | null;
  measurements: Measurement[];
  observation: string | null;
  comments: string | null;
}

/**
 * The caption table's row count varies: a plain photo has exactly 4 rows (Component,
 * Location, Observation/Classification, Comments); a measurement photo inserts 1+ extra rows
 * between Location and that row. Rather than assume a fixed row count, anchor on Component
 * (row 0), Location (row 1), and Comments (always last), then take the row immediately before
 * Comments as the observation/classification field -- everything between Location and that is
 * a measurement row. Deliberately doesn't check that row's own label text: one MDI software
 * version calls it "Observation", another calls it "Classification" (confirmed against a real
 * sample from each), and there's no reason a third variant couldn't use yet another word --
 * its position relative to Comments is the only thing that's stayed stable across both.
 */
function parseCaptionTable(captionTable: Element): ParsedCaption | null {
  const rows = directChildren(captionTable, "w:tr");
  if (rows.length < 4) return null;

  const labels = rows.map((row) => textOf(cellAt(row, 0)));
  if (labels[0] !== "Component" || labels[1] !== "Location") return null;
  const lastIndex = labels.length - 1;
  if (labels[lastIndex] !== "Comments") return null;

  const observationIndex = lastIndex - 1;

  const measurements: Measurement[] = [];
  for (let i = 2; i < observationIndex; i++) {
    const value = textOf(cellAt(rows[i], 1));
    if (labels[i]) measurements.push({ label: labels[i], value });
  }

  return {
    component: textOf(cellAt(rows[0], 1)) || null,
    location: textOf(cellAt(rows[1], 1)) || null,
    measurements,
    observation: textOf(cellAt(rows[observationIndex], 1)) || null,
    comments: textOf(cellAt(rows[lastIndex], 1)) || null,
  };
}

/**
 * The raw MDI export writes its caption tables with <w:tblGrid> before <w:tblPr> --
 * the OOXML schema requires the opposite order (CT_Tbl: tblPr, tblGrid, tr*). Real Word
 * opens the file fine (its parser is lenient), but a strict validator flags it, and we'd
 * rather not carry a source quirk into the assembled report. Cheap to normalize once here,
 * since we're already walking this element.
 */
function normalizeTblPrBeforeTblGrid(table: Element): void {
  const tblPr = firstChildTag(table, "w:tblPr");
  const tblGrid = firstChildTag(table, "w:tblGrid");
  if (!tblPr || !tblGrid) return;
  const children = Array.prototype.slice.call(table.childNodes) as Element[];
  if (children.indexOf(tblGrid) < children.indexOf(tblPr)) {
    table.insertBefore(tblPr, tblGrid);
  }
}

function cellAt(row: Element, i: number): Element | null {
  return directChildren(row, "w:tc")[i] ?? null;
}

function firstBlipEmbed(drawing: Element): string | null {
  const blips = drawing.getElementsByTagName("a:blip");
  if (blips.length === 0) return null;
  return blips[0].getAttribute("r:embed");
}
