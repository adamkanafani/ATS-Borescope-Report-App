import { findHeadingParagraph, findRowsInRange, findCellsInRange, extractText } from "./xmlTextUtils";
import type { ObservationsAggregation } from "./observationsAggregator";

export interface FillObservationsResult {
  documentXml: string;
  /** Row/group keys that got real text, e.g. "R1", "Stage 1 Nozzles". */
  filled: string[];
  /** Aggregation had data for these keys but no matching template row was found -- should be
   *  rare, since the keys come from the template's own rows, but surfaced rather than swallowed. */
  skipped: string[];
}

// The template names the shroud group inconsistently: "Shroud Blocks" for Stage 1, but
// "Honeycomb Shroud" for Stages 2 and 3 -- both mean the same row, normalized below.
const TURBINE_GROUP_PATTERN = /^Stage (\d+) (Nozzles|Buckets|Shroud Blocks|Honeycomb Shroud)$/;

function normalizeTurbineGroupName(name: string): string {
  return name === "Honeycomb Shroud" ? "Shroud Blocks" : name;
}

/** A Percent-Inspected cell's placeholder value -- a numeric percent ("100%", "75%") or Brett's
 *  own "still needs a number" placeholder. resetPercentPlaceholders() rewrites any of these
 *  still red at that point to a blank "XX%", rather than leaving one of Brett's specific-looking
 *  example numbers in a generated report. */
const PERCENT_PLACEHOLDER_PATTERN = /^\s*(\d+\s*%|Add Applicable\s*%)\s*$/i;

/**
 * Fills the Compressor and Turbine Observations tables' Condition cells from `aggregation`
 * (see observationsAggregator.ts). Deliberately does not touch Combustion, Exhaust, or any
 * Percent-Inspected cell -- see the plan file for why those aren't reliable to auto-fill from
 * this raw data. Leaves every cell it doesn't have data for untouched, and colors filled text
 * blue (0000FF) -- every AI-generated value in the report is blue, the template's own original
 * text is black, and red is reserved for Brett's own manual corrections (his existing
 * convention, which predates this app). This only resets Percent-Inspected placeholders to
 * "XX%" -- see recolorRemainingRedToBlack() (called by the caller, over the whole document) for
 * turning the remaining red itself black.
 */
export function fillObservations(documentXml: string, aggregation: ObservationsAggregation): FillObservationsResult {
  const xml = documentXml;
  const filled: string[] = [];
  const skipped: string[] = [];
  const edits: { start: number; end: number; replacement: string }[] = [];

  const obsHeading = findHeadingParagraph(xml, "Heading1", "Observations", 0);
  if (!obsHeading) throw new Error('Could not find an "Observations" (Heading1) heading in the template.');

  const compressorHeading = findHeadingParagraph(xml, "Heading2", "Compressor Section", obsHeading.paragraphStart);
  if (!compressorHeading) throw new Error('Could not find "Compressor Section" under Observations.');

  const combustionHeading = findHeadingParagraph(xml, "Heading2", "Combustion Section Cold Side", compressorHeading.paragraphStart);
  if (!combustionHeading) throw new Error('Could not find "Combustion Section Cold Side" under Observations.');

  const turbineHeading = findHeadingParagraph(xml, "Heading2", "Turbine Section", combustionHeading.paragraphStart);
  if (!turbineHeading) throw new Error('Could not find "Turbine Section" under Observations.');

  const exhaustHeading = findHeadingParagraph(xml, "Heading2", "Exhaust Section", turbineHeading.paragraphStart);
  if (!exhaustHeading) throw new Error('Could not find "Exhaust Section" under Observations.');

  // ---- Compressor: each row is [Stage, Condition, Percent] -- fill cells[1] where cells[0] matches. ----
  const compressorRemaining = new Map(aggregation.compressor);
  const compressorRows = findRowsInRange(xml, compressorHeading.paragraphStart, combustionHeading.paragraphStart);
  for (const row of compressorRows) {
    const cells = findCellsInRange(xml, row.start, row.end);
    if (cells.length < 2) continue;
    const label = extractText(xml, cells[0].start, cells[0].end);
    const text = compressorRemaining.get(label);
    if (text === undefined) continue;
    edits.push(buildCellReplacement(xml, cells[1], text));
    filled.push(label);
    compressorRemaining.delete(label);
  }
  for (const key of compressorRemaining.keys()) skipped.push(key);

  // ---- Turbine: anchor rows ("Stage N Nozzles"/"Buckets"/"Shroud Blocks") followed later by a
  // "Condition" cell -- same row (2 cells) or the next row's first cell (1-cell "Condition" label). ----
  const turbineByKey = new Map(aggregation.turbine.map((e) => [`${e.stage}:${e.group}`, e.text]));
  const turbineRows = findRowsInRange(xml, turbineHeading.paragraphStart, exhaustHeading.paragraphStart);
  let currentGroupKey: string | null = null;
  for (let i = 0; i < turbineRows.length; i++) {
    const cells = findCellsInRange(xml, turbineRows[i].start, turbineRows[i].end);
    if (cells.length === 0) continue;
    const firstCellText = extractText(xml, cells[0].start, cells[0].end);

    if (cells.length === 1) {
      const anchorMatch = TURBINE_GROUP_PATTERN.exec(firstCellText);
      if (anchorMatch) {
        currentGroupKey = `${anchorMatch[1]}:${normalizeTurbineGroupName(anchorMatch[2])}`;
        continue;
      }
      if (firstCellText === "Condition" && currentGroupKey) {
        const nextCells = i + 1 < turbineRows.length ? findCellsInRange(xml, turbineRows[i + 1].start, turbineRows[i + 1].end) : [];
        if (nextCells.length > 0) {
          fillTurbineTarget(currentGroupKey, nextCells[0]);
        }
      }
      continue;
    }

    if (firstCellText === "Condition" && currentGroupKey && cells.length >= 2) {
      fillTurbineTarget(currentGroupKey, cells[1]);
    }
  }

  function fillTurbineTarget(groupKey: string, targetCell: { start: number; end: number }) {
    const text = turbineByKey.get(groupKey);
    if (text === undefined) return;
    edits.push(buildCellReplacement(xml, targetCell, text));
    filled.push(`Stage ${groupKey.replace(":", " ")}`);
    turbineByKey.delete(groupKey);
  }

  for (const key of turbineByKey.keys()) skipped.push(`Stage ${key.replace(":", " ")}`);

  // Apply from the last edit backwards so earlier positions (computed against the original xml) stay valid.
  edits.sort((a, b) => b.start - a.start);
  let result = xml;
  for (const { start, end, replacement } of edits) {
    result = result.slice(0, start) + replacement + result.slice(end);
  }

  result = resetPercentPlaceholders(result, obsHeading.paragraphStart);

  return { documentXml: result, filled, skipped };
}

/**
 * Resets a Percent-Inspected cell's specific-looking placeholder value ("100%", "75%", "Add
 * Applicable %") to a blank "XX%", within the Observations section only -- scoped there rather
 * than run over the whole document since "XX%" is specifically what a Percent-Inspected cell
 * should show, not a general-purpose substitution. Only touches cells still red at this point
 * (our own fills above are already blue), so nothing this pass fills gets touched. The actual
 * red-to-black recolor happens separately, document-wide -- see recolorRemainingRedToBlack().
 */
function resetPercentPlaceholders(documentXml: string, obsHeadingStart: number): string {
  const photosHeading = findHeadingParagraph(documentXml, "Heading1", "Photos", obsHeadingStart);
  const regionEnd = photosHeading ? photosHeading.paragraphStart : documentXml.length;

  const before = documentXml.slice(0, obsHeadingStart);
  let region = documentXml.slice(obsHeadingStart, regionEnd);
  const after = documentXml.slice(regionEnd);

  region = region.replace(
    /(<w:r(?:\s[^>]*)?><w:rPr>(?:(?!<\/w:rPr>)[\s\S])*?<w:color w:val="FF0000"\/>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>)([^<]*)(<\/w:t>)/g,
    (match, open, text, close) => (PERCENT_PLACEHOLDER_PATTERN.test(text) ? `${open}XX%${close}` : match),
  );

  return before + region + after;
}

/**
 * Recolors every remaining instance of Brett's red "needs review" convention to black, across
 * the WHOLE document -- not just Observations. Adam's request: red is reserved for Brett's own
 * manual corrections when he reviews a generated report, so anything the template itself always
 * shipped in red (TOC entries, the Overall Assessment summary table, Inspection Details/TIL
 * tables, boilerplate notes, an unfilled Observations cell) needs to read as ordinary template
 * text -- black -- instead. Safe as a blanket replace because this app never writes FF0000
 * itself (only 0000FF, for real auto-filled values), so nothing generated gets touched.
 */
export function recolorRemainingRedToBlack(documentXml: string): string {
  return documentXml.replace(/FF0000/g, "000000");
}

/** Replaces a cell's paragraph content (from its first <w:p> to its last </w:p>) with a single
 *  fresh paragraph carrying `newText` in blue, reusing the original paragraph's properties
 *  (alignment/spacing) and font size where present. */
function buildCellReplacement(xml: string, cell: { start: number; end: number }, newText: string): { start: number; end: number; replacement: string } {
  const cellXml = xml.slice(cell.start, cell.end);
  const pPrMatch = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(cellXml);
  const pPr = pPrMatch ? pPrMatch[0] : "";
  const szMatch = /<w:sz w:val="(\d+)"\s*\/>/.exec(cellXml);
  const sz = szMatch ? szMatch[1] : "22";

  const firstP = xml.indexOf("<w:p", cell.start);
  const lastPClose = xml.lastIndexOf("</w:p>", cell.end);
  const start = firstP;
  const end = lastPClose + "</w:p>".length;

  const replacement =
    `<w:p>${pPr}<w:r><w:rPr><w:color w:val="0000FF"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>` +
    `<w:t xml:space="preserve">${escapeXmlText(newText)}</w:t></w:r></w:p>`;

  return { start, end, replacement };
}

function escapeXmlText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
