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

/**
 * Fills the Compressor and Turbine Observations tables' Condition cells from `aggregation`
 * (see observationsAggregator.ts). Deliberately does not touch Combustion, Exhaust, or any
 * Percent-Inspected cell -- see the plan file for why those aren't reliable to auto-fill from
 * this raw data; Percent-Inspected cells keep whichever default value the template itself
 * already shipped with (these vary per row -- not a single constant -- and a couple are literal
 * "Add Applicable %" placeholder text rather than a number; both are left exactly as-is). Leaves
 * every cell it doesn't have data for untouched, and colors filled text blue (0000FF) -- every
 * AI-generated value in the report is blue, the template's own original text is black, and red
 * marks a field the tech rep still needs to fill in or review themselves (Brett's own existing
 * convention, which predates this app and is left untouched everywhere this app doesn't have
 * real data to fill in).
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

  return { documentXml: result, filled, skipped };
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
