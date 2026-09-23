import type { RawPhotoUnit } from "./rawMdiParser";

/** Brett's formatting rule: "photos must be 3.2 by 4.3" (inches). 914400 EMU per inch. */
const PHOTO_EXTENT_CX = 3931920; // 4.3in
const PHOTO_EXTENT_CY = 2926080; // 3.2in

/** The caption-table column's width in the photo block table below -- the raw MDI caption
 *  table gets rescaled (see rescaleCaptionTable) to fit within this exactly, since different
 *  MDI software versions export their own caption table at slightly different absolute widths. */
const CAPTION_COLUMN_WIDTH = 4434;
/** Small safety margin subtracted from CAPTION_COLUMN_WIDTH before rescaling, so rounding across
 *  many nested cell widths can never push the result a hair over the real column width. */
const CAPTION_TARGET_WIDTH = CAPTION_COLUMN_WIDTH - 20;

export interface PhotoBlockOptions {
  /** New relationship id for this unit's own photo, already registered via addImageRelationship. */
  photoRelId: string;
  /** Unique per call within one assembly run, used for docPr/anchor ids Word expects to be distinct. */
  uniqueId: number;
}

/**
 * Raw MDI caption tables use an absolute-width (dxa), fixed-layout table for their outer
 * Section/Component/Observation/Comments grid. Different MDI software versions export this at
 * different native widths -- one sample measured 4375 dxa (fits fine), another 4618 dxa (wider
 * than our container column, so Word let it overflow past the cell, clipping/overlapping the
 * "Classification" column). Since w:type="dxa" widths never auto-shrink to their container the
 * way percentage widths do, every dxa width in the caption table (outer table, its cell widths,
 * cell margins, and any nested sub-tables like a measurement's own single-column table) is
 * scaled down by the same factor here, so the whole nested structure fits within our column
 * without ever touching tables that already fit.
 */
function rescaleCaptionTable(captionTableXml: string): string {
  const tblWMatch = /<w:tblW w:w="(\d+)" w:type="dxa"/.exec(captionTableXml);
  if (!tblWMatch) return captionTableXml;
  const actualWidth = parseInt(tblWMatch[1], 10);
  if (actualWidth <= CAPTION_TARGET_WIDTH) return captionTableXml;

  const scale = CAPTION_TARGET_WIDTH / actualWidth;
  return captionTableXml.replace(/w:w="(-?\d+)"/g, (match, value) => `w:w="${Math.trunc(parseInt(value, 10) * scale)}"`);
}

/**
 * Builds one top-level <w:tbl> for a single photo unit, matching the exact table shape
 * already used by the ATS template's own example photo blocks (see
 * reference/phase1-source-analysis.md, "What digging into the template found"):
 *   Row 0: cell 0 = main photo (standardized to 4.3in x 3.2in), cell 1 = the unit's own
 *          Component/Location/[measurements]/Observation/Comments caption table, rescaled to
 *          fit this column's width (fonts/sizes already match Brett's rules).
 * No row 1 -- the filename caption is dropped, matching what inspectors already do by hand.
 * No watermark overlay -- these photos come straight from the borescope's own capture
 * software, which already burns its "ADVANCED TURBINE SUPPORT" watermark into the image
 * itself, so stamping another one on top would just duplicate it.
 */
export function buildPhotoBlockXml(unit: RawPhotoUnit, opts: PhotoBlockOptions): string {
  const photoAnchorId = (opts.uniqueId * 4 + 2).toString(16).padStart(8, "0").toUpperCase();
  const photoEditId = (opts.uniqueId * 4 + 3).toString(16).padStart(8, "0").toUpperCase();
  const picId = 1000 + opts.uniqueId;

  const photo =
    `<w:r><w:rPr><w:rFonts w:ascii="Verdana" w:eastAsia="Verdana" w:hAnsi="Verdana" w:cs="Verdana"/><w:noProof/><w:sz w:val="20"/></w:rPr><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" wp14:anchorId="${photoAnchorId}" wp14:editId="${photoEditId}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
    `<wp:extent cx="${PHOTO_EXTENT_CX}" cy="${PHOTO_EXTENT_CY}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${picId + 1}" name="Picture ${picId + 1}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="${picId + 1}" name="Picture ${picId + 1}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${opts.photoRelId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${PHOTO_EXTENT_CX}" cy="${PHOTO_EXTENT_CY}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

  return (
    `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid-Copy"/><w:tblW w:w="10654" w:type="dxa"/><w:tblInd w:w="108" w:type="dxa"/>` +
    `<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>` +
    `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>` +
    `<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="6220"/><w:gridCol w:w="4434"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:trHeight w:val="457"/></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="6220" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:ind w:right="-738"/></w:pPr>${photo}</w:p></w:tc>` +
    // A table cell containing a <w:tbl> must still end with a <w:p> -- OOXML requires every
    // w:tc to close on a paragraph, even an empty one, after any nested table. Omitting it
    // parses as well-formed XML (generic parsers don't catch it) but Word itself refuses to
    // open the file, reporting it as corrupted.
    `<w:tc><w:tcPr><w:tcW w:w="${CAPTION_COLUMN_WIDTH}" w:type="dxa"/></w:tcPr>${rescaleCaptionTable(unit.captionTableXml)}<w:p/></w:tc>` +
    `</w:tr></w:tbl>`
  );
}
