import { NextRequest, NextResponse } from "next/server";
import { promises as fs } from "fs";
import path from "path";
import { loadDocx } from "@/lib/docx/zip";
import { detectTemplateFamily } from "@/lib/docx/templateFamilies";

/** ATS's shared library of blank cover-page templates -- one per unit/job type (35 and
 *  growing), maintained outside this project so new templates show up here without a deploy.
 *  See templateFamilies.ts for which of these the assembler actually knows how to fill in;
 *  picking one it doesn't recognize fails clearly at scan time rather than guessing wrong. */
const TEMPLATES_DIR = "C:\\Users\\akanafani\\OneDrive - Allied Power Group\\Strategy - Files\\ATS Reporting\\ATS Cover Page Templates";

interface TemplateOption {
  name: string;
  path: string;
}

/** Legacy Frame 7 unit variants (Frame 7, 7C, 7E, 7EA) share the same base report structure --
 *  grouped here so a raw file named for any one of them matches templates named for any other,
 *  instead of an exact-token match hiding a 7E raw file's own 7EA templates. 7FA is deliberately
 *  its own family (different Photos/Observations layout -- see templateFamilies.ts's
 *  FRAME_BEARING_FAMILY), so it's not in this group. No "Frame 7"/"7C"/plain-"7E" template exists
 *  in the library yet -- this just means one will be picked up correctly the day ATS adds it. */
const UNIT_TYPE_ALIAS_GROUPS: string[][] = [["7EA", "7E", "7C", "FRAME 7"]];

function unitAliasGroup(unit: string): string[] {
  const upper = unit.toUpperCase();
  return UNIT_TYPE_ALIAS_GROUPS.find((group) => group.includes(upper)) ?? [upper];
}

/** Curated list shown on the template-picker step, so the user doesn't have to browse folders
 *  and risk picking a finished sample report instead of a blank starter template. Only lists
 *  .docx files directly in the folder (no subfolders currently expected there).
 *
 *  With `?unit=`, narrows `templates` down further to the ones this raw file's own unit type
 *  actually needs and this app can actually fill in -- the full 35-template library mixes every
 *  job type ATS does (shim removal, TIL-specific one-offs, other frames) with every unit family,
 *  and most of that is irrelevant noise once you already know which raw file you're building
 *  from. Two filters, in order:
 *   1. Name match: only templates whose own name starts with the same unit-type token as the
 *      raw file (e.g. "7EA_Full_...docx" -> keep templates starting with "7EA "). Every raw
 *      export and every template ATS maintains follows this naming convention.
 *   2. Automation match: of those, only the ones detectTemplateFamily() actually recognizes --
 *      the name match alone still includes structurally unrelated one-off job templates (Shim
 *      Removal, TIL-specific inspections) that share the unit prefix but aren't borescope
 *      reports at all.
 *  `fullList` is always returned too (unfiltered, exactly what a bare GET would give), so the
 *  client can offer an escape hatch back to the complete library when this heuristic is wrong. */
export async function GET(request: NextRequest) {
  const unit = request.nextUrl.searchParams.get("unit");

  try {
    const entries = await fs.readdir(TEMPLATES_DIR, { withFileTypes: true });
    const fullList: TemplateOption[] = entries
      // "Sample Report" files are finished examples kept for style reference, not blank
      // templates -- excluded here since assembling into one throws (no fillable placeholders).
      .filter((e) => e.isFile() && e.name.toLowerCase().endsWith(".docx") && !e.name.toLowerCase().includes("sample report"))
      .map((e) => ({ name: e.name, path: path.join(TEMPLATES_DIR, e.name) }))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" }));

    if (!unit) {
      return NextResponse.json({ templates: fullList, fullList, unit: null, matchedCount: 0 });
    }

    const aliasGroup = unitAliasGroup(unit);
    const unitMatches = fullList.filter((t) => aliasGroup.some((alias) => t.name.toUpperCase().startsWith(`${alias} `)));

    const automated: TemplateOption[] = [];
    for (const t of unitMatches) {
      try {
        const buffer = await fs.readFile(t.path);
        const loaded = await loadDocx(buffer);
        if (detectTemplateFamily(loaded.documentXml)) automated.push(t);
      } catch {
        // Unreadable/corrupt template -- skip it rather than fail the whole list.
      }
    }

    return NextResponse.json({ templates: automated, fullList, unit, matchedCount: unitMatches.length });
  } catch (err) {
    return NextResponse.json({ templates: [], fullList: [], error: err instanceof Error ? err.message : "Failed to list templates" });
  }
}
