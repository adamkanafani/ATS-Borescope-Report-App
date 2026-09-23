import { NextResponse } from "next/server";
import { promises as fs } from "fs";
import path from "path";

/** ATS's shared library of blank cover-page templates -- one per unit/job type (35 and
 *  growing), maintained outside this project so new templates show up here without a deploy.
 *  See templateFamilies.ts for which of these the assembler actually knows how to fill in;
 *  picking one it doesn't recognize fails clearly at scan time rather than guessing wrong. */
const TEMPLATES_DIR = "C:\\Users\\akanafani\\OneDrive - Allied Power Group\\Strategy - Files\\ATS Reporting\\ATS Cover Page Templates";

/** Curated list shown on the template-picker step, so the user doesn't have to browse folders
 *  and risk picking a finished sample report instead of a blank starter template. Only lists
 *  .docx files directly in the folder (no subfolders currently expected there). */
export async function GET() {
  try {
    const entries = await fs.readdir(TEMPLATES_DIR, { withFileTypes: true });
    const templates = entries
      // "Sample Report" files are finished examples kept for style reference, not blank
      // templates -- excluded here since assembling into one throws (no fillable placeholders).
      .filter((e) => e.isFile() && e.name.toLowerCase().endsWith(".docx") && !e.name.toLowerCase().includes("sample report"))
      .map((e) => ({ name: e.name, path: path.join(TEMPLATES_DIR, e.name) }))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" }));
    return NextResponse.json({ templates });
  } catch (err) {
    return NextResponse.json({ templates: [], error: err instanceof Error ? err.message : "Failed to list templates" });
  }
}
