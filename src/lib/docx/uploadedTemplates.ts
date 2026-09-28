import { promises as fs } from "fs";
import path from "path";

/** Where manually-uploaded templates persist across sessions -- local to this machine, kept
 *  separate from ATS's shared, OneDrive-synced TEMPLATES_DIR (see api/templates/route.ts) so an
 *  ad-hoc upload never lands in the org-wide library other tech reps see. Gitignored; created on
 *  first upload (see api/templates/upload/route.ts). */
export const UPLOADED_TEMPLATES_DIR = path.join(process.cwd(), "uploaded-templates");

export interface TemplateOption {
  name: string;
  path: string;
}

export async function listUploadedTemplates(): Promise<TemplateOption[]> {
  try {
    const entries = await fs.readdir(UPLOADED_TEMPLATES_DIR, { withFileTypes: true });
    return entries
      .filter((e) => e.isFile() && e.name.toLowerCase().endsWith(".docx"))
      .map((e) => ({ name: e.name, path: path.join(UPLOADED_TEMPLATES_DIR, e.name) }))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" }));
  } catch {
    return [];
  }
}
