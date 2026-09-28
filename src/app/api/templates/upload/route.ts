import { NextRequest, NextResponse } from "next/server";
import { promises as fs } from "fs";
import path from "path";
import { UPLOADED_TEMPLATES_DIR } from "@/lib/docx/uploadedTemplates";

/** Picks a filename that won't clobber an earlier upload -- appends " (2)", " (3)", etc. until
 *  one doesn't already exist, rather than silently overwriting a same-named file from before. */
async function uniqueName(originalName: string): Promise<string> {
  const ext = path.extname(originalName);
  const base = path.basename(originalName, ext);
  let candidate = originalName;
  for (let n = 2; ; n++) {
    try {
      await fs.access(path.join(UPLOADED_TEMPLATES_DIR, candidate));
      candidate = `${base} (${n})${ext}`;
    } catch {
      return candidate;
    }
  }
}

/** Saves a manually-uploaded template to disk (see uploadedTemplates.ts) so it's available again
 *  next time without re-uploading -- entirely local, no network involved. */
export async function POST(request: NextRequest) {
  const form = await request.formData();
  const file = form.get("file");
  if (!(file instanceof File)) {
    return NextResponse.json({ error: "No file provided." }, { status: 400 });
  }
  if (!file.name.toLowerCase().endsWith(".docx")) {
    return NextResponse.json({ error: "Only .docx files can be uploaded as a template." }, { status: 400 });
  }

  try {
    await fs.mkdir(UPLOADED_TEMPLATES_DIR, { recursive: true });
    const name = await uniqueName(file.name);
    const savedPath = path.join(UPLOADED_TEMPLATES_DIR, name);
    await fs.writeFile(savedPath, Buffer.from(await file.arrayBuffer()));
    return NextResponse.json({ name, path: savedPath });
  } catch (err) {
    return NextResponse.json({ error: err instanceof Error ? err.message : "Upload failed" }, { status: 500 });
  }
}
