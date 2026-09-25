import { NextRequest, NextResponse } from "next/server";
import { saveDocx } from "@/lib/docx/zip";
import { assembleReport } from "@/lib/docx/templateAssembler";
import type { PhotosSection } from "@/lib/docx/photosSectionMap";
import { getSession } from "@/lib/reviewSession";

export async function POST(request: NextRequest) {
  const form = await request.formData();
  const sessionId = form.get("sessionId");
  const overridesRaw = form.get("overrides");
  const excludedSectionsRaw = form.get("excludedSections");
  if (typeof sessionId !== "string") return NextResponse.json({ error: "sessionId is required" }, { status: 400 });

  const session = getSession(sessionId);
  if (!session) {
    return NextResponse.json({ error: "This review session has expired -- please scan again." }, { status: 410 });
  }

  try {
    const overrides = new Map<number, PhotosSection | null>();
    if (typeof overridesRaw === "string") {
      const overridesInput: Record<string, PhotosSection | null> = JSON.parse(overridesRaw);
      for (const [key, value] of Object.entries(overridesInput)) {
        overrides.set(Number(key), value);
      }
    }

    const excludedSections = new Set<PhotosSection>();
    if (typeof excludedSectionsRaw === "string") {
      const excludedInput: PhotosSection[] = JSON.parse(excludedSectionsRaw);
      for (const section of excludedInput) excludedSections.add(section);
    }

    const result = assembleReport(session.template, session.units, {
      overrides,
      excludedSections,
    });
    const buffer = await saveDocx(session.template);

    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": 'attachment; filename="Assembled-Report.docx"',
        "X-Inserted-Count": String(result.inserted),
        "X-Unassigned-Count": String(result.unassigned.length),
      },
    });
  } catch (err) {
    console.error("generate failed:", err);
    return NextResponse.json({ error: err instanceof Error ? err.message : "Generation failed" }, { status: 500 });
  }
}
