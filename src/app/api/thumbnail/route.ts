import { NextRequest, NextResponse } from "next/server";
import { getSession } from "@/lib/reviewSession";

const CONTENT_TYPE_BY_EXT: Record<string, string> = {
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  png: "image/png",
  bmp: "image/bmp",
};

export async function GET(request: NextRequest) {
  const sessionId = request.nextUrl.searchParams.get("sessionId");
  const indexParam = request.nextUrl.searchParams.get("index");
  if (!sessionId || indexParam === null) {
    return NextResponse.json({ error: "sessionId and index are required" }, { status: 400 });
  }

  const session = getSession(sessionId);
  if (!session) {
    return NextResponse.json({ error: "This review session has expired -- please scan again." }, { status: 410 });
  }

  const unit = session.units[Number(indexParam)];
  if (!unit || !unit.imageBytes) {
    return NextResponse.json({ error: "No image for this unit" }, { status: 404 });
  }

  const ext = unit.imagePath?.split(".").pop()?.toLowerCase() ?? "jpg";
  const contentType = CONTENT_TYPE_BY_EXT[ext] ?? "image/jpeg";

  return new NextResponse(new Uint8Array(unit.imageBytes), {
    headers: {
      "Content-Type": contentType,
      "Cache-Control": "private, max-age=3600",
    },
  });
}
