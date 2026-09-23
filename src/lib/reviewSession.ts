import { randomUUID } from "crypto";
import type { LoadedDocx } from "./docx/zip";
import type { RawPhotoUnit } from "./docx/rawMdiParser";

export interface ReviewSession {
  template: LoadedDocx;
  units: RawPhotoUnit[];
  createdAt: number;
}

const SESSION_TTL_MS = 60 * 60 * 1000; // 1 hour -- a scan is cheap to redo, so no need to keep longer.
const sessions = new Map<string, ReviewSession>();

export function createSession(template: LoadedDocx, units: RawPhotoUnit[]): string {
  const id = randomUUID();
  sessions.set(id, { template, units, createdAt: Date.now() });
  return id;
}

export function getSession(id: string): ReviewSession | undefined {
  return sessions.get(id);
}

export function deleteSession(id: string): void {
  sessions.delete(id);
}

// Guard against re-registering this on every module reload in dev.
declare global {
  var __atsReviewSessionSweepStarted: boolean | undefined;
}

if (!globalThis.__atsReviewSessionSweepStarted) {
  globalThis.__atsReviewSessionSweepStarted = true;
  setInterval(() => {
    const cutoff = Date.now() - SESSION_TTL_MS;
    for (const [id, session] of sessions) {
      if (session.createdAt < cutoff) sessions.delete(id);
    }
  }, 10 * 60 * 1000).unref?.();
}
