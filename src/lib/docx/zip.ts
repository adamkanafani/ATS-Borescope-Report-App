import JSZip from "jszip";

export interface LoadedDocx {
  zip: JSZip;
  documentXml: string;
  /** The raw word/_rels/document.xml.rels text -- kept alongside the parsed map so new
   *  relationships added via addImageRelationship() can be serialized back out. */
  relsXml: string;
  /** Relationship id (e.g. "rId5") -> zip-absolute path (e.g. "word/media/image1.jpg"). */
  relationships: Map<string, string>;
  /** Zip-absolute path -> file bytes, for every image relationship target. */
  mediaFiles: Map<string, Buffer>;
}

const IMAGE_EXTENSION = /\.(jpg|jpeg|png|emf|bmp|gif|tiff?)$/i;
const IMAGE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/**
 * Loads a .docx, parsing word/document.xml and its relationships and pulling every
 * referenced image into memory. Doesn't touch headers/footers/other parts -- this app
 * only ever needs the main body.
 */
export async function loadDocx(buffer: Buffer): Promise<LoadedDocx> {
  const zip = await JSZip.loadAsync(buffer);

  const documentEntry = zip.file("word/document.xml");
  if (!documentEntry) {
    throw new Error("Not a valid .docx: missing word/document.xml");
  }
  const documentXml = await documentEntry.async("text");

  const relsEntry = zip.file("word/_rels/document.xml.rels");
  const relsXml = relsEntry ? await relsEntry.async("text") : "";
  const relationships = parseRelationships(relsXml, "word/document.xml");

  const mediaFiles = new Map<string, Buffer>();
  for (const zipPath of relationships.values()) {
    if (!IMAGE_EXTENSION.test(zipPath) || mediaFiles.has(zipPath)) continue;
    const entry = zip.file(zipPath);
    if (!entry) continue;
    mediaFiles.set(zipPath, await entry.async("nodebuffer"));
  }

  return { zip, documentXml, relsXml, relationships, mediaFiles };
}

export async function saveDocx(loaded: LoadedDocx): Promise<Buffer> {
  loaded.zip.file("word/document.xml", loaded.documentXml);
  loaded.zip.file("word/_rels/document.xml.rels", loaded.relsXml);
  return loaded.zip.generateAsync({ type: "nodebuffer" });
}

/**
 * Registers a new image as a fresh word/media/*.<ext> zip entry plus a fresh relationship
 * in word/_rels/document.xml.rels, and returns the new relationship id to embed via
 * <a:blip r:embed="...">. Mutates `loaded` in place (zip, relsXml, relationships, mediaFiles).
 */
export function addImageRelationship(loaded: LoadedDocx, imageBytes: Buffer, extension: string, namePrefix = "atsInserted"): string {
  const relId = `rId${nextRelationshipNumber(loaded.relationships)}`;
  const mediaName = `${namePrefix}${relId}.${extension}`;
  const zipPath = `word/media/${mediaName}`;

  const relationshipTag = `<Relationship Id="${relId}" Type="${IMAGE_RELATIONSHIP_TYPE}" Target="media/${mediaName}"/>`;
  if (loaded.relsXml.includes("</Relationships>")) {
    loaded.relsXml = loaded.relsXml.replace("</Relationships>", `${relationshipTag}</Relationships>`);
  } else {
    throw new Error("word/_rels/document.xml.rels doesn't look like a valid relationships file (no </Relationships>).");
  }

  loaded.relationships.set(relId, zipPath);
  loaded.mediaFiles.set(zipPath, imageBytes);
  loaded.zip.file(zipPath, imageBytes);

  return relId;
}

function nextRelationshipNumber(relationships: Map<string, string>): number {
  let max = 0;
  for (const id of relationships.keys()) {
    const match = /^rId(\d+)$/.exec(id);
    if (match) max = Math.max(max, parseInt(match[1], 10));
  }
  return max + 1;
}

/**
 * Parses a .rels file into id -> zip-absolute path, resolving each Target against the
 * part that owns the .rels file. A leading "/" on Target means "resolve from the package
 * root" (OOXML allows this, and the raw MDI export actually uses it -- e.g. "/media/image.jpg"
 * -- instead of the more common "media/image.jpg" relative to the owning part's own folder,
 * which is what the ATS templates use).
 */
function parseRelationships(relsXml: string, basePartPath: string): Map<string, string> {
  const map = new Map<string, string>();
  const baseDir = basePartPath.includes("/") ? basePartPath.slice(0, basePartPath.lastIndexOf("/")) : "";
  const relationshipTag = /<Relationship\b[^>]*\/>/g;
  let match: RegExpExecArray | null;
  while ((match = relationshipTag.exec(relsXml))) {
    const tag = match[0];
    const id = /\bId="([^"]+)"/.exec(tag)?.[1];
    const target = /\bTarget="([^"]+)"/.exec(tag)?.[1];
    if (!id || !target) continue;
    map.set(id, resolveTarget(baseDir, target));
  }
  return map;
}

function resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  const segments = (baseDir ? baseDir.split("/") : []).concat(target.split("/"));
  const resolved: string[] = [];
  for (const segment of segments) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") resolved.pop();
    else resolved.push(segment);
  }
  return resolved.join("/");
}
