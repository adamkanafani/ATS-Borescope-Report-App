import { DOMParser } from "@xmldom/xmldom";
import type { LoadedDocx } from "./zip";
import { firstChildTag, textOf } from "./domHelpers";

export interface ContentControlInfo {
  index: number;
  id: string | null;
  tag: string | null;
  alias: string | null;
  kind: "dropDownList" | "text" | "other";
  /** displayText of each <w:listItem>, for dropdowns only. */
  options: string[];
  currentText: string;
  isPlaceholder: boolean;
}

/**
 * Reads every Word content control (<w:sdt>) in a document -- ATS's template uses these for
 * every fillable field: inspection scope, TIL full/limited/none, unit-rotated yes/no, etc.
 * (see reference/phase1-source-analysis.md, "Docx structural findings"). Read-only: this
 * doesn't fill anything in, it's for mapping out what's fillable before we design that step.
 */
export function listContentControls(loaded: LoadedDocx): ContentControlInfo[] {
  const doc = new DOMParser().parseFromString(loaded.documentXml, "text/xml");
  const sdts = Array.from(doc.getElementsByTagName("w:sdt")) as unknown as Element[];

  return sdts.map((sdt, index) => {
    const sdtPr = firstChildTag(sdt, "w:sdtPr");
    const sdtContent = firstChildTag(sdt, "w:sdtContent");

    const id = sdtPr ? attrOf(firstChildTag(sdtPr, "w:id"), "w:val") : null;
    const tag = sdtPr ? attrOf(firstChildTag(sdtPr, "w:tag"), "w:val") : null;
    const alias = sdtPr ? attrOf(firstChildTag(sdtPr, "w:alias"), "w:val") : null;

    let kind: ContentControlInfo["kind"] = "other";
    const options: string[] = [];
    const dropdown = sdtPr ? firstChildTag(sdtPr, "w:dropDownList") : null;
    if (dropdown) {
      kind = "dropDownList";
      const items = dropdown.getElementsByTagName("w:listItem");
      for (let i = 0; i < items.length; i++) {
        const display = items[i].getAttribute("w:displayText") || items[i].getAttribute("w:value");
        if (display) options.push(display);
      }
    } else if (sdtPr && firstChildTag(sdtPr, "w:text")) {
      kind = "text";
    }

    const isPlaceholder = !!(sdtPr && firstChildTag(sdtPr, "w:showingPlcHdr"));
    const currentText = sdtContent ? textOf(sdtContent) : "";

    return { index, id, tag, alias, kind, options, currentText, isPlaceholder };
  });
}

function attrOf(el: Element | null, attr: string): string | null {
  return el ? el.getAttribute(attr) : null;
}
