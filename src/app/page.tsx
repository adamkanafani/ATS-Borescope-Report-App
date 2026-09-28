"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { PhotosSection } from "@/lib/docx/photosSectionMap";

interface Measurement {
  label: string;
  value: string;
}

interface UnitSummary {
  index: number;
  component: string | null;
  location: string | null;
  measurements: Measurement[];
  observation: string | null;
  comments: string | null;
  filenameCaption: string | null;
  hasImage: boolean;
  section: PhotosSection | null;
}

interface ScanResponse {
  sessionId: string;
  units: UnitSummary[];
  bySection: Record<PhotosSection, number>;
  unassignedCount: number;
  photosSubsections: PhotosSection[];
}

interface BrowseResponse {
  dir: string | null;
  parent: string | null;
  folders: string[];
  files: string[];
  isDriveList: boolean;
  error?: string;
}

const DRIVES_ROOT = "__DRIVES__";
const ONEDRIVE_ROOT = "__ONEDRIVE__";

/* Flat white folder glyph -- a plain emoji renders as a fixed-color image on Windows and can't
   be recolored with CSS, so the folder icon is inline SVG instead. */
function FolderIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 20 16" fill="none" aria-hidden="true">
      <path
        d="M1 2.5C1 1.67 1.67 1 2.5 1H7.2C7.68 1 8.13 1.23 8.42 1.62L9.4 3H17.5C18.33 3 19 3.67 19 4.5V13.5C19 14.33 18.33 15 17.5 15H2.5C1.67 15 1 14.33 1 13.5V2.5Z"
        fill="#ffffff"
      />
    </svg>
  );
}

/* Trash bin glyph, shown on the sidebar's Trash tab -- same reasoning as FolderIcon above. */
function TrashIcon() {
  return (
    <svg width="14" height="16" viewBox="0 0 14 16" fill="none" aria-hidden="true">
      <path
        d="M1 3.5H13M5 3.5V1.8C5 1.36 5.36 1 5.8 1H8.2C8.64 1 9 1.36 9 1.8V3.5M2.5 3.5L3.1 14.2C3.13 14.65 3.5 15 3.95 15H10.05C10.5 15 10.87 14.65 10.9 14.2L11.5 3.5"
        stroke="currentColor"
        strokeWidth="1.3"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

export default function Home() {
  const [rawPath, setRawPath] = useState<string | null>(null);
  // Remembers where the raw-file picker was last browsing, so backing out of a later step (which
  // unmounts it) reopens it there instead of always restarting at the home directory.
  const [rawBrowseDir, setRawBrowseDir] = useState<string | null>(null);
  const [templatePath, setTemplatePath] = useState<string | null>(null);
  const [scan, setScan] = useState<ScanResponse | null>(null);
  const [scanning, setScanning] = useState(false);
  const [scanError, setScanError] = useState<string | null>(null);
  const [selectedSection, setSelectedSection] = useState<PhotosSection | "Unassigned" | "Trash">("Compressor Section");
  const [overrides, setOverrides] = useState<Record<number, PhotosSection | null>>({});
  // Trashing a unit never touches `overrides` -- it's tracked separately so recovering it just
  // means removing it from this set, and it goes right back to whatever section (classifier
  // default or a prior manual reassignment) it already had.
  const [deletedIndices, setDeletedIndices] = useState<Set<number>>(new Set());
  // Shared between ReadyToGenerateScreen and ReviewScreen's GenerateSectionsModal so scoping the
  // report down on one screen (e.g. Compressor Only) isn't silently lost by detouring through
  // the other before actually generating.
  const [excludedSections, setExcludedSections] = useState<Set<PhotosSection>>(new Set());
  const [generating, setGenerating] = useState(false);
  const [generateError, setGenerateError] = useState<string | null>(null);
  // Whether the tech rep chose to open the full per-photo review screen instead of generating
  // straight from the "ready to generate" step -- see ReadyToGenerateScreen/ReviewScreen below.
  const [reviewing, setReviewing] = useState(false);

  function toggleExcludedSection(section: PhotosSection) {
    setExcludedSections((prev) => {
      const next = new Set(prev);
      if (next.has(section)) next.delete(section);
      else next.add(section);
      return next;
    });
  }

  async function runScan(raw: string, tmpl: string) {
    setScanning(true);
    setScanError(null);
    try {
      const res = await fetch("/api/scan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rawPath: raw, templatePath: tmpl }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Scan failed");
      setScan(data);
      setOverrides({});
      setDeletedIndices(new Set());
      setExcludedSections(new Set());
    } catch (err) {
      setScanError(err instanceof Error ? err.message : "Scan failed");
    } finally {
      setScanning(false);
    }
  }

  function resetFiles() {
    setRawPath(null);
    setTemplatePath(null);
    setScan(null);
    setScanError(null);
    setOverrides({});
    setDeletedIndices(new Set());
    setExcludedSections(new Set());
    setReviewing(false);
    setSelectedSection("Compressor Section");
  }

  async function postGenerate(sessionId: string) {
    // Trashed units are excluded the same way "Leave excluded" is (an explicit null override) --
    // computed here rather than stored in `overrides` itself, so recovering a trashed unit
    // before generating restores whatever section it actually had.
    const overridesWithTrash: Record<number, PhotosSection | null> = { ...overrides };
    for (const index of deletedIndices) overridesWithTrash[index] = null;

    const formData = new FormData();
    formData.set("sessionId", sessionId);
    formData.set("overrides", JSON.stringify(overridesWithTrash));
    formData.set("excludedSections", JSON.stringify([...excludedSections]));
    return fetch("/api/generate", { method: "POST", body: formData });
  }

  async function generateReport() {
    if (!scan || !rawPath || !templatePath) return;
    setGenerating(true);
    setGenerateError(null);
    try {
      let res = await postGenerate(scan.sessionId);
      if (res.status === 410) {
        // The server-side review session lives only in memory (see reviewSession.ts) and is
        // gone if the local dev server restarted since scanning -- but the raw file and template
        // are still the same files on disk, so silently rescan and retry once instead of forcing
        // a trip back through the file pickers (which would also lose overrides/deletedIndices).
        const rescan = await fetch("/api/scan", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ rawPath, templatePath }),
        });
        const rescanData = await rescan.json();
        if (!rescan.ok) throw new Error(rescanData.error || "The review session expired and could not be recovered -- please rescan.");
        setScan(rescanData);
        res = await postGenerate(rescanData.sessionId);
      }
      if (!res.ok) {
        const data = await res.json();
        throw new Error(data.error || "Generate failed");
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Assembled-Report.docx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
    } catch (err) {
      setGenerateError(err instanceof Error ? err.message : "Generate failed");
    } finally {
      setGenerating(false);
    }
  }

  if (!rawPath) {
    return (
      <FilePicker
        title="Select the raw borescope file"
        instructions="Pick the .docx generated directly by the borescope for this inspection."
        extensions="docx"
        onChosen={setRawPath}
        enableUseThisFolder
        initialDir={rawBrowseDir}
        onDirChange={setRawBrowseDir}
      />
    );
  }

  if (!templatePath) {
    return (
      <TemplateStep
        rawPath={rawPath}
        onChosen={(p) => {
          setTemplatePath(p);
          runScan(rawPath, p);
        }}
        onBack={() => setRawPath(null)}
      />
    );
  }

  if (scanning || (!scan && !scanError)) {
    return (
      <main className="placeholder">
        <h1>Scanning...</h1>
        <p>Parsing the raw borescope file and classifying every photo.</p>
      </main>
    );
  }

  if (scanError || !scan) {
    return (
      <main className="placeholder">
        <h1>Scan failed</h1>
        <p>{scanError}</p>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={() => runScan(rawPath, templatePath)}>Try again</button>
          <button className="secondary" onClick={resetFiles}>
            Choose different files
          </button>
        </div>
      </main>
    );
  }

  if (!reviewing) {
    return (
      <ReadyToGenerateScreen
        rawPath={rawPath}
        scan={scan}
        excludedSections={excludedSections}
        onToggleSection={toggleExcludedSection}
        generating={generating}
        generateError={generateError}
        onGenerate={generateReport}
        onReview={() => setReviewing(true)}
        onChangeFiles={resetFiles}
      />
    );
  }

  return (
    <ReviewScreen
      rawPath={rawPath}
      scan={scan}
      selectedSection={selectedSection}
      onSelectSection={setSelectedSection}
      overrides={overrides}
      onOverride={(index, section) => setOverrides((prev) => ({ ...prev, [index]: section }))}
      deletedIndices={deletedIndices}
      onDelete={(index) => setDeletedIndices((prev) => new Set(prev).add(index))}
      onRecover={(index) =>
        setDeletedIndices((prev) => {
          const next = new Set(prev);
          next.delete(index);
          return next;
        })
      }
      excludedSections={excludedSections}
      onToggleSection={toggleExcludedSection}
      generating={generating}
      generateError={generateError}
      onGenerate={generateReport}
      onChangeFiles={resetFiles}
    />
  );
}

function effectiveSection(unit: UnitSummary, overrides: Record<number, PhotosSection | null>): PhotosSection | null {
  return unit.index in overrides ? overrides[unit.index] : unit.section;
}

function ReviewScreen({
  rawPath,
  scan,
  selectedSection,
  onSelectSection,
  overrides,
  onOverride,
  deletedIndices,
  onDelete,
  onRecover,
  excludedSections,
  onToggleSection,
  generating,
  generateError,
  onGenerate,
  onChangeFiles,
}: {
  rawPath: string;
  scan: ScanResponse;
  selectedSection: PhotosSection | "Unassigned" | "Trash";
  onSelectSection: (s: PhotosSection | "Unassigned" | "Trash") => void;
  overrides: Record<number, PhotosSection | null>;
  onOverride: (index: number, section: PhotosSection | null) => void;
  deletedIndices: Set<number>;
  onDelete: (index: number) => void;
  onRecover: (index: number) => void;
  excludedSections: Set<PhotosSection>;
  onToggleSection: (section: PhotosSection) => void;
  generating: boolean;
  generateError: string | null;
  onGenerate: () => void;
  onChangeFiles: () => void;
}) {
  const [showGenerateModal, setShowGenerateModal] = useState(false);
  const live = useMemo(() => {
    const counts = Object.fromEntries(scan.photosSubsections.map((name) => [name, 0])) as Record<PhotosSection, number>;
    let unassigned = 0;
    for (const unit of scan.units) {
      if (deletedIndices.has(unit.index)) continue;
      const section = effectiveSection(unit, overrides);
      if (section) counts[section]++;
      else unassigned++;
    }
    return { counts, unassigned };
  }, [scan, overrides, deletedIndices]);

  const visibleUnits = scan.units.filter((unit) => {
    if (selectedSection === "Trash") return deletedIndices.has(unit.index);
    if (deletedIndices.has(unit.index)) return false;
    const section = effectiveSection(unit, overrides);
    return selectedSection === "Unassigned" ? section === null : section === selectedSection;
  });

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1 style={{ fontSize: 16, fontWeight: 700, color: "#ffffff" }}>ATS Borescope Report Builder</h1>
        <span className="job-path">{rawPath}</span>
        <button className="secondary" onClick={onChangeFiles}>
          Change Files
        </button>
      </header>
      <div className="app-body">
        <nav className="sidebar">
          <div className="sidebar-scroll">
            {scan.photosSubsections.map((name) => (
              <button
                key={name}
                className={`section-item ${selectedSection === name ? "active" : ""}`}
                onClick={() => onSelectSection(name)}
              >
                <div className="title-row">
                  <span>{name}</span>
                </div>
                <div className="confidence-tag">
                  {live.counts[name]} photo{live.counts[name] === 1 ? "" : "s"}
                </div>
              </button>
            ))}
            <button
              className={`section-item ${selectedSection === "Unassigned" ? "active" : ""}`}
              onClick={() => onSelectSection("Unassigned")}
            >
              <div className="title-row">
                <span>Unassigned</span>
                {live.unassigned > 0 && <span className="status-badge needs-attention">Needs Attention</span>}
              </div>
              <div className="confidence-tag">
                {live.unassigned} photo{live.unassigned === 1 ? "" : "s"}
              </div>
            </button>
            <button
              className={`section-item ${selectedSection === "Trash" ? "active" : ""}`}
              onClick={() => onSelectSection("Trash")}
            >
              <div className="title-row">
                <span>Trash</span>
                <TrashIcon />
              </div>
              <div className="confidence-tag">
                {deletedIndices.size} photo{deletedIndices.size === 1 ? "" : "s"}
              </div>
            </button>
          </div>
        </nav>
        <main className="main-panel">
          <h2 style={{ marginBottom: 8 }}>{selectedSection}</h2>
          {selectedSection === "Trash" ? (
            <p className="section-reason">
              Deleted photos land here instead of being removed outright -- they&apos;re left out of the
              generated report either way, but you can still recover one back to wherever it was before.
            </p>
          ) : selectedSection === "Unassigned" && live.unassigned > 0 ? (
            <p className="section-reason">
              The classifier couldn&apos;t place these by component name. Pick a section for each one below, or
              leave excluded to omit it from the generated report.
            </p>
          ) : null}
          {visibleUnits.length === 0 ? (
            <p>No photos here.</p>
          ) : (
            <ul className="unit-list">
              {visibleUnits.map((unit) => (
                    <li key={unit.index} className="unit-row">
                      {unit.hasImage ? (
                        <img
                          className="unit-row-thumb"
                          src={`/api/thumbnail?sessionId=${scan.sessionId}&index=${unit.index}`}
                          alt={unit.component ?? "photo"}
                          loading="lazy"
                        />
                      ) : (
                        <div className="unit-row-thumb unit-row-thumb-missing">No photo</div>
                      )}
                      <div className="unit-row-body">
                        <div className="unit-row-main">
                          <strong>{unit.component ?? "(unknown component)"}</strong>
                          {unit.location ? ` — ${unit.location}` : ""}
                        </div>
                        {unit.measurements.length > 0 && (
                          <div className="unit-row-measurements">
                            {unit.measurements.map((m) => `${m.label}: ${m.value}`).join(", ")}
                          </div>
                        )}
                        <div className="unit-row-observation">{unit.observation}</div>
                        {unit.comments && <div className="unit-row-comments">{unit.comments}</div>}
                        {selectedSection === "Trash" ? (
                          <button type="button" className="secondary unit-row-recover" onClick={() => onRecover(unit.index)}>
                            Recover
                          </button>
                        ) : (
                          <select
                            value={overrides[unit.index] ?? unit.section ?? ""}
                            onChange={(e) => onOverride(unit.index, (e.target.value as PhotosSection) || null)}
                          >
                            <option value="">Leave excluded</option>
                            {scan.photosSubsections.map((name) => (
                              <option key={name} value={name}>
                                {name}
                              </option>
                            ))}
                          </select>
                        )}
                      </div>
                      {selectedSection !== "Trash" && (
                        <button
                          type="button"
                          className="unit-row-delete"
                          onClick={() => onDelete(unit.index)}
                          title="Delete this photo (moves it to Trash)"
                          aria-label="Delete this photo"
                        >
                          ×
                        </button>
                      )}
                    </li>
              ))}
            </ul>
          )}
        </main>
      </div>
      <footer className="app-header">
        {generateError && <span className="folder-picker-error">{generateError}</span>}
        <button className="generate-report" disabled={generating} onClick={() => setShowGenerateModal(true)}>
          {generating ? "Generating..." : "Generate Report"}
        </button>
      </footer>
      {showGenerateModal && (
        <GenerateSectionsModal
          sections={scan.photosSubsections}
          excludedSections={excludedSections}
          onToggleSection={onToggleSection}
          onCancel={() => setShowGenerateModal(false)}
          onConfirm={() => {
            setShowGenerateModal(false);
            onGenerate();
          }}
        />
      )}
    </div>
  );
}

/** The fast path straight from "template picked" to "report downloaded" -- no per-photo review
 *  required. Tech reps are already fast in Word and don't want to review every photo/section
 *  inside the app first; the classifier already placed everything, so Generate here just needs
 *  the inspection scope (which sections this report should even include -- e.g. "Compressor
 *  Only" unchecks everything else) before handing off. "Review photos in detail first" is the
 *  escape hatch to the full ReviewScreen (sidebar, reassignment, delete/trash/recover) for anyone
 *  who wants it -- that screen is unchanged, just no longer mandatory. */
function ReadyToGenerateScreen({
  rawPath,
  scan,
  excludedSections,
  onToggleSection,
  generating,
  generateError,
  onGenerate,
  onReview,
  onChangeFiles,
}: {
  rawPath: string;
  scan: ScanResponse;
  excludedSections: Set<PhotosSection>;
  onToggleSection: (section: PhotosSection) => void;
  generating: boolean;
  generateError: string | null;
  onGenerate: () => void;
  onReview: () => void;
  onChangeFiles: () => void;
}) {
  return (
    <div className="folder-picker">
      <div className="folder-picker-card">
        <div className="app-logo">
          <h1>Ready to generate</h1>
        </div>
        <p className="folder-picker-subtitle">{rawPath}</p>
        <p className="folder-picker-subtitle">
          Every photo has already been classified and placed. Uncheck a section below to leave it
          (and its Observations table) out entirely -- for example, uncheck everything but
          Compressor Section for a compressor-only inspection.
        </p>
        <div className="modal-checklist">
          {scan.photosSubsections.map((section) => (
            <label key={section} className="modal-checkbox-row">
              <input type="checkbox" checked={!excludedSections.has(section)} onChange={() => onToggleSection(section)} />
              {section}
            </label>
          ))}
        </div>
        {generateError && <p className="folder-picker-error">{generateError}</p>}
        <div className="modal-actions" style={{ justifyContent: "space-between" }}>
          <div style={{ display: "flex", gap: 10 }}>
            <button className="secondary" onClick={onChangeFiles}>
              Change Files
            </button>
            <button className="secondary" onClick={onReview}>
              Review photos in detail first
            </button>
          </div>
          <button className="generate-report" disabled={generating} onClick={onGenerate}>
            {generating ? "Generating..." : "Generate Report"}
          </button>
        </div>
      </div>
    </div>
  );
}

/** Confirmation modal shown before generating -- some ATS customers only pay for part of the
 *  inspection, so unchecking a section here drops it (heading, photos, and its Observations
 *  table) from the generated report entirely rather than leaving it blank. Shares
 *  `excludedSections` with ReadyToGenerateScreen (see Home()) so scoping the report down there
 *  isn't lost by opening this modal instead. */
function GenerateSectionsModal({
  sections,
  excludedSections,
  onToggleSection,
  onCancel,
  onConfirm,
}: {
  sections: PhotosSection[];
  excludedSections: Set<PhotosSection>;
  onToggleSection: (section: PhotosSection) => void;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  return (
    <div className="modal-overlay">
      <div className="modal-card">
        <h2 className="modal-title">Which sections should this report include?</h2>
        <p className="section-reason">
          Unchecked sections are left out of the generated report entirely -- heading, photos, and Observations table
          -- instead of just being left blank. Useful when a customer only paid for part of the inspection.
        </p>
        <div className="modal-checklist">
          {sections.map((section) => (
            <label key={section} className="modal-checkbox-row">
              <input type="checkbox" checked={!excludedSections.has(section)} onChange={() => onToggleSection(section)} />
              {section}
            </label>
          ))}
        </div>
        <div className="modal-actions">
          <button className="secondary" onClick={onCancel}>
            Cancel
          </button>
          <button className="generate-report" onClick={onConfirm}>
            Generate Report
          </button>
        </div>
      </div>
    </div>
  );
}

interface TemplateOption {
  name: string;
  path: string;
}

/** Raw MDI exports and ATS's own templates both name themselves "{UnitType}_Full_..." /
 *  "{UnitType} Borescope Cover Page ...", so the raw file's own filename is enough to guess which
 *  unit type it's for -- e.g. "7EA_Full_022825-Batavia.docx" -> "7EA". Used to narrow the
 *  template list (see TemplateStep) instead of showing all ~35 templates for every unit/job type
 *  ATS does, most of which are irrelevant to whatever raw file was just picked. */
function deriveUnitTypeFromRawFilename(rawPath: string | null): string | null {
  if (!rawPath) return null;
  const base = rawPath.split(/[\\/]/).pop() ?? "";
  const match = /^([A-Za-z0-9.]+)_Full(?:_|\.|$)/i.exec(base);
  return match ? match[1] : null;
}

interface TemplatesResponse {
  templates: TemplateOption[];
  fullList: TemplateOption[];
  uploaded: TemplateOption[];
  unit: string | null;
  matchedCount: number;
  error?: string;
}

/** The template-picker step: a quick-select list of the blank templates found in ATS's shared
 *  template library (fetched from /api/templates, which already excludes finished sample
 *  reports), so picking the right file doesn't mean browsing folders and risking a finished
 *  report by mistake. Narrowed further to the raw file's own unit type and to templates this
 *  app's assembler actually recognizes (see /api/templates), with an escape hatch back to the
 *  full library since that's a heuristic, not a guarantee. Falls back to the plain folder browser
 *  for anything not in either list. */
function TemplateStep({ rawPath, onChosen, onBack }: { rawPath: string | null; onChosen: (path: string) => void; onBack: () => void }) {
  const unitType = useMemo(() => deriveUnitTypeFromRawFilename(rawPath), [rawPath]);
  const [result, setResult] = useState<TemplatesResponse | null>(null);
  const [templatesError, setTemplatesError] = useState<string | null>(null);
  const [showAll, setShowAll] = useState(false);
  const [browsing, setBrowsing] = useState(false);
  // Same as Home's rawBrowseDir -- remembers where "Browse for a different file..." was left,
  // so toggling back to the curated list and back to browsing again resumes there.
  const [browseDir, setBrowseDir] = useState<string | null>(null);
  const [uploading, setUploading] = useState(false);
  const uploadInputRef = useRef<HTMLInputElement>(null);

  function refresh() {
    const qs = unitType ? `?unit=${encodeURIComponent(unitType)}` : "";
    fetch(`/api/templates${qs}`)
      .then((res) => res.json())
      .then((json: TemplatesResponse) => {
        setResult(json);
        if (json.error) setTemplatesError(json.error);
      })
      .catch((err) => setTemplatesError(err instanceof Error ? err.message : "Failed to load templates"));
  }

  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setResult(null);
    setShowAll(false);
    refresh();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [unitType]);

  const templates = result ? (showAll || !unitType ? result.fullList : result.templates) : null;

  // Manually uploads a template that isn't in ATS's shared library (a brand-new one, or a
  // one-off a tech rep was handed directly) -- saved to disk so it's available again next time
  // (see /api/templates/upload), then used immediately for this report too.
  async function handleUpload(fileList: FileList | null) {
    const file = fileList?.[0];
    if (!file) return;
    setUploading(true);
    setTemplatesError(null);
    try {
      const formData = new FormData();
      formData.set("file", file);
      const res = await fetch("/api/templates/upload", { method: "POST", body: formData });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Upload failed");
      onChosen(json.path);
    } catch (err) {
      setTemplatesError(err instanceof Error ? err.message : "Upload failed");
      refresh();
    } finally {
      setUploading(false);
    }
  }

  if (browsing) {
    return (
      <FilePicker
        title="Select the ATS template"
        instructions="Pick the blank ATS Word template to assemble the photos into."
        extensions="docx"
        onChosen={onChosen}
        onBack={() => setBrowsing(false)}
        initialDir={browseDir}
        onDirChange={setBrowseDir}
      />
    );
  }

  return (
    <div className="folder-picker">
      <div className="folder-picker-card">
        <div className="app-logo">
          <h1>Select the ATS template</h1>
        </div>
        <p className="folder-picker-subtitle">Pick the blank ATS Word template to assemble the photos into.</p>
        <div className="toolbar">
          <button className="secondary" onClick={onBack}>
            Back
          </button>
        </div>
        {templatesError && <p className="folder-picker-error">{templatesError}</p>}
        {result && unitType && !showAll && (
          <p className="folder-picker-subtitle">
            {result.templates.length > 0
              ? `Showing "${unitType}" templates this app can auto-fill.`
              : result.matchedCount > 0
                ? `Found "${unitType}" template(s), but none are ones this app knows how to fill in yet.`
                : `No templates matching unit type "${unitType}" were found.`}{" "}
            <button className="secondary" onClick={() => setShowAll(true)} style={{ padding: "2px 10px", fontSize: 12 }}>
              Show all templates
            </button>
          </p>
        )}
        {result && unitType && showAll && (
          <p className="folder-picker-subtitle">
            Showing every template in the library.{" "}
            <button className="secondary" onClick={() => setShowAll(false)} style={{ padding: "2px 10px", fontSize: 12 }}>
              Show only &quot;{unitType}&quot; templates
            </button>
          </p>
        )}
        {templates === null ? (
          <div className="folder-list-loading">Loading...</div>
        ) : (
          <div className="folder-list folder-list-enter">
            {templates.length === 0 && <div className="folder-list-empty">No templates found.</div>}
            {templates.map((t) => (
              <button key={t.path} onClick={() => onChosen(t.path)}>
                📄 {t.name}
              </button>
            ))}
            {result && result.uploaded.length > 0 && (
              <>
                <div className="folder-list-section-label">Your uploaded templates</div>
                {result.uploaded.map((t) => (
                  <button key={t.path} onClick={() => onChosen(t.path)}>
                    📄 {t.name}
                  </button>
                ))}
              </>
            )}
            <button className="secondary" onClick={() => setBrowsing(true)}>
              Browse for a different file...
            </button>
            <input
              ref={uploadInputRef}
              type="file"
              accept=".docx"
              hidden
              onChange={(e) => {
                handleUpload(e.target.files);
                e.target.value = "";
              }}
            />
            <button className="secondary" disabled={uploading} onClick={() => uploadInputRef.current?.click()}>
              {uploading ? "Uploading..." : "Upload a new template..."}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

function FilePicker({
  title,
  instructions,
  extensions,
  onChosen,
  onBack,
  enableUseThisFolder,
  initialDir,
  onDirChange,
}: {
  title: string;
  instructions: string;
  extensions: string;
  onChosen: (path: string) => void;
  onBack?: () => void;
  /** Shows a "Use this folder" shortcut (mirroring Tech-Rep-Report-App's job-folder picker) that
   *  grabs the current folder's one matching file directly, for the common case of a folder that
   *  already contains exactly the file being looked for -- skipping the extra click on it. */
  enableUseThisFolder?: boolean;
  /** Where to start browsing -- lets a caller resume at wherever the user last was instead of
   *  always restarting at the home directory (see onDirChange). */
  initialDir?: string | null;
  /** Fired whenever the browsed directory changes, so a caller can remember it (in state that
   *  outlives this component -- e.g. across a "Back" that unmounts this picker) and hand it back
   *  as `initialDir` next time, instead of the picker always reopening at the home directory. */
  onDirChange?: (dir: string | null) => void;
}) {
  // Directory-visit history (like a browser's back/forward), separate from `data.parent`
  // ("Up one level", which walks toward the filesystem root, not visit order).
  const [history, setHistory] = useState<(string | null)[]>([initialDir ?? null]);
  const [historyIndex, setHistoryIndex] = useState(0);
  const dir = history[historyIndex];
  const [data, setData] = useState<BrowseResponse | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  function navigateTo(next: string | null) {
    const truncated = history.slice(0, historyIndex + 1);
    setHistory([...truncated, next]);
    setHistoryIndex(truncated.length);
  }

  useEffect(() => {
    onDirChange?.(dir);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dir]);

  useEffect(() => {
    let cancelled = false;
    // Resetting loading/error synchronously here (React's own documented pattern for
    // effect-based data fetching) so the picker shows a loading state immediately when `dir`
    // changes, rather than briefly showing the previous directory's stale listing.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setLoading(true);
    setError(null);
    const params = new URLSearchParams({ ext: extensions });
    if (dir) params.set("dir", dir);
    fetch(`/api/browse?${params.toString()}`)
      .then((res) => res.json())
      .then((json: BrowseResponse) => {
        if (cancelled) return;
        if (json.error) setError(json.error);
        else setData(json);
      })
      .catch((err) => {
        if (!cancelled) setError(err instanceof Error ? err.message : "Failed to browse");
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [dir, extensions]);

  const joinPath = (base: string, name: string) => (base.endsWith("\\") ? `${base}${name}` : `${base}\\${name}`);

  function useThisFolder() {
    if (!data?.dir) return;
    if (data.files.length === 1) {
      onChosen(joinPath(data.dir, data.files[0]));
    } else if (data.files.length === 0) {
      setError(`No .${extensions} file found in this folder.`);
    } else {
      setError(`This folder has more than one .${extensions} file -- pick one below.`);
    }
  }

  return (
    <div className="folder-picker">
      <div className="folder-picker-card">
        <div className="app-logo">
          <h1>{title}</h1>
        </div>
        <p className="folder-picker-subtitle">{instructions}</p>
        {data && !data.isDriveList && data.dir && <p className="job-path">{data.dir}</p>}
        <div className="toolbar">
          {onBack && (
            <button className="secondary" onClick={onBack}>
              Back
            </button>
          )}
          <button className="secondary" disabled={historyIndex === 0} onClick={() => setHistoryIndex((i) => i - 1)} title="Previous folder" aria-label="Previous folder">
            ◀
          </button>
          <button
            className="secondary"
            disabled={historyIndex === history.length - 1}
            onClick={() => setHistoryIndex((i) => i + 1)}
            title="Next folder"
            aria-label="Next folder"
          >
            ▶
          </button>
          <button className="secondary" onClick={() => navigateTo(DRIVES_ROOT)}>
            This PC
          </button>
          <button className="secondary" onClick={() => navigateTo(ONEDRIVE_ROOT)}>
            OneDrive
          </button>
          {data?.parent && (
            <button className="secondary" onClick={() => navigateTo(data.parent)}>
              Up one level
            </button>
          )}
          {enableUseThisFolder && data && !data.isDriveList && data.dir && (
            <button onClick={useThisFolder}>Use this folder</button>
          )}
        </div>
        {error && <p className="folder-picker-error">{error}</p>}
        {loading ? (
          <div className="folder-list-loading">Loading...</div>
        ) : (
          <div className="folder-list folder-list-enter">
            {data && data.folders.length === 0 && data.files.length === 0 && (
              <div className="folder-list-empty">Nothing here.</div>
            )}
            {data?.folders.map((name) => {
              const full = data.isDriveList ? name : joinPath(data.dir as string, name);
              return (
                <button key={full} onClick={() => navigateTo(full)}>
                  <FolderIcon /> {name}
                </button>
              );
            })}
            {data?.files.map((name) => {
              const full = joinPath(data.dir as string, name);
              return (
                <button key={full} onClick={() => onChosen(full)}>
                  📄 {name}
                </button>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}
