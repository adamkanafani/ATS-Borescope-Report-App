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
  const [selectedSection, setSelectedSection] = useState<PhotosSection | "Unassigned" | "Manual Photos" | "Trash">("Compressor Section");
  const [overrides, setOverrides] = useState<Record<number, PhotosSection | null>>({});
  // Trashing a unit never touches `overrides` -- it's tracked separately so recovering it just
  // means removing it from this set, and it goes right back to whatever section (classifier
  // default or a prior manual reassignment) it already had.
  const [deletedIndices, setDeletedIndices] = useState<Set<number>>(new Set());
  const [generating, setGenerating] = useState(false);
  const [generateError, setGenerateError] = useState<string | null>(null);
  const [operationalDataFiles, setOperationalDataFiles] = useState<File[]>([]);
  const [dataPlateFiles, setDataPlateFiles] = useState<File[]>([]);

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
    setOperationalDataFiles([]);
    setDataPlateFiles([]);
    setSelectedSection("Compressor Section");
  }

  async function generateReport(excludedSections: Set<PhotosSection>) {
    if (!scan) return;
    setGenerating(true);
    setGenerateError(null);
    try {
      // Trashed units are excluded the same way "Leave excluded" is (an explicit null override)
      // -- computed here rather than stored in `overrides` itself, so recovering a trashed unit
      // before generating restores whatever section it actually had.
      const overridesWithTrash: Record<number, PhotosSection | null> = { ...overrides };
      for (const index of deletedIndices) overridesWithTrash[index] = null;

      const formData = new FormData();
      formData.set("sessionId", scan.sessionId);
      formData.set("overrides", JSON.stringify(overridesWithTrash));
      formData.set("excludedSections", JSON.stringify([...excludedSections]));
      if (operationalDataFiles[0]) formData.set("operationalDataPhoto", operationalDataFiles[0]);
      for (const file of dataPlateFiles) formData.append("dataPlatePhoto", file);

      const res = await fetch("/api/generate", { method: "POST", body: formData });
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
      generating={generating}
      generateError={generateError}
      onGenerate={generateReport}
      operationalDataFiles={operationalDataFiles}
      dataPlateFiles={dataPlateFiles}
      onOperationalDataFiles={setOperationalDataFiles}
      onDataPlateFiles={setDataPlateFiles}
      onChangeFiles={resetFiles}
    />
  );
}

function effectiveSection(unit: UnitSummary, overrides: Record<number, PhotosSection | null>): PhotosSection | null {
  return unit.index in overrides ? overrides[unit.index] : unit.section;
}

/** A drag-and-drop (or click-to-browse) slot for up to `maxFiles` front-matter photos, with a
 *  live preview per file. Reads files straight from the browser's native picker/drop event --
 *  no server-side path resolution needed, since the actual bytes go up with the generate
 *  request. Some inspectors take more than one shot of the same thing (e.g. two data plate
 *  photos), and the template has a real placeholder slot for each one -- maxFiles matches
 *  however many slots actually exist. */
function PhotoDropZone({
  label,
  files,
  onFilesChange,
  maxFiles,
}: {
  label: string;
  files: File[];
  onFilesChange: (files: File[]) => void;
  maxFiles: number;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState(false);
  const canAddMore = files.length < maxFiles;

  function addFiles(incoming: FileList | null) {
    if (!incoming) return;
    const images = Array.from(incoming).filter((f) => f.type.startsWith("image/"));
    if (images.length === 0) return;
    onFilesChange([...files, ...images].slice(0, maxFiles));
  }

  return (
    <div className="photo-dropzone-group">
      <div className="photo-dropzone-label">
        {label} ({files.length}/{maxFiles})
      </div>
      {files.map((file, i) => (
        <PhotoPreviewCard key={`${file.name}-${i}`} file={file} onRemove={() => onFilesChange(files.filter((_, j) => j !== i))} />
      ))}
      {canAddMore && (
        <div
          className={`photo-dropzone ${isDragging ? "dragging" : ""}`}
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={(e) => {
            e.preventDefault();
            setIsDragging(false);
            addFiles(e.dataTransfer.files);
          }}
        >
          <input ref={inputRef} type="file" accept="image/*" multiple hidden onChange={(e) => addFiles(e.target.files)} />
          <div className="photo-dropzone-empty">
            Drag &amp; drop {files.length > 0 ? "another" : "an"} image here, or click to browse
          </div>
        </div>
      )}
    </div>
  );
}

/** One picked file's preview thumbnail + filename + remove button. */
function PhotoPreviewCard({ file, onRemove }: { file: File; onRemove: () => void }) {
  const previewUrl = useMemo(() => URL.createObjectURL(file), [file]);

  useEffect(() => {
    return () => URL.revokeObjectURL(previewUrl);
  }, [previewUrl]);

  return (
    <div className="photo-preview-card">
      <img src={previewUrl} alt={file.name} className="photo-dropzone-preview" />
      <div className="photo-dropzone-filename">{file.name}</div>
      <button type="button" className="secondary photo-dropzone-remove" onClick={onRemove}>
        Remove
      </button>
    </div>
  );
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
  generating,
  generateError,
  onGenerate,
  operationalDataFiles,
  dataPlateFiles,
  onOperationalDataFiles,
  onDataPlateFiles,
  onChangeFiles,
}: {
  rawPath: string;
  scan: ScanResponse;
  selectedSection: PhotosSection | "Unassigned" | "Manual Photos" | "Trash";
  onSelectSection: (s: PhotosSection | "Unassigned" | "Manual Photos" | "Trash") => void;
  overrides: Record<number, PhotosSection | null>;
  onOverride: (index: number, section: PhotosSection | null) => void;
  deletedIndices: Set<number>;
  onDelete: (index: number) => void;
  onRecover: (index: number) => void;
  generating: boolean;
  generateError: string | null;
  onGenerate: (excludedSections: Set<PhotosSection>) => void;
  operationalDataFiles: File[];
  dataPlateFiles: File[];
  onOperationalDataFiles: (files: File[]) => void;
  onDataPlateFiles: (files: File[]) => void;
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
            <button
              className={`section-item ${selectedSection === "Manual Photos" ? "active" : ""}`}
              onClick={() => onSelectSection("Manual Photos")}
            >
              <div className="title-row">
                <span>Manual Insert</span>
              </div>
              <div className="confidence-tag">{operationalDataFiles.length + dataPlateFiles.length}/3 photos set</div>
            </button>
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
          {selectedSection === "Manual Photos" ? (
            <>
              <h2 style={{ marginBottom: 8 }}>Manual Insert</h2>
              <p className="section-reason">
                The borescope never takes these shots itself -- drag them in (or click to browse) if you have them.
                Data Plate has two slots, since inspectors sometimes take a second one. Any slot left blank is
                removed from the generated report instead of showing &quot;Insert Photo Here&quot;.
              </p>
              <div className="front-matter-dropzones">
                <PhotoDropZone label="Operational Data Photo" files={operationalDataFiles} onFilesChange={onOperationalDataFiles} maxFiles={1} />
                <PhotoDropZone label="Data Plate Photo" files={dataPlateFiles} onFilesChange={onDataPlateFiles} maxFiles={2} />
              </div>
            </>
          ) : (
            <>
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
            </>
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
          onCancel={() => setShowGenerateModal(false)}
          onConfirm={(excludedSections) => {
            setShowGenerateModal(false);
            onGenerate(excludedSections);
          }}
        />
      )}
    </div>
  );
}

/** Confirmation modal shown before generating -- some ATS customers only pay for part of the
 *  inspection, so unchecking a section here drops it (heading, photos, and its Observations
 *  table) from the generated report entirely rather than leaving it blank. */
function GenerateSectionsModal({
  sections,
  onCancel,
  onConfirm,
}: {
  sections: PhotosSection[];
  onCancel: () => void;
  onConfirm: (excludedSections: Set<PhotosSection>) => void;
}) {
  const [checked, setChecked] = useState<Record<string, boolean>>(() => Object.fromEntries(sections.map((s) => [s, true])));

  function toggle(section: PhotosSection) {
    setChecked((prev) => ({ ...prev, [section]: !prev[section] }));
  }

  function handleConfirm() {
    onConfirm(new Set(sections.filter((s) => !checked[s])));
  }

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
              <input type="checkbox" checked={checked[section] ?? true} onChange={() => toggle(section)} />
              {section}
            </label>
          ))}
        </div>
        <div className="modal-actions">
          <button className="secondary" onClick={onCancel}>
            Cancel
          </button>
          <button className="generate-report" onClick={handleConfirm}>
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

/** The template-picker step: a quick-select list of the blank templates found in the project's
 *  reference/ folder (fetched from /api/templates, which already excludes finished sample
 *  reports), so picking the right file doesn't mean browsing folders and risking a finished
 *  report by mistake. Falls back to the plain folder browser for anything not in that list. */
function TemplateStep({ onChosen, onBack }: { onChosen: (path: string) => void; onBack: () => void }) {
  const [templates, setTemplates] = useState<TemplateOption[] | null>(null);
  const [templatesError, setTemplatesError] = useState<string | null>(null);
  const [browsing, setBrowsing] = useState(false);
  // Same as Home's rawBrowseDir -- remembers where "Browse for a different file..." was left,
  // so toggling back to the curated list and back to browsing again resumes there.
  const [browseDir, setBrowseDir] = useState<string | null>(null);

  useEffect(() => {
    fetch("/api/templates")
      .then((res) => res.json())
      .then((json: { templates: TemplateOption[]; error?: string }) => {
        setTemplates(json.templates);
        if (json.error) setTemplatesError(json.error);
      })
      .catch((err) => setTemplatesError(err instanceof Error ? err.message : "Failed to load templates"));
  }, []);

  // Same "grab the one obvious file" shortcut as the raw MDI step's FilePicker (see
  // enableUseThisFolder there) -- reuses the curated list already fetched above instead of a
  // second /api/browse call, since it's the same folder (TEMPLATES_DIR).
  function useThisFolder() {
    if (!templates) return;
    if (templates.length === 1) {
      onChosen(templates[0].path);
    } else if (templates.length === 0) {
      setTemplatesError("No templates found in this folder.");
    } else {
      setTemplatesError("This folder has more than one template -- pick one below.");
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
        enableUseThisFolder
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
          {templates && <button onClick={useThisFolder}>Use this folder</button>}
        </div>
        {templatesError && <p className="folder-picker-error">{templatesError}</p>}
        {templates === null ? (
          <div className="folder-list-loading">Loading...</div>
        ) : (
          <div className="folder-list folder-list-enter">
            {templates.length === 0 && <div className="folder-list-empty">No templates found in reference/.</div>}
            {templates.map((t) => (
              <button key={t.path} onClick={() => onChosen(t.path)}>
                📄 {t.name}
              </button>
            ))}
            <button className="secondary" onClick={() => setBrowsing(true)}>
              Browse for a different file...
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
