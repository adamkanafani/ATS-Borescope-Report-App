import { findHeadingParagraph } from "./xmlTextUtils";
import type { PhotosSection } from "./photosSectionMap";

/**
 * One Photos subsection this family's templates *can* have. Most are `required`: every
 * template in the family must have this exact Heading2, and detection fails if it's missing.
 * A few are genuinely optional -- some templates in the family have the heading, others don't
 * (e.g. only the plain 7FA template has its own Variable Inlet Guide Vanes bucket; 7HA and 9FA
 * don't). Optional entries are resolved per template instance (see resolvePhotosSubsections):
 * included only if their heading is actually found, and any raw unit classify() would have
 * routed to that section instead lands in `fallback` for a template that doesn't have it.
 * Getting this distinction right matters for more than cosmetics: assembleReport()'s splice
 * treats the gap between two subsections it does resolve as "belongs to the first one" -- an
 * unresolved subsection's own heading and content sit inside that gap and get overwritten right
 * along with it. A subsection some templates lack has to be modeled as optional, not just left
 * off the list, or generating from one of those templates silently destroys it.
 */
export interface PhotosSubsectionSpec {
  section: PhotosSection;
  optional?: { fallback: PhotosSection };
}

export interface TemplateFamily {
  id: string;
  label: string;
  /** Every Photos subsection this family's templates can have, in true document order --
   *  doubles as the literal Heading2 text to search for, since every PhotosSection name here is
   *  exactly the H2 text. See resolvePhotosSubsections() for turning this into the actual list
   *  for one specific template. */
  photosSubsectionSpecs: PhotosSubsectionSpec[];
  /** Photos subsection -> the Observations H2 heading text that covers it, for sections that
   *  have one. Omitted entries (e.g. Variable Inlet Guide Vanes, which reports through the
   *  Compressor Section table's own "VIGV" row instead of a heading of its own) have no
   *  Observations-side heading to find or remove. */
  observationsHeadingBySection: Partial<Record<PhotosSection, string>>;
  /** Maps a raw unit's `component` name to one of the sections named in `photosSubsectionSpecs`,
   *  or null to leave it for manual review rather than guess. Can return an optional section
   *  even for a template instance that doesn't have it -- resolveClassify() handles the
   *  fallback, callers should use that instead of calling this directly. */
  classify: (component: string) => PhotosSection | null;
}

interface Rule {
  section: PhotosSection;
  keywords: string[];
}

function makeClassifier(rules: Rule[]): (component: string) => PhotosSection | null {
  return (component: string) => {
    for (const rule of rules) {
      if (rule.keywords.some((keyword) => component.includes(keyword))) return rule.section;
    }
    return null;
  };
}

function requiredSections(family: TemplateFamily): PhotosSection[] {
  return family.photosSubsectionSpecs.filter((spec) => !spec.optional).map((spec) => spec.section);
}

/**
 * Ordered keyword rules for the standard 5-subsection family (7EA, ABB11N, Frame 5, GE10 --
 * confirmed structurally identical by comparing their Photos/Observations heading lists), built
 * from the 47 distinct `component` values actually observed across the 178 real photo units in
 * reference/raw-mdi-samples/7EA_Full_022825-Batavia.docx (see reference/phase1-source-analysis.md).
 * Order matters: Combustion's "Fuel Nozzle" must be checked before Turbine's bare "Nozzle" --
 * "Primary Fuel Nozzle" would otherwise false-match Turbine. Inlet's keywords come from the
 * template's own instructional example photos (Bellmouth/Floor) rather than this sample, since
 * this particular inspection had no inlet photos -- kept anyway since other inspections will.
 * "Venturi" and "Number Two Bearing" were added after a second raw sample turned them up
 * unassigned. "Outer/Inner Aft Mounting Bracket", "Outer Mount Bracket", "Impingement Sleeve",
 * "Steam Distributor", and "Flame Scanner" were added after cross-checking a real 7EA MDI
 * project's own component tree (Downloads/7EA_Full_110525.zip) turned them up unassigned too --
 * all combustion-casing hardware/instrumentation, folded into the one Combustion bucket here
 * since this family doesn't split Cold/Hot Side (see BEARING_RULES, which does and keeps the
 * mounting-hardware side of that split separately).
 */
const STANDARD_RULES: Rule[] = [
  {
    section: "Combustion Section",
    keywords: [
      "Combustion",
      "Crossfire",
      "Igniter",
      "Liner",
      "Fuel Nozzle",
      "TP ",
      "Transition",
      "Venturi",
      "Number Two Bearing",
      "Number 2 Bearing",
      "Outer Aft Mounting Bracket",
      "Outer Mount Bracket",
      "Inner Aft Mounting Bracket",
      "Impingement Sleeve",
      "Steam Distributor",
      "Flame Scanner",
    ],
  },
  { section: "Turbine Section", keywords: ["Nozzle", "Bucket", "Shroud"] },
  { section: "Compressor Section", keywords: ["Rotor Blade", "Stator Vane", "Guide Vane", "Compressor Case", "EGV", "Dovetail"] },
  { section: "Exhaust Section", keywords: ["Exhaust", "Flex Seal", "Strut", "Barrel", "Diffuser"] },
  { section: "Inlet Section", keywords: ["Bellmouth", "Inlet", "Floor"] },
];

export const FRAME_STANDARD_FAMILY: TemplateFamily = {
  id: "frame-standard",
  label: "Standard borescope (Inlet / Compressor / Combustion / Turbine / Exhaust)",
  photosSubsectionSpecs: [
    { section: "Inlet Section" },
    { section: "Compressor Section" },
    { section: "Combustion Section" },
    { section: "Turbine Section" },
    { section: "Exhaust Section" },
    // Only the "... with Generator" variant templates have these two -- nothing in `classify`
    // currently routes a raw unit to either one (no real generator MDI data verified yet), so
    // the fallback is effectively dead code today; they exist so a template that DOES have
    // these headings doesn't have them destroyed by assembleReport's splice (see
    // PhotosSubsectionSpec's own doc comment) and so they show up, always-empty, in the sidebar.
    { section: "Generator Turbine End", optional: { fallback: "Exhaust Section" } },
    { section: "Generator Exciter End", optional: { fallback: "Exhaust Section" } },
  ],
  observationsHeadingBySection: {
    "Compressor Section": "Compressor Section",
    "Combustion Section": "Combustion Section Cold Side",
    "Turbine Section": "Turbine Section",
    "Exhaust Section": "Exhaust Section",
  },
  classify: makeClassifier(STANDARD_RULES),
};

/**
 * The wider "full-size frame" family (7FA/7HA/9FA/7FA Dot 05 -- confirmed structurally identical
 * on their REQUIRED subsections by comparing heading lists) adds a Number One Bearing Area photo
 * group, and splits Combustion's photos into Cold Side/Hot Side instead of one bucket. Only the
 * plain 7FA template (and its "with Generator" variant) additionally has its own Variable Inlet
 * Guide Vanes bucket -- 7HA, 9FA, and 7FA Dot 05 do not, so it's modeled as optional rather than
 * required (see PhotosSubsectionSpec). 9FA also has a one-off "R-17 Rotor Blade Dovetail Slot
 * Trailing Edge" bucket no other template in the family has, also optional.
 *
 * The keywords (and the Hot/Cold split specifically) are built from a real 7FA MDI project's own
 * component tree (Downloads/MDI.zip, "7FA_Full_110525.mdz" -> Inspection/inspection.xml), not
 * guessed:
 *   - The bearing area's real component names are "Number 1/2 Bearing Right/Left Side Lift Oil
 *     Line" -- numeral "1"/"2", not spelled-out "One"/"Two" like 7EA's raw files use. Number 2
 *     Bearing still folds into Combustion (it physically sits at the combustor-case boundary,
 *     same reasoning as 7EA's own "Number Two Bearing" rule), only Number 1 Bearing gets its
 *     own Photos bucket.
 *   - "Variable Inlet Guide Vanes" is the only component with "Guide Vane" in its name that
 *     should get its own bucket -- "Exit Guide Vane" also contains that substring but belongs
 *     with Compressor, so the VIGV rule matches the full phrase, not the generic substring
 *     (unlike the Standard family, which buckets both together and can stay generic). For a
 *     template that doesn't have its own VIGV heading, resolveClassify() falls a VIGV match back
 *     to Compressor Section -- the same bucket the Standard family already uses for it.
 *   - Cold Side components are the combustion casing's exterior/mounting hardware (Combustion
 *     Casing Manifolds, CDC Inner Barrel, Quaternary Annulus, outer mounting brackets/zipper
 *     welds/spool bolts, impingement sleeve) -- everything an inspector reaches from outside the
 *     case. Hot Side is the actual flame-path hardware (fuel nozzles, igniters, crossfire tubes,
 *     liner/transition, transition seals) reached from inside the liner. A few older exports
 *     apparently named cold-side parts literally "Cold Side ___" (e.g. "Cold Side Outer TP"),
 *     kept as a first check in case an older raw file still uses it.
 *   - Nothing currently routes to "R-17 Rotor Blade Dovetail Slot Trailing Edge" specifically --
 *     dovetail-related components already classify to plain Compressor Section (its fallback),
 *     and no real 9FA raw sample has been checked yet to know whether they should be split out.
 */
const BEARING_RULES: Rule[] = [
  { section: "Variable Inlet Guide Vanes", keywords: ["Variable Inlet Guide Vane"] },
  { section: "Number One Bearing Area", keywords: ["Number 1 Bearing"] },
  { section: "Combustion Section Cold Side", keywords: ["Cold Side"] },
  {
    section: "Combustion Section Cold Side",
    keywords: [
      "Combustion Casing",
      "CDC",
      "Quaternary",
      "Outer Aft Mounting Bracket",
      "Outer Mount Bracket",
      "Outer Zipper Weld",
      "Bull Horn Bracket",
      "Spool Bolt",
      "Impingement Sleeve",
    ],
  },
  {
    section: "Combustion Section Hot Side",
    keywords: [
      "Combustion",
      "Crossfire",
      "Igniter",
      "Liner",
      "Fuel Nozzle",
      "TP ",
      "Transition",
      "Venturi",
      "Number Two Bearing",
      "Number 2 Bearing",
      "DLN",
      "Picture Frame",
      "Floating Seal",
      "Side Seal",
      "Inner Aft Mounting Bracket",
      "Inner Zipper Weld",
      "Steam Distributor",
      "Flame Scanner",
    ],
  },
  { section: "Turbine Section", keywords: ["Nozzle", "Bucket", "Shroud"] },
  { section: "Compressor Section", keywords: ["Rotor Blade", "Stator Vane", "Compressor Case", "EGV", "Exit Guide Vane", "Dovetail"] },
  { section: "Exhaust Section", keywords: ["Exhaust", "Flex Seal", "Strut", "Diffuser", "Barrel"] },
  { section: "Inlet Section", keywords: ["Bellmouth", "Inlet", "Floor"] },
];

export const FRAME_BEARING_FAMILY: TemplateFamily = {
  id: "frame-bearing",
  label: "Full borescope with bearing area (7FA / 7HA / 9FA)",
  photosSubsectionSpecs: [
    { section: "Number One Bearing Area" },
    { section: "Inlet Section" },
    { section: "Variable Inlet Guide Vanes", optional: { fallback: "Compressor Section" } },
    { section: "Compressor Section" },
    { section: "R-17 Rotor Blade Dovetail Slot Trailing Edge", optional: { fallback: "Compressor Section" } },
    { section: "Combustion Section Cold Side" },
    { section: "Combustion Section Hot Side" },
    { section: "Turbine Section" },
    { section: "Exhaust Section" },
    { section: "Generator Turbine End", optional: { fallback: "Exhaust Section" } },
    { section: "Generator Exciter End", optional: { fallback: "Exhaust Section" } },
  ],
  observationsHeadingBySection: {
    "Number One Bearing Area": "Bearing Areas",
    "Compressor Section": "Compressor Section",
    "Combustion Section Hot Side": "Combustion Section Hot Side",
    "Turbine Section": "Turbine Section",
    "Exhaust Section": "Exhaust Section",
  },
  classify: makeClassifier(BEARING_RULES),
};

export const TEMPLATE_FAMILIES: TemplateFamily[] = [FRAME_STANDARD_FAMILY, FRAME_BEARING_FAMILY];

/**
 * Auto-detects which family a loaded template matches, by checking whether every one of a
 * family's REQUIRED Photos subsections (not the optional ones -- see PhotosSubsectionSpec) can
 * be found as a Heading2 somewhere after the template's "Photos" Heading1. Tried largest
 * required-set first, so a template satisfying a smaller family's subset doesn't get
 * misdetected as the smaller/less specific family.
 */
export function detectTemplateFamily(documentXml: string): TemplateFamily | null {
  const photosHeading = findHeadingParagraph(documentXml, "Heading1", "Photos", 0);
  if (!photosHeading) return null;

  const candidates = [...TEMPLATE_FAMILIES].sort((a, b) => requiredSections(b).length - requiredSections(a).length);
  for (const family of candidates) {
    const allFound = requiredSections(family).every(
      (name) => !!findHeadingParagraph(documentXml, "Heading2", name, photosHeading.paragraphStart),
    );
    if (allFound) return family;
  }
  return null;
}

/**
 * The actual, ordered list of Photos subsections for ONE template instance: every required
 * subsection, plus each optional one whose own heading this specific template actually has.
 * This -- not `family.photosSubsectionSpecs` directly -- is what assembleReport() and
 * scan/route.ts must splice/report against, or an optional subsection this template doesn't
 * have gets destroyed rather than just skipped (see PhotosSubsectionSpec).
 */
export function resolvePhotosSubsections(family: TemplateFamily, documentXml: string, fromIndex: number): PhotosSection[] {
  return family.photosSubsectionSpecs
    .filter((spec) => !spec.optional || !!findHeadingParagraph(documentXml, "Heading2", spec.section, fromIndex))
    .map((spec) => spec.section);
}

/**
 * Wraps `family.classify` so it never returns a section this specific template doesn't actually
 * have a heading for -- redirects to that section's declared fallback instead. Always use this
 * (never `family.classify` directly) once a template's real subsections are known.
 */
export function resolveClassify(family: TemplateFamily, resolvedSections: Set<PhotosSection>): (component: string) => PhotosSection | null {
  return (component: string) => {
    const section = family.classify(component);
    if (!section || resolvedSections.has(section)) return section;
    const spec = family.photosSubsectionSpecs.find((s) => s.section === section);
    return spec?.optional?.fallback ?? section;
  };
}
