import type { RawPhotoUnit } from "./rawMdiParser";
import type { PhotosSection } from "./photosSectionMap";

export type TurbineGroup = "Nozzles" | "Buckets" | "Shroud Blocks";

export interface TurbineStageEntry {
  stage: number;
  group: TurbineGroup;
  text: string;
}

export interface ObservationsAggregation {
  /** Row key (e.g. "R1", "S1", "VIGV", "EGV1") -> Condition text. */
  compressor: Map<string, string>;
  turbine: TurbineStageEntry[];
  /** Units classified as Compressor/Turbine (by the same classifier the Photos step uses)
   *  whose component still didn't match a known Observations row -- a real gap worth
   *  surfacing, unlike Combustion/Exhaust/Inlet units, which are silently out of scope for
   *  this pass and never appear here. */
  unmatched: RawPhotoUnit[];
}

const TYPICAL = "Typical";

/**
 * Maps a raw unit's `component` to a Compressor Observations table row key, or null if it
 * doesn't belong there (Combustion/Exhaust/Turbine components, or anything unrecognized).
 * See reference/phase1-source-analysis.md and the plan file for how these were derived from
 * the 47 real component values in the Batavia sample.
 */
export function compressorRowKey(component: string): string | null {
  let m = /^Stage (\d+) Rotor Blade$/.exec(component);
  if (m) return `R${m[1]}`;
  m = /^Stage (\d+) Stator Vane$/.exec(component);
  if (m) return `S${m[1]}`;
  if (component === "Variable Inlet Guide Vanes" || component === "Variable Inlet Guide Vane") return "VIGV";
  m = /^Exit Guide Vane (\d+)$/.exec(component);
  if (m) return `EGV${m[1]}`;
  return null;
}

/** Maps a raw unit's `component` to a Turbine Observations stage/group, or null. */
export function turbineRowKey(component: string): { stage: number; group: TurbineGroup } | null {
  let m = /^Stage (\d+) Nozzle$/.exec(component);
  if (m) return { stage: Number(m[1]), group: "Nozzles" };
  m = /^Stage (\d+) Bucket$/.exec(component);
  if (m) return { stage: Number(m[1]), group: "Buckets" };
  m = /^Stage (\d+) Shroud$/.exec(component);
  if (m) return { stage: Number(m[1]), group: "Shroud Blocks" };
  return null;
}

/**
 * Deliberately mechanical, not a rewrite of Brett's stylistic rulebook: drops "Typical" units,
 * writes "No defects identified" if nothing's left, counts repeated distinct observations
 * (preserving the raw file's own front-to-back order), and joins with commas. A human is
 * expected to review/polish this -- see the red-text convention in observationsFiller.ts.
 */
function aggregateConditionText(units: RawPhotoUnit[]): string {
  const nonTypical = units.filter((u) => u.observation && u.observation !== TYPICAL);
  if (nonTypical.length === 0) return "No defects identified";

  const counts = new Map<string, number>();
  const order: string[] = [];
  for (const unit of nonTypical) {
    const text = unit.observation as string;
    if (!counts.has(text)) {
      counts.set(text, 0);
      order.push(text);
    }
    counts.set(text, counts.get(text)! + 1);
  }

  return order.map((text) => (counts.get(text)! > 1 ? `${text} (${counts.get(text)})` : text)).join(", ");
}

/**
 * `classify` decides which raw units are in scope for "unmatched" reporting -- a component the
 * family classifier puts in Compressor/Turbine Section but that neither row-key regex below
 * recognizes is a real gap worth surfacing. Row-key matching itself is tried on every unit
 * regardless of which Photos subsection it was classified into: a family that gives Variable
 * Inlet Guide Vanes its own Photos bucket (see templateFamilies.ts) still reports that unit's
 * condition through the Compressor Section Observations table's own "VIGV" row, same as a
 * family with no separate VIGV bucket at all.
 */
export function aggregateObservations(units: RawPhotoUnit[], classify: (component: string) => PhotosSection | null): ObservationsAggregation {
  const compressorGroups = new Map<string, RawPhotoUnit[]>();
  const turbineGroups = new Map<string, RawPhotoUnit[]>(); // key: `${stage}:${group}`
  const turbineMeta = new Map<string, { stage: number; group: TurbineGroup }>();
  const unmatched: RawPhotoUnit[] = [];

  for (const unit of units) {
    if (!unit.component) continue;

    const compressorKey = compressorRowKey(unit.component);
    if (compressorKey) {
      if (!compressorGroups.has(compressorKey)) compressorGroups.set(compressorKey, []);
      compressorGroups.get(compressorKey)!.push(unit);
      continue;
    }
    const turbine = turbineRowKey(unit.component);
    if (turbine) {
      const key = `${turbine.stage}:${turbine.group}`;
      if (!turbineGroups.has(key)) {
        turbineGroups.set(key, []);
        turbineMeta.set(key, turbine);
      }
      turbineGroups.get(key)!.push(unit);
      continue;
    }

    const section = classify(unit.component);
    if (section === "Compressor Section" || section === "Turbine Section") unmatched.push(unit); // a real gap
    // else: out of scope (Combustion/Exhaust/Inlet/Bearing Area/etc.), not a gap.
  }

  const compressor = new Map<string, string>();
  for (const [key, groupUnits] of compressorGroups) {
    compressor.set(key, aggregateConditionText(groupUnits));
  }

  const turbine: TurbineStageEntry[] = [];
  for (const [key, groupUnits] of turbineGroups) {
    const meta = turbineMeta.get(key)!;
    turbine.push({ stage: meta.stage, group: meta.group, text: aggregateConditionText(groupUnits) });
  }

  return { compressor, turbine, unmatched };
}
