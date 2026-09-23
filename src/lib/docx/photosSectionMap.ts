/**
 * Every Photos/Observations subsection name used across all supported template families (see
 * templateFamilies.ts). A given family only ever uses a subset of these -- the standard 5 for
 * 7EA/ABB11N-shaped templates, plus Bearing Area/VIGV/Cold-vs-Hot-split for the wider 7FA/7HA/9FA
 * family. Kept as one union (rather than a type per family) since the client and the API routes
 * pass sections around generically without caring which family they came from.
 */
export type PhotosSection =
  | "Inlet Section"
  | "Compressor Section"
  | "Combustion Section"
  | "Combustion Section Cold Side"
  | "Combustion Section Hot Side"
  | "Turbine Section"
  | "Exhaust Section"
  | "Number One Bearing Area"
  | "Variable Inlet Guide Vanes"
  | "R-17 Rotor Blade Dovetail Slot Trailing Edge"
  | "Generator Turbine End"
  | "Generator Exciter End";
