import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { loadDocx, saveDocx } from "../src/lib/docx/zip";
import { parseRawMdiPhotoUnits } from "../src/lib/docx/rawMdiParser";
import { assembleReport } from "../src/lib/docx/templateAssembler";

const PROJECT_ROOT = path.resolve(__dirname, "..");
const RAW_SAMPLE = process.env.ATS_TEST_RAW_SAMPLE || path.join(PROJECT_ROOT, "reference", "raw-mdi-samples", "7EA_Full_022825-Batavia.docx");
const TEMPLATE_SAMPLE = path.join(PROJECT_ROOT, "reference", "7EA Borescope Cover Page 111825.docx");
const OUT_DIR = process.env.ATS_TEST_OUT_DIR || path.join(os.tmpdir(), "ats-borescope-test-output");

async function main() {
  console.log(`Reading raw sample: ${RAW_SAMPLE}`);
  const rawLoaded = await loadDocx(fs.readFileSync(RAW_SAMPLE));
  const units = parseRawMdiPhotoUnits(rawLoaded);
  console.log(`Parsed ${units.length} raw photo units.`);

  console.log(`\nReading template: ${TEMPLATE_SAMPLE}`);
  const templateLoaded = await loadDocx(fs.readFileSync(TEMPLATE_SAMPLE));

  const result = assembleReport(templateLoaded, units, {});
  console.log(`\nInserted ${result.inserted} of ${units.length} units.`);
  console.log("By section:", result.bySection);
  console.log(`${result.unassigned.length} units unassigned:`);
  const unassignedComponents = new Set(result.unassigned.map((u) => u.component));
  console.log([...unassignedComponents].join(", ") || "(none)");

  console.log(`\nObservations: filled ${result.observationsFilled.length} rows/groups:`);
  console.log(result.observationsFilled.join(", ") || "(none)");
  console.log(`Observations: ${result.observationsSkipped.length} skipped (had data but no matching template row):`);
  console.log(result.observationsSkipped.join(", ") || "(none)");

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const outPath = path.join(OUT_DIR, "assembled-full.docx");
  const buffer = await saveDocx(templateLoaded);
  fs.writeFileSync(outPath, buffer);
  console.log(`\nSaved assembled report to ${outPath} (${buffer.length} bytes)`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
