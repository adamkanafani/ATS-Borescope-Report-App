import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { loadDocx } from "../src/lib/docx/zip";
import { parseRawMdiPhotoUnits } from "../src/lib/docx/rawMdiParser";
import { listContentControls } from "../src/lib/docx/templateInspector";

const PROJECT_ROOT = path.resolve(__dirname, "..");
const RAW_SAMPLE = process.env.ATS_TEST_RAW_SAMPLE || path.join(PROJECT_ROOT, "reference", "raw-mdi-samples", "7EA_Full_022825-Batavia.docx");
const TEMPLATE_SAMPLE = path.join(PROJECT_ROOT, "reference", "7EA Borescope Cover Page 111825.docx");
const OUT_DIR = process.env.ATS_TEST_OUT_DIR || path.join(os.tmpdir(), "ats-borescope-test-output");

async function main() {
  console.log("=== Raw MDI photo-unit parser ===");
  console.log(`Reading ${RAW_SAMPLE}`);
  const rawLoaded = await loadDocx(fs.readFileSync(RAW_SAMPLE));
  const units = parseRawMdiPhotoUnits(rawLoaded);
  console.log(`Parsed ${units.length} photo units.`);

  const withWarnings = units.filter((u) => u.warnings.length > 0);
  console.log(`${withWarnings.length} units have warnings:`);
  for (const u of withWarnings.slice(0, 15)) {
    console.log(`  #${u.index}: ${u.warnings.join(" | ")}`);
  }

  const missingImage = units.filter((u) => !u.imageBytes);
  console.log(`${missingImage.length} units are missing resolved image bytes.`);
  const missingFields = units.filter((u) => !u.component || !u.location || !u.observation);
  console.log(`${missingFields.length} units are missing component/location/observation.`);

  console.log("\nFirst 5 units:");
  for (const u of units.slice(0, 5)) {
    console.log(
      `  #${u.index} component=${JSON.stringify(u.component)} location=${JSON.stringify(u.location)} ` +
        `observation=${JSON.stringify(u.observation)} comments=${JSON.stringify(u.comments)} ` +
        `filenameCaption=${JSON.stringify(u.filenameCaption)} imagePath=${u.imagePath} ` +
        `imageBytes=${u.imageBytes ? u.imageBytes.length + " bytes" : "MISSING"}`,
    );
  }

  const withMeasurements = units.filter((u) => u.measurements.length > 0);
  console.log(`\n${withMeasurements.length} units have measurement rows. First 5:`);
  for (const u of withMeasurements.slice(0, 5)) {
    console.log(`  #${u.index} ${u.component} / ${u.location}: ${JSON.stringify(u.measurements)}`);
  }

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const sampleIndexes = [0, Math.floor(units.length / 2), units.length - 1].filter(
    (i, pos, arr) => i >= 0 && i < units.length && arr.indexOf(i) === pos,
  );
  for (const i of sampleIndexes) {
    const bytes = units[i].imageBytes;
    if (!bytes) continue;
    const outPath = path.join(OUT_DIR, `sample-photo-${i}.jpg`);
    fs.writeFileSync(outPath, bytes);
    console.log(`Saved ${outPath} (${bytes.length} bytes)`);
  }

  console.log("\n=== Template content controls ===");
  console.log(`Reading ${TEMPLATE_SAMPLE}`);
  const templateLoaded = await loadDocx(fs.readFileSync(TEMPLATE_SAMPLE));
  const controls = listContentControls(templateLoaded);
  console.log(`Found ${controls.length} content controls.`);
  const dropdowns = controls.filter((c) => c.kind === "dropDownList");
  console.log(`${dropdowns.length} are dropdown lists.`);
  for (const c of dropdowns) {
    console.log(
      `  #${c.index} alias=${JSON.stringify(c.alias)} tag=${JSON.stringify(c.tag)} ` +
        `current=${JSON.stringify(c.currentText)} placeholder=${c.isPlaceholder} options=${JSON.stringify(c.options)}`,
    );
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
