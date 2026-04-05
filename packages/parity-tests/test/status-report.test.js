import test from "node:test";
import assert from "node:assert/strict";

import {
  createDocumentationParityReport,
  createFormatParityStatusReport,
  listDocumentationStatusLines,
  summarizeFixtureCoverage,
} from "../src/index.js";

test("fixture-backed parity report links Word, Excel, and PowerPoint slices", () => {
  const formats = createFormatParityStatusReport();
  const word = formats.find((entry) => entry.format === "word");
  const excel = formats.find((entry) => entry.format === "excel");
  const powerpoint = formats.find((entry) => entry.format === "powerpoint");

  assert.ok(word);
  assert.ok(excel);
  assert.ok(powerpoint);

  assert.ok(word.fixtureIds.includes("word-formulas-script"));
  assert.ok(word.fixtureIds.includes("word-tables-script"));
  assert.ok(word.fixtureIds.includes("word-textbox-script"));
  assert.deepEqual(word.previewModes, ["html", "forms"]);

  assert.ok(excel.fixtureIds.includes("excel-beautiful-charts-script"));
  assert.ok(excel.fixtureIds.includes("excel-charts-demo-output"));
  assert.ok(excel.capabilityFamilies.calculations.includes("formulas"));
  assert.ok(excel.capabilityFamilies.calculations.includes("pivots"));

  assert.ok(powerpoint.fixtureIds.includes("ppt-beautiful-script"));
  assert.ok(powerpoint.fixtureIds.includes("ppt-animations-script"));
  assert.ok(powerpoint.fixtureIds.includes("ppt-3d-model-asset"));
  assert.ok(powerpoint.capabilityFamilies.renderingAndValidation.includes("animations"));
  assert.deepEqual(powerpoint.previewModes, ["html", "svg"]);
});

test("parity coverage summary keeps scaffold status and remaining implementation work explicit", () => {
  const coverage = summarizeFixtureCoverage();
  const documentation = createDocumentationParityReport();

  assert.equal(coverage.sourceProject, "../OfficeCLI");
  assert.ok(coverage.totalFixtures >= 20);
  assert.ok(coverage.referencedFixtures >= 1);
  assert.match(documentation.lineage, /migrated from OfficeCLI/i);

  for (const format of documentation.formats) {
    assert.equal(format.status, "scaffolded");
    assert.ok(format.fixtureCount > 0, `missing fixtures for ${format.format}`);
    assert.ok(format.canonicalPathCount > 0, `missing canonical paths for ${format.format}`);
    assert.ok(format.implementationMilestones.length > 0, `missing milestones for ${format.format}`);
    assert.ok(format.parityRisks.length > 0, `missing risks for ${format.format}`);
  }
});

test("documentation status file reports support evidence and remaining gaps", () => {
  const lines = listDocumentationStatusLines().join("\n");

  assert.match(lines, /Current implementation status/i);
  assert.match(lines, /Word/i);
  assert.match(lines, /Excel/i);
  assert.match(lines, /PowerPoint/i);
  assert.match(lines, /Remaining gaps/i);
  assert.match(lines, /fixture-backed evidence/i);
});
