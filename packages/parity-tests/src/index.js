import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { getLineageStatement } from "../../docs/src/index.mjs";
import {
  getWordAdapterManifest,
  summarizeWordAdapter,
  summarizeWordAdapterContract,
} from "../../word/src/index.js";
import {
  getExcelAdapterManifest,
  summarizeExcelAdapter,
  summarizeExcelAdapterContract,
} from "../../excel/src/index.js";
import {
  getPptAdapterManifest,
  summarizePptAdapter,
  summarizePptAdapterContract,
} from "../../ppt/src/index.js";

const moduleDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(moduleDir, "../../..");
const fixtureManifestPath = resolve(moduleDir, "../fixtures/manifest.json");
const documentationStatusPath = "docs/parity/implementation-status.md";

const formatDefinitions = Object.freeze([
  {
    format: "word",
    categoryPrefixes: ["word-"],
    getManifest: getWordAdapterManifest,
    getSummary: summarizeWordAdapter,
    getContractSummary: summarizeWordAdapterContract,
  },
  {
    format: "excel",
    categoryPrefixes: ["excel-"],
    getManifest: getExcelAdapterManifest,
    getSummary: summarizeExcelAdapter,
    getContractSummary: summarizeExcelAdapterContract,
  },
  {
    format: "powerpoint",
    categoryPrefixes: ["ppt-"],
    getManifest: getPptAdapterManifest,
    getSummary: summarizePptAdapter,
    getContractSummary: summarizePptAdapterContract,
  },
]);

export function loadFixtureManifest() {
  return JSON.parse(readFileSync(fixtureManifestPath, "utf8"));
}

export function listDocumentationStatusLines() {
  return readFileSync(resolve(repoRoot, documentationStatusPath), "utf8")
    .split("\n")
    .filter((line) => line.trim().length > 0);
}

function collectFixtures(fixtures, prefixes) {
  return fixtures.filter((fixture) =>
    prefixes.some(
      (prefix) => fixture.id.startsWith(prefix) || String(fixture.category ?? "").startsWith(prefix),
    ),
  );
}

function summarizeVerificationModes(fixtures) {
  return fixtures.reduce((acc, fixture) => {
    for (const mode of fixture.verification ?? []) {
      acc[mode] = (acc[mode] ?? 0) + 1;
    }
    return acc;
  }, {});
}

export function createFormatParityStatusReport() {
  const fixtureManifest = loadFixtureManifest();

  return formatDefinitions.map((definition) => {
    const manifest = definition.getManifest();
    const summary = definition.getSummary();
    const contractSummary = definition.getContractSummary();
    const fixtures = collectFixtures(fixtureManifest.fixtures, definition.categoryPrefixes);

    return {
      format: definition.format,
      packageName: manifest.packageName,
      status: "scaffolded",
      fixtureCount: fixtures.length,
      fixtureIds: fixtures.map((fixture) => fixture.id),
      copiedFixtureCount: fixtures.filter((fixture) => fixture.mode === "copied").length,
      referencedFixtureCount: fixtures.filter((fixture) => fixture.mode === "referenced").length,
      verificationModes: summarizeVerificationModes(fixtures),
      publicSurfaceCount: summary.surfaceCount,
      previewModes: [...contractSummary.previewModes],
      canonicalPathCount: contractSummary.canonicalPathCount,
      capabilityFamilies: manifest.capabilityFamilies,
      implementationMilestones: [...manifest.implementationMilestones],
      parityRisks: [...manifest.parityRisks],
    };
  });
}

export function createDocumentationParityReport() {
  return {
    lineage: getLineageStatement(),
    statusDocument: documentationStatusPath,
    formats: createFormatParityStatusReport(),
  };
}

export function summarizeFixtureCoverage() {
  const fixtureManifest = loadFixtureManifest();
  const formats = createFormatParityStatusReport();

  return {
    sourceProject: fixtureManifest.sourceProject,
    totalFixtures: fixtureManifest.fixtures.length,
    copiedFixtures: fixtureManifest.fixtures.filter((fixture) => fixture.mode === "copied").length,
    referencedFixtures: fixtureManifest.fixtures.filter((fixture) => fixture.mode === "referenced").length,
    formats,
  };
}
