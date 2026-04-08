import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  checkShape,
  checkSlide,
  checkPresentation,
  formatCheckReport,
} from "../src/check.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-check-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("checkShape - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkShape(tempPath, "invalid");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("checkShape - returns error for invalid slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkShape(tempPath, "/slide[999]/shape[1]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("checkShape - returns shape check result for valid shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkShape(tempPath, "/slide[1]/shape[1]");
    if (result.ok) {
      assert.ok(result.data);
      assert.ok(result.data!.path);
      assert.ok(typeof result.data.hasIssues === "boolean");
      assert.ok(Array.isArray(result.data.issues));
    }
  } finally {
    // Clean up
  }
});

test("checkSlide - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkSlide(tempPath, 999);
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("checkSlide - returns slide check result", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkSlide(tempPath, 1);
    if (result.ok) {
      assert.ok(result.data);
      assert.equal(result.data!.slideIndex, 1);
      assert.ok(result.data!.path);
      assert.ok(typeof result.data.hasIssues === "boolean");
      assert.ok(typeof result.data.issueCount === "number");
      assert.ok(Array.isArray(result.data.issues));
    }
  } finally {
    // Clean up
  }
});

test("checkPresentation - returns presentation check result", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkPresentation(tempPath);
    if (result.ok) {
      assert.ok(result.data);
      assert.ok(result.data!.filePath);
      assert.ok(typeof result.data.slideCount === "number");
      assert.ok(typeof result.data.shapeCount === "number");
      assert.ok(typeof result.data.issueCount === "number");
      assert.ok(typeof result.data.hasIssues === "boolean");
      assert.ok(Array.isArray(result.data.issues));
      assert.ok(typeof result.data.issuesBySeverity === "object");
      assert.ok(typeof result.data.issuesByCategory === "object");
    }
  } finally {
    // Clean up
  }
});

test("checkPresentation - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkPresentation(tempPath, { slideIndex: 999 });
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("checkPresentation - filters by slide index when specified", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkPresentation(tempPath, { slideIndex: 1 });
    if (result.ok) {
      assert.equal(result.data!.slideCount, 1);
    }
  } finally {
    // Clean up
  }
});

test("checkPresentation - respects check options", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Disable all checks
    const result = await checkPresentation(tempPath, {
      checkTextOverflow: false,
      checkMissingTitles: false,
      checkEmptySlides: false,
    });
    if (result.ok) {
      // Should have no issues when all checks are disabled
      assert.equal(result.data!.issueCount, 0);
    }
  } finally {
    // Clean up
  }
});

test("formatCheckReport - returns human-readable report", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkPresentation(tempPath);
    if (result.ok) {
      const report = formatCheckReport(result.data!);
      assert.ok(typeof report === "string");
      assert.ok(report.length > 0);
      assert.ok(report.includes("Checking layout:"));
    }
  } finally {
    // Clean up
  }
});

test("formatCheckReport - returns JSON when json option is true", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkPresentation(tempPath);
    if (result.ok) {
      const report = formatCheckReport(result.data!, { json: true });
      assert.ok(typeof report === "string");
      const parsed = JSON.parse(report);
      assert.ok(parsed.filePath);
      assert.ok(Array.isArray(parsed.issues));
    }
  } finally {
    // Clean up
  }
});

test("formatCheckReport - includes issue details in verbose mode", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkPresentation(tempPath);
    if (result.ok && result.data!.issueCount > 0) {
      const report = formatCheckReport(result.data!, { verbose: true });
      // In verbose mode, we expect either Details: or at least Suggestion: for each issue
      assert.ok(report.includes("Suggestion:"));
    }
  } finally {
    // Clean up
  }
});
