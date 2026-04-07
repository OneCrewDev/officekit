import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  viewAsHtml,
  generatePreview,
} from "../src/preview-html.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

const DATA_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/data_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-preview-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

// ============================================================================
// viewAsHtml Tests
// ============================================================================

test("viewAsHtml - returns HTML for all slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsHtml(tempPath);
    assert.ok(result.ok, `viewAsHtml failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.slideCount === "number");
    assert.ok(result.data.slideCount > 0);
    assert.ok(typeof result.data.html === "string");
    assert.ok(result.data.html.length > 0);

    // Check HTML structure
    assert.ok(result.data.html.includes("<!DOCTYPE html>"));
    assert.ok(result.data.html.includes("<html"));
    assert.ok(result.data.html.includes("</html>"));
    assert.ok(result.data.html.includes("<head>"));
    assert.ok(result.data.html.includes("<body>"));
  } finally {
    // Clean up
  }
});

test("viewAsHtml - returns HTML for specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsHtml(tempPath, 1);
    assert.ok(result.ok, `viewAsHtml with slideIndex failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.slideCount, 1);
    assert.ok(typeof result.data.html === "string");
    assert.ok(result.data.html.includes("Slide 1"));
  } finally {
    // Clean up
  }
});

test("viewAsHtml - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsHtml(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("viewAsHtml - contains slide containers", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsHtml(tempPath);
    assert.ok(result.ok, `viewAsHtml failed: ${result.error?.message}`);
    assert.ok(result.data);

    // Check for slide containers
    assert.ok(result.data.html.includes('class="slide-container"'));
    assert.ok(result.data.html.includes('class="slide"'));
    assert.ok(result.data.html.includes('class="slide-label"'));
  } finally {
    // Clean up
  }
});

test("viewAsHtml - handles data_presentation with tables", async () => {
  const tempPath = await copyToTemp(DATA_PPTX);
  try {
    const result = await viewAsHtml(tempPath);
    assert.ok(result.ok, `viewAsHtml on data presentation failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data.html.length > 0);
  } finally {
    // Clean up
  }
});

// ============================================================================
// viewAsSvg Tests
// ============================================================================

test("viewAsSvg - returns SVG for all slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Dynamic import to test the SVG module
    const { viewAsSvg } = await import("../src/preview-svg.js");
    const result = await viewAsSvg(tempPath);
    assert.ok(result.ok, `viewAsSvg failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.slideCount === "number");
    assert.ok(result.data.slideCount > 0);
    assert.ok(typeof result.data.svg === "string");
    assert.ok(result.data.svg.length > 0);

    // Check SVG structure
    assert.ok(result.data.svg.includes("<svg"));
    assert.ok(result.data.svg.includes("xmlns="));
    assert.ok(result.data.svg.includes("viewBox="));
  } finally {
    // Clean up
  }
});

test("viewAsSvg - returns SVG for specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const { viewAsSvg } = await import("../src/preview-svg.js");
    const result = await viewAsSvg(tempPath, 1);
    assert.ok(result.ok, `viewAsSvg with slideIndex failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.slideCount, 1);
    assert.ok(typeof result.data.svg === "string");
    assert.ok(result.data.svg.includes('class="slide"'));
  } finally {
    // Clean up
  }
});

test("viewAsSvg - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const { viewAsSvg } = await import("../src/preview-svg.js");
    const result = await viewAsSvg(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("viewAsSvg - contains SVG elements", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const { viewAsSvg } = await import("../src/preview-svg.js");
    const result = await viewAsSvg(tempPath);
    assert.ok(result.ok, `viewAsSvg failed: ${result.error?.message}`);
    assert.ok(result.data);

    // Check for SVG elements
    assert.ok(result.data.svg.includes("<svg"));
    assert.ok(result.data.svg.includes("viewBox="));
  } finally {
    // Clean up
  }
});

// ============================================================================
// generatePreview Tests
// ============================================================================

test("generatePreview - generates HTML preview by default", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await generatePreview(tempPath);
    assert.ok(result.ok, `generatePreview failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.format, "html");
    assert.ok(typeof result.data.output === "string");
    assert.ok(result.data.output.length > 0);
  } finally {
    // Clean up
  }
});

test("generatePreview - generates HTML preview explicitly", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await generatePreview(tempPath, { format: "html" });
    assert.ok(result.ok, `generatePreview with html format failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.format, "html");
    assert.ok(result.data.output.includes("<!DOCTYPE html>"));
  } finally {
    // Clean up
  }
});

test("generatePreview - generates SVG preview", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await generatePreview(tempPath, { format: "svg" });
    assert.ok(result.ok, `generatePreview with svg format failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.format, "svg");
    assert.ok(typeof result.data.output === "string");
    assert.ok(result.data.output.includes("<svg"));
  } finally {
    // Clean up
  }
});

test("generatePreview - generates preview for specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await generatePreview(tempPath, { format: "html", slides: [1] });
    assert.ok(result.ok, `generatePreview with slides option failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.slideCount, 1);
  } finally {
    // Clean up
  }
});

test("generatePreview - handles non-existent file", async () => {
  const result = await generatePreview("/non/existent/path.pptx");
  assert.ok(!result.ok);
  assert.ok(result.error);
});

// ============================================================================
// Edge Cases
// ============================================================================

test("viewAsHtml - handles file without valid slides", async () => {
  // This test checks that the function returns proper error for invalid files
  const result = await viewAsHtml("/non/existent/path.pptx");
  assert.ok(!result.ok);
  assert.ok(result.error);
});

test("viewAsSvg - handles file without valid slides", async () => {
  const { viewAsSvg } = await import("../src/preview-svg.js");
  const result = await viewAsSvg("/non/existent/path.pptx");
  assert.ok(!result.ok);
  assert.ok(result.error);
});
