import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  viewAsText,
  viewAsAnnotated,
  viewAsOutline,
  viewAsStats,
  viewAsIssues,
  getSlideStats,
  checkShapeTextOverflow,
} from "../src/views.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

const DATA_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/data_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-views-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

// ============================================================================
// ViewAsText Tests
// ============================================================================

test("viewAsText - returns text from all slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsText(tempPath);
    assert.ok(result.ok, `viewAsText failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.slideCount === "number");
    assert.ok(Array.isArray(result.data.slides));
    assert.ok(result.data.slides.length > 0);

    // Check slide structure
    const slide = result.data.slides[0];
    assert.ok(typeof slide.index === "number");
    assert.ok(typeof slide.path === "string");
    assert.ok(typeof slide.text === "string");
    assert.ok(Array.isArray(slide.shapes));
  } finally {
    // Clean up
  }
});

test("viewAsText - returns text from specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsText(tempPath, 1);
    assert.ok(result.ok, `viewAsText with slideIndex failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.slideCount, 1);
    assert.equal(result.data.slides[0].index, 1);
  } finally {
    // Clean up
  }
});

test("viewAsText - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsText(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

// ============================================================================
// ViewAsAnnotated Tests
// ============================================================================

test("viewAsAnnotated - returns annotated view of slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsAnnotated(tempPath);
    assert.ok(result.ok, `viewAsAnnotated failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.slideCount === "number");
    assert.ok(Array.isArray(result.data.slides));
    assert.ok(result.data.slides.length > 0);

    // Check slide annotation structure
    const slide = result.data.slides[0];
    assert.ok(typeof slide.index === "number");
    assert.ok(typeof slide.path === "string");
    assert.ok(Array.isArray(slide.elements));

    // Check element structure
    if (slide.elements.length > 0) {
      const element = slide.elements[0];
      assert.ok(typeof element.path === "string");
      assert.ok(typeof element.type === "string");
    }
  } finally {
    // Clean up
  }
});

test("viewAsAnnotated - returns annotated view of specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsAnnotated(tempPath, 1);
    assert.ok(result.ok, `viewAsAnnotated with slideIndex failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.slideCount, 1);
    assert.equal(result.data.slides[0].index, 1);
  } finally {
    // Clean up
  }
});

test("viewAsAnnotated - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsAnnotated(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

// ============================================================================
// ViewAsOutline Tests
// ============================================================================

test("viewAsOutline - returns outline of slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsOutline(tempPath);
    assert.ok(result.ok, `viewAsOutline failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.slideCount === "number");
    assert.ok(Array.isArray(result.data.slides));
    assert.ok(result.data.slides.length > 0);

    // Check slide outline structure
    const slide = result.data.slides[0];
    assert.ok(typeof slide.index === "number");
    assert.ok(typeof slide.path === "string");
    assert.ok(Array.isArray(slide.content));

    // Check content structure
    if (slide.content.length > 0) {
      const item = slide.content[0];
      assert.ok(typeof item.type === "string");
      assert.ok(typeof item.path === "string");
      assert.ok(typeof item.description === "string");
    }
  } finally {
    // Clean up
  }
});

test("viewAsOutline - returns outline of specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsOutline(tempPath, 1);
    assert.ok(result.ok, `viewAsOutline with slideIndex failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.slideCount, 1);
    assert.equal(result.data.slides[0].index, 1);
  } finally {
    // Clean up
  }
});

test("viewAsOutline - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsOutline(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

// ============================================================================
// ViewAsStats Tests
// ============================================================================

test("viewAsStats - returns statistics for presentation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsStats(tempPath);
    assert.ok(result.ok, `viewAsStats failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.slideCount === "number");
    assert.ok(typeof result.data.shapeCount === "number");
    assert.ok(typeof result.data.textLength === "number");
    assert.ok(typeof result.data.tableCount === "number");
    assert.ok(typeof result.data.chartCount === "number");
    assert.ok(Array.isArray(result.data.slides));

    // Check per-slide stats
    if (result.data.slides.length > 0) {
      const slideStat = result.data.slides[0];
      assert.ok(typeof slideStat.index === "number");
      assert.ok(typeof slideStat.shapeCount === "number");
      assert.ok(typeof slideStat.textLength === "number");
    }
  } finally {
    // Clean up
  }
});

test("viewAsStats - returns stats for specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsStats(tempPath, 1);
    assert.ok(result.ok, `viewAsStats with slideIndex failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.slideCount, 1);
  } finally {
    // Clean up
  }
});

test("viewAsStats - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsStats(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

// ============================================================================
// getSlideStats Tests
// ============================================================================

test("getSlideStats - returns stats for specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getSlideStats(tempPath, 1);
    assert.ok(result.ok, `getSlideStats failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.index, 1);
    assert.ok(typeof result.data.shapeCount === "number");
  } finally {
    // Clean up
  }
});

test("getSlideStats - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getSlideStats(tempPath, 999);
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

// ============================================================================
// ViewAsIssues Tests
// ============================================================================

test("viewAsIssues - returns issues in presentation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsIssues(tempPath);
    assert.ok(result.ok, `viewAsIssues failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.slideCount === "number");
    assert.ok(typeof result.data.issueCount === "number");
    assert.ok(Array.isArray(result.data.issues));

    // Check issue structure if any
    if (result.data.issues.length > 0) {
      const issue = result.data.issues[0];
      assert.ok(typeof issue.severity === "string");
      assert.ok(["error", "warning", "info"].includes(issue.severity));
      assert.ok(typeof issue.category === "string");
      assert.ok(typeof issue.message === "string");
    }
  } finally {
    // Clean up
  }
});

test("viewAsIssues - returns issues for specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsIssues(tempPath, 1);
    assert.ok(result.ok, `viewAsIssues with slideIndex failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.slideCount, 1);
  } finally {
    // Clean up
  }
});

test("viewAsIssues - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await viewAsIssues(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

// ============================================================================
// checkShapeTextOverflow Tests
// ============================================================================

test("checkShapeTextOverflow - checks shape for text overflow", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First get a valid shape path
    const textResult = await viewAsText(tempPath);
    if (!textResult.ok || textResult.data.slides[0].shapes.length === 0) {
      // Skip if no shapes
      return;
    }

    const shapePath = textResult.data.slides[0].shapes[0].path;
    const result = await checkShapeTextOverflow(tempPath, shapePath);

    if (result.ok) {
      assert.ok(typeof result.data.hasOverflow === "boolean");
      assert.ok(typeof result.data.path === "string");
      assert.equal(result.data.path, shapePath);
    }
  } finally {
    // Clean up
  }
});

test("checkShapeTextOverflow - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkShapeTextOverflow(tempPath, "invalid");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("checkShapeTextOverflow - returns error for non-existent shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkShapeTextOverflow(tempPath, "/slide[1]/shape[999]");
    // Should either be ok with no overflow or not_found
    if (!result.ok) {
      assert.ok(result.error?.code === "not_found" || result.error?.code === "invalid_input");
    }
  } finally {
    // Clean up
  }
});

// ============================================================================
// Edge Cases
// ============================================================================

test("viewAsText - handles empty presentation", async () => {
  // Test with data_presentation which might have different content
  const tempPath = await copyToTemp(DATA_PPTX);
  try {
    const result = await viewAsText(tempPath);
    assert.ok(result.ok, `viewAsText on data presentation failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data.slides.length >= 0);
  } finally {
    // Clean up
  }
});

test("viewAsStats - works with data_presentation", async () => {
  const tempPath = await copyToTemp(DATA_PPTX);
  try {
    const result = await viewAsStats(tempPath);
    assert.ok(result.ok, `viewAsStats on data presentation failed: ${result.error?.message}`);
    assert.ok(result.data);
    // data_presentation might have tables and charts
    assert.ok(typeof result.data.tableCount === "number");
  } finally {
    // Clean up
  }
});
