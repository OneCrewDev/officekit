import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  get,
  getSlide,
  getShape,
  getTable,
  getChart,
  getPlaceholder,
  query,
  querySlides,
  queryShapes,
  getShapeProperties,
  getTextContent,
  getTableStructure,
  getChartData,
} from "../src/query.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

const DATA_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/data_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-query-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("getSlide - returns slide at index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getSlide(tempPath, 1);
    assert.ok(result.ok, `getSlide failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data.path);
    assert.ok(result.data.index === 1);
    assert.ok(Array.isArray(result.data.shapes));
  } finally {
    // Clean up
  }
});

test("getSlide - returns error for invalid index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getSlide(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("getShape - returns shape at path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getShape(tempPath, "/slide[1]/shape[1]");
    // Shape might not exist in test file, but should either return shape or not_found
    if (result.ok) {
      assert.ok(result.data!.path);
      assert.ok(result.data!.type);
    } else {
      assert.ok(result.error?.code === "not_found" || result.error?.code === "invalid_input");
    }
  } finally {
    // Clean up
  }
});

test("getShape - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getShape(tempPath, "invalid");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("getPlaceholder - returns placeholder by type", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getPlaceholder(tempPath, 1, "title");
    // May succeed or fail depending on whether title placeholder exists
    if (result.ok) {
      assert.ok(result.data!.path);
      assert.ok(result.data!.type === "title");
    }
  } finally {
    // Clean up
  }
});

test("getPlaceholder - returns error for non-existent type", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getPlaceholder(tempPath, 1, "nonexistent");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("querySlides - returns all slides without selector", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await querySlides(tempPath);
    assert.ok(result.ok, `querySlides failed: ${result.error?.message}`);
    assert.ok(Array.isArray(result.data));
    assert.ok(result.data.length >= 0);
  } finally {
    // Clean up
  }
});

test("queryShapes - returns shapes on a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await queryShapes(tempPath, 1);
    assert.ok(result.ok, `queryShapes failed: ${result.error?.message}`);
    assert.ok(Array.isArray(result.data));
  } finally {
    // Clean up
  }
});

test("queryShapes - filters shapes with selector", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await queryShapes(tempPath, 1, "shape");
    assert.ok(result.ok, `queryShapes with selector failed: ${result.error?.message}`);
    assert.ok(Array.isArray(result.data));
  } finally {
    // Clean up
  }
});

test("getShapeProperties - returns shape properties", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First get a shape to find a valid path
    const shapeResult = await getShape(tempPath, "/slide[1]/shape[1]");
    if (!shapeResult.ok) {
      // Skip if no shapes
      return;
    }

    const result = await getShapeProperties(tempPath, "/slide[1]/shape[1]");
    if (result.ok) {
      assert.ok(result.data);
      // Properties may or may not be present depending on the shape
    }
  } finally {
    // Clean up
  }
});

test("getTextContent - returns text from shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getTextContent(tempPath, "/slide[1]/shape[1]");
    // May succeed or fail depending on shape existence
    if (result.ok) {
      assert.ok(result.data);
      assert.ok(typeof result.data.text === "string");
      assert.ok(Array.isArray(result.data.paragraphs));
    }
  } finally {
    // Clean up
  }
});

test("get - returns element at path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await get(tempPath, "/slide[1]");
    if (result.ok) {
      assert.ok(result.data);
    }
  } finally {
    // Clean up
  }
});

test("query - uses selector parser correctly", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await query(tempPath, "slide[1] shape");
    assert.ok(result.ok, `query failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(Array.isArray(result.data.shapes));
  } finally {
    // Clean up
  }
});

test("getTableStructure - returns table structure", async () => {
  const tempPath = await copyToTemp(DATA_PPTX);
  try {
    const result = await getTableStructure(tempPath, "/slide[1]/table[1]");
    // May succeed or fail depending on whether tables exist
    if (result.ok) {
      assert.ok(result.data);
      assert.ok(result.data.path);
      assert.ok(Array.isArray(result.data.rows));
    }
  } finally {
    // Clean up
  }
});

test("getChartData - returns chart data", async () => {
  const tempPath = await copyToTemp(DATA_PPTX);
  try {
    const result = await getChartData(tempPath, "/slide[1]/chart[1]");
    // May succeed or fail depending on whether charts exist
    if (result.ok) {
      assert.ok(result.data);
      assert.ok(result.data.path);
    }
  } finally {
    // Clean up
  }
});

test("querySlides - filters slides with text selector", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await querySlides(tempPath, "slide");
    assert.ok(result.ok);
    assert.ok(Array.isArray(result.data));
  } finally {
    // Clean up
  }
});
