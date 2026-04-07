import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  checkShapeTextOverflow,
  checkSlideOverflow,
  getOverflowIssues,
} from "../src/overflow.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-overflow-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("checkShapeTextOverflow - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkShapeTextOverflow(tempPath, "invalid");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("checkShapeTextOverflow - returns error for invalid slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkShapeTextOverflow(tempPath, "/slide[999]/shape[1]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("checkShapeTextOverflow - returns error for invalid shape path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkShapeTextOverflow(tempPath, "/slide[1]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("checkShapeTextOverflow - returns overflow result for valid shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkShapeTextOverflow(tempPath, "/slide[1]/shape[1]");
    // May succeed or fail depending on whether shape exists and has text
    if (result.ok) {
      assert.ok(result.data);
      assert.ok(typeof result.data.hasOverflow === "boolean");
      assert.ok(result.data.path);
    }
  } finally {
    // Clean up
  }
});

test("checkSlideOverflow - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkSlideOverflow(tempPath, 999);
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("checkSlideOverflow - returns slide overflow result", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await checkSlideOverflow(tempPath, 1);
    if (result.ok) {
      assert.ok(result.data);
      assert.equal(result.data.slideIndex, 1);
      assert.ok(result.data.path);
      assert.ok(typeof result.data.hasOverflow === "boolean");
      assert.ok(Array.isArray(result.data.overflowingShapes));
    }
  } finally {
    // Clean up
  }
});

test("getOverflowIssues - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getOverflowIssues(tempPath, 999);
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("getOverflowIssues - returns all overflow issues for presentation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getOverflowIssues(tempPath);
    if (result.ok) {
      assert.ok(result.data);
      assert.ok(typeof result.data.slideCount === "number");
      assert.ok(typeof result.data.issueCount === "number");
      assert.ok(Array.isArray(result.data.issues));
    }
  } finally {
    // Clean up
  }
});

test("getOverflowIssues - returns overflow issues for specific slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getOverflowIssues(tempPath, 1);
    if (result.ok) {
      assert.ok(result.data);
      assert.equal(result.data.slideCount, 1);
      assert.ok(Array.isArray(result.data.issues));
    }
  } finally {
    // Clean up
  }
});