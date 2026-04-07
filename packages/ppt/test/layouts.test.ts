import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getSlideLayout,
  setSlideLayout,
  getLayouts,
} from "../src/layouts.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("getLayouts - returns all layouts in presentation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getLayouts(tempPath);
    assert.ok(result.ok, `getLayouts failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(Array.isArray(result.data.layouts));
    assert.ok(result.data.layouts.length > 0, "Expected at least one layout");
  } finally {
    // Clean up
  }
});

test("getLayouts - layout entries have required properties", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getLayouts(tempPath);
    assert.ok(result.ok);
    for (const layout of result.data!.layouts) {
      assert.ok(typeof layout.index === "number");
      assert.ok(typeof layout.name === "string");
    }
  } finally {
    // Clean up
  }
});

test("getSlideLayout - gets layout for a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getSlideLayout(tempPath, 1);
    assert.ok(result.ok, `getSlideLayout failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.layoutName === "string");
  } finally {
    // Clean up
  }
});

test("getSlideLayout - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getSlideLayout(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setSlideLayout - changes slide layout", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First get available layouts
    const layoutsResult = await getLayouts(tempPath);
    assert.ok(layoutsResult.ok);
    if (layoutsResult.data!.layouts.length < 2) {
      // Skip test if only one layout available
      return;
    }

    // Set to second layout
    const setResult = await setSlideLayout(tempPath, 1, 2);
    assert.ok(setResult.ok, `setSlideLayout failed: ${setResult.error?.message}`);

    // Verify the change
    const getResult = await getSlideLayout(tempPath, 1);
    assert.ok(getResult.ok);
    assert.equal(getResult.data!.layoutIndex, 2);
  } finally {
    // Clean up
  }
});

test("setSlideLayout - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setSlideLayout(tempPath, 999, 1);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setSlideLayout - returns error for invalid layout index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setSlideLayout(tempPath, 1, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});
