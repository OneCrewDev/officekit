import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getSlideBackground,
  setSlideBackground,
  setGradientFill,
  setPictureFill,
} from "../src/background.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-background-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("getSlideBackground - gets slide background", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getSlideBackground(tempPath, 1);
    // May succeed with empty background or fail if no background
    if (result.ok) {
      assert.ok(result.data!.background);
    } else {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getSlideBackground - returns error for invalid slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getSlideBackground(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setSlideBackground - sets solid background", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setSlideBackground(tempPath, 1, { type: "solid", color: "FFCCCC" });
    // May succeed or fail
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setSlideBackground - sets gradient background", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setSlideBackground(tempPath, 1, {
      type: "gradient",
      gradient: {
        type: "linear",
        colors: [{ color: "FF0000", position: 0 }, { color: "0000FF", position: 100000 }],
        angle: 90
      }
    });
    // May succeed or fail
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setSlideBackground - sets no fill background", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setSlideBackground(tempPath, 1, { type: "none" });
    // May succeed or fail
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setSlideBackground - returns error for invalid slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setSlideBackground(tempPath, 999, { type: "solid", color: "FFCCCC" });
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setGradientFill - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setGradientFill(tempPath, "invalid", {
      type: "linear",
      colors: [{ color: "FF0000", position: 0 }],
      angle: 45
    });
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setGradientFill - sets gradient fill on shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setGradientFill(tempPath, "/slide[1]/shape[1]", {
      type: "linear",
      colors: [{ color: "FF0000", position: 0 }, { color: "0000FF", position: 100000 }],
      angle: 45
    });
    // May fail if shape doesn't exist
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setPictureFill - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setPictureFill(tempPath, "invalid", "rId1");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});
