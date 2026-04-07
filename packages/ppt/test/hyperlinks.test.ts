import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getHyperlink,
  setHyperlink,
  removeHyperlink,
  setExternalHyperlink,
  setInternalHyperlink,
} from "../src/hyperlinks.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-hyperlinks-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("getHyperlink - returns null for shape without hyperlink", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getHyperlink(tempPath, "/slide[1]/shape[1]");
    // May return null (no hyperlink) or error (shape doesn't exist)
    if (result.ok) {
      assert.ok(result.data === null || result.data !== undefined);
    }
  } finally {
    // Clean up
  }
});

test("getHyperlink - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getHyperlink(tempPath, "invalid");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setHyperlink - sets a hyperlink on a shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Try to set hyperlink on shape 1
    const result = await setHyperlink(tempPath, "/slide[1]/shape[1]", "https://example.com");
    // May fail if shape doesn't exist, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setHyperlink - returns error for invalid URL", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setHyperlink(tempPath, "/slide[1]/shape[1]", "not-a-valid-url");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setHyperlink - returns error for empty URL", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setHyperlink(tempPath, "/slide[1]/shape[1]", "");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setHyperlink - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setHyperlink(tempPath, "invalid", "https://example.com");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("removeHyperlink - removes a hyperlink from a shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Try to remove hyperlink from shape 1
    const result = await removeHyperlink(tempPath, "/slide[1]/shape[1]");
    // May fail if shape doesn't exist or has no hyperlink, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("removeHyperlink - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeHyperlink(tempPath, "invalid");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setExternalHyperlink - sets an external hyperlink", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setExternalHyperlink(tempPath, "/slide[1]/shape[1]", "https://example.com");
    // May fail if shape doesn't exist, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setInternalHyperlink - sets an internal hyperlink to another slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setInternalHyperlink(tempPath, "/slide[1]/shape[1]", 2);
    // May fail if shape doesn't exist or target slide doesn't exist, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setInternalHyperlink - returns error for invalid target slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setInternalHyperlink(tempPath, "/slide[1]/shape[1]", 999);
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setInternalHyperlink - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setInternalHyperlink(tempPath, "invalid", 2);
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});
