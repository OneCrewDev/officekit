import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getNotes,
  setNotes,
  removeNotes,
} from "../src/notes.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("getNotes - returns empty string for slide without notes", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getNotes(tempPath, 1);
    assert.ok(result.ok, `getNotes failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.equal(result.data.text, "");
  } finally {
    // Clean up
  }
});

test("getNotes - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getNotes(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setNotes - sets notes text on a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const testNotes = "These are test notes for the slide";

    const setResult = await setNotes(tempPath, 1, testNotes);
    assert.ok(setResult.ok, `setNotes failed: ${setResult.error?.message}`);

    const getResult = await getNotes(tempPath, 1);
    assert.ok(getResult.ok);
    assert.ok(getResult.data!.text.includes("test notes"));
  } finally {
    // Clean up
  }
});

test("setNotes - supports multiline notes", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const testNotes = "Line 1\nLine 2\nLine 3";

    await setNotes(tempPath, 1, testNotes);

    const getResult = await getNotes(tempPath, 1);
    assert.ok(getResult.ok);
    assert.ok(getResult.data!.text.includes("Line 1"));
    assert.ok(getResult.data!.text.includes("Line 2"));
    assert.ok(getResult.data!.text.includes("Line 3"));
  } finally {
    // Clean up
  }
});

test("setNotes - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setNotes(tempPath, 999, "test");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeNotes - removes notes from a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First set notes
    await setNotes(tempPath, 1, "Test notes to be removed");

    // Verify notes exist
    const beforeResult = await getNotes(tempPath, 1);
    assert.ok(beforeResult.ok);
    assert.ok(beforeResult.data!.text.length > 0);

    // Remove notes
    const removeResult = await removeNotes(tempPath, 1);
    assert.ok(removeResult.ok, `removeNotes failed: ${removeResult.error?.message}`);

    // Verify notes are removed
    const afterResult = await getNotes(tempPath, 1);
    assert.ok(afterResult.ok);
    assert.equal(afterResult.data!.text, "");
  } finally {
    // Clean up
  }
});

test("removeNotes - handles slide without notes gracefully", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Try to remove notes from a slide that has none
    const result = await removeNotes(tempPath, 1);
    assert.ok(result.ok, `removeNotes failed: ${result.error?.message}`);
  } finally {
    // Clean up
  }
});

test("removeNotes - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeNotes(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});
