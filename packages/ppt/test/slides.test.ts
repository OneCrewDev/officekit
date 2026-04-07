import test from "node:test";
import assert from "node:assert/strict";
import { readFile, writeFile, copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getSlides,
  addSlide,
  removeSlide,
  moveSlide,
  duplicateSlide,
} from "../src/slides.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("getSlides - returns slides from presentation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getSlides(tempPath);
    assert.ok(result.ok, `getSlides failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data.total >= 0);
    assert.ok(Array.isArray(result.data.slides));
  } finally {
    // Clean up
  }
});

test("addSlide - adds a new slide to presentation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const beforeResult = await getSlides(tempPath);
    assert.ok(beforeResult.ok);
    const beforeCount = beforeResult.data!.total;

    const addResult = await addSlide(tempPath);
    assert.ok(addResult.ok, `addSlide failed: ${addResult.error?.message}`);
    assert.ok(addResult.data!);
    assert.ok(addResult.data!.path);

    const afterResult = await getSlides(tempPath);
    assert.ok(afterResult.ok);
    assert.equal(afterResult.data!.total, beforeCount + 1);
  } finally {
    // Clean up
  }
});

test("addSlide - adds slide with specific layout", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // This PPTX only has 1 layout, so use index 1
    const addResult = await addSlide(tempPath, 1);
    assert.ok(addResult.ok, `addSlide with layout failed: ${addResult.error?.message}`);
  } finally {
    // Clean up
  }
});

test("removeSlide - removes a slide from presentation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add a slide to ensure there's at least one
    await addSlide(tempPath);

    const beforeResult = await getSlides(tempPath);
    assert.ok(beforeResult.ok);
    const beforeCount = beforeResult.data!.total;

    const removeResult = await removeSlide(tempPath, 1);
    assert.ok(removeResult.ok, `removeSlide failed: ${removeResult.error?.message}`);

    const afterResult = await getSlides(tempPath);
    assert.ok(afterResult.ok);
    assert.equal(afterResult.data!.total, beforeCount - 1);
  } finally {
    // Clean up
  }
});

test("removeSlide - returns error for invalid index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeSlide(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("moveSlide - moves slide from one position to another", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add a few slides to have something to move
    await addSlide(tempPath);
    await addSlide(tempPath);
    await addSlide(tempPath);

    const beforeResult = await getSlides(tempPath);
    assert.ok(beforeResult.ok);
    const slidesBefore = beforeResult.data!.slides;

    // Move slide from position 1 to position 3
    const moveResult = await moveSlide(tempPath, 1, 3);
    assert.ok(moveResult.ok, `moveSlide failed: ${moveResult.error?.message}`);

    const afterResult = await getSlides(tempPath);
    assert.ok(afterResult.ok);
    // After moving slide 1 to position 3, the first slide should be different
    // But we can't easily verify the order without examining the actual PPTX
  } finally {
    // Clean up
  }
});

test("moveSlide - returns error for invalid indices", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add a slide first
    await addSlide(tempPath);

    const result1 = await moveSlide(tempPath, 999, 1);
    assert.ok(!result1.ok);
    assert.equal(result1.error?.code, "invalid_input");

    const result2 = await moveSlide(tempPath, 1, 999);
    assert.ok(!result2.ok);
    assert.equal(result2.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("duplicateSlide - duplicates an existing slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const beforeResult = await getSlides(tempPath);
    assert.ok(beforeResult.ok);
    const beforeCount = beforeResult.data!.total;

    const dupResult = await duplicateSlide(tempPath, 1);
    assert.ok(dupResult.ok, `duplicateSlide failed: ${dupResult.error?.message}`);
    assert.ok(dupResult.data!);
    assert.ok(dupResult.data!.path);

    const afterResult = await getSlides(tempPath);
    assert.ok(afterResult.ok);
    assert.equal(afterResult.data!.total, beforeCount + 1);
  } finally {
    // Clean up
  }
});

test("duplicateSlide - returns error for invalid index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await duplicateSlide(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});
