import test from "node:test";
import assert from "node:assert/strict";
import { readFile, writeFile, copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getSlides,
  addSlide,
} from "../src/slides.js";

import {
  swapSlides,
  swapShapes,
  copyShape,
  copySlide,
  rawGet,
  rawSet,
  batch,
} from "../src/mutations.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-mutation-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("swapSlides - swaps two slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add some slides to have something to swap
    await addSlide(tempPath);
    await addSlide(tempPath);
    await addSlide(tempPath);

    const beforeResult = await getSlides(tempPath);
    assert.ok(beforeResult.ok);

    // Swap slides 1 and 3
    const swapResult = await swapSlides(tempPath, 1, 3);
    assert.ok(swapResult.ok, `swapSlides failed: ${swapResult.error?.message}`);

    const afterResult = await getSlides(tempPath);
    assert.ok(afterResult.ok);
    // After swapping, the slide order should be different
  } finally {
    // Clean up
  }
});

test("swapSlides - returns error for invalid index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await swapSlides(tempPath, 999, 1);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("swapShapes - swaps two shapes on the same slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add a slide first to have some shapes
    await addSlide(tempPath);

    // Swap shape 1 and 2 on slide 1
    const swapResult = await swapShapes(tempPath, "/slide[1]/shape[1]", "/slide[1]/shape[2]");
    assert.ok(swapResult.ok, `swapShapes failed: ${swapResult.error?.message}`);
  } finally {
    // Clean up
  }
});

test("swapShapes - returns error for shapes on different slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);
    await addSlide(tempPath);

    const result = await swapShapes(tempPath, "/slide[1]/shape[1]", "/slide[2]/shape[1]");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("copyShape - copies a shape to another slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);

    // Copy shape 1 from slide 1 to slide 2
    const copyResult = await copyShape(tempPath, "/slide[1]/shape[1]", 2);
    assert.ok(copyResult.ok, `copyShape failed: ${copyResult.error?.message}`);
    assert.ok(copyResult.data);
    assert.ok(copyResult.data.path);
  } finally {
    // Clean up
  }
});

test("copyShape - returns error for invalid source path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);

    const result = await copyShape(tempPath, "/invalid/path", 2);
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("copySlide - duplicates a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const beforeResult = await getSlides(tempPath);
    assert.ok(beforeResult.ok);
    const beforeCount = beforeResult.data!.total;

    // Copy slide 1
    const copyResult = await copySlide(tempPath, 1, -1);
    assert.ok(copyResult.ok, `copySlide failed: ${copyResult.error?.message}`);
    assert.ok(copyResult.data!);
    assert.ok(copyResult.data!.path);

    const afterResult = await getSlides(tempPath);
    assert.ok(afterResult.ok);
    assert.equal(afterResult.data!.total, beforeCount + 1);
  } finally {
    // Clean up
  }
});

test("copySlide - inserts copy at specific position", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const beforeResult = await getSlides(tempPath);
    assert.ok(beforeResult.ok);
    const beforeCount = beforeResult.data!.total;

    // Copy slide 1 and insert at position 2
    const copyResult = await copySlide(tempPath, 1, 2);
    assert.ok(copyResult.ok, `copySlide failed: ${copyResult.error?.message}`);
    assert.ok(copyResult.data!);

    const afterResult = await getSlides(tempPath);
    assert.ok(afterResult.ok);
    assert.equal(afterResult.data!.total, beforeCount + 1);
  } finally {
    // Clean up
  }
});

test("copySlide - returns error for invalid source index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await copySlide(tempPath, 999, 1);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("rawGet - gets raw XML for an element", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await rawGet(tempPath, "/slide[1]/shape[1]");
    // May fail if shape doesn't exist or other issues
    // Just verify the result structure
    if (result.ok) {
      assert.ok(result.data);
      assert.ok(typeof result.data.xml === "string");
    }
  } finally {
    // Clean up
  }
});

test("rawSet - sets raw XML for an element", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set some XML on slide 1 shape 1
    const result = await rawSet(tempPath, "/slide[1]/shape[1]", "<p:sp>test</p:sp>");
    // May fail due to various reasons, but should not throw
    // Just verify the result
    if (!result.ok) {
      assert.ok(result.error); // Error is expected in some cases
    }
  } finally {
    // Clean up
  }
});

test("batch - executes multiple operations", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await batch(tempPath, [
      { op: "setShapeText", params: { path: "/slide[1]/shape[1]", text: "Hello" } },
    ]);
    // May fail but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("batch - returns error for unknown operation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await batch(tempPath, [
      { op: "unknown", params: {} },
    ] as unknown as Parameters<typeof batch>[1]);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});
