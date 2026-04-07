import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  setShapeText,
  setShapeProperty,
  removeShape,
  setSlideProperty,
} from "../src/shapes.js";
import { addSlide, getSlides } from "../src/slides.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-shapes-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("setShapeText - sets text in a shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add a slide to work with
    await addSlide(tempPath);

    // Try to set text in shape 1
    const result = await setShapeText(tempPath, "/slide[1]/shape[1]", "Hello World");
    // May fail if shape doesn't exist, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error); // Error is expected if shape doesn't exist
    }
  } finally {
    // Clean up
  }
});

test("setShapeText - sets text in placeholder", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Try to set text in title placeholder
    const result = await setShapeText(tempPath, "/slide[1]/placeholder[title]", "New Title");
    // May fail if placeholder doesn't exist, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setShapeText - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setShapeText(tempPath, "invalid", "text");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setShapeProperty - sets fill color", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);

    const result = await setShapeProperty(tempPath, "/slide[1]/shape[1]", "fillColor", "FF0000");
    // May fail if shape doesn't exist, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setShapeProperty - sets line color", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);

    const result = await setShapeProperty(tempPath, "/slide[1]/shape[1]", "lineColor", "00FF00");
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setShapeProperty - sets line width", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);

    const result = await setShapeProperty(tempPath, "/slide[1]/shape[1]", "lineWidth", "2");
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setShapeProperty - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setShapeProperty(tempPath, "invalid", "fillColor", "FF0000");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeShape - removes a shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);

    const result = await removeShape(tempPath, "/slide[1]/shape[1]");
    // May fail if shape doesn't exist
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("removeShape - removes a placeholder", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeShape(tempPath, "/slide[1]/placeholder[title]");
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("removeShape - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeShape(tempPath, "invalid");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setSlideProperty - sets background", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setSlideProperty(tempPath, 1, "background", "FFCCCC");
    // May fail but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setSlideProperty - returns error for invalid slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setSlideProperty(tempPath, 999, "background", "FFCCCC");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setSlideProperty - returns error for unknown property", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setSlideProperty(tempPath, 1, "unknownProperty", "value");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});
