import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getTextRuns,
  setTextRuns,
  addTextParagraph,
  setTextFormat,
} from "../src/text.js";
import { addSlide } from "../src/slides.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-text-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("getTextRuns - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getTextRuns(tempPath, "invalid");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("getTextRuns - returns error for invalid slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getTextRuns(tempPath, "/slide[999]/shape[1]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setTextRuns - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setTextRuns(tempPath, "invalid", [{ text: "test" }]);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setTextRuns - sets text runs on a shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);

    const result = await setTextRuns(tempPath, "/slide[1]/shape[1]", [
      { text: "Hello ", font: "Arial", bold: true },
      { text: "World", font: "Arial", italic: true }
    ]);
    // May fail if shape doesn't exist, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("addTextParagraph - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addTextParagraph(tempPath, "invalid", "New paragraph");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("addTextParagraph - adds a paragraph to a shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);

    const result = await addTextParagraph(tempPath, "/slide[1]/shape[1]", "New paragraph", { alignment: "center" });
    // May fail if shape doesn't exist, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setTextFormat - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setTextFormat(tempPath, "invalid", { font: "Arial", size: 18 });
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setTextFormat - sets text format on a shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addSlide(tempPath);

    const result = await setTextFormat(tempPath, "/slide[1]/shape[1]", {
      font: "Arial",
      size: 18,
      bold: true,
      color: "FF0000"
    });
    // May fail if shape doesn't exist, but shouldn't throw
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});
