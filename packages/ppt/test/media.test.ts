import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getMedia,
  addPicture,
  removeMedia,
  replacePicture,
  getMediaData,
} from "../src/media.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-media-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

// Create a simple PNG image buffer (1x1 pixel)
function createTestImage(): Buffer {
  // Minimal valid PNG (1x1 transparent pixel)
  const pngData = Buffer.from([
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
    0x08, 0xD7, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x05, 0x00, 0x01,
    0x87, 0xA1, 0x4E, 0xD4,
    0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND chunk
    0xAE, 0x42, 0x60, 0x82
  ]);
  return pngData;
}

test("getMedia - returns empty array for slide without media", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getMedia(tempPath, 1);
    assert.ok(result.ok, `getMedia failed: ${result.error?.message}`);
    assert.ok(Array.isArray(result.data?.media));
  } finally {
    // Clean up
  }
});

test("getMedia - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getMedia(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("addPicture - adds a picture to a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const imageData = createTestImage();
    const result = await addPicture(
      tempPath,
      1,
      { data: imageData, contentType: "image/png" },
      { x: 1000000, y: 1000000, width: 2000000, height: 2000000 }
    );
    assert.ok(result.ok, `addPicture failed: ${result.error?.message}`);
    assert.ok(result.data?.path);
    assert.ok(result.data?.path.includes("/slide[1]/picture["));
  } finally {
    // Clean up
  }
});

test("addPicture - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const imageData = createTestImage();
    const result = await addPicture(tempPath, 999, { data: imageData, contentType: "image/png" });
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeMedia - removes a picture from a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add a picture
    const imageData = createTestImage();
    const addResult = await addPicture(tempPath, 1, { data: imageData, contentType: "image/png" });
    if (!addResult.ok) {
      assert.fail(`addPicture failed: ${addResult.error?.message}`);
    }

    const picPath = addResult.data!.path;

    // Now remove it
    const removeResult = await removeMedia(tempPath, picPath);
    assert.ok(removeResult.ok, `removeMedia failed: ${removeResult.error?.message}`);
  } finally {
    // Clean up
  }
});

test("removeMedia - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeMedia(tempPath, "invalid");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("replacePicture - replaces an existing picture", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add a picture
    const imageData = createTestImage();
    const addResult = await addPicture(tempPath, 1, { data: imageData, contentType: "image/png" });
    if (!addResult.ok) {
      assert.fail(`addPicture failed: ${addResult.error?.message}`);
    }

    const picPath = addResult.data!.path;

    // Now replace it
    const newImageData = createTestImage();
    const replaceResult = await replacePicture(tempPath, picPath, { data: newImageData, contentType: "image/png" });
    assert.ok(replaceResult.ok, `replacePicture failed: ${replaceResult.error?.message}`);
  } finally {
    // Clean up
  }
});

test("getMediaData - returns data for a picture", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add a picture
    const imageData = createTestImage();
    const addResult = await addPicture(tempPath, 1, { data: imageData, contentType: "image/png" });
    if (!addResult.ok) {
      assert.fail(`addPicture failed: ${addResult.error?.message}`);
    }

    const picPath = addResult.data!.path;

    // Now get its data
    const dataResult = await getMediaData(tempPath, picPath);
    assert.ok(dataResult.ok, `getMediaData failed: ${dataResult.error?.message}`);
    assert.ok(dataResult.data?.data);
    assert.ok(dataResult.data?.contentType);
  } finally {
    // Clean up
  }
});

test("getMediaData - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getMediaData(tempPath, "invalid");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});
