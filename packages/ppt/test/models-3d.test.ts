import test from "node:test";
import assert from "node:assert/strict";
import { copyFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  get3DModels,
  add3DModel,
  remove3DModel,
  set3DModelRotation,
} from "../src/models-3d.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

// Create a minimal valid GLB file (GL Binary format)
// This is a simplified GLB header for testing purposes
function createTestGlb(): Buffer {
  // GLB header structure (12 bytes)
  // Magic number: 0x46546C67 (glTF)
  // Version: 2 (little-endian)
  // Length: total file length (little-endian)
  const header = Buffer.alloc(12);
  header.writeUInt32LE(0x46546C67, 0); // magic
  header.writeUInt32LE(2, 4); // version
  const totalLength = 12 + 12 + 8 + 8; // header + minimal JSON + minimal BIN + minimal CHUNK
  header.writeUInt32LE(totalLength, 8); // length

  // Minimal JSON chunk (12 bytes header + JSON content)
  const jsonContent = JSON.stringify({
    asset: { version: "2.0", generator: "Test" },
    scene: 0,
    scenes: [{ nodes: [0] }],
    nodes: [{ mesh: 0 }],
    meshes: [{ primitives: [{}] }],
  });
  const jsonChunk = Buffer.alloc(8 + jsonContent.length);
  jsonChunk.writeUInt32LE(jsonContent.length, 0); // chunk length
  jsonChunk.writeUInt32LE(0x4E4F534A, 4); // chunk type (JSON)
  jsonChunk.write(jsonContent, 8);

  // Minimal BIN chunk (empty)
  const binChunk = Buffer.alloc(8); // 8 bytes header, no data
  binChunk.writeUInt32LE(0, 0); // chunk length
  binChunk.writeUInt32LE(0x004E4942, 4); // chunk type (BIN)

  return Buffer.concat([header, jsonChunk, binChunk]);
}

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-3dmodel-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("get3DModels - returns empty array for slide without 3D models", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await get3DModels(tempPath, 1);
    assert.ok(result.ok, `get3DModels failed: ${result.error?.message}`);
    assert.ok(Array.isArray(result.data?.models));
    assert.equal(result.data?.models.length, 0);
  } finally {
    // Clean up
  }
});

test("get3DModels - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await get3DModels(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("add3DModel - adds a 3D model to a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Create a temporary GLB file
    const glbPath = path.join(tmpdir(), `test-model-${Date.now()}.glb`);
    const glbData = createTestGlb();
    await writeFile(glbPath, glbData);

    const result = await add3DModel(
      tempPath,
      1,
      glbPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 3000000 },
      { x: 0, y: 0, z: 45 }
    );
    assert.ok(result.ok, `add3DModel failed: ${result.error?.message}`);
    assert.ok(result.data?.path);
    assert.ok(result.data?.path.includes("/slide[1]/model3d["));

    // Verify the model was added
    const getResult = await get3DModels(tempPath, 1);
    assert.ok(getResult.ok, `get3DModels failed: ${getResult.error?.message}`);
    assert.ok(getResult.data?.models.length === 1);

    // Clean up temp GLB
    await require("node:fs/promises").unlink(glbPath);
  } finally {
    // Clean up
  }
});

test("add3DModel - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const glbPath = path.join(tmpdir(), `test-model-${Date.now()}.glb`);
    const glbData = createTestGlb();
    await writeFile(glbPath, glbData);

    const result = await add3DModel(
      tempPath,
      999,
      glbPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 3000000 }
    );
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");

    // Clean up temp GLB
    await require("node:fs/promises").unlink(glbPath);
  } finally {
    // Clean up
  }
});

test("add3DModel - returns error for non-glb file", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Create a temporary non-GLB file
    const wrongPath = path.join(tmpdir(), `test-model-${Date.now()}.obj`);
    await writeFile(wrongPath, "not a glb file");

    const result = await add3DModel(
      tempPath,
      1,
      wrongPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 3000000 }
    );
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");

    // Clean up temp file
    await require("node:fs/promises").unlink(wrongPath);
  } finally {
    // Clean up
  }
});

test("remove3DModel - removes a 3D model from a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Create and add a temporary GLB file
    const glbPath = path.join(tmpdir(), `test-model-${Date.now()}.glb`);
    const glbData = createTestGlb();
    await writeFile(glbPath, glbData);

    const addResult = await add3DModel(
      tempPath,
      1,
      glbPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 3000000 }
    );
    if (!addResult.ok) {
      assert.fail(`add3DModel failed: ${addResult.error?.message}`);
    }

    const modelPath = addResult.data!.path;

    // Now remove it
    const removeResult = await remove3DModel(tempPath, modelPath);
    assert.ok(removeResult.ok, `remove3DModel failed: ${removeResult.error?.message}`);

    // Verify it's gone
    const getResult = await get3DModels(tempPath, 1);
    assert.ok(getResult.ok);
    assert.equal(getResult.data?.models.length, 0);

    // Clean up temp GLB
    await require("node:fs/promises").unlink(glbPath);
  } finally {
    // Clean up
  }
});

test("remove3DModel - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await remove3DModel(tempPath, "invalid");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("remove3DModel - returns error for path without model3d index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await remove3DModel(tempPath, "/slide[1]/shape[1]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("set3DModelRotation - updates rotation of a 3D model", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Create and add a temporary GLB file
    const glbPath = path.join(tmpdir(), `test-model-${Date.now()}.glb`);
    const glbData = createTestGlb();
    await writeFile(glbPath, glbData);

    const addResult = await add3DModel(
      tempPath,
      1,
      glbPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 3000000 }
    );
    if (!addResult.ok) {
      assert.fail(`add3DModel failed: ${addResult.error?.message}`);
    }

    const modelPath = addResult.data!.path;

    // Update rotation
    const rotResult = await set3DModelRotation(
      tempPath,
      modelPath,
      { x: 45, y: 30, z: 60 }
    );
    assert.ok(rotResult.ok, `set3DModelRotation failed: ${rotResult.error?.message}`);

    // Clean up temp GLB
    await require("node:fs/promises").unlink(glbPath);
  } finally {
    // Clean up
  }
});

test("set3DModelRotation - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await set3DModelRotation(
      tempPath,
      "invalid",
      { x: 45, y: 30, z: 60 }
    );
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("set3DModelRotation - returns error for path without model3d index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await set3DModelRotation(
      tempPath,
      "/slide[1]/shape[1]",
      { x: 45, y: 30, z: 60 }
    );
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});
