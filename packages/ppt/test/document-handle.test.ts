import test from "node:test";
import assert from "node:assert/strict";
import { copyFile, unlink } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  open,
  close,
  isOpen,
  getInfo,
  getZip,
  getFilePath,
  isDirty,
  setDirty,
} from "../src/document-handle.js";
import { getOpenCount, clearRegistry } from "../src/registry.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-handle-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("open - returns a document handle", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await open(tempPath);
    assert.ok(result.ok, `open failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data.handle);
    assert.ok(typeof result.data.handle === "string");
    assert.ok(result.data.handle.startsWith("doc-"));
    assert.equal(result.data.filePath, tempPath);

    // Clean up
    await close(result.data.handle);
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("open - returns error for non-existent file", async () => {
  const result = await open("/non/existent/path.pptx");
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "not_found");
});

test("open - returns error for non-pptx file", async () => {
  const result = await open("/some/path.txt");
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "invalid_input");
});

test("close - closes an open document without saving", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const openResult = await open(tempPath);
    assert.ok(openResult.ok);

    const handle = openResult.data!.handle;
    assert.ok(isOpen(handle));

    const closeResult = await close(handle);
    assert.ok(closeResult.ok);
    assert.ok(!isOpen(handle));
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("close - closes and saves an open document", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const openResult = await open(tempPath);
    assert.ok(openResult.ok);

    const handle = openResult.data!.handle;
    assert.ok(isOpen(handle));

    // Mark as dirty
    setDirty(handle);
    assert.ok(isDirty(handle));

    // Close with save
    const closeResult = await close(handle, true);
    assert.ok(closeResult.ok);
    assert.ok(!isOpen(handle));
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("close - returns error for non-existent handle", async () => {
  const result = await close("non-existent-handle");
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "not_found");
});

test("isOpen - returns true for open document", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await open(tempPath);
    assert.ok(result.ok);
    assert.ok(isOpen(result.data!.handle));
    await close(result.data!.handle);
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("isOpen - returns false for closed document", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await open(tempPath);
    assert.ok(result.ok);
    const handle = result.data!.handle;
    await close(handle);
    assert.ok(!isOpen(handle));
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("isOpen - returns false for non-existent handle", async () => {
  assert.ok(!isOpen("non-existent-handle"));
});

test("getInfo - returns document info", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const openResult = await open(tempPath);
    assert.ok(openResult.ok);

    const handle = openResult.data!.handle;
    const infoResult = getInfo(handle);

    assert.ok(infoResult.ok);
    assert.equal(infoResult.data!.handle, handle);
    assert.equal(infoResult.data!.filePath, tempPath);
    assert.equal(infoResult.data!.dirty, false);
    assert.ok(infoResult.data!.openedAt instanceof Date);

    await close(handle);
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("getInfo - returns error for non-existent handle", async () => {
  const result = getInfo("non-existent-handle");
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "not_found");
});

test("getZip - returns the zip contents", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const openResult = await open(tempPath);
    assert.ok(openResult.ok);

    const handle = openResult.data!.handle;
    const zipResult = getZip(handle);

    assert.ok(zipResult.ok);
    assert.ok(zipResult.data! instanceof Map);
    assert.ok(zipResult.data!.size > 0);
    // ppt/presentation.xml should be present
    assert.ok(zipResult.data!.has("ppt/presentation.xml"));

    await close(handle);
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("getZip - returns error for non-existent handle", async () => {
  const result = getZip("non-existent-handle");
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "not_found");
});

test("getFilePath - returns the file path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const openResult = await open(tempPath);
    assert.ok(openResult.ok);

    const handle = openResult.data!.handle;
    const pathResult = getFilePath(handle);

    assert.ok(pathResult.ok);
    assert.equal(pathResult.data!, tempPath);

    await close(handle);
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("isDirty - returns false initially", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const openResult = await open(tempPath);
    assert.ok(openResult.ok);

    const handle = openResult.data!.handle;
    assert.ok(!isDirty(handle));

    await close(handle);
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("setDirty - marks document as dirty", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const openResult = await open(tempPath);
    assert.ok(openResult.ok);

    const handle = openResult.data!.handle;
    assert.ok(!isDirty(handle));

    setDirty(handle);
    assert.ok(isDirty(handle));

    await close(handle);
  } finally {
    try {
      await unlink(tempPath);
    } catch { /* ignore */ }
  }
});

test("multiple documents can be open simultaneously", async () => {
  const tempPath1 = await copyToTemp(TEST_PPTX);
  const tempPath2 = await copyToTemp(TEST_PPTX);
  try {
    clearRegistry();

    const result1 = await open(tempPath1);
    const result2 = await open(tempPath2);

    assert.ok(result1.ok);
    assert.ok(result2.ok);
    assert.notEqual(result1.data!.handle, result2.data!.handle);

    assert.equal(getOpenCount(), 2);
    assert.ok(isOpen(result1.data!.handle));
    assert.ok(isOpen(result2.data!.handle));

    await close(result1.data!.handle);
    assert.ok(!isOpen(result1.data!.handle));
    assert.ok(isOpen(result2.data!.handle));
    assert.equal(getOpenCount(), 1);

    await close(result2.data!.handle);
    assert.equal(getOpenCount(), 0);
  } finally {
    try {
      await unlink(tempPath1);
    } catch { /* ignore */ }
    try {
      await unlink(tempPath2);
    } catch { /* ignore */ }
    clearRegistry();
  }
});

test("closing one document does not affect others", async () => {
  const tempPath1 = await copyToTemp(TEST_PPTX);
  const tempPath2 = await copyToTemp(TEST_PPTX);
  try {
    clearRegistry();

    const result1 = await open(tempPath1);
    const result2 = await open(tempPath2);

    assert.ok(result1.ok);
    assert.ok(result2.ok);

    // Close first document
    await close(result1.data!.handle);

    // Second document should still be open
    assert.ok(!isOpen(result1.data!.handle));
    assert.ok(isOpen(result2.data!.handle));

    // Clean up
    await close(result2.data!.handle);
  } finally {
    try {
      await unlink(tempPath1);
    } catch { /* ignore */ }
    try {
      await unlink(tempPath2);
    } catch { /* ignore */ }
    clearRegistry();
  }
});
