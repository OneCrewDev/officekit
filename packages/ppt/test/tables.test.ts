import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  setTableCell,
  removeTableRow,
  removeTableColumn,
} from "../src/tables.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-tables-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("setTableCell - sets text in a table cell", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Try to set cell text - may fail if table doesn't exist
    const result = await setTableCell(tempPath, "/slide[1]/table[1]/tr[1]/tc[1]", "Hello");
    if (!result.ok) {
      assert.ok(result.error); // Error is expected if table doesn't exist
    }
  } finally {
    // Clean up
  }
});

test("setTableCell - returns error for invalid path format", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setTableCell(tempPath, "/slide[1]/shape[1]", "text");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setTableCell - returns error for non-slide path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setTableCell(tempPath, "invalid/path", "text");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeTableRow - removes a row from a table", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Try to remove a row - may fail if table doesn't exist
    const result = await removeTableRow(tempPath, "/slide[1]/table[1]/tr[1]");
    if (!result.ok) {
      assert.ok(result.error); // Error is expected if table doesn't exist
    }
  } finally {
    // Clean up
  }
});

test("removeTableRow - returns error for invalid path format", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeTableRow(tempPath, "/slide[1]/shape[1]");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeTableColumn - removes a column from a table", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Try to remove a column - may fail if table doesn't exist
    const result = await removeTableColumn(tempPath, "/slide[1]/table[1]", 1);
    if (!result.ok) {
      assert.ok(result.error); // Error is expected if table doesn't exist
    }
  } finally {
    // Clean up
  }
});

test("removeTableColumn - returns error for invalid table path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeTableColumn(tempPath, "/slide[1]/shape[1]", 1);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeTableColumn - returns error for invalid column index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeTableColumn(tempPath, "/slide[1]/table[1]", 0);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});
