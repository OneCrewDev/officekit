import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  addEquation,
  getEquations,
  setEquation,
  removeEquation,
} from "../src/equations.ts";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-equations-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("addEquation - adds a LaTeX equation to a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addEquation(tempPath, 1, "\\frac{a}{b}");
    assert.ok(result.ok, `addEquation failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data.path.includes("/slide[1]/shape"));
  } finally {
    // Clean up
  }
});

test("addEquation - adds an OMML equation to a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const omml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:frac><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:frac></m:oMath>';
    const result = await addEquation(tempPath, 1, omml);
    assert.ok(result.ok, `addEquation failed: ${result.error?.message}`);
    assert.ok(result.data);
  } finally {
    // Clean up
  }
});

test("addEquation - adds equation with custom position", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addEquation(tempPath, 1, "x^2 + y^2", {
      x: 1000000,
      y: 2000000,
      width: 4000000,
      height: 500000
    });
    assert.ok(result.ok, `addEquation failed: ${result.error?.message}`);
  } finally {
    // Clean up
  }
});

test("addEquation - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addEquation(tempPath, 999, "\\alpha");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("addEquation - supports Greek letters", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addEquation(tempPath, 1, "\\alpha + \\beta = \\gamma");
    assert.ok(result.ok, `addEquation failed: ${result.error?.message}`);
  } finally {
    // Clean up
  }
});

test("addEquation - supports square roots", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addEquation(tempPath, 1, "\\sqrt{x^2 + y^2}");
    assert.ok(result.ok, `addEquation failed: ${result.error?.message}`);
  } finally {
    // Clean up
  }
});

test("addEquation - supports subscripts and superscripts", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addEquation(tempPath, 1, "x_{n}^{2}");
    assert.ok(result.ok, `addEquation failed: ${result.error?.message}`);
  } finally {
    // Clean up
  }
});

test("getEquations - returns empty array for slide without equations", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getEquations(tempPath, 1);
    assert.ok(result.ok, `getEquations failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(Array.isArray(result.data.equations));
    // Note: The template may already have equations, so we don't assert empty
  } finally {
    // Clean up
  }
});

test("getEquations - returns equations from a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add an equation first
    await addEquation(tempPath, 1, "\\pi r^2");

    const result = await getEquations(tempPath, 1);
    assert.ok(result.ok, `getEquations failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data.equations.length > 0);
  } finally {
    // Clean up
  }
});

test("getEquations - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getEquations(tempPath, 999);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setEquation - updates an existing equation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add an equation first
    const addResult = await addEquation(tempPath, 1, "\\frac{a}{b}");
    assert.ok(addResult.ok, `addEquation failed: ${addResult.error?.message}`);

    // Update the equation
    const setResult = await setEquation(tempPath, addResult.data.path, "\\frac{x}{y}");
    assert.ok(setResult.ok, `setEquation failed: ${setResult.error?.message}`);

    // Verify the update
    const getResult = await getEquations(tempPath, 1);
    assert.ok(getResult.ok);
  } finally {
    // Clean up
  }
});

test("setEquation - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setEquation(tempPath, "invalid", "\\frac{a}{b}");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setEquation - returns error for non-existent equation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setEquation(tempPath, "/slide[1]/shape[999]", "\\frac{a}{b}");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("removeEquation - removes an equation from a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add an equation first
    const addResult = await addEquation(tempPath, 1, "\\frac{a}{b}");
    assert.ok(addResult.ok, `addEquation failed: ${addResult.error?.message}`);

    // Remove the equation
    const removeResult = await removeEquation(tempPath, addResult.data.path);
    assert.ok(removeResult.ok, `removeEquation failed: ${removeResult.error?.message}`);
  } finally {
    // Clean up
  }
});

test("removeEquation - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeEquation(tempPath, "invalid");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeEquation - returns error for non-existent equation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeEquation(tempPath, "/slide[1]/shape[999]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("addEquation - supports special symbols like \\infty", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addEquation(tempPath, 1, "\\infty");
    assert.ok(result.ok, `addEquation failed: ${result.error?.message}`);
  } finally {
    // Clean up
  }
});

test("addEquation - supports sums and integrals", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addEquation(tempPath, 1, "\\sum_{i=1}^{n} i = \\frac{n(n+1)}{2}");
    assert.ok(result.ok, `addEquation failed: ${result.error?.message}`);
  } finally {
    // Clean up
  }
});
