import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import { merge } from "../src/merge.js";
import { setShapeText } from "../src/shapes.js";
import { getTextRuns } from "../src/text.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-merge-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

async function getShapeTextContent(filePath: string, shapePath: string): Promise<string> {
  const result = await getTextRuns(filePath, shapePath);
  if (!result.ok) {
    throw new Error(result.error?.message);
  }
  return result.data!.paragraphs?.map(p => p.text).join("") ?? result.data!.runs.map(r => r.text).join("");
}

test("merge - replaces simple {{key}} placeholders", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set shape text with a placeholder
    const setResult = await setShapeText(tempPath, "/slide[1]/shape[1]", "Hello, {{name}}!");
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge data into the template
    const mergeResult = await merge(tempPath, { name: "John" }, tempPath);
    assert.ok(mergeResult.ok, `merge failed: ${mergeResult.error?.message}`);
    assert.ok(mergeResult.data);
    assert.ok(mergeResult.data.replacements >= 1, "Should have at least 1 replacement");

    // Verify the placeholder was replaced
    const text = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(text.includes("Hello, John!"));
  } finally {
    // Clean up
  }
});

test("merge - replaces multiple placeholders", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set shape text with multiple placeholders
    const setResult = await setShapeText(
      tempPath,
      "/slide[1]/shape[1]",
      "{{greeting}}, {{name}}! Today is {{date}}."
    );
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge data into the template
    const mergeResult = await merge(
      tempPath,
      { greeting: "Hello", name: "Alice", date: "2024-01-15" },
      tempPath
    );
    assert.ok(mergeResult.ok, `merge failed: ${mergeResult.error?.message}`);
    assert.ok(mergeResult.data);
    assert.ok(mergeResult.data.replacements >= 3, "Should have at least 3 replacements");

    // Verify all placeholders were replaced
    const text = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(text.includes("Hello, Alice!"));
    assert.ok(text.includes("2024-01-15"));
  } finally {
    // Clean up
  }
});

test("merge - handles nested key access with dot notation", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set shape text with a nested placeholder
    const setResult = await setShapeText(
      tempPath,
      "/slide[1]/shape[1]",
      "User: {{user.name}}, Email: {{user.email}}"
    );
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge data into the template
    const mergeResult = await merge(
      tempPath,
      { user: { name: "Bob", email: "bob@example.com" } },
      tempPath
    );
    assert.ok(mergeResult.ok, `merge failed: ${mergeResult.error?.message}`);
    assert.ok(mergeResult.data);
    assert.ok(mergeResult.data.replacements >= 2, "Should have at least 2 replacements");

    // Verify the nested placeholders were replaced
    const text = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(text.includes("User: Bob"));
    assert.ok(text.includes("Email: bob@example.com"));
  } finally {
    // Clean up
  }
});

test("merge - processes conditional blocks {{#if}}...{{/if}}", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set shape text with conditional block
    const setResult = await setShapeText(
      tempPath,
      "/slide[1]/shape[1]",
      "Hello{{#if showName}}, {{name}}{{/if}}!"
    );
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge with showName = true - conditional content should be included
    const mergeResult1 = await merge(
      tempPath,
      { showName: true, name: "Charlie" },
      tempPath
    );
    assert.ok(mergeResult1.ok);
    assert.ok(mergeResult1.data);
    assert.ok(mergeResult1.data.conditionals >= 1);

    let text = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(text.includes("Hello, Charlie!"), `Expected "Hello, Charlie!", got "${text}"`);
  } finally {
    // Clean up
  }
});

test("merge - conditional block with false condition removes content", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set shape text with conditional block
    const setResult = await setShapeText(
      tempPath,
      "/slide[1]/shape[1]",
      "Hello{{#if showName}}, {{name}}{{/if}}!"
    );
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge with showName = false - conditional content should be removed
    const mergeResult = await merge(
      tempPath,
      { showName: false, name: "Charlie" },
      tempPath
    );
    assert.ok(mergeResult.ok);
    assert.ok(mergeResult.data);
    assert.ok(mergeResult.data.conditionals >= 1);

    const text = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(text === "Hello!", `Expected "Hello!", got "${text}"`);
  } finally {
    // Clean up
  }
});

test("merge - processes loop blocks {{#each}}...{{/each}}", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set shape text with loop block
    const setResult = await setShapeText(
      tempPath,
      "/slide[1]/shape[1]",
      "Items:{{#each items}} {{name}}{{/each}}"
    );
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge data with items array
    const mergeResult = await merge(
      tempPath,
      { items: [{ name: "Apple" }, { name: "Banana" }, { name: "Cherry" }] },
      tempPath
    );
    assert.ok(mergeResult.ok, `merge failed: ${mergeResult.error?.message}`);
    assert.ok(mergeResult.data);
    assert.ok(mergeResult.data.loops >= 1);

    // Verify the loop was expanded
    const text = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(text.includes("Apple"));
    assert.ok(text.includes("Banana"));
    assert.ok(text.includes("Cherry"));
  } finally {
    // Clean up
  }
});

test("merge - handles formatted date placeholders", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set shape text with date format
    const setResult = await setShapeText(
      tempPath,
      "/slide[1]/shape[1]",
      "Date: {{date:yyyy-mm-dd}}"
    );
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge with a date
    const mergeResult = await merge(
      tempPath,
      { date: "2024-12-25" },
      tempPath
    );
    assert.ok(mergeResult.ok, `merge failed: ${mergeResult.error?.message}`);
    assert.ok(mergeResult.data);
    assert.ok(mergeResult.data.replacements >= 1);

    // Verify the date was formatted
    const text = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(text.includes("2024-12-25"));
  } finally {
    // Clean up
  }
});

test("merge - handles empty values gracefully", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set shape text with placeholders
    const setResult = await setShapeText(
      tempPath,
      "/slide[1]/shape[1]",
      "Name: {{name}}, Title: {{title}}"
    );
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge with missing/empty values
    const mergeResult = await merge(
      tempPath,
      { name: "", title: undefined },
      tempPath
    );
    assert.ok(mergeResult.ok, `merge failed: ${mergeResult.error?.message}`);
    assert.ok(mergeResult.data);

    // Verify empty values result in empty strings
    const text = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(text.includes("Name: , Title:"));
  } finally {
    // Clean up
  }
});

test("merge - returns error for invalid template path", async () => {
  const result = await merge(
    "/nonexistent/path/template.pptx",
    { name: "John" },
    "/tmp/output.pptx"
  );
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "operation_failed");
});

test("merge - processes multiple slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set text on different slides
    await setShapeText(tempPath, "/slide[1]/shape[1]", "Slide 1: {{name}}");
    await setShapeText(tempPath, "/slide[2]/shape[1]", "Slide 2: {{name}}");

    // Merge data
    const mergeResult = await merge(tempPath, { name: "TestUser" }, tempPath);
    assert.ok(mergeResult.ok, `merge failed: ${mergeResult.error?.message}`);
    assert.ok(mergeResult.data);
    assert.ok(mergeResult.data.slidesProcessed >= 2);
  } finally {
    // Clean up
  }
});

test("merge - handles text without placeholders", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Set text without any placeholders
    const setResult = await setShapeText(
      tempPath,
      "/slide[1]/shape[1]",
      "This is plain text with no placeholders."
    );
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge should still succeed (no changes needed)
    const mergeResult = await merge(tempPath, { name: "John" }, tempPath);
    assert.ok(mergeResult.ok, `merge failed: ${mergeResult.error?.message}`);
    assert.ok(mergeResult.data);
    assert.equal(mergeResult.data.replacements, 0);

    // Verify text is unchanged
    const text = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(text.includes("plain text"));
  } finally {
    // Clean up
  }
});

test("merge - writes to specified output path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const outputPath = path.join(tmpdir(), `ppt-merge-output-${Date.now()}.pptx`);
  try {
    // Set text with placeholder
    const setResult = await setShapeText(
      tempPath,
      "/slide[1]/shape[1]",
      "Hello, {{name}}!"
    );
    if (!setResult.ok) {
      assert.fail(`setShapeText failed: ${setResult.error?.message}`);
    }

    // Merge to different output path
    const mergeResult = await merge(tempPath, { name: "OutputTest" }, outputPath);
    assert.ok(mergeResult.ok, `merge failed: ${mergeResult.error?.message}`);

    // Verify output file exists and has merged content
    const text = await getShapeTextContent(outputPath, "/slide[1]/shape[1]");
    assert.ok(text.includes("Hello, OutputTest!"));

    // Original file should be unchanged (still has placeholder)
    const originalText = await getShapeTextContent(tempPath, "/slide[1]/shape[1]");
    assert.ok(originalText.includes("{{name}}"));
  } finally {
    // Clean up
  }
});
