import test from "node:test";
import assert from "node:assert/strict";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

import {
  getWordNode,
  queryWordNodes,
  getDocumentInfo,
  viewWordDocument,
  parsePath,
  buildPath,
  parseSelector,
  validatePath,
  validateSelector
} from "../src/index.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const TEST_DOC = join(__dirname, "../../../../OfficeCLI/assets/showcase/restaurant-menu.docx");
const OUTPUT_DIR = join(__dirname, "output");

// Ensure output directory exists
async function ensureOutputDir() {
  try {
    await mkdir(OUTPUT_DIR, { recursive: true });
  } catch {}
}

// Copy test file to temp location for write tests
async function copyTestFile() {
  const dest = join(OUTPUT_DIR, "test-copy.docx");
  const content = await readFile(TEST_DOC);
  await writeFile(dest, content);
  return dest;
}

await ensureOutputDir();

test("parsePath parses basic paths", () => {
  const result = parsePath("/body/p[1]");
  assert.equal(result.ok, true);
  if (result.data?.segments) {
    assert.equal(result.data.segments[0].name, "body");
    assert.equal(result.data.segments[1].name, "p");
    assert.equal(result.data.segments[1].index, 1);
  }
});

test("parsePath parses stable ID paths", () => {
  const result = parsePath("/body/p[@paraId=1A2B3C4D]");
  assert.equal(result.ok, true);
  if (result.data?.segments) {
    assert.equal(result.data.segments[1].stringIndex, "@paraId=1A2B3C4D");
  }
});

test("buildPath reconstructs paths", () => {
  const result = buildPath([
    { name: "body" },
    { name: "p", index: 1 }
  ]);
  assert.equal(result, "/body/p[1]");
});

test("parseSelector parses element selectors", () => {
  const result = parseSelector("p");
  assert.equal(result.ok, true);
  if (result.data) {
    assert.equal(result.data.element, "p");
  }
});

test("parseSelector parses indexed selectors", () => {
  const result = parseSelector("p[1]");
  assert.equal(result.ok, true);
  if (result.data) {
    assert.equal(result.data.element, "p");
    assert.equal(result.data.attributes?.index, "1");
  }
});

test("parseSelector parses :contains selector", () => {
  const result = parseSelector('p:contains("Hello")');
  assert.equal(result.ok, true);
  if (result.data) {
    assert.equal(result.data.containsText, "Hello");
  }
});

test("validatePath validates correct paths", () => {
  assert.equal(validatePath("/body/p[1]").ok, true);
  assert.equal(validatePath("/body").ok, true);
  assert.equal(validatePath("/header[1]").ok, true);
});

test("validatePath rejects invalid paths", () => {
  assert.equal(validatePath("body/p[1]").ok, false);  // relative path
  assert.equal(validatePath("").ok, false);
});

test("validateSelector validates correct selectors", () => {
  assert.equal(validateSelector("p").ok, true);
  assert.equal(validateSelector("p[1]").ok, true);
});

test("getWordNode returns document root", async () => {
  const result = await getWordNode(TEST_DOC, "/", 0);
  assert.equal(result.ok, true);
  if (result.data) {
    assert.equal(result.data.path, "/");
  }
});

test("getWordNode returns body info", async () => {
  const result = await getWordNode(TEST_DOC, "/body", 0);
  assert.equal(result.ok, true);
  if (result.data) {
    assert.equal(result.data.path, "/body");
  }
});

test("queryWordNodes finds paragraphs", async () => {
  const result = await queryWordNodes(TEST_DOC, "p");
  assert.equal(result.ok, true);
  assert.ok(result.data && result.data.length > 0);
  assert.equal(result.data?.[0].type, "paragraph");
});

test("queryWordNodes finds by text content", async () => {
  const result = await queryWordNodes(TEST_DOC, 'p:contains("Menu")');
  assert.equal(result.ok, true);
});

test("getDocumentInfo returns document metadata", async () => {
  const result = await getDocumentInfo(TEST_DOC);
  assert.equal(result.ok, true);
});

test("viewWordDocument text mode", async () => {
  const result = await viewWordDocument(TEST_DOC, "text", { maxLines: 5 });
  assert.equal(result.mode, "text");
  assert.ok(typeof result.output === "string");
  assert.ok(result.output.length > 0);
});

test("viewWordDocument annotated mode", async () => {
  const result = await viewWordDocument(TEST_DOC, "annotated", { maxLines: 3 });
  assert.equal(result.mode, "annotated");
  assert.ok(typeof result.output === "string");
});

test("viewWordDocument outline mode", async () => {
  const result = await viewWordDocument(TEST_DOC, "outline");
  assert.equal(result.mode, "outline");
  assert.ok(typeof result.output === "string");
});

test("viewWordDocument stats mode", async () => {
  const result = await viewWordDocument(TEST_DOC, "stats");
  assert.equal(result.mode, "stats");
  assert.ok(typeof result.output === "string");
});

test("viewWordDocument issues mode", async () => {
  const result = await viewWordDocument(TEST_DOC, "issues");
  assert.equal(result.mode, "issues");
  assert.ok(typeof result.output === "string");
});

test("viewWordDocument html mode", async () => {
  const result = await viewWordDocument(TEST_DOC, "html");
  assert.equal(result.mode, "html");
  assert.ok(result.output.includes("<"));
});

test("viewWordDocument json mode", async () => {
  const result = await viewWordDocument(TEST_DOC, "json");
  assert.equal(result.mode, "json");
  assert.ok(result.output.includes("{"));
});

test("viewWordDocument forms mode", async () => {
  const result = await viewWordDocument(TEST_DOC, "forms");
  assert.equal(result.mode, "forms");
  assert.ok(typeof result.output === "string");
});

test("viewWordDocument rejects invalid mode", async () => {
  await assert.rejects(
    async () => viewWordDocument(TEST_DOC, "invalid"),
    /Unsupported view mode/
  );
});

test("viewWordDocument with startLine/endLine", async () => {
  const result = await viewWordDocument(TEST_DOC, "text", {
    startLine: 1,
    endLine: 5,
    maxLines: 10
  });
  assert.ok(typeof result.output === "string");
});
