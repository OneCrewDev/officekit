import test from "node:test";
import assert from "node:assert/strict";

import {
  parsePath,
  buildPath,
  slidePath,
  shapePath,
  tablePath,
  placeholderPath,
  chartPath,
  cellPath,
  getSlideIndex,
  isPresentationPath,
  isSlidePath,
  isCellPath,
  isPlaceholderPath,
  parentPath,
  lastSegmentName,
  validatePath,
  isValidPath,
  getElementType,
  isElementType,
} from "../src/path.ts";

test("parsePath - parses root paths", () => {
  const result = parsePath("/");
  assert.ok(result.ok);
  assert.deepEqual(result.data, { isAbsolute: true, segments: [], original: "/" });

  const presResult = parsePath("/presentation");
  assert.ok(presResult.ok);
  assert.deepEqual(presResult.data, {
    isAbsolute: true,
    segments: [{ name: "presentation" }],
    original: "/presentation",
  });

  const themeResult = parsePath("/theme");
  assert.ok(themeResult.ok);
  assert.deepEqual(themeResult.data, {
    isAbsolute: true,
    segments: [{ name: "theme" }],
    original: "/theme",
  });
});

test("parsePath - parses slide paths", () => {
  const result = parsePath("/slide[1]");
  assert.ok(result.ok);
  assert.deepEqual(result.data, {
    isAbsolute: true,
    segments: [{ name: "slide", index: 1 }],
    original: "/slide[1]",
  });

  const result5 = parsePath("/slide[5]");
  assert.ok(result5.ok);
  assert.equal(result5.data?.segments[0].index, 5);
});

test("parsePath - parses nested paths", () => {
  const result = parsePath("/slide[1]/shape[1]");
  assert.ok(result.ok);
  assert.deepEqual(result.data, {
    isAbsolute: true,
    segments: [
      { name: "slide", index: 1 },
      { name: "shape", index: 1 },
    ],
    original: "/slide[1]/shape[1]",
  });

  const tableResult = parsePath("/slide[1]/table[1]");
  assert.ok(tableResult.ok);
  assert.deepEqual(tableResult.data, {
    isAbsolute: true,
    segments: [
      { name: "slide", index: 1 },
      { name: "table", index: 1 },
    ],
    original: "/slide[1]/table[1]",
  });
});

test("parsePath - parses placeholder paths with named selectors", () => {
  const result = parsePath("/slide[1]/placeholder[title]");
  assert.ok(result.ok);
  assert.deepEqual(result.data, {
    isAbsolute: true,
    segments: [
      { name: "slide", index: 1 },
      { name: "placeholder", nameSelector: "title" },
    ],
    original: "/slide[1]/placeholder[title]",
  });

  const bodyResult = parsePath("/slide[1]/placeholder[body]");
  assert.ok(bodyResult.ok);
  assert.equal(bodyResult.data?.segments[1].nameSelector, "body");
});

test("parsePath - parses chart paths", () => {
  const result = parsePath("/slide[1]/chart[1]");
  assert.ok(result.ok);
  assert.deepEqual(result.data, {
    isAbsolute: true,
    segments: [
      { name: "slide", index: 1 },
      { name: "chart", index: 1 },
    ],
    original: "/slide[1]/chart[1]",
  });
});

test("parsePath - parses table cell paths", () => {
  const result = parsePath("/slide[1]/table[1]/tr[2]/tc[3]");
  assert.ok(result.ok);
  assert.deepEqual(result.data, {
    isAbsolute: true,
    segments: [
      { name: "slide", index: 1 },
      { name: "table", index: 1 },
      { name: "tr", index: 2 },
      { name: "tc", index: 3 },
    ],
    original: "/slide[1]/table[1]/tr[2]/tc[3]",
  });
});

test("parsePath - parses slidemaster and slidelayout paths", () => {
  const masterResult = parsePath("/slidemaster[1]");
  assert.ok(masterResult.ok);
  assert.deepEqual(masterResult.data, {
    isAbsolute: true,
    segments: [{ name: "slidemaster", index: 1 }],
    original: "/slidemaster[1]",
  });

  const layoutResult = parsePath("/slidelayout[1]");
  assert.ok(layoutResult.ok);
  assert.deepEqual(layoutResult.data, {
    isAbsolute: true,
    segments: [{ name: "slidelayout", index: 1 }],
    original: "/slidelayout[1]",
  });
});

test("parsePath - rejects invalid paths", () => {
  // Empty path
  const emptyResult = parsePath("");
  assert.ok(!emptyResult.ok);

  // Non-absolute path
  const relativeResult = parsePath("slide[1]");
  assert.ok(!relativeResult.ok);
  assert.equal(relativeResult.error?.code, "invalid_path");

  // Invalid root segment
  const invalidRootResult = parsePath("/invalid[1]");
  assert.ok(!invalidRootResult.ok);

  // Invalid child segment
  const invalidChildResult = parsePath("/slide[1]/invalid[1]");
  assert.ok(!invalidChildResult.ok);

  // Zero or negative index
  const zeroIndexResult = parsePath("/slide[0]");
  assert.ok(!zeroIndexResult.ok);

  const negativeIndexResult = parsePath("/slide[-1]");
  assert.ok(!negativeIndexResult.ok);
});

test("parsePath - rejects invalid segment sequences", () => {
  // Cannot have shape after table (must go through tr/tc)
  const invalidSequenceResult = parsePath("/slide[1]/table[1]/shape[1]");
  assert.ok(!invalidSequenceResult.ok);
});

test("buildPath - builds paths from segments", () => {
  const path = buildPath([{ name: "slide", index: 1 }]);
  assert.equal(path, "/slide[1]");

  const nestedPath = buildPath([
    { name: "slide", index: 1 },
    { name: "shape", index: 2 },
  ]);
  assert.equal(nestedPath, "/slide[1]/shape[2]");

  const cellPathBuilt = buildPath([
    { name: "slide", index: 1 },
    { name: "table", index: 1 },
    { name: "tr", index: 2 },
    { name: "tc", index: 3 },
  ]);
  assert.equal(cellPathBuilt, "/slide[1]/table[1]/tr[2]/tc[3]");
});

test("buildPath - builds named selectors", () => {
  const path = buildPath([
    { name: "slide", index: 1 },
    { name: "placeholder", nameSelector: "title" },
  ]);
  assert.equal(path, "/slide[1]/placeholder[title]");
});

test("buildPath - builds empty path for no segments", () => {
  const path = buildPath([]);
  assert.equal(path, "/");
});

test("slidePath - creates slide path", () => {
  assert.equal(slidePath(1), "/slide[1]");
  assert.equal(slidePath(5), "/slide[5]");
});

test("shapePath - creates shape path", () => {
  assert.equal(shapePath(1, 1), "/slide[1]/shape[1]");
  assert.equal(shapePath(3, 5), "/slide[3]/shape[5]");
});

test("tablePath - creates table path", () => {
  assert.equal(tablePath(1, 1), "/slide[1]/table[1]");
});

test("placeholderPath - creates placeholder path", () => {
  assert.equal(placeholderPath(1, "title"), "/slide[1]/placeholder[title]");
  assert.equal(placeholderPath(1, "body"), "/slide[1]/placeholder[body]");
});

test("chartPath - creates chart path", () => {
  assert.equal(chartPath(1, 1), "/slide[1]/chart[1]");
});

test("cellPath - creates cell path", () => {
  assert.equal(cellPath(1, 1, 2, 3), "/slide[1]/table[1]/tr[2]/tc[3]");
});

test("getSlideIndex - extracts slide index", () => {
  assert.equal(getSlideIndex("/slide[5]"), 5);
  assert.equal(getSlideIndex("/slide[1]/shape[2]"), 1);
  assert.equal(getSlideIndex("/slide[1]/table[1]/tr[2]/tc[3]"), 1);
  assert.equal(getSlideIndex("/presentation"), null);
});

test("isPresentationPath - checks for presentation root", () => {
  assert.ok(isPresentationPath("/"));
  assert.ok(isPresentationPath("/presentation"));
  assert.ok(!isPresentationPath("/slide[1]"));
  assert.ok(!isPresentationPath("/theme"));
});

test("isSlidePath - checks for slide paths", () => {
  assert.ok(isSlidePath("/slide[1]"));
  assert.ok(isSlidePath("/slide[1]/shape[1]"));
  assert.ok(isSlidePath("/slide[1]/table[1]/tr[1]/tc[1]"));
  assert.ok(!isSlidePath("/presentation"));
  assert.ok(!isSlidePath("/theme"));
});

test("isCellPath - checks for cell paths", () => {
  assert.ok(isCellPath("/slide[1]/table[1]/tr[1]/tc[1]"));
  assert.ok(!isCellPath("/slide[1]/table[1]"));
  assert.ok(!isCellPath("/slide[1]/shape[1]"));
});

test("isPlaceholderPath - checks for placeholder paths", () => {
  assert.ok(isPlaceholderPath("/slide[1]/placeholder[title]"));
  assert.ok(isPlaceholderPath("/slide[1]/placeholder[body]"));
  assert.ok(!isPlaceholderPath("/slide[1]/shape[1]"));
});

test("parentPath - gets parent path", () => {
  assert.equal(parentPath("/slide[1]/shape[2]"), "/slide[1]");
  assert.equal(parentPath("/slide[1]"), "/");
  assert.equal(parentPath("/slide[1]/table[1]/tr[2]"), "/slide[1]/table[1]");
});

test("lastSegmentName - gets last segment name", () => {
  assert.equal(lastSegmentName("/slide[1]/shape[2]"), "shape");
  assert.equal(lastSegmentName("/slide[1]/table[1]/tr[2]/tc[3]"), "tc");
  assert.equal(lastSegmentName("/slide[1]"), "slide");
});

test("validatePath - validates paths", () => {
  assert.ok(validatePath("/slide[1]").ok);
  assert.ok(validatePath("/slide[1]/shape[1]").ok);
  assert.ok(validatePath("/slide[1]/placeholder[title]").ok);
  assert.ok(!validatePath("invalid").ok);
  assert.ok(!validatePath("/slide[0]").ok);
});

test("isValidPath - checks path validity", () => {
  assert.ok(isValidPath("/slide[1]"));
  assert.ok(isValidPath("/slide[1]/shape[1]"));
  assert.ok(isValidPath("/slide[1]/table[1]/tr[1]/tc[1]"));
  assert.ok(!isValidPath(""));
  assert.ok(!isValidPath("slide[1]"));
  assert.ok(!isValidPath("/slide[0]"));
});

test("getElementType - gets element type", () => {
  assert.equal(getElementType("/slide[1]/shape[2]"), "shape");
  assert.equal(getElementType("/slide[1]/table[1]/tr[2]/tc[3]"), "tc");
  assert.equal(getElementType("/presentation"), "presentation");
  assert.equal(getElementType("/theme"), "theme");
  assert.equal(getElementType("/"), "presentation");
});

test("isElementType - checks element type", () => {
  assert.ok(isElementType("/slide[1]/shape[2]", "shape"));
  assert.ok(isElementType("/slide[1]/table[1]", "table"));
  assert.ok(!isElementType("/slide[1]/shape[2]", "table"));
});
