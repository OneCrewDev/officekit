import test from "node:test";
import assert from "node:assert/strict";

import {
  parseSelector,
  parseSlideSelector,
  buildSelector,
  validateSelector,
  isValidSelector,
  slideSelector,
  typeSelector,
  textSelector,
  hasTextFilter,
  hasAttributeFilter,
  isElementType,
  isPlaceholderSelector,
  isCellSelector,
  isRowSelector,
  normalizePlaceholderTypeName,
} from "../src/selectors.ts";

test("parseSelector - parses basic selectors", () => {
  const result = parseSelector("slide[1]");
  assert.ok(result.ok);
  assert.equal(result.data?.slideNum, 1);
});

test("parseSelector - parses element type selectors", () => {
  const result = parseSelector("shape");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "shape");
});

test("parseSelector - parses indexed selectors", () => {
  const result = parseSelector("shape[2]");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "shape");
  assert.equal(result.data?.attributes.index, "2");
});

test("parseSelector - parses placeholder selectors", () => {
  const result = parseSelector("placeholder[title]");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "placeholder");
  assert.equal(result.data?.attributes.name, "title");
});

test("parseSelector - parses placeholder body selector", () => {
  const result = parseSelector("placeholder[body]");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "placeholder");
  assert.equal(result.data?.attributes.name, "body");
});

test("parseSelector - parses :contains selector", () => {
  const result = parseSelector('shape:contains("Hello")');
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "shape");
  assert.equal(result.data?.textContains, "Hello");
});

test("parseSelector - parses :contains with single quotes", () => {
  const result = parseSelector("shape:contains('World')");
  assert.ok(result.ok);
  assert.equal(result.data?.textContains, "World");
});

test("parseSelector - parses :contains without quotes", () => {
  const result = parseSelector("shape:contains(text)");
  assert.ok(result.ok);
  assert.equal(result.data?.textContains, "text");
});

test("parseSelector - parses :empty selector", () => {
  const result = parseSelector("shape:empty");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "shape");
  assert.equal(result.data?.attributes.empty, "true");
});

test("parseSelector - parses :no-alt selector", () => {
  const result = parseSelector("shape:no-alt");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "shape");
  assert.equal(result.data?.attributes.noAlt, "true");
});

test("parseSelector - parses compound selectors", () => {
  const result = parseSelector("slide[1] shape");
  assert.ok(result.ok);
  assert.equal(result.data?.slideNum, 1);
  assert.equal(result.data?.elementType, "shape");
});

test("parseSelector - parses child combinator", () => {
  const result = parseSelector("slide[1] > shape[2]");
  assert.ok(result.ok);
  assert.equal(result.data?.slideNum, 1);
  assert.equal(result.data?.elementType, "shape");
  assert.equal(result.data?.attributes.index, "2");
  assert.equal(result.data?.childCombinator, ">");
});

test("parseSelector - parses text shorthand", () => {
  const result = parseSelector("shape:Find me");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "shape");
  assert.equal(result.data?.textContains, "Find me");
});

test("parseSelector - parses attribute filters", () => {
  const result = parseSelector("shape[@type=text]");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "shape");
  assert.equal(result.data?.attributes.type, "text");
});

test("parseSelector - parses chart selector", () => {
  const result = parseSelector("chart");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "chart");
});

test("parseSelector - parses picture selector", () => {
  const result = parseSelector("picture");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "picture");
});

test("parseSelector - parses table selector", () => {
  const result = parseSelector("table");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "table");
});

test("parseSelector - parses media selector", () => {
  const result = parseSelector("media");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "media");
});

test("parseSelector - parses notes selector", () => {
  const result = parseSelector("notes");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "notes");
});

test("parseSelector - parses group selector", () => {
  const result = parseSelector("group");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "group");
});

test("parseSelector - parses zoom selector", () => {
  const result = parseSelector("zoom");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "zoom");
});

test("parseSelector - parses table cell selector", () => {
  const result = parseSelector("tc");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "tc");
});

test("parseSelector - parses table row selector", () => {
  const result = parseSelector("tr");
  assert.ok(result.ok);
  assert.equal(result.data?.elementType, "tr");
});

test("parseSelector - rejects empty selector", () => {
  const result = parseSelector("");
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "invalid_selector");
});

test("parseSlideSelector - same as parseSelector", () => {
  const result = parseSlideSelector("slide[1] shape");
  assert.ok(result.ok);
  assert.equal(result.data?.slideNum, 1);
  assert.equal(result.data?.elementType, "shape");
});

test("buildSelector - builds basic selectors", () => {
  const selector = buildSelector({ elementType: "shape", attributes: {} });
  assert.equal(selector, "shape");
});

test("buildSelector - builds indexed selectors", () => {
  const selector = buildSelector({
    elementType: "shape",
    slideNum: 1,
    attributes: { index: "2" },
  });
  assert.equal(selector, "slide[1] shape[2]");
});

test("buildSelector - builds :contains selectors", () => {
  const selector = buildSelector({
    elementType: "shape",
    attributes: {},
    textContains: "Hello",
  });
  assert.equal(selector, 'shape:contains("Hello")');
});

test("buildSelector - builds :empty selectors", () => {
  const selector = buildSelector({
    elementType: "shape",
    attributes: { empty: "true" },
  });
  assert.equal(selector, "shape:empty");
});

test("buildSelector - builds placeholder selectors", () => {
  const selector = buildSelector({
    elementType: "placeholder",
    attributes: { name: "title" },
  });
  assert.equal(selector, "placeholder[title]");
});

test("validateSelector - validates selectors", () => {
  assert.ok(validateSelector("slide[1]").ok);
  assert.ok(validateSelector("shape:contains(text)").ok);
  assert.ok(!validateSelector("").ok);
});

test("isValidSelector - checks selector validity", () => {
  assert.ok(isValidSelector("slide[1]"));
  assert.ok(isValidSelector("shape:empty"));
  assert.ok(!isValidSelector(""));
});

test("slideSelector - creates slide selector", () => {
  const selector = slideSelector(1, "shape");
  assert.equal(selector.slideNum, 1);
  assert.equal(selector.elementType, "shape");
});

test("typeSelector - creates type selector", () => {
  const selector = typeSelector("chart");
  assert.equal(selector.elementType, "chart");
});

test("textSelector - creates text selector", () => {
  const selector = textSelector("shape", "Hello");
  assert.equal(selector.elementType, "shape");
  assert.equal(selector.textContains, "Hello");
});

test("hasTextFilter - checks for text filter", () => {
  assert.ok(hasTextFilter({ elementType: "shape", attributes: {}, textContains: "test" }));
  assert.ok(!hasTextFilter({ elementType: "shape", attributes: {} }));
});

test("hasAttributeFilter - checks for attribute filter", () => {
  assert.ok(hasAttributeFilter({ elementType: "shape", attributes: { index: "1" } }));
  assert.ok(!hasAttributeFilter({ elementType: "shape", attributes: {} }));
});

test("isElementType - checks element type", () => {
  const selector = { elementType: "shape", attributes: {} };
  assert.ok(isElementType(selector, "shape"));
  assert.ok(!isElementType(selector, "table"));
});

test("isPlaceholderSelector - checks placeholder", () => {
  const selector = { elementType: "placeholder", attributes: { name: "title" } };
  assert.ok(isPlaceholderSelector(selector));
  assert.ok(!isPlaceholderSelector({ elementType: "shape", attributes: {} }));
});

test("isCellSelector - checks cell", () => {
  assert.ok(isCellSelector({ elementType: "tc", attributes: {} }));
  assert.ok(isCellSelector({ elementType: "cell", attributes: {} }));
  assert.ok(!isCellSelector({ elementType: "tr", attributes: {} }));
});

test("isRowSelector - checks row", () => {
  assert.ok(isRowSelector({ elementType: "tr", attributes: {} }));
  assert.ok(isRowSelector({ elementType: "row", attributes: {} }));
  assert.ok(!isRowSelector({ elementType: "tc", attributes: {} }));
});

test("normalizePlaceholderTypeName - normalizes placeholder types", () => {
  assert.equal(normalizePlaceholderTypeName("CenterTitle"), "title");
  assert.equal(normalizePlaceholderTypeName("centeredtitle"), "title");
  assert.equal(normalizePlaceholderTypeName("ctitle"), "title");
  assert.equal(normalizePlaceholderTypeName("Subtitle"), "subtitle");
  assert.equal(normalizePlaceholderTypeName("sub"), "subtitle");
  assert.equal(normalizePlaceholderTypeName("Date"), "date");
  assert.equal(normalizePlaceholderTypeName("datetime"), "date");
  assert.equal(normalizePlaceholderTypeName("dt"), "date");
  assert.equal(normalizePlaceholderTypeName("SlideNum"), "slidenum");
  assert.equal(normalizePlaceholderTypeName("slidenumber"), "slidenum");
  assert.equal(normalizePlaceholderTypeName("sldnum"), "slidenum");
  assert.equal(normalizePlaceholderTypeName("Object"), "object");
  assert.equal(normalizePlaceholderTypeName("obj"), "object");
  assert.equal(normalizePlaceholderTypeName("diagram"), "diagram");
  assert.equal(normalizePlaceholderTypeName("dgm"), "diagram");
});
