/**
 * Selector grammar parser for @officekit/ppt.
 *
 * Selectors are used to query elements within a presentation using a CSS-like syntax.
 * They can filter by element type, position, attributes, and text content.
 *
 * Selector Syntax:
 * ---------------
 * [basic]     := elementtype[index]
 * [type]      := elementtype[index][@attr=value]  (attribute filter)
 * [text]      := elementtype:contains("text")
 * [combinator]:= ancestor descendant
 *              | parent > child
 *              | prev + next  (adjacent sibling)
 * [compound]  := selector:contains("text")[@attr=value]
 *
 * Examples:
 * ---------
 * - `slide[1]` - First slide
 * - `slide[1] shape` - Any shape inside first slide
 * - `slide[1] > shape[1]` - First direct child shape of first slide
 * - `shape[type=text]` - Shapes with type attribute = "text"
 * - `placeholder[title]` - Title placeholders
 * - `shape:contains("Hello")` - Shapes containing "Hello"
 * - `shape:empty` - Shapes with no text
 * - `shape:no-alt` - Shapes without alt text
 * - `picture + picture` - Consecutive pictures
 *
 * This is a simplified selector parser focused on the PPT use cases
 * observed in OfficeCLI PowerPointHandler.Query.cs.
 */

import { err, ok, andThen } from "./result.js";
import type { ParsedSelector, Result } from "./types.js";

// ============================================================================
// Selector Pattern Constants
// ============================================================================

/**
 * Regex patterns for parsing selectors.
 */
const SELECTOR_PATTERNS = {
  /** Matches element[index] or element[index] followed by rest */
  INDEXED: /^([a-zA-Z]+)\[(\d+)\]/,
  /** Matches element[name] for named selectors like placeholder[title] */
  NAMED: /^([a-zA-Z]+)\[([a-zA-Z][a-zA-Z0-9_]*)\]/,
  /** Matches element[@attr=value] for attribute filters */
  ATTRIBUTE: /^([a-zA-Z]+)\[@([a-zA-Z]+)=([^\]]+)\]/,
  /** Matches :contains("text") or :contains(text) */
  CONTAINS: /^:contains\(\s*(["']?)(.*?)\1\s*\)/,
  /** Matches :empty pseudo-selector */
  EMPTY: /^:empty\b/,
  /** Matches :no-alt pseudo-selector (shapes without alt text) */
  NO_ALT: /^:no-alt\b/,
  /** Matches :text shorthand for :contains on previous selector */
  TEXT_SHORTHAND: /^:(["'])(.*?)\1/,
  /** Matches child combinator > */
  CHILD_COMBINATOR: /^\s*>\s*/,
  /** Matches descendant whitespace */
  DESCENDANT_COMBINATOR: /^\s+/,
  /** Matches adjacent sibling + */
  ADJACENT_SIBLING: /^\s*\+\s*/,
} as const;

/**
 * Known element types in selectors.
 */
const KNOWN_ELEMENT_TYPES = new Set([
  "shape",
  "textbox",
  "title",
  "picture",
  "pic",
  "video",
  "audio",
  "equation",
  "math",
  "formula",
  "table",
  "chart",
  "placeholder",
  "notes",
  "connector",
  "connection",
  "group",
  "zoom",
  "slidemaster",
  "slidelayout",
  "media",
  "image",
  "tc",
  "cell",
  "tr",
  "row",
  "paragraph",
  "run",
]);

// ============================================================================
// Selector Parsing
// ============================================================================

/**
 * Parses a selector string into a ParsedSelector structure.
 *
 * @example
 * parseSelector("slide[1] shape")
 * // Returns: { elementType: "shape", slideNum: 1, attributes: {}, childCombinator: " " }
 *
 * @example
 * parseSelector('shape:contains("Hello")')
 * // Returns: { elementType: "shape", attributes: {}, textContains: "Hello" }
 */
export function parseSelector(selector: string): Result<ParsedSelector> {
  if (!selector || typeof selector !== "string") {
    return err("invalid_selector", "Selector must be a non-empty string");
  }

  const result: ParsedSelector = {
    attributes: {},
  };

  let remaining = selector.trim();

  // Handle :text shorthand before other processing
  // e.g., "shape:Find me" should become elementType=shape, textContains="Find me"
  // BUT NOT :empty, :no-alt, or :contains(...)
  const textShorthandMatch = remaining.match(/:([a-zA-Z][^"']*)$/);
  if (textShorthandMatch && !remaining.includes(":contains(") && !remaining.includes(":empty") && !remaining.includes(":no-alt")) {
    // This is the text shorthand syntax
    result.textContains = textShorthandMatch[1];
    remaining = remaining.slice(0, -textShorthandMatch[0].length);
  }

  // Check for adjacent sibling combinator at the start
  if (SELECTOR_PATTERNS.ADJACENT_SIBLING.test(remaining)) {
    result.adjacentSibling = true;
    remaining = remaining.replace(SELECTOR_PATTERNS.ADJACENT_SIBLING, "");
  }

  // Parse element type with optional index
  // e.g., "slide[1]" -> elementType="slide", slideNum=1
  // e.g., "shape[2]" -> elementType="shape", index=2 (stored as slideNum for slide filter)
  const indexedMatch = remaining.match(SELECTOR_PATTERNS.INDEXED);
  if (indexedMatch) {
    const type = indexedMatch[1].toLowerCase();
    const index = parseInt(indexedMatch[2], 10);

    // If it's a slide selector, store the slide number
    if (type === "slide") {
      result.slideNum = index;
    } else {
      result.elementType = type;
      // Store index in attributes if not slide
      result.attributes.index = String(index);
    }

    remaining = remaining.slice(indexedMatch[0].length);
  }

  // Parse named selector
  // e.g., "placeholder[title]" -> elementType="placeholder", name="title"
  const namedMatch = remaining.match(SELECTOR_PATTERNS.NAMED);
  if (namedMatch) {
    const type = namedMatch[1].toLowerCase();
    const name = namedMatch[2];

    result.elementType = type;
    result.attributes.name = name;

    remaining = remaining.slice(namedMatch[0].length);
  }

  // Parse attribute filter
  // e.g., "shape[@type=text]" -> elementType="shape", attributes.type="text"
  const attrMatch = remaining.match(SELECTOR_PATTERNS.ATTRIBUTE);
  if (attrMatch) {
    const type = attrMatch[1].toLowerCase();
    const attrName = attrMatch[2];
    const attrValue = attrMatch[3];

    if (!result.elementType) {
      result.elementType = type;
    }
    result.attributes[attrName] = attrValue;

    remaining = remaining.slice(attrMatch[0].length);
  }

  // If there's remaining content and no element type yet, extract it
  // e.g., "shape:contains(...)" -> elementType=shape
  // Trim leading whitespace first (for cases like "slide[1] shape")
  remaining = remaining.trimStart();
  if (remaining.length > 0 && !result.elementType) {
    const typeMatch = remaining.match(/^([a-zA-Z]+)/);
    if (typeMatch) {
      result.elementType = typeMatch[1].toLowerCase();
      remaining = remaining.slice(typeMatch[0].length);
    }
  }

  // Parse :contains pseudo-selector (AFTER element type is extracted)
  const containsMatch = remaining.match(SELECTOR_PATTERNS.CONTAINS);
  if (containsMatch) {
    result.textContains = containsMatch[2];
    remaining = remaining.slice(containsMatch[0].length);
  }

  // Parse :empty pseudo-selector
  if (SELECTOR_PATTERNS.EMPTY.test(remaining)) {
    result.attributes.empty = "true";
    remaining = remaining.replace(SELECTOR_PATTERNS.EMPTY, "");
  }

  // Parse :no-alt pseudo-selector
  if (SELECTOR_PATTERNS.NO_ALT.test(remaining)) {
    result.attributes.noAlt = "true";
    remaining = remaining.replace(SELECTOR_PATTERNS.NO_ALT, "");
  }

  // Parse combinators
  if (SELECTOR_PATTERNS.CHILD_COMBINATOR.test(remaining)) {
    result.childCombinator = ">";
    remaining = remaining.replace(SELECTOR_PATTERNS.CHILD_COMBINATOR, "");
  } else if (SELECTOR_PATTERNS.DESCENDANT_COMBINATOR.test(remaining)) {
    result.childCombinator = " ";
    remaining = remaining.replace(SELECTOR_PATTERNS.DESCENDANT_COMBINATOR, "");
  }

  // After parsing a combinator, continue parsing remaining selector parts
  // This handles cases like "slide[1] > shape[2]"
  if (result.childCombinator && remaining.length > 0) {
    // Trim and try to parse the next selector
    remaining = remaining.trimStart();

    // Parse indexed selector after combinator (e.g., "shape[2]")
    const indexedAfterCombinator = remaining.match(/^([a-zA-Z]+)\[(\d+)\]/);
    if (indexedAfterCombinator) {
      result.elementType = indexedAfterCombinator[1].toLowerCase();
      result.attributes.index = indexedAfterCombinator[2];
      remaining = remaining.slice(indexedAfterCombinator[0].length);
    } else {
      // Try to extract just the element type
      const typeMatch = remaining.match(/^([a-zA-Z]+)/);
      if (typeMatch) {
        result.elementType = typeMatch[1].toLowerCase();
        remaining = remaining.slice(typeMatch[0].length);
      }
    }
  }

  // Handle attribute filters after element type and pseudo-selectors
  const postAttrMatch = remaining.match(SELECTOR_PATTERNS.ATTRIBUTE);
  if (postAttrMatch) {
    const attrName = postAttrMatch[2];
    const attrValue = postAttrMatch[3];
    result.attributes[attrName] = attrValue;
    remaining = remaining.slice(postAttrMatch[0].length);
  }

  // If we still have content, we may have a partial parse
  // For forward compatibility, we just note any unparsed content
  if (remaining.length > 0) {
    // Could be additional selectors for compound queries
    // For now, we don't handle this but we don't error either
  }

  return ok(result);
}

/**
 * Parses a selector that may include a slide filter prefix.
 *
 * @example
 * parseSlideSelector("slide[1] shape")
 * // Returns: { elementType: "shape", slideNum: 1, attributes: {} }
 */
export function parseSlideSelector(selector: string): Result<ParsedSelector> {
  return parseSelector(selector);
}

// ============================================================================
// Selector Building
// ============================================================================

/**
 * Builds a selector string from a ParsedSelector.
 *
 * @example
 * buildSelector({ elementType: "shape", slideNum: 1, textContains: "Hello" })
 * // Returns: 'slide[1] shape:contains("Hello")'
 */
export function buildSelector(parsed: ParsedSelector): string {
  const parts: string[] = [];

  // Add slide filter if present
  if (parsed.slideNum !== undefined) {
    parts.push(`slide[${parsed.slideNum}]`);
  }

  // Add element type or default
  const elementType = parsed.elementType || "shape";
  parts.push(elementType);

  // Add index attribute if present
  if (parsed.attributes.index !== undefined) {
    parts[parts.length - 1] += `[${parsed.attributes.index}]`;
  }

  // Add name attribute if present
  if (parsed.attributes.name !== undefined) {
    parts[parts.length - 1] += `[${parsed.attributes.name}]`;
  }

  // Add attribute filters
  for (const [key, value] of Object.entries(parsed.attributes)) {
    if (key !== "index" && key !== "name" && key !== "empty" && key !== "noAlt") {
      parts[parts.length - 1] += `[@${key}=${value}]`;
    }
  }

  // Add :empty
  if (parsed.attributes.empty === "true") {
    parts[parts.length - 1] += ":empty";
  }

  // Add :no-alt
  if (parsed.attributes.noAlt === "true") {
    parts[parts.length - 1] += ":no-alt";
  }

  // Add :contains
  if (parsed.textContains !== undefined) {
    parts[parts.length - 1] += `:contains("${parsed.textContains}")`;
  }

  // Add child combinator if present
  if (parsed.childCombinator === ">") {
    // Insert > between slide filter and element
    if (parts.length > 1) {
      parts.splice(1, 0, ">");
    } else {
      parts.push(">");
    }
  }

  // Add adjacent sibling
  if (parsed.adjacentSibling) {
    parts.push("+");
  }

  return parts.join(" ");
}

// ============================================================================
// Selector Validation
// ============================================================================

/**
 * Validates that a selector is well-formed.
 *
 * @example
 * validateSelector("slide[1] shape")
 * // Returns: { ok: true } or { ok: false, error: { code: "invalid_selector", message: "..." } }
 */
export function validateSelector(selector: string): Result<void> {
  return andThen(parseSelector(selector), () => ok(void 0));
}

/**
 * Checks if a selector is valid without returning detailed error.
 */
export function isValidSelector(selector: string): boolean {
  return parseSelector(selector).ok;
}

// ============================================================================
// Selector Helpers
// ============================================================================

/**
 * Creates a selector for a specific slide's elements.
 *
 * @example
 * slideSelector(1, "shape")
 * // Returns: { elementType: "shape", slideNum: 1, attributes: {} }
 */
export function slideSelector(slideNum: number, elementType = "shape"): ParsedSelector {
  return {
    elementType,
    slideNum,
    attributes: {},
  };
}

/**
 * Creates a selector for a specific element type across all slides.
 *
 * @example
 * typeSelector("chart")
 * // Returns: { elementType: "chart", attributes: {} }
 */
export function typeSelector(elementType: string): ParsedSelector {
  return {
    elementType,
    attributes: {},
  };
}

/**
 * Creates a selector with a text filter.
 *
 * @example
 * textSelector("shape", "Hello")
 * // Returns: { elementType: "shape", attributes: {}, textContains: "Hello" }
 */
export function textSelector(elementType: string, text: string): ParsedSelector {
  return {
    elementType,
    attributes: {},
    textContains: text,
  };
}

/**
 * Checks if a parsed selector has a text filter.
 */
export function hasTextFilter(parsed: ParsedSelector): boolean {
  return parsed.textContains !== undefined;
}

/**
 * Checks if a parsed selector has an attribute filter.
 */
export function hasAttributeFilter(parsed: ParsedSelector): boolean {
  return Object.keys(parsed.attributes).length > 0;
}

/**
 * Checks if a parsed selector targets a specific element type.
 */
export function isElementType(parsed: ParsedSelector, type: string): boolean {
  return parsed.elementType === type.toLowerCase();
}

/**
 * Checks if a parsed selector targets a placeholder.
 */
export function isPlaceholderSelector(parsed: ParsedSelector): boolean {
  return isElementType(parsed, "placeholder");
}

/**
 * Checks if a parsed selector targets a table cell.
 */
export function isCellSelector(parsed: ParsedSelector): boolean {
  return isElementType(parsed, "tc") || isElementType(parsed, "cell");
}

/**
 * Checks if a parsed selector targets a table row.
 */
export function isRowSelector(parsed: ParsedSelector): boolean {
  return isElementType(parsed, "tr") || isElementType(parsed, "row");
}

/**
 * Normalizes placeholder type names in selectors.
 *
 * @example
 * normalizePlaceholderTypeName("CenterTitle")  // Returns: "title"
 * normalizePlaceholderTypeName("ctitle")  // Returns: "title"
 */
export function normalizePlaceholderTypeName(name: string): string {
  const lower = name.toLowerCase();
  switch (lower) {
    case "centertitle":
    case "centeredtitle":
    case "ctitle":
      return "title";
    case "subtitle":
    case "sub":
      return "subtitle";
    case "date":
    case "datetime":
    case "dt":
      return "date";
    case "slidenum":
    case "slidenumber":
    case "sldnum":
      return "slidenum";
    case "object":
    case "obj":
      return "object";
    case "clipart":
      return "clipart";
    case "diagram":
    case "dgm":
      return "diagram";
    case "picture":
    case "pic":
      return "picture";
    default:
      return lower;
  }
}
