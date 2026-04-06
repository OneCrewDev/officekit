/**
 * Path parsing and resolution for @officekit/ppt.
 *
 * PPT paths are absolute paths that identify elements within a presentation:
 * - `/presentation` - the root presentation
 * - `/slide[N]` - the Nth slide (1-indexed)
 * - `/slide[N]/shape[M]` - the Mth shape on slide N
 * - `/slide[N]/table[M]` - the Mth table on slide N
 * - `/slide[N]/placeholder[title]` - placeholder by type name
 * - `/slide[N]/chart[M]` - the Mth chart on slide N
 * - `/slide[N]/picture[M]` or `/slide[N]/media[M]` - media elements
 * - `/slide[N]/notes` - notes for a slide
 * - `/slide[N]/connector[M]` - the Mth connector
 * - `/slide[N]/group[M]` - the Mth group
 * - `/slide[N]/zoom[M]` - the Mth zoom element
 * - `/slide[N]/tr[R]/tc[C]` - table row and cell
 * - `/slidemaster[N]` - the Nth slide master
 * - `/slidelayout[N]` - the Nth slide layout
 * - `/theme` - the presentation theme
 *
 * Syntax:
 * - `/element[index]` - indexed access (1-indexed)
 * - `/element[name]` - named access by type (e.g., placeholder[title])
 * - `/element[@attr=value]` - attribute filtering (future)
 *
 * @example
 * parsePath("/slide[1]/shape[2]")
 * // Returns: { isAbsolute: true, segments: [{ name: "slide", index: 1 }, { name: "shape", index: 2 }], original: "/slide[1]/shape[2]" }
 */

import { err, ok, andThen } from "./result.js";
import type { ParsedPath, PathSegment, Result } from "./types.js";

// ============================================================================
// Path Pattern Constants
// ============================================================================

/**
 * Regex patterns for parsing path segments.
 */
const SEGMENT_PATTERNS = {
  /** Matches /element[index] or /element[index]/... */
  INDEXED: /^\/([a-zA-Z]+)\[(\d+)\](.*)$/,
  /** Matches /element[name] where name is a word (placeholder types, etc.) */
  NAMED: /^\/([a-zA-Z]+)\[([a-zA-Z][a-zA-Z0-9_]*)\](.*)$/,
  /** Matches /element[@attr=value] for attribute filtering */
  ATTRIBUTE: /^\/([a-zA-Z]+)\[@([a-zA-Z]+)=([^\]]+)\](.*)$/,
} as const;

/**
 * Valid top-level path segments.
 */
const VALID_ROOT_SEGMENTS = new Set([
  "presentation",
  "slide",
  "slidemaster",
  "slidelayout",
  "theme",
]);

/**
 * Valid element types that can appear after a slide.
 */
const VALID_SLIDE_CHILDREN = new Set([
  "shape",
  "table",
  "placeholder",
  "chart",
  "picture",
  "pic",
  "media",
  "image",
  "notes",
  "connector",
  "connection",
  "group",
  "zoom",
  "audio",
  "video",
  "tr",
  "tc",
  "paragraph",
  "run",
  "animation",
]);

/**
 * Valid element types that can appear after a table.
 */
const VALID_TABLE_CHILDREN = new Set(["tr", "tc"]);

/**
 * Placeholder type names.
 */
const PLACEHOLDER_TYPES = new Set([
  "title",
  "body",
  "subtitle",
  "centertitle",
  "centeredtitle",
  "ctitle",
  "date",
  "datetime",
  "dt",
  "footer",
  "slidenum",
  "slidenumber",
  "sldnum",
  "object",
  "obj",
  "chart",
  "table",
  "clipart",
  "diagram",
  "dgm",
  "media",
  "picture",
  "pic",
  "header",
]);

// ============================================================================
// Path Parsing
// ============================================================================

/**
 * Parses a PPT path string into a structured ParsedPath.
 *
 * @example
 * parsePath("/slide[1]/shape[2]")
 * // Returns: { isAbsolute: true, segments: [{ name: "slide", index: 1 }, { name: "shape", index: 2 }], original: "/slide[1]/shape[2]" }
 *
 * @example
 * parsePath("/slide[1]/placeholder[title]")
 * // Returns: { isAbsolute: true, segments: [{ name: "slide", index: 1 }, { name: "placeholder", nameSelector: "title" }], original: "/slide[1]/placeholder[title]" }
 */
export function parsePath(path: string): Result<ParsedPath> {
  if (!path || typeof path !== "string") {
    return err("invalid_path", "Path must be a non-empty string");
  }

  const original = path;
  const isAbsolute = path.startsWith("/");

  if (!isAbsolute) {
    return err("invalid_path", "PPT paths must be absolute (start with /)", "Paths should start with /presentation, /slide[N], etc.");
  }

  // Handle root paths
  if (path === "/") {
    return ok({ isAbsolute: true, segments: [], original });
  }

  if (path === "/presentation") {
    return ok({ isAbsolute: true, segments: [{ name: "presentation" }], original });
  }

  if (path === "/theme") {
    return ok({ isAbsolute: true, segments: [{ name: "theme" }], original });
  }

  // Handle /slidemaster[N] and /slidelayout[N]
  const masterMatch = path.match(/^\/slidemaster\[(\d+)\]$/i);
  if (masterMatch) {
    const index = parseInt(masterMatch[1], 10);
    if (index < 1) {
      return err("invalid_path", "Slide master index must be at least 1");
    }
    return ok({
      isAbsolute: true,
      segments: [{ name: "slidemaster", index }],
      original,
    });
  }

  const layoutMatch = path.match(/^\/slidelayout\[(\d+)\]$/i);
  if (layoutMatch) {
    const index = parseInt(layoutMatch[1], 10);
    if (index < 1) {
      return err("invalid_path", "Slide layout index must be at least 1");
    }
    return ok({
      isAbsolute: true,
      segments: [{ name: "slidelayout", index }],
      original,
    });
  }

  // Parse sequential segments
  const segments: PathSegment[] = [];
  let remaining = path;

  while (remaining.length > 0 && remaining !== "/") {
    const segmentResult = parseSegment(remaining);
    if (!segmentResult.ok) {
      return segmentResult;
    }

    const { segment, rest } = segmentResult.data as { segment: PathSegment; rest: string };
    segments.push(segment);
    remaining = rest;
  }

  if (segments.length === 0) {
    return err("invalid_path", "Path contains no valid segments");
  }

  // Validate segment sequence
  const validationResult = validateSegmentSequence(segments);
  if (!validationResult.ok) {
    return validationResult;
  }

  return ok({ isAbsolute, segments, original });
}

/**
 * Parses a single path segment from the beginning of a path string.
 */
function parseSegment(path: string): Result<{ segment: PathSegment; rest: string }> {
  // Try indexed segment: /element[index]
  let match = path.match(SEGMENT_PATTERNS.INDEXED);
  if (match) {
    const name = match[1].toLowerCase();
    const index = parseInt(match[2], 10);
    const rest = match[3] || "";

    if (index < 1) {
      return err("invalid_path", `${name} index must be at least 1`);
    }

    return ok({
      segment: { name, index },
      rest,
    });
  }

  // Try named segment: /element[name] (for placeholders like placeholder[title])
  match = path.match(SEGMENT_PATTERNS.NAMED);
  if (match) {
    const name = match[1].toLowerCase();
    const nameSelector = match[2];
    const rest = match[3] || "";

    // Check if this is a valid named segment
    if (name === "placeholder") {
      if (!PLACEHOLDER_TYPES.has(nameSelector.toLowerCase())) {
        // It might be a numeric index being treated as a name - let the indexed pattern handle it
        // But first check if it's a valid placeholder type alias
        const normalizedName = normalizePlaceholderType(nameSelector);
        return ok({
          segment: { name, nameSelector: normalizedName },
          rest,
        });
      }
      return ok({
        segment: { name, nameSelector: nameSelector.toLowerCase() },
        rest,
      });
    }

    // For other elements, treat as named selector
    return ok({
      segment: { name, nameSelector },
      rest,
    });
  }

  // Try attribute filter segment: /element[@attr=value]
  match = path.match(SEGMENT_PATTERNS.ATTRIBUTE);
  if (match) {
    const name = match[1].toLowerCase();
    const attrName = match[2];
    const attrValue = match[3];
    const rest = match[4] || "";

    return ok({
      segment: { name, typeFilter: attrValue },
      rest,
    });
  }

  return err("invalid_path", `Invalid path segment: ${path.slice(0, 20)}...`);
}

/**
 * Validates that a sequence of path segments forms a valid path.
 */
function validateSegmentSequence(segments: PathSegment[]): Result<void> {
  if (segments.length === 0) {
    return ok(void 0);
  }

  // First segment must be a valid root segment
  const first = segments[0];
  if (!VALID_ROOT_SEGMENTS.has(first.name)) {
    return err(
      "invalid_path",
      `Invalid root segment: ${first.name}`,
      "Paths must start with /presentation, /slide[N], /slidemaster[N], or /slidelayout[N]",
    );
  }

  // Validate segment-specific rules
  for (let i = 0; i < segments.length; i++) {
    const current = segments[i];
    const prev = i > 0 ? segments[i - 1] : null;

    // After slide, slidemaster, or slidelayout, valid children differ
    if (prev && (prev.name === "slide" || prev.name === "slidemaster" || prev.name === "slidelayout")) {
      if (!VALID_SLIDE_CHILDREN.has(current.name)) {
        return err(
          "invalid_path",
          `Invalid child segment '${current.name}' after ${prev.name}`,
          `Valid children of ${prev.name} include: ${[...VALID_SLIDE_CHILDREN].join(", ")}`,
        );
      }
    }

    // After table, only tr/tc are valid
    if (prev && prev.name === "table") {
      if (!VALID_TABLE_CHILDREN.has(current.name)) {
        return err(
          "invalid_path",
          `Invalid child segment '${current.name}' after table`,
          "Valid children of table include: tr, tc",
        );
      }
    }

    // After tr, only tc is valid
    if (prev && prev.name === "tr") {
      if (current.name !== "tc") {
        return err("invalid_path", "Invalid child segment after table row", "Only tc (table cell) is valid after tr");
      }
    }
  }

  return ok(void 0);
}

/**
 * Normalizes placeholder type names to their canonical form.
 */
function normalizePlaceholderType(name: string): string {
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

// ============================================================================
// Path Building
// ============================================================================

/**
 * Builds a path string from segments.
 *
 * @example
 * buildPath([{ name: "slide", index: 1 }, { name: "shape", index: 2 }])
 * // Returns: "/slide[1]/shape[2]"
 */
export function buildPath(segments: PathSegment[]): string {
  if (segments.length === 0) {
    return "/";
  }

  return segments
    .map((seg) => {
      if (seg.index !== undefined) {
        return `/${seg.name}[${seg.index}]`;
      }
      if (seg.nameSelector !== undefined) {
        return `/${seg.name}[${seg.nameSelector}]`;
      }
      if (seg.typeFilter !== undefined) {
        return `/${seg.name}[@type=${seg.typeFilter}]`;
      }
      return `/${seg.name}`;
    })
    .join("");
}

/**
 * Builds a slide path.
 *
 * @example
 * slidePath(3)  // Returns: "/slide[3]"
 */
export function slidePath(slideIndex: number): string {
  return `/slide[${slideIndex}]`;
}

/**
 * Builds a shape path.
 *
 * @example
 * shapePath(3, 5)  // Returns: "/slide[3]/shape[5]"
 */
export function shapePath(slideIndex: number, shapeIndex: number): string {
  return `/slide[${slideIndex}]/shape[${shapeIndex}]`;
}

/**
 * Builds a table path.
 *
 * @example
 * tablePath(3, 1)  // Returns: "/slide[3]/table[1]"
 */
export function tablePath(slideIndex: number, tableIndex: number): string {
  return `/slide[${slideIndex}]/table[${tableIndex}]`;
}

/**
 * Builds a placeholder path by type name.
 *
 * @example
 * placeholderPath(3, "title")  // Returns: "/slide[3]/placeholder[title]"
 */
export function placeholderPath(slideIndex: number, placeholderType: string): string {
  return `/slide[${slideIndex}]/placeholder[${placeholderType}]`;
}

/**
 * Builds a chart path.
 *
 * @example
 * chartPath(3, 1)  // Returns: "/slide[3]/chart[1]"
 */
export function chartPath(slideIndex: number, chartIndex: number): string {
  return `/slide[${slideIndex}]/chart[${chartIndex}]`;
}

/**
 * Builds a table cell path.
 *
 * @example
 * cellPath(3, 1, 2, 3)  // Returns: "/slide[3]/table[1]/tr[2]/tc[3]"
 */
export function cellPath(slideIndex: number, tableIndex: number, rowIndex: number, cellIndex: number): string {
  return `/slide[${slideIndex}]/table[${tableIndex}]/tr[${rowIndex}]/tc[${cellIndex}]`;
}

// ============================================================================
// Path Queries
// ============================================================================

/**
 * Extracts the slide index from a path, if present.
 *
 * @example
 * getSlideIndex("/slide[5]/shape[1]")  // Returns: 5
 * getSlideIndex("/presentation")  // Returns: null
 */
export function getSlideIndex(path: string): number | null {
  const match = path.match(/^\/slide\[(\d+)\]/i);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * Checks if a path refers to the presentation root.
 *
 * @example
 * isPresentationPath("/")  // Returns: true
 * isPresentationPath("/presentation")  // Returns: true
 */
export function isPresentationPath(path: string): boolean {
  return path === "/" || path === "/presentation";
}

/**
 * Checks if a path refers to a slide.
 *
 * @example
 * isSlidePath("/slide[1]")  // Returns: true
 * isSlidePath("/slide[1]/shape[2]")  // Returns: true
 */
export function isSlidePath(path: string): boolean {
  return /^\/slide\[\d+\]/i.test(path);
}

/**
 * Checks if a path refers to a table cell.
 *
 * @example
 * isCellPath("/slide[1]/table[1]/tr[2]/tc[3]")  // Returns: true
 */
export function isCellPath(path: string): boolean {
  return /\/tc\[\d+\]$/i.test(path);
}

/**
 * Checks if a path refers to a placeholder.
 *
 * @example
 * isPlaceholderPath("/slide[1]/placeholder[title]")  // Returns: true
 */
export function isPlaceholderPath(path: string): boolean {
  return /\/placeholder\[/i.test(path);
}

/**
 * Gets the parent path of a given path.
 *
 * @example
 * parentPath("/slide[1]/shape[2]")  // Returns: "/slide[1]"
 * parentPath("/slide[1]")  // Returns: "/"
 */
export function parentPath(path: string): string {
  const lastSlash = path.lastIndexOf("/");
  if (lastSlash <= 0) {
    return "/";
  }
  return path.slice(0, lastSlash) || "/";
}

/**
 * Gets the last segment name from a path.
 *
 * @example
 * lastSegmentName("/slide[1]/shape[2]")  // Returns: "shape"
 */
export function lastSegmentName(path: string): string | null {
  // Use global match to find all segments, then get the last one
  const matches = path.match(/\/([a-zA-Z]+)\[/g);
  if (!matches || matches.length === 0) {
    return null;
  }
  const lastMatch = matches[matches.length - 1];
  const match = lastMatch.match(/\/([a-zA-Z]+)\[/);
  return match ? match[1].toLowerCase() : null;
}

// ============================================================================
// Validation
// ============================================================================

/**
 * Validates that a path is well-formed.
 *
 * @example
 * validatePath("/slide[1]/shape[2]")
 * // Returns: { ok: true } or { ok: false, error: { code: "invalid_path", message: "..." } }
 */
export function validatePath(path: string): Result<void> {
  return andThen(parsePath(path), () => ok(void 0));
}

/**
 * Checks if a path is valid without returning detailed error.
 *
 * @example
 * isValidPath("/slide[1]/shape[2]")  // Returns: true
 * isValidPath("invalid")  // Returns: false
 */
export function isValidPath(path: string): boolean {
  return parsePath(path).ok;
}

// ============================================================================
// Element Type Detection
// ============================================================================

/**
 * Known element types in PPT paths.
 */
export type PptElementType =
  | "presentation"
  | "slide"
  | "slidemaster"
  | "slidelayout"
  | "theme"
  | "shape"
  | "textbox"
  | "table"
  | "placeholder"
  | "chart"
  | "picture"
  | "pic"
  | "media"
  | "image"
  | "notes"
  | "connector"
  | "connection"
  | "group"
  | "zoom"
  | "audio"
  | "video"
  | "tr"
  | "tc"
  | "paragraph"
  | "run"
  | "animation";

/**
 * Gets the element type of the final segment in a path.
 *
 * @example
 * getElementType("/slide[1]/shape[2]")  // Returns: "shape"
 * getElementType("/slide[1]/table[1]/tr[2]/tc[3]")  // Returns: "tc"
 */
export function getElementType(path: string): PptElementType | null {
  // Handle root paths first
  if (path === "/" || path === "/presentation") return "presentation";
  if (path === "/theme") return "theme";

  // Use global match to find all segments, then get the last one
  const matches = path.match(/\/([a-zA-Z]+)\[/g);
  if (!matches || matches.length === 0) {
    return null;
  }
  const lastMatch = matches[matches.length - 1];
  const match = lastMatch.match(/\/([a-zA-Z]+)\[/);
  return match ? match[1].toLowerCase() as PptElementType : null;
}

/**
 * Checks if a path targets a specific element type.
 *
 * @example
 * isElementType("/slide[1]/shape[2]", "shape")  // Returns: true
 * isElementType("/slide[1]/shape[2]", "table")  // Returns: false
 */
export function isElementType(path: string, type: PptElementType): boolean {
  return getElementType(path) === type;
}
