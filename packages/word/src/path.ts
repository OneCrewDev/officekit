/**
 * Path parsing and resolution for @officekit/word.
 *
 * Word paths are absolute paths that identify elements within a document:
 * - `/` - the root document
 * - `/body` - the document body
 * - `/body/p[N]` - the Nth paragraph
 * - `/body/p[@paraId=XXX]` - paragraph by stable ID
 * - `/body/tbl[N]` - the Nth table
 * - `/body/tbl[N]/tr[N]` - the Nth row of table N
 * - `/body/tbl[N]/tr[N]/tc[N]` - the Nth cell of row N
 * - `/body/r[N]` - the Nth run
 * - `/header[N]` - the Nth header
 * - `/footer[N]` - the Nth footer
 * - `/footnote[N]` or `/footnote[@footnoteId=N]` - footnote by index or ID
 * - `/endnote[N]` or `/endnote[@endnoteId=N]` - endnote by index or ID
 * - `/styles` - style definitions
 * - `/styles/StyleId` - a specific style
 * - `/numbering` - numbering definitions
 * - `/settings` - document settings
 * - `/comments` - comments
 * - `/toc[N]` - the Nth TOC
 * - `/field[N]` - the Nth field
 * - `/formfield[N]` or `/formfield[name]` - form field by index or name
 * - `/section[N]` - the Nth section
 * - `/chart[N]` - the Nth chart
 * - `/watermark` - watermark text
 * - `/bookmark[Name]` - bookmark by name
 *
 * Syntax:
 * - `/element[index]` - indexed access (1-indexed)
 * - `/element[@attr=value]` - stable ID access
 * - `/element[name]` - named access
 *
 * @example
 * parsePath("/body/p[1]")
 * // Returns: { isAbsolute: true, segments: [{ name: "body" }, { name: "p", index: 1 }], original: "/body/p[1]" }
 */

import { err, ok, andThen } from "./result.js";
import type { ParsedPath, PathSegment, Result } from "./types.js";

// ============================================================================
// Path Pattern Constants
// ============================================================================

const SEGMENT_PATTERNS = {
  INDEXED: /^\/([a-zA-Z]+)\[(\d+)\](.*)$/,
  NAMED: /^\/([a-zA-Z]+)\[([a-zA-Z][a-zA-Z0-9_]*)\](.*)$/,
  ATTRIBUTE: /^\/([a-zA-Z]+)\[@([a-zA-Z]+)=([^\]]+)\](.*)$/,
} as const;

const VALID_ROOT_SEGMENTS = new Set([
  "body",
  "styles",
  "header",
  "footer",
  "numbering",
  "settings",
  "comments",
  "footnote",
  "endnote",
  "toc",
  "field",
  "formfield",
  "section",
  "chart",
  "watermark",
  "bookmark",
]);

const VALID_BODY_CHILDREN = new Set([
  "p",
  "paragraph",
  "tbl",
  "table",
  "r",
  "run",
  "sdt",
  "oMathPara",
  "hyperlink",
]);

const VALID_TABLE_CHILDREN = new Set(["tr", "row"]);
const VALID_ROW_CHILDREN = new Set(["tc", "cell"]);

const PATH_ALIASES: Record<string, string> = {
  paragraph: "p",
  table: "tbl",
  row: "tr",
  cell: "tc",
};

// ============================================================================
// Path Parsing
// ============================================================================

/**
 * Parses a Word path string into a structured ParsedPath.
 */
export function parsePath(path: string): Result<ParsedPath> {
  if (!path || typeof path !== "string") {
    return err("invalid_path", "Path must be a non-empty string");
  }

  const original = path;
  const isAbsolute = path.startsWith("/");

  if (!isAbsolute) {
    return err("invalid_path", "Word paths must be absolute (start with /)");
  }

  if (path === "/") {
    return ok({ isAbsolute: true, segments: [], original });
  }

  const segments: PathSegment[] = [];
  let remaining = path;

  while (remaining.length > 0 && remaining !== "/") {
    const segmentResult = parseSegment(remaining);
    if (!segmentResult.ok) {
      return err(segmentResult.error?.code ?? "invalid_path", segmentResult.error?.message ?? "Failed to parse segment");
    }

    const { segment, rest } = segmentResult.data as { segment: PathSegment; rest: string };
    segments.push(segment);
    remaining = rest;
  }

  if (segments.length === 0) {
    return err("invalid_path", "Path contains no valid segments");
  }

  return ok({ isAbsolute, segments, original });
}

/**
 * Parses a single path segment from the beginning of a path string.
 */
function parseSegment(path: string): Result<{ segment: PathSegment; rest: string }> {
  let match = path.match(SEGMENT_PATTERNS.INDEXED);
  if (match) {
    const name = normalizeElementName(match[1]);
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

  match = path.match(SEGMENT_PATTERNS.ATTRIBUTE);
  if (match) {
    const name = normalizeElementName(match[1]);
    const attrName = match[2];
    const attrValue = match[3];
    const rest = match[4] || "";

    if (attrName === "paraId" || attrName === "textId" || attrName === "commentId" || attrName === "sdtId" || attrName === "footnoteId" || attrName === "endnoteId") {
      return ok({
        segment: { name, stringIndex: `@${attrName}=${attrValue}` },
        rest,
      });
    }

    return ok({
      segment: { name, stringIndex: attrValue },
      rest,
    });
  }

  match = path.match(SEGMENT_PATTERNS.NAMED);
  if (match) {
    const name = normalizeElementName(match[1]);
    const nameSelector = match[2];
    const rest = match[3] || "";

    return ok({
      segment: { name, stringIndex: nameSelector },
      rest,
    });
  }

  const simpleMatch = path.match(/^\/([a-zA-Z]+)(.*)$/);
  if (simpleMatch) {
    const name = normalizeElementName(simpleMatch[1]);
    const rest = simpleMatch[2] || "";
    return ok({
      segment: { name },
      rest,
    });
  }

  return err("invalid_path", `Invalid path segment: ${path.slice(0, 20)}...`);
}

/**
 * Normalizes element names to their canonical form.
 */
function normalizeElementName(name: string): string {
  const lower = name.toLowerCase();
  return PATH_ALIASES[lower] || lower;
}

// ============================================================================
// Path Building
// ============================================================================

/**
 * Builds a path string from segments.
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
      if (seg.stringIndex !== undefined) {
        if (seg.stringIndex.startsWith("@")) {
          return `/${seg.name}[${seg.stringIndex}]`;
        }
        return `/${seg.name}[${seg.stringIndex}]`;
      }
      return `/${seg.name}`;
    })
    .join("");
}

/**
 * Builds a body path.
 */
export function bodyPath(): string {
  return "/body";
}

/**
 * Builds a paragraph path.
 */
export function paragraphPath(index: number): string {
  return `/body/p[${index}]`;
}

/**
 * Builds a paragraph path using stable ID.
 */
export function paragraphPathById(paraId: string): string {
  return `/body/p[@paraId=${paraId}]`;
}

/**
 * Builds a table path.
 */
export function tablePath(index: number): string {
  return `/body/tbl[${index}]`;
}

/**
 * Builds a table row path.
 */
export function tableRowPath(tableIndex: number, rowIndex: number): string {
  return `/body/tbl[${tableIndex}]/tr[${rowIndex}]`;
}

/**
 * Builds a table cell path.
 */
export function tableCellPath(tableIndex: number, rowIndex: number, cellIndex: number): string {
  return `/body/tbl[${tableIndex}]/tr[${rowIndex}]/tc[${cellIndex}]`;
}

/**
 * Builds a run path.
 */
export function runPath(paraIndex: number, runIndex: number): string {
  return `/body/p[${paraIndex}]/r[${runIndex}]`;
}

/**
 * Builds a header path.
 */
export function headerPath(index: number): string {
  return `/header[${index}]`;
}

/**
 * Builds a footer path.
 */
export function footerPath(index: number): string {
  return `/footer[${index}]`;
}

// ============================================================================
// Path Queries
// ============================================================================

/**
 * Extracts the body path if present (strips leading segments to get /body).
 */
export function getBodyPath(path: string): string | null {
  if (path === "/body") return "/body";
  if (path.startsWith("/body/")) return "/body";
  if (path.startsWith("/body[")) return "/body";
  return null;
}

/**
 * Checks if a path refers to the document root.
 */
export function isDocumentRoot(path: string): boolean {
  return path === "/" || path === "";
}

/**
 * Checks if a path refers to the body.
 */
export function isBodyPath(path: string): boolean {
  return path === "/body";
}

/**
 * Checks if a path refers to a paragraph.
 */
export function isParagraphPath(path: string): boolean {
  return /\/p(\[@|$|\[)/.test(path) || /\/paragraph(\[@|$|\[)/.test(path);
}

/**
 * Checks if a path refers to a table.
 */
export function isTablePath(path: string): boolean {
  return /\/tbl(\[@|$|\[)/.test(path) || /\/table(\[@|$|\[)/.test(path);
}

/**
 * Checks if a path refers to a table cell.
 */
export function isCellPath(path: string): boolean {
  return /\/tc\[/.test(path) || /\/cell\[/.test(path);
}

/**
 * Checks if a path refers to a header.
 */
export function isHeaderPath(path: string): boolean {
  return /^\/header\[/.test(path);
}

/**
 * Checks if a path refers to a footer.
 */
export function isFooterPath(path: string): boolean {
  return /^\/footer\[/.test(path);
}

/**
 * Checks if a path refers to a footnote.
 */
export function isFootnotePath(path: string): boolean {
  return /^\/footnote\[/.test(path);
}

/**
 * Checks if a path refers to an endnote.
 */
export function isEndnotePath(path: string): boolean {
  return /^\/endnote\[/.test(path);
}

/**
 * Gets the parent path of a given path.
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
 */
export function lastSegmentName(path: string): string | null {
  const matches = path.match(/\/([a-zA-Z]+)\[/g);
  if (!matches || matches.length === 0) {
    return null;
  }
  const lastMatch = matches[matches.length - 1];
  const match = lastMatch.match(/\/([a-zA-Z]+)\[/);
  return match ? match[1].toLowerCase() : null;
}

/**
 * Gets the element type of the final segment in a path.
 */
export function getElementType(path: string): string | null {
  if (path === "/" || path === "") return "document";
  if (path === "/body") return "body";
  if (path === "/styles") return "styles";
  if (path === "/numbering") return "numbering";
  if (path === "/settings") return "settings";
  if (path === "/comments") return "comments";
  if (path === "/watermark") return "watermark";

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
 */
export function validatePath(path: string): Result<void> {
  return andThen(parsePath(path), () => ok(void 0));
}

/**
 * Checks if a path is valid without returning detailed error.
 */
export function isValidPath(path: string): boolean {
  return parsePath(path).ok;
}
