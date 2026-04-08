/**
 * Layout checking operations for @officekit/ppt.
 *
 * Provides comprehensive functions to detect and report layout issues:
 * - checkPresentation: Scan entire presentation for layout issues
 * - checkShapeTextOverflow: Check if text overflows a specific shape's bounds
 * - checkSlide: Check all shapes on a slide for layout issues
 * - formatCheckReport: Format check results as human-readable report
 *
 * Reference: OfficeCli's CheckShapeTextOverflow implementation
 */

import { readFile } from "node:fs/promises";
import path from "node:path";
import { readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput, notFound } from "./result.js";
import type { Result, ShapeModel } from "./types.js";
import { getSlideIndex } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Issue severity levels.
 */
export type IssueSeverity = "info" | "warning" | "error";

/**
 * Issue categories for layout problems.
 */
export type IssueCategory =
  | "text_overflow"
  | "text_truncation"
  | "shape_too_small"
  | "missing_title"
  | "empty_slide"
  | "layout_mismatch";

/**
 * A detected layout issue.
 */
export interface LayoutIssue {
  /** Unique identifier for this issue */
  id: string;
  /** Issue severity */
  severity: IssueSeverity;
  /** Issue category */
  category: IssueCategory;
  /** Human-readable message */
  message: string;
  /** Path to the affected element */
  path: string;
  /** Suggested fix if available */
  suggestion?: string;
  /** Additional details about the issue */
  details?: {
    /** Shape name if applicable */
    shapeName?: string;
    /** Text length if relevant */
    textLength?: number;
    /** Shape dimensions if relevant */
    shapeWidth?: number;
    shapeHeight?: number;
    /** Estimated overflow amount in EMUs */
    overflowAmount?: number;
    /** Font size in points if relevant */
    fontSize?: number;
    /** Number of lines if relevant */
    lineCount?: number;
    /** Required height/width if relevant */
    requiredSize?: number;
  };
}

/**
 * Result from checking a presentation.
 */
export interface CheckPresentationResult {
  /** Absolute file path */
  filePath: string;
  /** Total slides checked */
  slideCount: number;
  /** Total shapes checked */
  shapeCount: number;
  /** Total issues found */
  issueCount: number;
  /** Issues grouped by severity */
  issuesBySeverity: {
    error: number;
    warning: number;
    info: number;
  };
  /** Issues grouped by category */
  issuesByCategory: Record<IssueCategory, number>;
  /** All detected issues */
  issues: LayoutIssue[];
  /** Whether any issues were found */
  hasIssues: boolean;
}

/**
 * Result from checking a slide.
 */
export interface CheckSlideResult {
  /** Slide index (1-based) */
  slideIndex: number;
  /** Slide path */
  path: string;
  /** Total shapes on slide */
  shapeCount: number;
  /** Issues found on this slide */
  issueCount: number;
  /** Whether any issues were detected */
  hasIssues: boolean;
  /** All detected issues */
  issues: LayoutIssue[];
}

/**
 * Result from checking a shape.
 */
export interface CheckShapeResult {
  /** Shape path */
  path: string;
  /** Shape name */
  name?: string;
  /** Whether any issues were detected */
  hasIssues: boolean;
  /** All detected issues */
  issues: LayoutIssue[];
}

/**
 * Options for checking a presentation.
 */
export interface CheckOptions {
  /** Check text overflow issues (default: true) */
  checkTextOverflow?: boolean;
  /** Check for missing titles (default: true) */
  checkMissingTitles?: boolean;
  /** Check for empty slides (default: true) */
  checkEmptySlides?: boolean;
  /** Severity threshold for reporting (default: 'info') */
  minSeverity?: IssueSeverity;
  /** Slide index to check (optional, checks all slides if not specified) */
  slideIndex?: number;
}

// ============================================================================
// Constants
// ============================================================================

/** EMUs per point (1 pt = 12700 EMUs) */
const EMU_PER_PT = 12700.0;

/** Default left/right inset in EMUs (0.1 inch = 91440 EMUs) */
const DEFAULT_LR_INSET = 91440;

/** Default top/bottom inset in EMUs (0.05 inch = 45720 EMUs) */
const DEFAULT_TB_INSET = 45720;

/** Default font size for textboxes in points */
const DEFAULT_FONT_SIZE_PT = 18.0;

/** Tolerance percentage for overflow detection */
const OVERFLOW_TOLERANCE = 0.05; // 5%

// ============================================================================
// Helpers
// ============================================================================

/**
 * Parses relationship entries from a .rels XML string.
 */
function parseRelationshipEntries(xml: string): Array<{ id: string; target: string; type?: string }> {
  const relationships: Array<{ id: string; target: string; type?: string }> = [];
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const attributes = match[1];
    const id = /Id="([^"]+)"/.exec(attributes)?.[1];
    const target = /Target="([^"]+)"/.exec(attributes)?.[1];
    const type = /Type="([^"]+)"/.exec(attributes)?.[1];
    if (id && target) {
      relationships.push({ id, target, type });
    }
  }
  return relationships;
}

/**
 * Normalizes a zip path relative to a base directory.
 */
function normalizeZipPath(baseDir: string, target: string): string {
  const normalized = target.replace(/\\/g, "/");
  if (normalized.startsWith("/")) {
    return path.posix.normalize(normalized.slice(1));
  }
  return path.posix.normalize(path.posix.join(baseDir, normalized));
}

/**
 * Reads an entry from the zip as a string.
 */
function requireEntry(zip: Map<string, Buffer>, entryName: string): string {
  const buffer = zip.get(entryName);
  if (!buffer) {
    throw new Error(`OOXML entry '${entryName}' is missing`);
  }
  return buffer.toString("utf8");
}

/**
 * Loads a presentation and returns its zip contents.
 */
async function loadPresentation(filePath: string): Promise<Result<Map<string, Buffer>>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);
    return ok(zip);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets the slide IDs from presentation.xml.
 */
function getSlideIds(presentationXml: string): Array<{ id: string; relId: string }> {
  const slideIds: Array<{ id: string; relId: string }> = [];
  for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*\bid="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g)) {
    slideIds.push({ id: match[1], relId: match[2] });
  }
  for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"[^>]*\bid="([^"]+)"[^>]*\/?>/g)) {
    const relId = match[1];
    const id = match[2];
    if (!slideIds.some(s => s.relId === relId)) {
      slideIds.push({ id, relId });
    }
  }
  return slideIds;
}

/**
 * Gets the slide entry path from the zip by slide index.
 */
function getSlideEntryPath(zip: Map<string, Buffer>, slideIndex: number): Result<string> {
  const presentationXml = requireEntry(zip, "ppt/presentation.xml");
  const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
  const relationships = parseRelationshipEntries(relsXml);
  const slideIds = getSlideIds(presentationXml);

  if (slideIndex < 1 || slideIndex > slideIds.length) {
    return invalidInput(`Slide index ${slideIndex} is out of range (1-${slideIds.length})`);
  }

  const slide = slideIds[slideIndex - 1];
  const slideRel = relationships.find(r => r.id === slide.relId);
  const slidePath = normalizeZipPath("ppt", slideRel?.target ?? "");

  return ok(slidePath);
}

/**
 * Gets all slide entries with their info.
 */
function getAllSlideEntries(zip: Map<string, Buffer>): Result<Array<{ index: number; path: string; entryPath: string }>> {
  try {
    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const slideIds = getSlideIds(presentationXml);

    const slides = slideIds.map((slide, idx) => {
      const slideRel = relationships.find(r => r.id === slide.relId);
      const entryPath = normalizeZipPath("ppt", slideRel?.target ?? "");
      return {
        index: idx + 1,
        path: `/slide[${idx + 1}]`,
        entryPath,
      };
    });

    return ok(slides);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Extracts shape name from shape XML.
 */
function extractShapeName(shapeXml: string): string | undefined {
  const nameMatch = shapeXml.match(/<p:nvCxnSpPr[^>]*>[\s\S]*?<p:cNvPr[^>]*name="([^"]*)"[^>]*>/);
  if (nameMatch) return nameMatch[1];

  const altMatch = shapeXml.match(/<p:cNvPr[^>]*name="([^"]*)"[^>]*>/);
  return altMatch ? altMatch[1] : undefined;
}

/**
 * Extracts text from a shape.
 */
function extractTextFromShape(shapeXml: string): string {
  const textRuns: string[] = [];
  for (const match of shapeXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) {
    textRuns.push(match[1]);
  }
  return textRuns.join("");
}

/**
 * Extracts shape properties (position, size) from shape XML.
 */
function extractShapeProperties(shapeXml: string): {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
} {
  const props: { x?: number; y?: number; width?: number; height?: number } = {};

  const spPrMatch = shapeXml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);
  if (spPrMatch) {
    const spPrContent = spPrMatch[1];

    const xfrmMatch = spPrContent.match(/<a:xfrm(?:[^>]*)>([\s\S]*?)<\/a:xfrm>/);
    if (xfrmMatch) {
      const xfrmContent = xfrmMatch[1];

      const offMatch = xfrmContent.match(/<a:off[^>]*x="([^"]*)"[^>]*y="([^"]*)"[^>]*>/);
      if (offMatch) {
        props.x = parseInt(offMatch[1], 10);
        props.y = parseInt(offMatch[2], 10);
      }

      const extMatch = xfrmContent.match(/<a:ext[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"[^>]*>/);
      if (extMatch) {
        props.width = parseInt(extMatch[1], 10);
        props.height = parseInt(extMatch[2], 10);
      }
    }
  }

  return props;
}

/**
 * Extracts body properties (margins) from shape XML.
 */
function extractBodyProperties(shapeXml: string): {
  leftInset?: number;
  rightInset?: number;
  topInset?: number;
  bottomInset?: number;
} {
  const props: { leftInset?: number; rightInset?: number; topInset?: number; bottomInset?: number } = {};

  const txBodyMatch = shapeXml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
  if (txBodyMatch) {
    const txBodyContent = txBodyMatch[1];
    const bpMatch = txBodyContent.match(/<a:bodyPr([^>]*)\/?>/);
    if (bpMatch) {
      const attrs = bpMatch[1];
      const leftMatch = attrs.match(/l="([^"]*)"/);
      const rightMatch = attrs.match(/r="([^"]*)"/);
      const topMatch = attrs.match(/t="([^"]*)"/);
      const bottomMatch = attrs.match(/b="([^"]*)"/);

      if (leftMatch) props.leftInset = parseInt(leftMatch[1], 10);
      if (rightMatch) props.rightInset = parseInt(rightMatch[1], 10);
      if (topMatch) props.topInset = parseInt(topMatch[1], 10);
      if (bottomMatch) props.bottomInset = parseInt(bottomMatch[1], 10);
    }
  }

  return props;
}

/**
 * Extracts font size from shape XML (from run properties or body properties).
 */
function extractFontSize(shapeXml: string): number {
  // Check run properties first
  const runFontMatch = shapeXml.match(/<a:rPr[^>]*sz="([^"]*)"[^>]*>/);
  if (runFontMatch) {
    return parseInt(runFontMatch[1], 10) / 100.0; // Font size is in hundredths of points
  }

  // Check body properties default font size
  const bodyFontMatch = shapeXml.match(/<a:bodyPr[^>]*fontsize="([^"]*)"[^>]*>/);
  if (bodyFontMatch) {
    return parseInt(bodyFontMatch[1], 10) / 100.0;
  }

  return DEFAULT_FONT_SIZE_PT;
}

/**
 * Checks if a character is CJK or full-width.
 * Reference: OfficeCli's ParseHelpers.IsCjkOrFullWidth
 */
function isCjkOrFullWidth(char: string): boolean {
  const code = char.charCodeAt(0);

  // CJK Unified Ideographs (4E00-9FFF)
  if (code >= 0x4E00 && code <= 0x9FFF) return true;

  // CJK Compatibility Ideographs (F900-FAFF)
  if (code >= 0xF900 && code <= 0xFAFF) return true;

  // Halfwidth and Fullwidth Forms (FF00-FFEF)
  if (code >= 0xFF00 && code <= 0xFFEF) return true;

  // Hiragana (3040-309F)
  if (code >= 0x3040 && code <= 0x309F) return true;

  // Katakana (30A0-30FF)
  if (code >= 0x30A0 && code <= 0x30FF) return true;

  // Hangul Compatibility Jamo (3130-318F)
  if (code >= 0x3130 && code <= 0x318F) return true;

  // Hangul Syllables (AC00-D7AF)
  if (code >= 0xAC00 && code <= 0xD7AF) return true;

  //CJK Unified Ideographs Extension A (3400-4DBF)
  if (code >= 0x3400 && code <= 0x4DBF) return true;

  return false;
}

/**
 * Generates a unique issue ID.
 */
function generateIssueId(): string {
  return `issue_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
}

/**
 * Parses a shape from XML.
 */
function parseShapeFromXml(shapeXml: string, slideIndex: number, shapeIndex: number): ShapeModel & {
  bodyProps: ReturnType<typeof extractBodyProperties>;
  fontSize: number;
} {
  const shapePathStr = `/slide[${slideIndex}]/shape[${shapeIndex}]`;

  const name = extractShapeName(shapeXml);
  const text = extractTextFromShape(shapeXml);
  const props = extractShapeProperties(shapeXml);
  const bodyProps = extractBodyProperties(shapeXml);
  const fontSize = extractFontSize(shapeXml);

  return {
    path: shapePathStr,
    name,
    text,
    type: "shape",
    x: props.x,
    y: props.y,
    width: props.width,
    height: props.height,
    bodyProps,
    fontSize,
  };
}

/**
 * Parses all shapes from slide XML.
 */
function parseShapesFromSlideXml(
  slideXml: string,
  slideIndex: number
): Array<ShapeModel & { bodyProps: ReturnType<typeof extractBodyProperties>; fontSize: number }> {
  const shapes: Array<ShapeModel & { bodyProps: ReturnType<typeof extractBodyProperties>; fontSize: number }> = [];

  const shapePattern = /<p:sp(?:[\s\S]*?)<\/p:sp>/g;
  let shapeIndex = 0;

  for (const shapeMatch of slideXml.matchAll(shapePattern)) {
    shapeIndex++;
    const shapeXml = shapeMatch[0];
    shapes.push(parseShapeFromXml(shapeXml, slideIndex, shapeIndex));
  }

  return shapes;
}

/**
 * Checks text overflow for a shape using sophisticated estimation.
 * Reference: OfficeCli's CheckTextOverflow function
 */
function checkTextOverflowForShape(
  shape: ShapeModel & { bodyProps: ReturnType<typeof extractBodyProperties>; fontSize: number }
): LayoutIssue | null {
  const text = shape.text;
  if (!text || !shape.width || !shape.height) {
    return null;
  }

  // Get shape dimensions in points
  const shapeWidthPt = shape.width / EMU_PER_PT;
  const shapeHeightPt = shape.height / EMU_PER_PT;

  // Get margins (with defaults)
  const leftEmu = shape.bodyProps.leftInset ?? DEFAULT_LR_INSET;
  const rightEmu = shape.bodyProps.rightInset ?? DEFAULT_LR_INSET;
  const topEmu = shape.bodyProps.topInset ?? DEFAULT_TB_INSET;
  const bottomEmu = shape.bodyProps.bottomInset ?? DEFAULT_TB_INSET;

  // Calculate usable dimensions
  const marginPt = (topEmu + bottomEmu) / EMU_PER_PT;
  const usableWidth = shapeWidthPt - (leftEmu + rightEmu) / EMU_PER_PT;
  const usableHeight = shapeHeightPt - (topEmu + bottomEmu) / EMU_PER_PT;

  // Check if shape is too small for its margins
  if (usableWidth <= 0 || usableHeight <= 0) {
    const defaultLinePt = 18.0;
    const needPt = marginPt + defaultLinePt;
    const minHeightCm = (needPt / 72.0) * 2.54;
    const minHeightEmu = Math.round(minHeightCm * 360000.0);

    return {
      id: generateIssueId(),
      severity: "error",
      category: "shape_too_small",
      message: `Text box is too narrow: need ${needPt.toFixed(0)}pt for margins + 1 line, but usable area is 0pt`,
      path: shape.path,
      suggestion: `Try expanding the text box height to at least ${minHeightEmu} EMUs`,
      details: {
        shapeName: shape.name,
        shapeWidth: shape.width,
        shapeHeight: shape.height,
        overflowAmount: Math.round((needPt - shapeHeightPt) * EMU_PER_PT),
      },
    };
  }

  const fontSizePt = shape.fontSize || DEFAULT_FONT_SIZE_PT;

  // Estimate text width per line using per-character measurement
  const textLines = text.replace(/\\n/g, "\n").split("\n");
  let totalLines = 0;

  for (const line of textLines) {
    if (line.length === 0) {
      totalLines += 1;
      continue;
    }

    // Walk characters, accumulate width, wrap when exceeding usable width
    let linesForSegment = 1;
    let currentLineWidth = 0;

    for (const char of line) {
      // CJK characters are approximately full-width (same as font size)
      // Latin characters are approximately 0.55 times font size
      const charWidth = isCjkOrFullWidth(char) ? fontSizePt : fontSizePt * 0.55;

      if (currentLineWidth + charWidth > usableWidth && currentLineWidth > 0) {
        linesForSegment++;
        currentLineWidth = charWidth;
      } else {
        currentLineWidth += charWidth;
      }
    }

    totalLines += linesForSegment;
  }

  // Estimate total height needed
  const estimatedHeight = totalLines * fontSizePt;

  // Check for overflow with tolerance
  if (estimatedHeight > usableHeight * (1 + OVERFLOW_TOLERANCE)) {
    // Calculate minimum height
    const minHeightCm = (estimatedHeight + marginPt) / 72.0 * 2.54;
    const minHeightEmu = Math.round(Math.ceil(minHeightCm * 20) / 20.0 * 360000.0);

    const overflowPercent = Math.round((estimatedHeight / usableHeight) * 100 - 100);

    return {
      id: generateIssueId(),
      severity: overflowPercent > 50 ? "error" : overflowPercent > 20 ? "warning" : "info",
      category: overflowPercent > 100 ? "text_truncation" : "text_overflow",
      message: `Text overflow: ${totalLines} lines at ${fontSizePt.toFixed(1)}pt need ${estimatedHeight.toFixed(0)}pt, but only ${usableHeight.toFixed(0)}pt available (${overflowPercent}% overflow)`,
      path: shape.path,
      suggestion: `Expand the text box height to at least ${minHeightEmu} EMUs, or reduce font size or text content`,
      details: {
        shapeName: shape.name,
        textLength: text.length,
        shapeWidth: shape.width,
        shapeHeight: shape.height,
        overflowAmount: Math.round((estimatedHeight - usableHeight) * EMU_PER_PT),
        fontSize: fontSizePt,
        lineCount: totalLines,
        requiredSize: minHeightEmu,
      },
    };
  }

  return null;
}

/**
 * Checks for missing title on a slide.
 */
function checkMissingTitle(
  slideXml: string,
  slideIndex: number,
  shapes: Array<ShapeModel & { bodyProps: ReturnType<typeof extractBodyProperties>; fontSize: number }>
): LayoutIssue | null {
  // Look for a title placeholder
  const hasTitle = shapes.some(
    (s) =>
      s.placeholderType !== undefined ||
      (s.name && /title/i.test(s.name)) ||
      (s.text && s.path.includes("placeholder"))
  );

  if (!hasTitle && shapes.length > 0) {
    return {
      id: generateIssueId(),
      severity: "info",
      category: "missing_title",
      message: `Slide ${slideIndex} does not have a title placeholder`,
      path: `/slide[${slideIndex}]`,
      suggestion: "Consider adding a title placeholder to improve slide structure",
    };
  }

  return null;
}

/**
 * Checks if a slide is empty.
 */
function checkEmptySlide(
  slideXml: string,
  slideIndex: number,
  shapes: Array<ShapeModel & { bodyProps: ReturnType<typeof extractBodyProperties>; fontSize: number }>
): LayoutIssue | null {
  // Check if slide has any text content
  const hasContent = shapes.some((s) => s.text && s.text.trim().length > 0);

  if (!hasContent && shapes.length > 0) {
    return {
      id: generateIssueId(),
      severity: "warning",
      category: "empty_slide",
      message: `Slide ${slideIndex} has shapes but no text content`,
      path: `/slide[${slideIndex}]`,
      suggestion: "Consider adding text content or removing empty shapes",
    };
  }

  return null;
}

// ============================================================================
// Public API
// ============================================================================

/**
 * Checks a specific shape for layout issues.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @returns Result with shape check
 *
 * @example
 * const result = await checkShape("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 * if (result.ok && result.data.hasIssues) {
 *   console.log(`Issues found: ${result.data.issues}`);
 * }
 */
export async function checkShape(
  filePath: string,
  pptPath: string
): Promise<Result<CheckShapeResult>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("checkShape requires a slide path like /slide[1]/shape[1]");
  }

  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error!.code, zipResult.error!.message);
  }
  const zip = zipResult.data!;

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error!.code, slidePathResult.error!.message);
  }

  const slideEntry = slidePathResult.data!;
  const slideXml = requireEntry(zip, slideEntry);

  // Extract shape index
  const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
  if (!shapeIndexMatch) {
    return invalidInput("Invalid shape path");
  }
  const shapeIndex = parseInt(shapeIndexMatch[1], 10);

  // Parse shapes and find the target
  const shapes = parseShapesFromSlideXml(slideXml, slideIndex);
  const shape = shapes.find((s) => s.path === pptPath);

  if (!shape) {
    return notFound("Shape", String(shapeIndex), `Shape not found at path ${pptPath}`);
  }

  const issues: LayoutIssue[] = [];

  // Check text overflow
  const overflowIssue = checkTextOverflowForShape(shape);
  if (overflowIssue) {
    issues.push(overflowIssue);
  }

  return ok({
    path: pptPath,
    name: shape.name,
    hasIssues: issues.length > 0,
    issues,
  });
}

/**
 * Checks all shapes on a slide for layout issues.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param options - Check options
 * @returns Result with slide check
 *
 * @example
 * const result = await checkSlide("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(`Slide ${result.data.slideIndex} has ${result.data.issueCount} issues`);
 * }
 */
export async function checkSlide(
  filePath: string,
  slideIndex: number,
  options: CheckOptions = {}
): Promise<Result<CheckSlideResult>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error!.code, zipResult.error!.message);
  }
  const zip = zipResult.data!;

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error!.code, slidePathResult.error!.message);
  }

  const slideEntry = slidePathResult.data!;
  const slideXml = requireEntry(zip, slideEntry);

  const shapes = parseShapesFromSlideXml(slideXml, slideIndex);
  const issues: LayoutIssue[] = [];

  // Check each shape for text overflow
  if (options.checkTextOverflow !== false) {
    for (const shape of shapes) {
      const overflowIssue = checkTextOverflowForShape(shape);
      if (overflowIssue) {
        issues.push(overflowIssue);
      }
    }
  }

  // Check for missing title
  if (options.checkMissingTitles !== false) {
    const titleIssue = checkMissingTitle(slideXml, slideIndex, shapes);
    if (titleIssue) {
      issues.push(titleIssue);
    }
  }

  // Check for empty slide
  if (options.checkEmptySlides !== false) {
    const emptyIssue = checkEmptySlide(slideXml, slideIndex, shapes);
    if (emptyIssue) {
      issues.push(emptyIssue);
    }
  }

  return ok({
    slideIndex,
    path: `/slide[${slideIndex}]`,
    shapeCount: shapes.length,
    issueCount: issues.length,
    hasIssues: issues.length > 0,
    issues,
  });
}

/**
 * Checks an entire presentation for layout issues.
 *
 * @param filePath - Path to the PPTX file
 * @param options - Check options
 * @returns Result with comprehensive check report
 *
 * @example
 * const result = await checkPresentation("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(`Found ${result.data.issueCount} issues across ${result.data.slideCount} slides`);
 *   for (const issue of result.data.issues) {
 *     console.log(`[${issue.severity}] ${issue.path}: ${issue.message}`);
 *   }
 * }
 */
export async function checkPresentation(
  filePath: string,
  options: CheckOptions = {}
): Promise<Result<CheckPresentationResult>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error!.code, zipResult.error!.message);
  }
  const zip = zipResult.data!;

  const slidesInfoResult = getAllSlideEntries(zip);
  if (!slidesInfoResult.ok) {
    return err(slidesInfoResult.error!.code, slidesInfoResult.error!.message);
  }
  const slidesInfo = slidesInfoResult.data!;

  // Filter to specific slide if requested
  const targetSlides = options.slideIndex
    ? slidesInfo.filter((s) => s.index === options.slideIndex)
    : slidesInfo;

  if (options.slideIndex && targetSlides.length === 0) {
    return invalidInput(
      `Slide index ${options.slideIndex} is out of range (1-${slidesInfo.length})`
    );
  }

  const allIssues: LayoutIssue[] = [];
  let totalShapeCount = 0;

  // Check each slide
  for (const slideInfo of targetSlides) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);
    const shapes = parseShapesFromSlideXml(slideXml, slideInfo.index);
    totalShapeCount += shapes.length;

    // Check each shape for text overflow
    if (options.checkTextOverflow !== false) {
      for (const shape of shapes) {
        const overflowIssue = checkTextOverflowForShape(shape);
        if (overflowIssue) {
          allIssues.push(overflowIssue);
        }
      }
    }

    // Check for missing title
    if (options.checkMissingTitles !== false) {
      const titleIssue = checkMissingTitle(slideXml, slideInfo.index, shapes);
      if (titleIssue) {
        allIssues.push(titleIssue);
      }
    }

    // Check for empty slide
    if (options.checkEmptySlides !== false) {
      const emptyIssue = checkEmptySlide(slideXml, slideInfo.index, shapes);
      if (emptyIssue) {
        allIssues.push(emptyIssue);
      }
    }
  }

  // Filter by severity if needed
  const severityOrder: IssueSeverity[] = ["info", "warning", "error"];
  const minSeverityIndex = options.minSeverity
    ? severityOrder.indexOf(options.minSeverity)
    : 0;

  const filteredIssues = allIssues.filter((issue) => {
    const issueSeverityIndex = severityOrder.indexOf(issue.severity);
    return issueSeverityIndex >= minSeverityIndex;
  });

  // Count by severity
  const issuesBySeverity = {
    error: filteredIssues.filter((i) => i.severity === "error").length,
    warning: filteredIssues.filter((i) => i.severity === "warning").length,
    info: filteredIssues.filter((i) => i.severity === "info").length,
  };

  // Count by category
  const issuesByCategory: Record<IssueCategory, number> = {
    text_overflow: 0,
    text_truncation: 0,
    shape_too_small: 0,
    missing_title: 0,
    empty_slide: 0,
    layout_mismatch: 0,
  };

  for (const issue of filteredIssues) {
    issuesByCategory[issue.category]++;
  }

  return ok({
    filePath,
    slideCount: targetSlides.length,
    shapeCount: totalShapeCount,
    issueCount: filteredIssues.length,
    issuesBySeverity,
    issuesByCategory,
    issues: filteredIssues,
    hasIssues: filteredIssues.length > 0,
  });
}

/**
 * Formats a check result as a human-readable report.
 *
 * @param result - The check result to format
 * @param options - Formatting options
 * @returns Formatted report string
 *
 * @example
 * const result = await checkPresentation("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(formatCheckReport(result.data));
 * }
 */
export function formatCheckReport(
  result: CheckPresentationResult,
  options: { verbose?: boolean; json?: boolean } = {}
): string {
  if (options.json) {
    return JSON.stringify(result, null, 2);
  }

  const lines: string[] = [];

  lines.push(`Checking layout: ${result.filePath}`);
  lines.push(`Scanned ${result.slideCount} slide(s) and ${result.shapeCount} shape(s)`);

  if (result.issueCount === 0) {
    lines.push("No layout issues found.");
    return lines.join("\n");
  }

  lines.push(`Found ${result.issueCount} layout issue(s):`);

  // Group issues by severity
  if (result.issuesBySeverity.error > 0) {
    lines.push(`  [ERROR] ${result.issuesBySeverity.error} issue(s)`);
  }
  if (result.issuesBySeverity.warning > 0) {
    lines.push(`  [WARNING] ${result.issuesBySeverity.warning} issue(s)`);
  }
  if (result.issuesBySeverity.info > 0) {
    lines.push(`  [INFO] ${result.issuesBySeverity.info} issue(s)`);
  }

  lines.push("");

  // List each issue
  for (const issue of result.issues) {
    const severityTag = `[${issue.severity.toUpperCase()}]`;
    lines.push(`  ${severityTag} ${issue.path}: ${issue.message}`);

    if (options.verbose && issue.suggestion) {
      lines.push(`      Suggestion: ${issue.suggestion}`);
    }

    if (options.verbose && issue.details) {
      const details: string[] = [];
      if (issue.details.shapeName) details.push(`name="${issue.details.shapeName}"`);
      if (issue.details.textLength !== undefined)
        details.push(`textLength=${issue.details.textLength}`);
      if (issue.details.shapeWidth !== undefined && issue.details.shapeHeight !== undefined) {
        details.push(`size=${issue.details.shapeWidth}x${issue.details.shapeHeight}`);
      }
      if (issue.details.fontSize !== undefined) details.push(`fontSize=${issue.details.fontSize}pt`);
      if (issue.details.lineCount !== undefined) details.push(`lines=${issue.details.lineCount}`);
      if (issue.details.overflowAmount !== undefined)
        details.push(`overflow=${issue.details.overflowAmount} EMUs`);

      if (details.length > 0) {
        lines.push(`      Details: ${details.join(", ")}`);
      }
    }
  }

  return lines.join("\n");
}

/**
 * Re-export overflow checking functions for backward compatibility.
 */
export {
  checkShapeTextOverflow,
  checkSlideOverflow,
  getOverflowIssues,
  type ShapeOverflowResult,
  type SlideOverflowResult,
  type OverflowIssue,
  type OverflowIssuesResult,
} from "./overflow.js";
