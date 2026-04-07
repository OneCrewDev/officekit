/**
 * Overflow checking operations for @officekit/ppt.
 *
 * Provides functions to detect and report text overflow issues in shapes:
 * - checkShapeTextOverflow: Check if text overflows a specific shape's bounds
 * - checkSlideOverflow: Check all shapes on a slide for overflow
 * - getOverflowIssues: Get all overflow issues in the presentation
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
 * Result from checking shape text overflow.
 */
export interface ShapeOverflowResult {
  /** Whether text overflow is detected */
  hasOverflow: boolean;
  /** Shape path */
  path: string;
  /** Expected text box width in EMUs */
  expectedWidth?: number;
  /** Actual text content width estimate */
  estimatedTextWidth?: number;
  /** Excess amount in EMUs */
  excessAmount?: number;
  /** Shape properties at time of check */
  shape?: {
    x?: number;
    y?: number;
    width?: number;
    height?: number;
  };
}

/**
 * Result from checking a slide for overflow.
 */
export interface SlideOverflowResult {
  /** Slide index (1-based) */
  slideIndex: number;
  /** Slide path */
  path: string;
  /** Whether any overflow was detected */
  hasOverflow: boolean;
  /** Shapes with overflow issues */
  overflowingShapes: ShapeOverflowResult[];
}

/**
 * Overflow issue with severity and suggestions.
 */
export interface OverflowIssue {
  /** Issue severity */
  severity: "info" | "warning" | "error";
  /** Issue category */
  category: "text_overflow" | "text_truncation" | "shape_too_small";
  /** Human-readable message */
  message: string;
  /** Path to the affected element */
  path: string;
  /** Suggested fix if available */
  suggestion?: string;
}

/**
 * Result from getting all overflow issues.
 */
export interface OverflowIssuesResult {
  /** Total slide count */
  slideCount: number;
  /** Total issue count */
  issueCount: number;
  /** All detected issues */
  issues: OverflowIssue[];
}

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
  const props: ReturnType<typeof extractShapeProperties> = {};

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
 * Parses a shape from XML.
 */
function parseShapeFromXml(shapeXml: string, slideIndex: number, shapeIndex: number): ShapeModel {
  const shapePathStr = `/slide[${slideIndex}]/shape[${shapeIndex}]`;

  const name = extractShapeName(shapeXml);
  const text = extractTextFromShape(shapeXml);
  const props = extractShapeProperties(shapeXml);

  return {
    path: shapePathStr,
    name,
    text,
    type: "shape",
    x: props.x,
    y: props.y,
    width: props.width,
    height: props.height,
  };
}

/**
 * Parses all shapes from slide XML.
 */
function parseShapesFromSlideXml(slideXml: string, slideIndex: number): ShapeModel[] {
  const shapes: ShapeModel[] = [];

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
 * Estimates if text might overflow a shape's text box.
 * Uses a heuristic based on character count and average char width.
 */
function estimateTextOverflow(shape: ShapeModel, text: string): boolean {
  if (!shape.width || !text) {
    return false;
  }

  // Average character width in EMUs (approximate for typical fonts)
  const avgCharWidth = 1000;
  const padding = 2000;

  const estimatedTextWidth = text.length * avgCharWidth + padding;
  return estimatedTextWidth > shape.width;
}

/**
 * Estimates the severity of overflow based on how much it exceeds the bounds.
 */
function estimateOverflowSeverity(shape: ShapeModel, text: string): "info" | "warning" | "error" {
  if (!shape.width || !text) {
    return "info";
  }

  const avgCharWidth = 1000;
  const padding = 2000;
  const estimatedTextWidth = text.length * avgCharWidth + padding;
  const overflowRatio = estimatedTextWidth / shape.width;

  if (overflowRatio > 2) {
    return "error";
  } else if (overflowRatio > 1.5) {
    return "warning";
  }
  return "info";
}

// ============================================================================
// Public API
// ============================================================================

/**
 * Checks if text overflows in a specific shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @returns Result with overflow check
 *
 * @example
 * const result = await checkShapeTextOverflow("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 * if (result.ok && result.data.hasOverflow) {
 *   console.log(`Text overflow detected: ${result.data.excessAmount} EMUs excess`);
 * }
 */
export async function checkShapeTextOverflow(
  filePath: string,
  pptPath: string,
): Promise<Result<ShapeOverflowResult>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("checkShapeTextOverflow requires a slide path");
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

  // Find the shape
  const shapes = parseShapesFromSlideXml(slideXml, slideIndex);
  const shape = shapes.find(s => s.path === pptPath);

  if (!shape) {
    return notFound("Shape", String(shapeIndex), `Shape not found at path ${pptPath}`);
  }

  if (!shape.text) {
    return ok({
      hasOverflow: false,
      path: pptPath,
      shape: {
        x: shape.x,
        y: shape.y,
        width: shape.width,
        height: shape.height,
      },
    });
  }

  if (!shape.width) {
    return ok({
      hasOverflow: false,
      path: pptPath,
      shape: {
        x: shape.x,
        y: shape.y,
        width: shape.width,
        height: shape.height,
      },
    });
  }

  // Estimate text width
  const avgCharWidth = 1000;
  const padding = 2000;
  const estimatedTextWidth = shape.text.length * avgCharWidth + padding;
  const hasOverflow = estimatedTextWidth > shape.width;

  return ok({
    hasOverflow,
    path: pptPath,
    expectedWidth: shape.width,
    estimatedTextWidth,
    excessAmount: hasOverflow ? estimatedTextWidth - shape.width : undefined,
    shape: {
      x: shape.x,
      y: shape.y,
      width: shape.width,
      height: shape.height,
    },
  });
}

/**
 * Checks all shapes on a slide for overflow issues.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with slide overflow check
 *
 * @example
 * const result = await checkSlideOverflow("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(`Slide ${result.data.slideIndex} has ${result.data.overflowingShapes.length} overflowing shapes`);
 * }
 */
export async function checkSlideOverflow(
  filePath: string,
  slideIndex: number,
): Promise<Result<SlideOverflowResult>> {
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
  const overflowingShapes: ShapeOverflowResult[] = [];

  for (const shape of shapes) {
    if (!shape.text || !shape.width) {
      continue;
    }

    const avgCharWidth = 1000;
    const padding = 2000;
    const estimatedTextWidth = shape.text.length * avgCharWidth + padding;
    const hasOverflow = estimatedTextWidth > shape.width;

    if (hasOverflow) {
      overflowingShapes.push({
        hasOverflow,
        path: shape.path,
        expectedWidth: shape.width,
        estimatedTextWidth,
        excessAmount: estimatedTextWidth - shape.width,
        shape: {
          x: shape.x,
          y: shape.y,
          width: shape.width,
          height: shape.height,
        },
      });
    }
  }

  return ok({
    slideIndex,
    path: `/slide[${slideIndex}]`,
    hasOverflow: overflowingShapes.length > 0,
    overflowingShapes,
  });
}

/**
 * Gets all overflow issues in a presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - Optional 1-based slide index to check specific slide
 * @returns Result with all overflow issues
 *
 * @example
 * const result = await getOverflowIssues("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(`Found ${result.data.issueCount} overflow issues`);
 *   for (const issue of result.data.issues) {
 *     console.log(`[${issue.severity}] ${issue.message}`);
 *   }
 * }
 */
export async function getOverflowIssues(
  filePath: string,
  slideIndex?: number,
): Promise<Result<OverflowIssuesResult>> {
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
  const targetSlides = slideIndex
    ? slidesInfo.filter(s => s.index === slideIndex)
    : slidesInfo;

  if (slideIndex && targetSlides.length === 0) {
    return invalidInput(`Slide index ${slideIndex} is out of range (1-${slidesInfo.length})`);
  }

  const issues: OverflowIssue[] = [];

  for (const slideInfo of targetSlides) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);
    const shapes = parseShapesFromSlideXml(slideXml, slideInfo.index);

    for (const shape of shapes) {
      if (!shape.text || !shape.width) {
        continue;
      }

      const avgCharWidth = 1000;
      const padding = 2000;
      const estimatedTextWidth = shape.text.length * avgCharWidth + padding;
      const hasOverflow = estimatedTextWidth > shape.width;

      if (hasOverflow) {
        const severity = estimateOverflowSeverity(shape, shape.text);
        const excessRatio = estimatedTextWidth / shape.width;

        let category: OverflowIssue["category"] = "text_overflow";
        let message = `Text in ${shape.path} may overflow the text box`;
        let suggestion = "Consider expanding the text box or reducing font size";

        if (excessRatio > 2) {
          category = "text_truncation";
          message = `Text in ${shape.path} is likely truncated (${Math.round(excessRatio * 100 - 100)}% overflow)`;
          suggestion = "Expand the text box significantly or reduce text content";
        } else if (shape.width < 50000) {
          category = "shape_too_small";
          message = `Text box ${shape.path} is very narrow (${shape.width} EMUs) and may cause overflow`;
          suggestion = "Consider widening the text box";
        }

        issues.push({
          severity,
          category,
          message,
          path: shape.path,
          suggestion,
        });
      }
    }
  }

  return ok({
    slideCount: targetSlides.length,
    issueCount: issues.length,
    issues,
  });
}