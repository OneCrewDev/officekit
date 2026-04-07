/**
 * View and query output operations for @officekit/ppt.
 *
 * Provides functions to extract different views of a PowerPoint presentation:
 * - ViewAsText: Plain text extraction
 * - ViewAsAnnotated: Element annotations with types and positions
 * - ViewAsOutline: Slide outline/summary
 * - ViewAsStats: Statistics about the presentation
 * - ViewAsIssues: Potential problems detection
 */

import { readFile } from "node:fs/promises";
import path from "node:path";
import { readStoredZip } from "../../core/src/zip.js";
import { err, ok, andThen, map, notFound, invalidInput } from "./result.js";
import type {
  Result,
  SlideModel,
  ShapeModel,
  ParagraphModel,
} from "./types.js";
import {
  parsePath,
  getSlideIndex,
  isPlaceholderPath,
  slidePath,
} from "./path.js";
import {
  get,
  getSlide,
  getShape,
  querySlides,
  queryShapes,
  getSlide as loadSlide,
} from "./query.js";

// ============================================================================
// Types for View Outputs
// ============================================================================

/**
 * Result from ViewAsText - structured text extraction.
 */
export interface ViewTextResult {
  /** Total slide count */
  slideCount: number;
  /** Text extracted from each slide */
  slides: SlideText[];
}

/**
 * Text content from a single slide.
 */
export interface SlideText {
  /** Slide index (1-based) */
  index: number;
  /** Slide path */
  path: string;
  /** Title text if present */
  title?: string;
  /** All text content concatenated */
  text: string;
  /** Text organized by shape */
  shapes: ShapeText[];
}

/**
 * Text content from a single shape.
 */
export interface ShapeText {
  /** Shape path */
  path: string;
  /** Shape name if available */
  name?: string;
  /** Shape type */
  type: string;
  /** Text content */
  text: string;
  /** Paragraphs */
  paragraphs: ParagraphModel[];
}

/**
 * Result from ViewAsAnnotated - annotated view of elements.
 */
export interface ViewAnnotatedResult {
  /** Total slide count */
  slideCount: number;
  /** Annotated slides */
  slides: SlideAnnotation[];
}

/**
 * Annotation for a single slide.
 */
export interface SlideAnnotation {
  /** Slide index (1-based) */
  index: number;
  /** Slide path */
  path: string;
  /** Title if present */
  title?: string;
  /** Element annotations */
  elements: ElementAnnotation[];
}

/**
 * Annotation for a single element.
 */
export interface ElementAnnotation {
  /** Element path */
  path: string;
  /** Element type (shape, table, chart, placeholder) */
  type: string;
  /** Element name if available */
  name?: string;
  /** Position and size in EMUs */
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  /** Placeholder type if applicable */
  placeholderType?: string;
  /** Text content preview (truncated) */
  textPreview?: string;
  /** Fill color if present */
  fill?: string;
  /** Line color if present */
  line?: string;
}

/**
 * Result from ViewAsOutline - presentation outline.
 */
export interface ViewOutlineResult {
  /** Total slide count */
  slideCount: number;
  /** Slide outlines */
  slides: SlideOutline[];
}

/**
 * Outline for a single slide.
 */
export interface SlideOutline {
  /** Slide index (1-based) */
  index: number;
  /** Slide path */
  path: string;
  /** Slide title */
  title?: string;
  /** Slide layout type */
  layoutType?: string;
  /** Content hierarchy */
  content: OutlineContent[];
}

/**
 * Content item in the outline.
 */
export interface OutlineContent {
  /** Content type */
  type: "title" | "placeholder" | "shape" | "table" | "chart" | "media";
  /** Path to the element */
  path: string;
  /** Content description */
  description: string;
  /** Child items for complex elements */
  children?: OutlineContent[];
}

/**
 * Result from ViewAsStats - presentation statistics.
 */
export interface ViewStatsResult {
  /** Total slide count */
  slideCount: number;
  /** Total shape count across all slides */
  shapeCount: number;
  /** Total text length across all slides */
  textLength: number;
  /** Total table count */
  tableCount: number;
  /** Total chart count */
  chartCount: number;
  /** Total picture count */
  pictureCount: number;
  /** Total media count */
  mediaCount: number;
  /** Per-slide statistics */
  slides: SlideStats[];
}

/**
 * Statistics for a single slide.
 */
export interface SlideStats {
  /** Slide index (1-based) */
  index: number;
  /** Slide path */
  path: string;
  /** Title if present */
  title?: string;
  /** Shape count on this slide */
  shapeCount: number;
  /** Text length on this slide */
  textLength: number;
  /** Table count on this slide */
  tableCount: number;
  /** Chart count on this slide */
  chartCount: number;
  /** Picture count on this slide */
  pictureCount: number;
  /** Media count on this slide */
  mediaCount: number;
  /** Placeholder count on this slide */
  placeholderCount: number;
}

/**
 * Result from ViewAsIssues - detected issues.
 */
export interface ViewIssuesResult {
  /** Total slide count */
  slideCount: number;
  /** Total issue count */
  issueCount: number;
  /** Issues found */
  issues: Issue[];
}

/**
 * A detected issue in the presentation.
 */
export interface Issue {
  /** Issue severity */
  severity: "error" | "warning" | "info";
  /** Issue category */
  category: string;
  /** Human-readable message */
  message: string;
  /** Path to the affected element */
  path?: string;
  /** Suggested fix if available */
  suggestion?: string;
}

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

// ============================================================================
// Helper Functions
// ============================================================================

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
 * Gets all slide entries from the presentation.
 */
function getAllSlideEntries(zip: Map<string, Buffer>): Result<Array<{ index: number; path: string; entryPath: string }>> {
  try {
    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const slideIds = getSlideIds(presentationXml);

    const slides = slideIds.map((s, idx) => {
      const rel = relationships.find(r => r.id === s.relId);
      const target = rel?.target ?? "";
      const entryPath = normalizeZipPath("ppt", target);
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
 * Gets the zip entry path for a slide by its 1-based index.
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
 * Extracts text content from a shape's XML.
 */
function extractTextFromShape(shapeXml: string): string {
  const textRuns: string[] = [];
  for (const match of shapeXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) {
    textRuns.push(match[1]);
  }
  return textRuns.join("");
}

/**
 * Extracts placeholder type from a shape.
 */
function extractPlaceholderType(shapeXml: string): string | undefined {
  const phMatch = shapeXml.match(/<p:ph[^>]*type="([^"]*)"[^>]*>/);
  return phMatch ? phMatch[1] : undefined;
}

/**
 * Extracts placeholder index from a shape.
 */
function extractPlaceholderIndex(shapeXml: string): number | undefined {
  const idxMatch = shapeXml.match(/<p:ph[^>]*idx="([^"]*)"[^>]*>/);
  return idxMatch ? parseInt(idxMatch[1], 10) : undefined;
}

/**
 * Extracts shape name from shape XML.
 */
function extractShapeName(shapeXml: string): string | undefined {
  const nameMatch = shapeXml.match(/<p:nvCxnSpPr[^>]*>[\s\S]*?<p:cNvPr[^>]*name="([^"]*)"[^>]*>/);
  if (nameMatch) return nameMatch[1];

  // Try alternate pattern
  const altMatch = shapeXml.match(/<p:cNvPr[^>]*name="([^"]*)"[^>]*>/);
  return altMatch ? altMatch[1] : undefined;
}

/**
 * Extracts shape properties (position, size, fill, line) from shape XML.
 */
function extractShapeProperties(shapeXml: string): {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rotation?: number;
  fill?: string;
  line?: string;
  lineWidth?: number;
  alt?: string;
} {
  const props: ReturnType<typeof extractShapeProperties> = {};

  // Extract position and size from spPr
  const spPrMatch = shapeXml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);
  if (spPrMatch) {
    const spPrContent = spPrMatch[1];

    // Extract xfrm values
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

      const rotMatch = xfrmContent.match(/<a:xfrm[^>]*rot="([^"]*)"[^>]*>/);
      if (rotMatch) {
        props.rotation = parseInt(rotMatch[1], 10) / 60000; // OOXML rotation is in 60000ths of a degree
      }
    }

    // Extract fill color
    const solidFillMatch = spPrContent.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
    if (solidFillMatch) {
      const colorMatch = solidFillMatch[1].match(/<a:srgbClr[^>]*val="([^"]*)"[^>]*>/);
      if (colorMatch) {
        props.fill = colorMatch[1];
      }
    }

    // Extract line color
    const lnMatch = spPrContent.match(/<a:ln(?:[^>]*)>([\s\S]*?)<\/a:ln>/);
    if (lnMatch) {
      const lnContent = lnMatch[1];
      const lnColorMatch = lnContent.match(/<a:solidFill[^>]*>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:solidFill>/);
      if (lnColorMatch) {
        props.line = lnColorMatch[1];
      }

      const lnWidthMatch = lnContent.match(/<a:ln[^>]*w="([^"]*)"[^>]*>/);
      if (lnWidthMatch) {
        props.lineWidth = parseInt(lnWidthMatch[1], 10) / 12700; // Convert EMUs to points
      }
    }
  }

  // Extract alt text
  const altMatch = shapeXml.match(/<p:cNvPr[^>]*descr="([^"]*)"[^>]*>/);
  if (altMatch) {
    props.alt = altMatch[1];
  }

  return props;
}

/**
 * Parses all shapes from slide XML.
 */
function parseShapesFromSlideXml(slideXml: string, slideIndex: number): ShapeModel[] {
  const shapes: ShapeModel[] = [];

  // Match shape elements
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
 * Parses a shape from XML.
 */
function parseShapeFromXml(shapeXml: string, slideIndex: number, shapeIndex: number): ShapeModel {
  const shapePathStr = `/slide[${slideIndex}]/shape[${shapeIndex}]`;

  const name = extractShapeName(shapeXml);
  const text = extractTextFromShape(shapeXml);
  const placeholderType = extractPlaceholderType(shapeXml);
  const placeholderIndex = extractPlaceholderIndex(shapeXml);
  const props = extractShapeProperties(shapeXml);

  // Determine shape type
  let type = "shape";
  if (placeholderType) {
    type = "placeholder";
  } else if (shapeXml.includes("<p:sp>")) {
    type = "shape";
  } else if (shapeXml.includes("<p:pic>")) {
    type = "picture";
  } else if (shapeXml.includes("<p:graphicFrame>")) {
    type = "graphicFrame";
  }

  return {
    path: shapePathStr,
    name,
    text,
    type,
    alt: props.alt,
    x: props.x,
    y: props.y,
    width: props.width,
    height: props.height,
    rotation: props.rotation,
    fill: props.fill,
    line: props.line,
    lineWidth: props.lineWidth,
    placeholderType: placeholderType as ShapeModel["placeholderType"],
    placeholderIndex,
    paragraphs: [],
    childCount: 0,
  };
}

/**
 * Parses all tables from slide XML.
 */
function parseTablesFromSlideXml(slideXml: string, slideIndex: number): Array<{ path: string; name?: string; rowCount?: number; columnCount?: number }> {
  const tables: Array<{ path: string; name?: string; rowCount?: number; columnCount?: number }> = [];

  // Match graphic frames containing tables
  const graphicFramePattern = /<p:graphicFrame(?:[\s\S]*?)<\/p:graphicFrame>/g;

  let tableIndex = 0;
  for (const frameMatch of slideXml.matchAll(graphicFramePattern)) {
    const frameContent = frameMatch[0];

    // Check if it contains a table
    if (!frameContent.includes("<a:tbl>")) {
      continue;
    }

    tableIndex++;
    const path = `/slide[${slideIndex}]/table[${tableIndex}]`;

    // Extract table name
    const nameMatch = frameContent.match(/<a:tbl(?:[^>]*)>[\s\S]*?<a:tblPr[^>]*>[\s\S]*?<a:nvCxnSpPr[^>]*>[\s\S]*?<p:cNvPr[^>]*name="([^"]*)"[^>]*>/);
    const name = nameMatch ? nameMatch[1] : undefined;

    // Extract grid columns to count columns
    const gridColPattern = /<a:gridCol[^>]*>[\s\S]*?<\/a:gridCol>/g;
    const gridCols = frameContent.match(gridColPattern) || [];
    const columnCount = gridCols.length;

    // Count rows
    const rowPattern = /<a:tr(?:[^>]*)>([\s\S]*?)<\/a:tr>/g;
    const rowMatches = frameContent.matchAll(rowPattern);
    let rowCount = 0;
    for (const _ of rowMatches) {
      rowCount++;
    }

    tables.push({ path, name, rowCount, columnCount });
  }

  return tables;
}

/**
 * Parses all charts from slide XML.
 */
function parseChartsFromSlideXml(zip: Map<string, Buffer>, slideXml: string, slideIndex: number): Array<{ path: string; title?: string; type?: string }> {
  const charts: Array<{ path: string; title?: string; type?: string }> = [];

  // Find chart relationships in slide
  const chartPattern = /<p:graphicFrame(?:[\s\S]*?)<\/p:graphicFrame>/g;
  let chartIndex = 0;

  for (const frameMatch of slideXml.matchAll(chartPattern)) {
    const frameContent = frameMatch[0];

    // Check if it contains a chart reference
    if (!frameContent.includes("<c:chart")) {
      continue;
    }

    chartIndex++;
    const path = `/slide[${slideIndex}]/chart[${chartIndex}]`;
    charts.push({ path });
  }

  return charts;
}

/**
 * Parses all placeholders from slide XML.
 */
function parsePlaceholdersFromSlideXml(slideXml: string, slideIndex: number): Array<{ path: string; type: string; name?: string; text?: string; index?: number }> {
  const placeholders: Array<{ path: string; type: string; name?: string; text?: string; index?: number }> = [];

  // Match placeholder shapes
  const placeholderPattern = /<p:sp(?:[\s\S]*?)<\/p:sp>/g;

  for (const shapeMatch of slideXml.matchAll(placeholderPattern)) {
    const shapeXml = shapeMatch[0];

    // Check if this is a placeholder
    if (!shapeXml.includes("<p:ph")) {
      continue;
    }

    const phType = extractPlaceholderType(shapeXml);
    if (!phType) {
      continue;
    }

    const phIndex = extractPlaceholderIndex(shapeXml);
    const name = extractShapeName(shapeXml);
    const text = extractTextFromShape(shapeXml);
    const path = `/slide[${slideIndex}]/placeholder[${phType}]`;

    placeholders.push({
      path,
      type: phType,
      index: phIndex,
      name,
      text,
    });
  }

  return placeholders;
}

/**
 * Estimates if text might overflow a shape's text box.
 * This is a rough heuristic based on character count and average char width.
 */
function estimateTextOverflow(shape: ShapeModel, text: string): boolean {
  if (!shape.width || !text) {
    return false;
  }

  // Average character width in EMUs (approximate for typical fonts)
  const avgCharWidth = 1000; // EMUs
  const padding = 2000; // EMUs for margins

  const estimatedTextWidth = text.length * avgCharWidth + padding;
  return estimatedTextWidth > shape.width;
}

// ============================================================================
// ViewAsText Implementation
// ============================================================================

/**
 * Extracts plain text from a presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - Optional 1-based slide index to get text from specific slide
 * @returns Result with structured text extraction
 *
 * @example
 * // Get all text from presentation
 * const result = await viewAsText("/path/to/presentation.pptx");
 *
 * // Get text from specific slide
 * const result = await viewAsText("/path/to/presentation.pptx", 1);
 */
export async function viewAsText(
  filePath: string,
  slideIndex?: number,
): Promise<Result<ViewTextResult>> {
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

  const slides: SlideText[] = [];

  for (const slideInfo of targetSlides!) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);
    const { title, shapes: shapeTexts } = extractTextFromSlide(slideXml, slideInfo.index);

    // Parse shapes for text content
    const parsedShapes = parseShapesFromSlideXml(slideXml, slideInfo.index);
    const shapes: ShapeText[] = parsedShapes.map(shape => ({
      path: shape.path,
      name: shape.name,
      type: shape.type,
      text: shape.text || "",
      paragraphs: shape.paragraphs || [],
    }));

    // Concatenate all text
    const allText = shapes.map(s => s.text).filter(t => t).join("\n");

    slides.push({
      index: slideInfo.index,
      path: slideInfo.path,
      title,
      text: allText,
      shapes,
    });
  }

  return ok({
    slideCount: slides.length,
    slides,
  });
}

/**
 * Extracts text from a slide's XML.
 */
function extractTextFromSlide(slideXml: string, slideIndex: number): { title?: string; titlePath?: string; shapes: Array<{ path: string; text: string }> } {
  let title: string | undefined;
  let titlePath: string | undefined;
  const shapes: Array<{ path: string; text: string }> = [];

  // Parse all shapes
  const parsedShapes = parseShapesFromSlideXml(slideXml, slideIndex);

  for (const shape of parsedShapes) {
    if (shape.text) {
      // Check if this is a title placeholder, otherwise use first shape as fallback title
      if (shape.placeholderType === "title" && !title) {
        title = shape.text;
        titlePath = shape.path;
      } else if (!title) {
        // Fallback: use first shape's text as title
        title = shape.text;
        titlePath = shape.path;
      } else {
        // Only add to shapes array if not used as title
        shapes.push({
          path: shape.path,
          text: shape.text,
        });
      }
    }
  }

  return { title, titlePath, shapes };
}

// ============================================================================
// ViewAsAnnotated Implementation
// ============================================================================

/**
 * Gets an annotated view of the presentation showing element types and positions.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - Optional 1-based slide index to annotate specific slide
 * @returns Result with annotated view
 *
 * @example
 * const result = await viewAsAnnotated("/path/to/presentation.pptx");
 */
export async function viewAsAnnotated(
  filePath: string,
  slideIndex?: number,
): Promise<Result<ViewAnnotatedResult>> {
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

  const slides: SlideAnnotation[] = [];

  for (const slideInfo of targetSlides!) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);
    const { title } = extractTextFromSlide(slideXml, slideInfo.index);

    const elements: ElementAnnotation[] = [];

    // Parse shapes
    const shapes = parseShapesFromSlideXml(slideXml, slideInfo.index);
    for (const shape of shapes) {
      elements.push({
        path: shape.path,
        type: shape.type,
        name: shape.name,
        x: shape.x,
        y: shape.y,
        width: shape.width,
        height: shape.height,
        placeholderType: shape.placeholderType,
        textPreview: shape.text ? (shape.text.length > 50 ? shape.text.slice(0, 50) + "..." : shape.text) : undefined,
        fill: shape.fill,
        line: shape.line,
      });
    }

    // Parse tables
    const tables = parseTablesFromSlideXml(slideXml, slideInfo.index);
    for (const table of tables) {
      elements.push({
        path: table.path,
        type: "table",
        name: table.name,
      });
    }

    // Parse charts
    const charts = parseChartsFromSlideXml(zip, slideXml, slideInfo.index);
    for (const chart of charts) {
      elements.push({
        path: chart.path,
        type: "chart",
        name: chart.title,
      });
    }

    // Parse placeholders
    const placeholders = parsePlaceholdersFromSlideXml(slideXml, slideInfo.index);
    for (const placeholder of placeholders) {
      elements.push({
        path: placeholder.path,
        type: "placeholder",
        name: placeholder.name,
        placeholderType: placeholder.type,
        textPreview: placeholder.text ? (placeholder.text.length > 50 ? placeholder.text.slice(0, 50) + "..." : placeholder.text) : undefined,
      });
    }

    slides.push({
      index: slideInfo.index,
      path: slideInfo.path,
      title,
      elements,
    });
  }

  return ok({
    slideCount: slides.length,
    slides,
  });
}

// ============================================================================
// ViewAsOutline Implementation
// ============================================================================

/**
 * Gets an outline/summary view of the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - Optional 1-based slide index to get outline for specific slide
 * @returns Result with outline view
 *
 * @example
 * const result = await viewAsOutline("/path/to/presentation.pptx");
 */
export async function viewAsOutline(
  filePath: string,
  slideIndex?: number,
): Promise<Result<ViewOutlineResult>> {
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

  const slides: SlideOutline[] = [];

  for (const slideInfo of targetSlides!) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);
    const { title, titlePath } = extractTextFromSlide(slideXml, slideInfo.index);

    const content: OutlineContent[] = [];

    // Parse placeholders first as they are typically main content
    const placeholders = parsePlaceholdersFromSlideXml(slideXml, slideInfo.index);
    for (const placeholder of placeholders) {
      content.push({
        type: "placeholder",
        path: placeholder.path,
        description: `${placeholder.type}${placeholder.text ? `: "${placeholder.text.slice(0, 30)}${placeholder.text.length > 30 ? "..." : ""}"` : ""}`,
      });
    }

    // Parse shapes
    const parsedShapes = parseShapesFromSlideXml(slideXml, slideInfo.index);
    let shapeIndex = 0;
    for (const shape of parsedShapes) {
      // Skip shapes that are placeholders (already added)
      if (shape.placeholderType) {
        continue;
      }
      // Skip the title shape (used as slide title)
      if (titlePath && shape.path === titlePath) {
        continue;
      }
      shapeIndex++;

      const children: OutlineContent[] = [];

      // Add paragraphs as children if there are multiple
      if (shape.paragraphs && shape.paragraphs.length > 1) {
        for (let i = 0; i < shape.paragraphs.length; i++) {
          const para = shape.paragraphs[i];
          if (para.text) {
            children.push({
              type: "shape",
              path: `${shape.path}/para[${i + 1}]`,
              description: para.text.slice(0, 50) + (para.text.length > 50 ? "..." : ""),
            });
          }
        }
      }

      content.push({
        type: shape.type === "textbox" ? "shape" : (shape.type as OutlineContent["type"]),
        path: shape.path,
        description: `Shape ${shapeIndex}:${shape.name ? ` (${shape.name})` : ""}${shape.text ? ` ${shape.text.slice(0, 30)}${shape.text.length > 30 ? "..." : ""}` : ""}`,
        children: children.length > 0 ? children : undefined,
      });
    }

    // Parse tables
    const tables = parseTablesFromSlideXml(slideXml, slideInfo.index);
    for (const table of tables) {
      content.push({
        type: "table",
        path: table.path,
        description: `table${table.name ? ` (${table.name})` : ""} - ${table.rowCount || 0} rows x ${table.columnCount || 0} cols`,
      });
    }

    // Parse charts
    const charts = parseChartsFromSlideXml(zip, slideXml, slideInfo.index);
    for (const chart of charts) {
      content.push({
        type: "chart",
        path: chart.path,
        description: `chart${chart.title ? `: "${chart.title}"` : ""}`,
      });
    }

    slides.push({
      index: slideInfo.index,
      path: slideInfo.path,
      title,
      layoutType: undefined, // Would need to parse layout info
      content,
    });
  }

  return ok({
    slideCount: slides.length,
    slides,
  });
}

// ============================================================================
// ViewAsStats Implementation
// ============================================================================

/**
 * Gets statistics about the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - Optional 1-based slide index to get stats for specific slide
 * @returns Result with presentation statistics
 *
 * @example
 * const result = await viewAsStats("/path/to/presentation.pptx");
 */
export async function viewAsStats(
  filePath: string,
  slideIndex?: number,
): Promise<Result<ViewStatsResult>> {
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

  let totalShapeCount = 0;
  let totalTextLength = 0;
  let totalTableCount = 0;
  let totalChartCount = 0;
  let totalPictureCount = 0;
  let totalMediaCount = 0;

  const slideStats: SlideStats[] = [];

  for (const slideInfo of targetSlides!) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);
    const { title, shapes: shapeTexts } = extractTextFromSlide(slideXml, slideInfo.index);

    const shapeCount = shapeTexts.length;
    const textLength = shapeTexts.reduce((acc, s) => acc + (s.text?.length || 0), 0);
    const tableCount = parseTablesFromSlideXml(slideXml, slideInfo.index).length;
    const chartCount = parseChartsFromSlideXml(zip, slideXml, slideInfo.index).length;
    const placeholderCount = parsePlaceholdersFromSlideXml(slideXml, slideInfo.index).length;

    // Count pictures and media
    let pictureCount = 0;
    let mediaCount = 0;
    const parsedShapes = parseShapesFromSlideXml(slideXml, slideInfo.index);
    for (const shape of parsedShapes) {
      if (shape.type === "picture") {
        pictureCount++;
      }
    }

    totalShapeCount += shapeCount;
    totalTextLength += textLength;
    totalTableCount += tableCount;
    totalChartCount += chartCount;
    totalPictureCount += pictureCount;
    totalMediaCount += mediaCount;

    slideStats.push({
      index: slideInfo.index,
      path: slideInfo.path,
      title,
      shapeCount,
      textLength,
      tableCount,
      chartCount,
      pictureCount,
      mediaCount,
      placeholderCount,
    });
  }

  return ok({
    slideCount: slideStats.length,
    shapeCount: totalShapeCount,
    textLength: totalTextLength,
    tableCount: totalTableCount,
    chartCount: totalChartCount,
    pictureCount: totalPictureCount,
    mediaCount: totalMediaCount,
    slides: slideStats,
  });
}

/**
 * Gets statistics for a specific slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with slide statistics
 *
 * @example
 * const result = await getSlideStats("/path/to/presentation.pptx", 1);
 */
export async function getSlideStats(
  filePath: string,
  slideIndex: number,
): Promise<Result<SlideStats>> {
  const statsResult = await viewAsStats(filePath, slideIndex);
  if (!statsResult.ok) {
    return err(statsResult.error!.code, statsResult.error!.message);
  }

  const slideStats = statsResult.data!.slides[0];
  if (!slideStats) {
    return invalidInput(`Slide index ${slideIndex} is out of range`);
  }

  return ok(slideStats);
}

// ============================================================================
// ViewAsIssues Implementation
// ============================================================================

/**
 * Finds potential problems in the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - Optional 1-based slide index to check specific slide
 * @returns Result with detected issues
 *
 * @example
 * const result = await viewAsIssues("/path/to/presentation.pptx");
 */
export async function viewAsIssues(
  filePath: string,
  slideIndex?: number,
): Promise<Result<ViewIssuesResult>> {
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

  const issues: Issue[] = [];

  for (const slideInfo of targetSlides!) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);

    // Check for missing title
    const placeholders = parsePlaceholdersFromSlideXml(slideXml, slideInfo.index);
    const titlePlaceholder = placeholders.find(p => p.type === "title");
    if (!titlePlaceholder || !titlePlaceholder.text || titlePlaceholder.text.trim() === "") {
      issues.push({
        severity: "warning",
        category: "missing_title",
        message: `Slide ${slideInfo.index} is missing a title`,
        path: slideInfo.path,
        suggestion: "Add a title to this slide for better accessibility and navigation",
      });
    }

    // Check shapes for potential text overflow
    const shapes = parseShapesFromSlideXml(slideXml, slideInfo.index);
    for (const shape of shapes) {
      // Check for shapes without alt text
      if (!shape.alt || shape.alt.trim() === "") {
        // Only warn for non-placeholder shapes that might be images or complex shapes
        if (shape.type !== "placeholder" && shape.type !== "shape") {
          issues.push({
            severity: "info",
            category: "missing_alt_text",
            message: `Shape ${shape.path} is missing alternative text`,
            path: shape.path,
            suggestion: "Add descriptive alt text for accessibility",
          });
        }
      }

      // Check for potential text overflow in text boxes
      if (shape.text && shape.width && shape.width < 50000) { // Very narrow text box
        if (estimateTextOverflow(shape, shape.text)) {
          issues.push({
            severity: "warning",
            category: "text_overflow_risk",
            message: `Text in ${shape.path} may overflow the text box`,
            path: shape.path,
            suggestion: "Consider expanding the text box width or reducing font size",
          });
        }
      }

      // Check for shapes with missing names
      if (!shape.name || shape.name.trim() === "") {
        issues.push({
          severity: "info",
          category: "unnamed_shape",
          message: `Shape ${shape.path} does not have a name`,
          path: shape.path,
          suggestion: "Name shapes for easier identification and maintenance",
        });
      }
    }

    // Check for empty slides
    if (shapes.length === 0) {
      issues.push({
        severity: "warning",
        category: "empty_slide",
        message: `Slide ${slideInfo.index} has no shapes`,
        path: slideInfo.path,
        suggestion: "Add content to this slide or remove it if not needed",
      });
    }

    // Check tables
    const tables = parseTablesFromSlideXml(slideXml, slideInfo.index);
    for (const table of tables) {
      if (!table.name || table.name.trim() === "") {
        issues.push({
          severity: "info",
          category: "unnamed_table",
          message: `Table ${table.path} does not have a name`,
          path: table.path,
          suggestion: "Name tables for easier identification",
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

/**
 * Checks if text overflows in a specific shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @returns Result with overflow check
 *
 * @example
 * const result = await checkShapeTextOverflow("/path/to/presentation.pptx", "/slide[1]/shape[1]");
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
  const avgCharWidth = 1000; // EMUs
  const padding = 2000; // EMUs
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
