/**
 * Query and get operations for @officekit/ppt.
 *
 * Provides functions to retrieve and query elements within a PowerPoint
 * presentation using path-based and selector-based approaches.
 */

import { readFile } from "node:fs/promises";
import path from "node:path";
import { readStoredZip } from "../../core/src/zip.js";
import { err, ok, andThen, map, notFound, invalidInput } from "./result.js";
import type {
  Result,
  SlideModel,
  ShapeModel,
  TableModel,
  TableRowModel,
  TableCellModel,
  ChartModel,
  PlaceholderModel,
  ParagraphModel,
  RunModel,
  PlaceholderType,
  ParsedSelector,
} from "./types.js";
import {
  parsePath,
  getSlideIndex,
  isSlidePath,
  isPlaceholderPath,
  buildPath,
  slidePath,
  shapePath,
  tablePath,
  placeholderPath,
  chartPath,
} from "./path.js";
import { parseSelector } from "./selectors.js";

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
 * Gets the relationships entry name for a given entry.
 */
function getRelationshipsEntryName(entryName: string): string {
  const directory = path.posix.dirname(entryName);
  const basename = path.posix.basename(entryName);
  return path.posix.join(directory, "_rels", `${basename}.rels`);
}

/**
 * Extracts layout information from layout XML.
 */
function parseLayoutInfo(layoutXml: string): { name: string; type?: string } {
  const name = /<p:cSld\b[^>]*name="([^"]*)"/.exec(layoutXml)?.[1] ?? "";
  const type = /<p:sldLayout\b[^>]*type="([^"]+)"/.exec(layoutXml)?.[1];
  return { name, type };
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
 * Gets slide IDs with their full information.
 */
function getSlideInfo(
  zip: Map<string, Buffer>,
): Result<Array<{ index: number; path: string; relId: string; entryPath: string }>> {
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
        relId: s.relId,
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

// ============================================================================
// XML Parsing Helpers
// ============================================================================

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
 * Extracts paragraphs from a shape's text body.
 */
function extractParagraphsFromShape(shapeXml: string): ParagraphModel[] {
  const paragraphs: ParagraphModel[] = [];

  // Find the text body
  const txBodyMatch = shapeXml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
  if (!txBodyMatch) {
    return paragraphs;
  }

  const txBody = txBodyMatch[1];

  // Match individual paragraphs
  const paraPattern = /<a:p(?:[^>]*)>([\s\S]*?)<\/a:p>/g;
  let paraIndex = 0;

  for (const paraMatch of txBody.matchAll(paraPattern)) {
    paraIndex++;
    const paraContent = paraMatch[1];

    // Extract alignment
    const alignmentMatch = paraContent.match(/<a:pPr[^>]*algn="([^"]*)"[^>]*>/);
    const alignment = alignmentMatch ? (alignmentMatch[1] as ParagraphModel["alignment"]) : undefined;

    // Extract margin left
    const marginLeftMatch = paraContent.match(/<a:pPr[^>]*marL="([^"]*)"[^>]*>/);
    const marginLeft = marginLeftMatch ? parseInt(marginLeftMatch[1], 10) : undefined;

    // Extract margin right
    const marginRightMatch = paraContent.match(/<a:pPr[^>]*marR="([^"]*)"[^>]*>/);
    const marginRight = marginRightMatch ? parseInt(marginRightMatch[1], 10) : undefined;

    // Extract line spacing
    const lineSpacingMatch = paraContent.match(/<a:pPr[^>]*spc="([^"]*)"[^>]*>/);
    const lineSpacing = lineSpacingMatch ? lineSpacingMatch[1] : undefined;

    // Extract space before/after
    const spaceBeforeMatch = paraContent.match(/<a:pPr[^>]*spcBef="([^"]*)"[^>]*>/);
    const spaceBefore = spaceBeforeMatch ? spaceBeforeMatch[1] : undefined;

    const spaceAfterMatch = paraContent.match(/<a:pPr[^>]*spcAft="([^"]*)"[^>]*>/);
    const spaceAfter = spaceAfterMatch ? spaceAfterMatch[1] : undefined;

    // Extract runs
    const runs = extractRunsFromParagraph(paraContent);

    // Get concatenated text
    const text = runs.map(r => r.text).join("");

    paragraphs.push({
      index: paraIndex,
      text,
      alignment,
      marginLeft,
      marginRight,
      lineSpacing,
      spaceBefore,
      spaceAfter,
      runs,
      childCount: runs.length,
    });
  }

  return paragraphs;
}

/**
 * Extracts runs from a paragraph's content.
 */
function extractRunsFromParagraph(paraContent: string): RunModel[] {
  const runs: RunModel[] = [];
  let runIndex = 0;

  // Match run elements - they can be <a:r> or <a:rPr> followed by <a:t>
  const runPattern = /<a:r(?:[^>]*)>([\s\S]*?)<\/a:r>/g;

  for (const runMatch of paraContent.matchAll(runPattern)) {
    runIndex++;
    const runContent = runMatch[1];

    // Extract text
    const textMatch = runContent.match(/<a:t(?:[^>]*)>([^<]*)<\/a:t>/);
    const text = textMatch ? textMatch[1] : "";

    // Extract run properties
    const rPrMatch = runContent.match(/<a:rPr(?:[^>]*)>([\s\S]*?)<\/a:rPr>/);
    if (!rPrMatch) {
      runs.push({ index: runIndex, text });
      continue;
    }

    const rPrContent = rPrMatch[1];

    // Extract font typeface
    const typefaceMatch = rPrContent.match(/<a:latin(?:[^>]*typeface="([^"]*)"[^>]*)?\/>/);
    const font = typefaceMatch ? typefaceMatch[1] : undefined;

    // Extract font size
    const sizeMatch = rPrContent.match(/<a:latin[^>]*sz="([^"]*)"[^>]*\/>/);
    const size = sizeMatch ? sizeMatch[1] : undefined;

    // Extract bold
    const boldMatch = rPrContent.match(/<a:latin[^>]*b="1"[^>]*\/>/);
    const bold = boldMatch ? true : undefined;

    // Extract italic
    const italicMatch = rPrContent.match(/<a:latin[^>]*i="1"[^>]*\/>/);
    const italic = italicMatch ? true : undefined;

    // Extract underline
    const underlineMatch = rPrContent.match(/<a:latin[^>]*u="([^"]*)"[^>]*\/>/);
    const underline = underlineMatch ? underlineMatch[1] : undefined;

    // Extract strikethrough
    const strikeMatch = rPrContent.match(/<a:latin[^>]*strike="([^"]*)"[^>]*\/>/);
    const strike = strikeMatch ? strikeMatch[1] : undefined;

    // Extract color
    const colorMatch = rPrContent.match(/<a:solidFill[^>]*>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:solidFill>/);
    const color = colorMatch ? colorMatch[1] : undefined;

    runs.push({
      index: runIndex,
      text,
      font,
      size,
      bold,
      italic,
      underline,
      strike,
      color,
    });
  }

  return runs;
}

/**
 * Extracts placeholder type from a shape.
 */
function extractPlaceholderType(shapeXml: string): PlaceholderType | undefined {
  const phMatch = shapeXml.match(/<p:ph[^>]*type="([^"]*)"[^>]*>/);
  return phMatch ? (phMatch[1] as PlaceholderType) : undefined;
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
 * Parses a table from XML.
 */
function parseTable(tableXml: string, slideIndex: number, tableIndex: number): TableModel {
  const pathPrefix = `/slide[${slideIndex}]/table[${tableIndex}]`;

  // Extract table name
  const nameMatch = tableXml.match(/<a:tbl(?:[^>]*)>[\s\S]*?<a:tblPr[^>]*>[\s\S]*?<a:nvCxnSpPr[^>]*>[\s\S]*?<p:cNvPr[^>]*name="([^"]*)"[^>]*>/);
  const name = nameMatch ? nameMatch[1] : undefined;

  // Extract grid columns
  const gridColPattern = /<a:gridCol[^>]*>[\s\S]*?<\/a:gridCol>/g;
  const gridCols = tableXml.match(gridColPattern) || [];
  const columnCount = gridCols.length;

  // Extract rows
  const rowPattern = /<a:tr(?:[^>]*)>([\s\S]*?)<\/a:tr>/g;
  const rows: TableRowModel[] = [];
  let rowIndex = 0;

  for (const rowMatch of tableXml.matchAll(rowPattern)) {
    rowIndex++;
    const rowContent = rowMatch[1];
    const rowPath = `${pathPrefix}/tr[${rowIndex}]`;

    // Extract cells
    const cellPattern = /<a:tc(?:[^>]*)>([\s\S]*?)<\/a:tc>/g;
    const cells: TableCellModel[] = [];
    let cellIndex = 0;

    for (const cellMatch of rowContent.matchAll(cellPattern)) {
      cellIndex++;
      const cellContent = cellMatch[1];
      const cellPath = `${rowPath}/tc[${cellIndex}]`;

      // Extract cell text
      const textRuns: string[] = [];
      for (const tMatch of cellContent.matchAll(/<a:t>([^<]*)<\/a:t>/g)) {
        textRuns.push(tMatch[1]);
      }
      const text = textRuns.join("");

      // Extract grid span
      const gridSpanMatch = cellContent.match(/<a:tcPr[^>]*gridSpan="([^"]*)"[^>]*>/);
      const gridSpan = gridSpanMatch ? parseInt(gridSpanMatch[1], 10) : undefined;

      // Extract row span
      const rowSpanMatch = cellContent.match(/<a:tcPr[^>]*rowSpan="([^"]*)"[^>]*>/);
      const rowSpan = rowSpanMatch ? parseInt(rowSpanMatch[1], 10) : undefined;

      // Extract vmerge
      const vmergeMatch = cellContent.match(/<a:tcPr[^>]*vMerge="([^"]*)"[^>]*>/);
      const vmerge = vmergeMatch ? vmergeMatch[1] === "1" : undefined;

      // Extract hmerge
      const hmergeMatch = cellContent.match(/<a:tcPr[^>]*hMerge="([^"]*)"[^>]*>/);
      const hmerge = hmergeMatch ? hmergeMatch[1] === "1" : undefined;

      // Extract fill color
      const fillMatch = cellContent.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
      let fill: string | undefined;
      if (fillMatch) {
        const colorMatch = fillMatch[1].match(/<a:srgbClr[^>]*val="([^"]*)"[^>]*>/);
        if (colorMatch) {
          fill = colorMatch[1];
        }
      }

      // Extract vertical alignment
      const anchorMatch = cellContent.match(/<a:tcPr[^>]*anchor="([^"]*)"[^>]*>/);
      const valign = anchorMatch ? (anchorMatch[1] as TableCellModel["valign"]) : undefined;

      // Extract horizontal alignment
      const alignMatch = cellContent.match(/<a:pPr[^>]*algn="([^"]*)"[^>]*>/);
      const alignment = alignMatch ? alignMatch[1] : undefined;

      cells.push({
        index: cellIndex,
        path: cellPath,
        text,
        gridSpan,
        rowSpan,
        vmerge,
        hmerge,
        fill,
        valign,
        alignment,
      });
    }

    rows.push({
      index: rowIndex,
      path: rowPath,
      cellCount: cells.length,
      cells,
    });
  }

  return {
    path: pathPrefix,
    name,
    columnCount,
    rowCount: rows.length,
    rows,
    hasHeaderRow: rows.length > 0,
  };
}

/**
 * Parses a chart from slide XML.
 */
function parseChart(chartXml: string, chartRelId: string, slideIndex: number, chartIndex: number): ChartModel {
  const chartPathStr = `/slide[${slideIndex}]/chart[${chartIndex}]`;

  // This is a simplified chart parser - actual chart data is in separate chart files
  // For now, return basic chart info

  return {
    path: chartPathStr,
    type: "chart", // Would need to parse actual chart type from chart XML
  };
}

// ============================================================================
// Get Operations
// ============================================================================

/**
 * Gets an element at the specified path.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the element (e.g., "/slide[1]", "/slide[1]/shape[2]")
 * @returns Result with the element at the path
 *
 * @example
 * const result = await get("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 * if (result.ok) {
 *   console.log(result.data); // ShapeModel
 * }
 */
export async function get(filePath: string, pptPath: string): Promise<Result<unknown>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("get requires a slide path");
  }

  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code || "operation_failed", zipResult.error?.message || "Failed to load presentation");
  }
  const zip = zipResult.data!;

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code || "operation_failed", slidePathResult.error?.message || "Failed to get slide entry path");
  }

  const slideEntry = slidePathResult.data;
  const slideXml = requireEntry(zip, slideEntry!);

  // Check if it's a placeholder
  if (isPlaceholderPath(pptPath)) {
    return getPlaceholder(filePath, slideIndex, pptPath.includes("[title]") ? "title" : "body");
  }

  // Parse the path to get the element type and index
  const pathResult = parsePath(pptPath);
  if (!pathResult.ok) {
    return err(pathResult.error?.code || "invalid_path", pathResult.error?.message || "Failed to parse path");
  }

  const segments = pathResult.data!.segments;
  if (segments.length === 1 && segments[0].name === "slide") {
    // Return slide
    return getSlide(filePath, slideIndex);
  }

  if (segments.length >= 2) {
    const childSegment = segments[1];

    switch (childSegment.name) {
      case "shape":
        return getShape(filePath, pptPath);
      case "table":
        return getTable(filePath, pptPath);
      case "chart":
        return getChart(filePath, pptPath);
      case "placeholder":
        if (childSegment.nameSelector) {
          return getPlaceholder(filePath, slideIndex, childSegment.nameSelector);
        }
        break;
    }
  }

  return invalidInput(`Unsupported path: ${pptPath}`);
}

/**
 * Gets a slide at the specified index.
 *
 * @param filePath - Path to the PPTX file
 * @param index - 1-based slide index
 * @returns Result with the slide model
 *
 * @example
 * const result = await getSlide("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.shapes);
 * }
 */
export async function getSlide(filePath: string, index: number): Promise<Result<SlideModel>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code || "operation_failed", zipResult.error?.message || "Failed to load presentation");
  }
  const zip = zipResult.data!;

  const slidePathResult = getSlideEntryPath(zip, index);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code || "operation_failed", slidePathResult.error?.message || "Failed to get slide entry path");
  }

  const slideEntry = slidePathResult.data!;
  const slideXml = requireEntry(zip, slideEntry);

  // Get all slides info for layout name lookup
  const slidesInfoResult = getSlideInfo(zip);
  if (!slidesInfoResult.ok) {
    return err("operation_failed", "Could not get slides info");
  }
  const slidesInfo = slidesInfoResult.data!;
  const slideInfo = slidesInfo[index - 1];

  // Parse shapes
  const shapes = parseShapesFromSlideXml(slideXml, index);
  const tables = parseTablesFromSlideXml(slideXml, index);
  const charts = parseChartsFromSlideXml(zip, slideXml, index);
  const placeholders = parsePlaceholdersFromSlideXml(slideXml, index);

  // Extract layout and theme info by following relationships
  const slideRelsPath = getRelationshipsEntryName(slideEntry);
  const slideRelsXml = zip.get(slideRelsPath)?.toString("utf8") ?? "";
  const slideRels = parseRelationshipEntries(slideRelsXml);
  const layoutRel = slideRels.find(r => r.type?.endsWith("/slideLayout"));

  let layoutName = "";
  let layoutType: string | undefined;
  let themeName: string | undefined;

  if (layoutRel) {
    const layoutPath = normalizeZipPath(path.posix.dirname(slideEntry), layoutRel.target);
    const layoutXml = zip.get(layoutPath)?.toString("utf8") ?? "";
    const layoutInfo = parseLayoutInfo(layoutXml);
    layoutName = layoutInfo.name;
    layoutType = layoutInfo.type;

    // Follow to master and then to theme
    const layoutRelsPath = getRelationshipsEntryName(layoutPath);
    const layoutRelsXml = zip.get(layoutRelsPath)?.toString("utf8") ?? "";
    const layoutRels = parseRelationshipEntries(layoutRelsXml);
    const masterRel = layoutRels.find(r => r.type?.endsWith("/slideMaster"));

    if (masterRel) {
      const masterPath = normalizeZipPath(path.posix.dirname(layoutPath), masterRel.target);
      const masterRelsPath = getRelationshipsEntryName(masterPath);
      const masterRelsXml = zip.get(masterRelsPath)?.toString("utf8") ?? "";
      const masterRels = parseRelationshipEntries(masterRelsXml);
      const themeRel = masterRels.find(r => r.type?.endsWith("/theme"));

      if (themeRel) {
        const themePath = normalizeZipPath(path.posix.dirname(masterPath), themeRel.target);
        const themeXml = zip.get(themePath)?.toString("utf8") ?? "";
        const themeNameMatch = /<a:theme\b[^>]*name="([^"]*)"/.exec(themeXml);
        themeName = themeNameMatch?.[1];
      }
    }
  }

  // Get title from first title placeholder, or first shape with title placeholder type, or first shape with text
  const titlePlaceholder = placeholders.find(p => p.type === "title");
  const titleFromPlaceholder = titlePlaceholder?.text || shapes.find(s => s.placeholderType === "title")?.text;
  // Fallback to first shape with text if no title placeholder found
  const title = titleFromPlaceholder || shapes.find(s => s.text && s.text.trim().length > 0)?.text;

  // Get notes
  const notesResult = await getSlideNotes(zip, index);
  const notes = notesResult.ok ? notesResult.data : undefined;

  return ok({
    index,
    path: `/slide[${index}]`,
    title,
    notes,
    layout: layoutName || undefined,
    layoutType,
    themeName,
    shapes,
    tables,
    charts,
    pictures: [],
    media: [],
    placeholders,
    childCount: shapes.length + tables.length + charts.length,
  });
}

/**
 * Gets the notes for a slide.
 */
async function getSlideNotes(zip: Map<string, Buffer>, slideIndex: number): Promise<Result<string>> {
  try {
    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const slideIds = getSlideIds(presentationXml);

    if (slideIndex < 1 || slideIndex > slideIds.length) {
      return invalidInput(`Slide index ${slideIndex} is out of range`);
    }

    const slide = slideIds[slideIndex - 1];
    const slideRel = relationships.find(r => r.id === slide.relId);
    const slidePath = normalizeZipPath("ppt", slideRel?.target ?? "");
    const slideRelsPath = `ppt/slides/_rels/${path.posix.basename(slidePath)}.rels`;

    const slideRelsXml = zip.get(slideRelsPath)?.toString("utf8") || "";

    // Find notes relationship
    const notesRelMatch = slideRelsXml.match(/<Relationship[^>]*Type="[^"]*notesMaster[^"]*"[^>]*Id="([^"]*)"[^>]*>/);
    if (!notesRelMatch) {
      return ok(""); // No notes
    }

    // Find notes slide
    const notesRelId = notesRelMatch[1];
    const notesTargetMatch = slideRelsXml.match(new RegExp(`<Relationship[^>]*Id="${notesRelId}"[^>]*Target="([^"]*)"[^>]*>`));
    if (!notesTargetMatch) {
      return ok("");
    }

    let notesTarget = notesTargetMatch[1];
    if (!notesTarget.startsWith("../")) {
      notesTarget = `ppt/notesSlides/${path.posix.basename(notesTarget.replace("notes/", ""))}`;
    } else {
      notesTarget = notesTarget.replace("../", "ppt/");
    }

    const notesXml = zip.get(notesTarget)?.toString("utf8") || "";
    if (!notesXml) {
      return ok("");
    }

    // Extract text from notes
    const textRuns: string[] = [];
    for (const match of notesXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) {
      textRuns.push(match[1]);
    }

    return ok(textRuns.join(""));
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets a shape at the specified path.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @returns Result with the shape model
 *
 * @example
 * const result = await getShape("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 */
export async function getShape(filePath: string, pptPath: string): Promise<Result<ShapeModel>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("getShape requires a slide path");
  }

  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code || "operation_failed", zipResult.error?.message || "Failed to load presentation");
  }
  const zip = zipResult.data!;

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code || "operation_failed", slidePathResult.error?.message || "Failed to get slide entry path");
  }

  const slideEntry = slidePathResult.data!;
  const slideXml = requireEntry(zip, slideEntry);

  // Extract shape index
  const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
  if (!shapeIndexMatch) {
    return invalidInput("Invalid shape path");
  }
  const shapeIndex = parseInt(shapeIndexMatch[1], 10);

  // Find all shapes
  const shapePattern = /<p:sp(?:[\s\S]*?)<\/p:sp>/g;
  const shapes = slideXml.match(shapePattern);

  if (!shapes || shapeIndex < 1 || shapeIndex > shapes.length) {
    return notFound("Shape", String(shapeIndex), `Slide ${slideIndex} has ${shapes?.length || 0} shapes`);
  }

  const shapeXml = shapes[shapeIndex - 1];
  return ok(parseShapeFromXml(shapeXml, slideIndex, shapeIndex));
}

/**
 * Parses a shape from XML.
 */
function parseShapeFromXml(shapeXml: string, slideIndex: number, shapeIndex: number): ShapeModel {
  const shapePathStr = `/slide[${slideIndex}]/shape[${shapeIndex}]`;

  const name = extractShapeName(shapeXml);
  const text = extractTextFromShape(shapeXml);
  const paragraphs = extractParagraphsFromShape(shapeXml);
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
    placeholderType,
    placeholderIndex,
    paragraphs,
    childCount: paragraphs.length,
  };
}

/**
 * Gets a table at the specified path.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the table (e.g., "/slide[1]/table[1]")
 * @returns Result with the table model
 *
 * @example
 * const result = await getTable("/path/to/presentation.pptx", "/slide[1]/table[1]");
 */
export async function getTable(filePath: string, pptPath: string): Promise<Result<TableModel>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("getTable requires a slide path");
  }

  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code || "operation_failed", zipResult.error?.message || "Failed to load presentation");
  }
  const zip = zipResult.data!;

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code || "operation_failed", slidePathResult.error?.message || "Failed to get slide entry path");
  }

  const slideEntry = slidePathResult.data!;
  const slideXml = requireEntry(zip, slideEntry);

  // Extract table index
  const tableIndexMatch = pptPath.match(/\/table\[(\d+)\]/i);
  if (!tableIndexMatch) {
    return invalidInput("Invalid table path");
  }
  const tableIndex = parseInt(tableIndexMatch[1], 10);

  // Find all tables
  const tablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  const tables = slideXml.match(tablePattern);

  if (!tables || tableIndex < 1 || tableIndex > tables.length) {
    return notFound("Table", String(tableIndex), `Slide ${slideIndex} has ${tables?.length || 0} tables`);
  }

  const tableXml = tables[tableIndex - 1];
  return ok(parseTable(tableXml, slideIndex, tableIndex));
}

/**
 * Gets a chart at the specified path.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the chart (e.g., "/slide[1]/chart[1]")
 * @returns Result with the chart model
 *
 * @example
 * const result = await getChart("/path/to/presentation.pptx", "/slide[1]/chart[1]");
 */
export async function getChart(filePath: string, pptPath: string): Promise<Result<ChartModel>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("getChart requires a slide path");
  }

  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code || "operation_failed", zipResult.error?.message || "Failed to load presentation");
  }
  const zip = zipResult.data!;

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code || "operation_failed", slidePathResult.error?.message || "Failed to get slide entry path");
  }

  const slideEntry = slidePathResult.data!;
  const slideXml = requireEntry(zip, slideEntry);

  // Extract chart index
  const chartIndexMatch = pptPath.match(/\/chart\[(\d+)\]/i);
  if (!chartIndexMatch) {
    return invalidInput("Invalid chart path");
  }
  const chartIndex = parseInt(chartIndexMatch[1], 10);

  // Find chart relationships in slide
  const slideRelsPath = `ppt/slides/_rels/${path.posix.basename(slideEntry)}.rels`;
  const slideRelsXml = zip.get(slideRelsPath)?.toString("utf8") || "";

  // Find all chart relationships
  const chartRels: Array<{ index: number; id: string; target: string }> = [];
  const chartRelPattern = /<Relationship[^>]*Type="[^"]*chart[^"]*"[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*>/g;
  let idx = 0;
  for (const match of slideRelsXml.matchAll(chartRelPattern)) {
    idx++;
    chartRels.push({ index: idx, id: match[1], target: match[2] });
  }

  if (chartIndex < 1 || chartIndex > chartRels.length) {
    return notFound("Chart", String(chartIndex), `Slide ${slideIndex} has ${chartRels.length} charts`);
  }

  const chartRel = chartRels[chartIndex - 1];
  const chartTarget = chartRel.target.startsWith("../")
    ? chartRel.target.replace("../", "ppt/")
    : `ppt/charts/${path.posix.basename(chartRel.target)}`;

  // Load chart XML
  let chartXml: string;
  try {
    chartXml = requireEntry(zip, chartTarget);
  } catch {
    // Try alternate path format
    const altTarget = `ppt/charts/chart${chartIndex}.xml`;
    chartXml = requireEntry(zip, altTarget);
  }

  return ok(parseChartFromXml(chartXml, slideIndex, chartIndex));
}

/**
 * Parses a chart from chart XML.
 */
function parseChartFromXml(chartXml: string, slideIndex: number, chartIndex: number): ChartModel {
  const chartPathStr = `/slide[${slideIndex}]/chart[${chartIndex}]`;

  // Extract title
  const titleMatch = chartXml.match(/<c:title[^>]*>[\s\S]*?<c:tx[^>]*>[\s\S]*?<a:t>([^<]*)<\/a:t>[\s\S]*?<\/c:tx>[\s\S]*?<\/c:title>/);
  const title = titleMatch ? titleMatch[1] : undefined;

  // Extract chart type
  const typeMatch = chartXml.match(/<c:barChart>|<c:lineChart>|<c:pieChart>|<c:scatterChart>|<c:areaChart>/);
  let chartType = "chart";
  if (typeMatch) {
    if (typeMatch[0].includes("bar")) chartType = "bar";
    else if (typeMatch[0].includes("line")) chartType = "line";
    else if (typeMatch[0].includes("pie")) chartType = "pie";
    else if (typeMatch[0].includes("scatter")) chartType = "scatter";
    else if (typeMatch[0].includes("area")) chartType = "area";
  }

  // Extract series
  const series: ChartModel["series"] = [];
  const seriesPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g;
  for (const seriesMatch of chartXml.matchAll(seriesPattern)) {
    const seriesContent = seriesMatch[1];

    const seriesNameMatch = seriesContent.match(/<c:tx[^>]*>[\s\S]*?<a:t>([^<]*)<\/a:t>[\s\S]*?<\/c:tx>/);
    const seriesName = seriesNameMatch ? seriesNameMatch[1] : undefined;

    const categoriesMatch = seriesContent.match(/<c:cat[^>]*>[\s\S]*?<c:strRef[^>]*>[\s\S]*?<c:strCache[^>]*>[\s\S]*?<c:ptCount[^>]*val="([^"]*)"[^>]*>/);
    const categories = categoriesMatch ? categoriesMatch[1] : undefined;

    const valuesMatch = seriesContent.match(/<c:val[^>]*>[\s\S]*?<c:numRef[^>]*>[\s\S]*?<c:numCache[^>]*>[\s\S]*?<c:ptCount[^>]*val="([^"]*)"[^>]*>/);
    const values = valuesMatch ? valuesMatch[1] : undefined;

    // Extract series color
    const colorMatch = seriesContent.match(/<c:spPr[^>]*>[\s\S]*?<a:solidFill[^>]*>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>/);
    const color = colorMatch ? colorMatch[1] : undefined;

    series.push({ name: seriesName, categories, values, color });
  }

  // Extract legend
  const legendMatch = chartXml.match(/<c:legend[^>]*>[\s\S]*?<\/c:legend>/);
  const legend = legendMatch ? true : undefined;

  // Extract data labels
  const dataLabelsMatch = chartXml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/);
  const dataLabels = dataLabelsMatch ? "shown" : undefined;

  return {
    path: chartPathStr,
    title,
    type: chartType as ChartModel["type"],
    series,
    legend,
    dataLabels,
  };
}

/**
 * Gets a placeholder by type.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param type - Placeholder type (e.g., "title", "body")
 * @returns Result with the placeholder model
 *
 * @example
 * const result = await getPlaceholder("/path/to/presentation.pptx", 1, "title");
 */
export async function getPlaceholder(
  filePath: string,
  slideIndex: number,
  type: string,
): Promise<Result<PlaceholderModel>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code || "operation_failed", zipResult.error?.message || "Failed to load presentation");
  }
  const zip = zipResult.data!;

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code || "operation_failed", slidePathResult.error?.message || "Failed to get slide entry path");
  }

  const slideEntry = slidePathResult.data!;
  const slideXml = requireEntry(zip, slideEntry);

  // Find placeholder with matching type
  const placeholderPattern = new RegExp(`<p:sp[\\s\\S]*?<p:ph[^>]*type="${type}"[^>]*>[\\s\\S]*?</p:sp>`, "g");
  const placeholders = slideXml.match(placeholderPattern);

  if (!placeholders || placeholders.length === 0) {
    return notFound("Placeholder", type, `No placeholder of type '${type}' found on slide ${slideIndex}`);
  }

  const placeholderXml = placeholders[0];
  const placeholderIndex = 1;

  // Find all placeholders to determine index
  const allPlaceholderPattern = /<p:sp[\s\S]*?<p:ph[^>]*>[\s\S]*?<\/p:sp>/g;
  const allPlaceholders = slideXml.match(allPlaceholderPattern) || [];
  let idx = 0;
  for (const ph of allPlaceholders) {
    const phType = extractPlaceholderType(ph);
    if (phType === type) {
      idx++;
      if (ph === placeholderXml) {
        break;
      }
    }
  }

  const shapeIndexMatch = placeholderXml.match(/<p:nvCxnSpPr[^>]*>[\s\S]*?<p:cNvPr[^>]*idx="([^"]*)"[^>]*>/);
  const idxVal = shapeIndexMatch ? parseInt(shapeIndexMatch[1], 10) : idx;

  const path = `/slide[${slideIndex}]/placeholder[${type}]`;
  const name = extractShapeName(placeholderXml);
  const text = extractTextFromShape(placeholderXml);
  const shape = parseShapeFromXml(placeholderXml, slideIndex, idx);

  return ok({
    path,
    type: type as PlaceholderType,
    index: idxVal,
    name,
    text,
    shape,
  });
}

// ============================================================================
// Query Operations
// ============================================================================

/**
 * Queries elements using a CSS-like selector.
 *
 * @param filePath - Path to the PPTX file
 * @param selector - CSS-like selector (e.g., "slide[1] shape", "shape[type=text]")
 * @returns Result with matching elements
 *
 * @example
 * const result = await query("/path/to/presentation.pptx", "slide[1] shape");
 * if (result.ok) {
 *   console.log(result.data.shapes);
 * }
 */
export async function query(
  filePath: string,
  selector: string,
): Promise<Result<{ shapes: ShapeModel[]; tables: TableModel[]; charts: ChartModel[]; placeholders: PlaceholderModel[] }>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code || "operation_failed", zipResult.error?.message || "Failed to load presentation");
  }
  const zip = zipResult.data!;

  const parsedSelector = parseSelector(selector);
  if (!parsedSelector.ok) {
    return invalidInput(`Invalid selector: ${selector}`);
  }

  const sel = parsedSelector.data!;

  // If slide number is specified, query within that slide
  if (sel.slideNum !== undefined) {
    const slideResult = await queryShapesOnSlide(filePath, sel.slideNum, selector);
    if (!slideResult.ok) {
      return err(slideResult.error?.code || "operation_failed", slideResult.error?.message || "Failed to query shapes");
    }
    const data = slideResult.data!;
    return ok({
      shapes: data.filter((e): e is ShapeModel => "type" in e && e.type === "shape"),
      tables: data.filter((e): e is TableModel => "rows" in e),
      charts: data.filter((e): e is ChartModel => "title" in e || "series" in e),
      placeholders: data.filter((e): e is PlaceholderModel => "index" in e || "name" in e),
    });
  }

  // Query across all slides
  const slidesInfoResult = getSlideInfo(zip);
  if (!slidesInfoResult.ok) {
    return err(slidesInfoResult.error?.code || "operation_failed", slidesInfoResult.error?.message || "Failed to get slides info");
  }
  const slidesInfo = slidesInfoResult.data!;

  const allShapes: ShapeModel[] = [];
  const allTables: TableModel[] = [];
  const allCharts: ChartModel[] = [];
  const allPlaceholders: PlaceholderModel[] = [];

  for (const slideInfo of slidesInfo) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);

    // Parse based on element type in selector
    const elementType = sel.elementType || "shape";

    if (elementType === "shape" || elementType === "textbox") {
      const shapes = parseShapesFromSlideXml(slideXml, slideInfo.index);
      for (const shape of shapes) {
        if (matchesSelector(shape, sel, elementType)) {
          allShapes.push(shape);
        }
      }
    }

    if (elementType === "table") {
      const tables = parseTablesFromSlideXml(slideXml, slideInfo.index);
      for (const table of tables) {
        allTables.push(table);
      }
    }

    if (elementType === "chart") {
      const charts = parseChartsFromSlideXml(zip, slideXml, slideInfo.index);
      for (const chart of charts) {
        allCharts.push(chart);
      }
    }

    if (elementType === "placeholder") {
      const placeholders = parsePlaceholdersFromSlideXml(slideXml, slideInfo.index);
      for (const placeholder of placeholders) {
        if (sel.attributes.name && placeholder.type !== sel.attributes.name) {
          continue;
        }
        allPlaceholders.push(placeholder);
      }
    }
  }

  return ok({
    shapes: allShapes,
    tables: allTables,
    charts: allCharts,
    placeholders: allPlaceholders,
  });
}

/**
 * Queries slides with optional selector filter.
 *
 * @param filePath - Path to the PPTX file
 * @param selector - Optional CSS-like selector to filter slides
 * @returns Result with matching slides
 *
 * @example
 * const result = await querySlides("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(result.data); // SlideModel[]
 * }
 */
export async function querySlides(
  filePath: string,
  selector?: string,
): Promise<Result<SlideModel[]>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code || "operation_failed", zipResult.error?.message || "Failed to load presentation");
  }
  const zip = zipResult.data!;

  const slidesInfoResult = getSlideInfo(zip);
  if (!slidesInfoResult.ok) {
    return err(slidesInfoResult.error?.code || "operation_failed", slidesInfoResult.error?.message || "Failed to get slides info");
  }
  const slidesInfo = slidesInfoResult.data!;

  const slides: SlideModel[] = [];

  for (const slideInfo of slidesInfo) {
    const slideResult = await getSlide(filePath, slideInfo.index);
    if (slideResult.ok) {
      if (selector) {
        const parsed = parseSelector(selector);
        if (parsed.ok) {
          // Filter based on selector
          const sel = parsed.data!;
          if (sel.textContains) {
            // Filter by text content
            const text = slideResult.data!.title || "";
            if (!text.includes(sel.textContains)) {
              continue;
            }
          }
        }
      }
      slides.push(slideResult.data!);
    }
  }

  return ok(slides);
}

/**
 * Queries shapes on a specific slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param selector - Optional CSS-like selector to filter shapes
 * @returns Result with matching shapes
 *
 * @example
 * const result = await queryShapes("/path/to/presentation.pptx", 1, "shape[type=text]");
 */
export async function queryShapes(
  filePath: string,
  slideIndex: number,
  selector?: string,
): Promise<Result<Array<ShapeModel | TableModel | ChartModel | PlaceholderModel>>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code || "operation_failed", zipResult.error?.message || "Failed to load presentation");
  }
  const zip = zipResult.data!;

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code || "operation_failed", slidePathResult.error?.message || "Failed to get slide entry path");
  }

  const slideEntry = slidePathResult.data!;
  const slideXml = requireEntry(zip, slideEntry);

  const results: Array<ShapeModel | TableModel | ChartModel | PlaceholderModel> = [];

  // Parse selector if provided
  let parsedSelector: ParsedSelector | undefined;
  if (selector) {
    const parseResult = parseSelector(selector);
    if (!parseResult.ok) {
      return invalidInput(`Invalid selector: ${selector}`);
    }
    parsedSelector = parseResult.data;
  }

  // Parse shapes
  const shapes = parseShapesFromSlideXml(slideXml, slideIndex);
  for (const shape of shapes) {
    if (!parsedSelector || matchesSelector(shape, parsedSelector, "shape")) {
      results.push(shape);
    }
  }

  // Parse tables
  const tables = parseTablesFromSlideXml(slideXml, slideIndex);
  for (const table of tables) {
    if (!parsedSelector || matchesSelector(table, parsedSelector, "table")) {
      results.push(table);
    }
  }

  // Parse charts
  const charts = parseChartsFromSlideXml(zip, slideXml, slideIndex);
  for (const chart of charts) {
    if (!parsedSelector || matchesSelector(chart, parsedSelector, "chart")) {
      results.push(chart);
    }
  }

  // Parse placeholders
  const placeholders = parsePlaceholdersFromSlideXml(slideXml, slideIndex);
  for (const placeholder of placeholders) {
    if (!parsedSelector || matchesSelector(placeholder, parsedSelector, "placeholder")) {
      results.push(placeholder);
    }
  }

  return ok(results);
}

/**
 * Internal query shapes on slide (uses selector string directly).
 */
async function queryShapesOnSlide(
  filePath: string,
  slideIndex: number,
  selector: string,
): Promise<Result<Array<ShapeModel | TableModel | ChartModel | PlaceholderModel>>> {
  return queryShapes(filePath, slideIndex, selector);
}

/**
 * Checks if an element matches a selector.
 */
function matchesSelector(
  element: ShapeModel | TableModel | ChartModel | PlaceholderModel,
  selector: ParsedSelector,
  elementType: string,
): boolean {
  // Check element type
  if (selector.elementType && selector.elementType !== elementType) {
    return false;
  }

  // Check index attribute
  if (selector.attributes.index) {
    const idx = parseInt(selector.attributes.index, 10);
    if ("index" in element && element.index !== idx) {
      return false;
    }
  }

  // Check name attribute (for placeholders)
  if (selector.attributes.name) {
    if ("type" in element && element.type !== selector.attributes.name) {
      return false;
    }
  }

  // Check text contains
  if (selector.textContains) {
    let text = "";
    if ("text" in element) {
      text = element.text ?? "";
    } else if ("title" in element) {
      text = (element as ChartModel).title ?? "";
    }
    if (!text.includes(selector.textContains)) {
      return false;
    }
  }

  // Check empty
  if (selector.attributes.empty === "true") {
    const text = "text" in element ? element.text : "title" in element ? element.title : "";
    if (text && text.length > 0) {
      return false;
    }
  }

  // Check no-alt
  if (selector.attributes.noAlt === "true") {
    if (!("alt" in element) || !element.alt) {
      return true; // no-alt means shapes without alt text
    }
    return false;
  }

  return true;
}

// ============================================================================
// Element Inspection Operations
// ============================================================================

/**
 * Gets all properties of a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @returns Result with shape properties
 *
 * @example
 * const result = await getShapeProperties("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 * if (result.ok) {
 *   console.log(result.data.x, result.data.y, result.data.fill);
 * }
 */
export async function getShapeProperties(
  filePath: string,
  pptPath: string,
): Promise<Result<{
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rotation?: number;
  fill?: string;
  line?: string;
  lineWidth?: number;
  alt?: string;
  name?: string;
}>> {
  const shapeResult = await getShape(filePath, pptPath);
  if (!shapeResult.ok) {
    return err(shapeResult.error?.code || "operation_failed", shapeResult.error?.message || "Failed to get shape");
  }

  const shape = shapeResult.data!;
  return ok({
    x: shape.x,
    y: shape.y,
    width: shape.width,
    height: shape.height,
    rotation: shape.rotation,
    fill: shape.fill,
    line: shape.line,
    lineWidth: shape.lineWidth,
    alt: shape.alt,
    name: shape.name,
  });
}

/**
 * Gets text content from a shape or placeholder.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the element (e.g., "/slide[1]/shape[1]" or "/slide[1]/placeholder[title]")
 * @returns Result with text content
 *
 * @example
 * const result = await getTextContent("/path/to/presentation.pptx", "/slide[1]/placeholder[title]");
 * if (result.ok) {
 *   console.log(result.data.text);
 * }
 */
export async function getTextContent(
  filePath: string,
  pptPath: string,
): Promise<Result<{
  text: string;
  paragraphs: ParagraphModel[];
}>> {
  // Check if it's a placeholder path
  if (isPlaceholderPath(pptPath)) {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("Invalid path");
    }

    // Extract placeholder type from path
    const typeMatch = pptPath.match(/\/placeholder\[([^\]]+)\]/i);
    if (!typeMatch) {
      return invalidInput("Invalid placeholder path");
    }
    const type = typeMatch[1];

    const placeholderResult = await getPlaceholder(filePath, slideIndex, type);
    if (!placeholderResult.ok) {
      return err(placeholderResult.error?.code || "operation_failed", placeholderResult.error?.message || "Failed to get placeholder");
    }

    const placeholder = placeholderResult.data!;
    return ok({
      text: placeholder.text || "",
      paragraphs: placeholder.shape?.paragraphs || [],
    });
  }

  // Otherwise, treat as shape
  const shapeResult = await getShape(filePath, pptPath);
  if (!shapeResult.ok) {
    return err(shapeResult.error?.code || "operation_failed", shapeResult.error?.message || "Failed to get shape");
  }

  const shape = shapeResult.data!;
  return ok({
    text: shape.text || "",
    paragraphs: shape.paragraphs || [],
  });
}

/**
 * Gets the structure of a table.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the table (e.g., "/slide[1]/table[1]")
 * @returns Result with table structure
 *
 * @example
 * const result = await getTableStructure("/path/to/presentation.pptx", "/slide[1]/table[1]");
 * if (result.ok) {
 *   console.log(result.data.rowCount, result.data.columnCount);
 *   console.log(result.data.rows[0].cells[0].text);
 * }
 */
export async function getTableStructure(
  filePath: string,
  pptPath: string,
): Promise<Result<{
  path: string;
  name?: string;
  rowCount?: number;
  columnCount?: number;
  hasHeaderRow?: boolean;
  rows: TableRowModel[];
}>> {
  const tableResult = await getTable(filePath, pptPath);
  if (!tableResult.ok) {
    return err(tableResult.error?.code || "operation_failed", tableResult.error?.message || "Failed to get table");
  }

  const table = tableResult.data!;
  return ok({
    path: table.path,
    name: table.name,
    rowCount: table.rowCount,
    columnCount: table.columnCount,
    hasHeaderRow: table.hasHeaderRow,
    rows: table.rows,
  });
}

/**
 * Gets chart data and series information.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the chart (e.g., "/slide[1]/chart[1]")
 * @returns Result with chart data
 *
 * @example
 * const result = await getChartData("/path/to/presentation.pptx", "/slide[1]/chart[1]");
 * if (result.ok) {
 *   console.log(result.data.type);
 *   console.log(result.data.series);
 * }
 */
export async function getChartData(
  filePath: string,
  pptPath: string,
): Promise<Result<{
  path: string;
  title?: string;
  type?: string;
  series?: ChartModel["series"];
  legend?: string | boolean;
  dataLabels?: string;
}>> {
  const chartResult = await getChart(filePath, pptPath);
  if (!chartResult.ok) {
    return err(chartResult.error?.code || "operation_failed", chartResult.error?.message || "Failed to get chart");
  }

  const chart = chartResult.data!;
  return ok({
    path: chart.path,
    title: chart.title,
    type: chart.type,
    series: chart.series,
    legend: chart.legend,
    dataLabels: chart.dataLabels,
  });
}

// ============================================================================
// Internal Parsing Functions
// ============================================================================

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
 * Parses all tables from slide XML.
 */
function parseTablesFromSlideXml(slideXml: string, slideIndex: number): TableModel[] {
  const tables: TableModel[] = [];

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
    const tableMatch = frameContent.match(/<a:tbl>([\s\S]*?)<\/a:tbl>/);
    if (tableMatch) {
      tables.push(parseTable(tableMatch[0], slideIndex, tableIndex));
    }
  }

  // Also match tables directly in shapes (rare but possible)
  const directTablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  for (const tableMatch of slideXml.matchAll(directTablePattern)) {
    // Check if this table is already captured in a graphic frame
    if (slideXml.includes("<p:graphicFrame") && tableMatch[0].includes("<a:tbl>")) {
      // Already handled
      continue;
    }
    tableIndex++;
    tables.push(parseTable(tableMatch[0], slideIndex, tableIndex));
  }

  return tables;
}

/**
 * Parses all charts from slide XML.
 */
function parseChartsFromSlideXml(
  zip: Map<string, Buffer>,
  slideXml: string,
  slideIndex: number,
): ChartModel[] {
  const charts: ChartModel[] = [];

  // Find chart relationships in slide
  const slideEntryPath = slideXml; // We need the actual path, but we don't have it here
  // This is a limitation - we'd need to pass the slide entry path

  // For now, we'll look for chart placeholders
  const chartPattern = /<p:graphicFrame(?:[\s\S]*?)<\/p:graphicFrame>/g;
  let chartIndex = 0;

  for (const frameMatch of slideXml.matchAll(chartPattern)) {
    const frameContent = frameMatch[0];

    // Check if it contains a chart reference
    if (!frameContent.includes("<c:chart")) {
      continue;
    }

    chartIndex++;
    charts.push({
      path: `/slide[${slideIndex}]/chart[${chartIndex}]`,
      type: "chart",
    });
  }

  return charts;
}

/**
 * Parses all placeholders from slide XML.
 */
function parsePlaceholdersFromSlideXml(slideXml: string, slideIndex: number): PlaceholderModel[] {
  const placeholders: PlaceholderModel[] = [];

  // Match placeholder shapes
  const placeholderPattern = /<p:sp(?:[\s\S]*?)<\/p:sp>/g;
  let placeholderIndex = 0;

  for (const shapeMatch of slideXml.matchAll(placeholderPattern)) {
    const shapeXml = shapeMatch[0];

    // Check if this is a placeholder
    if (!shapeXml.includes("<p:ph")) {
      continue;
    }

    placeholderIndex++;

    const phType = extractPlaceholderType(shapeXml);
    if (!phType) {
      continue;
    }

    const phIndex = extractPlaceholderIndex(shapeXml);
    const name = extractShapeName(shapeXml);
    const text = extractTextFromShape(shapeXml);
    const path = `/slide[${slideIndex}]/placeholder[${phType}]`;

    // Find shape index for the underlying shape
    const allShapes = slideXml.match(/<p:sp(?:[\s\S]*?)<\/p:sp>/g) || [];
    let shapeIdx = 0;
    for (const s of allShapes) {
      shapeIdx++;
      if (s === shapeXml) {
        break;
      }
    }

    placeholders.push({
      path,
      type: phType as PlaceholderType,
      index: phIndex,
      name,
      text,
      shape: parseShapeFromXml(shapeXml, slideIndex, shapeIdx),
    });
  }

  return placeholders;
}
