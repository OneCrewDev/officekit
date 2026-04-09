/**
 * Word HTML Preview - Main Entry Point
 *
 * Renders a Word document as full-featured HTML with page layout,
 * styles, tables, images, headers/footers, and footnotes/endnotes.
 */

import JSZip from "jszip";

import { createHtmlRenderContext, type HtmlRenderContext, type HtmlPreviewOptions } from "./context.js";
import { generateWordCss } from "./css.js";
import { renderParagraphHtml } from "./text.js";
import { renderTableHtml } from "./tables.js";
import { renderDrawingHtml, resolveImageFromRelId } from "./shapes.js";
import { renderFootnotesHtml, renderEndnotesHtml, renderHeaderHtml, renderFooterHtml } from "./notes.js";

// ============================================================================
// Main Entry Point
// ============================================================================

/**
 * Render a Word document as HTML preview.
 *
 * @param zip - JSZip containing the docx contents
 * @param documentXml - The document.xml content
 * @param stylesXml - The styles.xml content (optional)
 * @param options - Rendering options
 * @returns Complete HTML document as string
 */
export async function renderHtmlPreview(
  zip: JSZip,
  documentXml: string,
  stylesXml: string = "",
  options: HtmlPreviewOptions = {}
): Promise<string> {
  const {
    pageFilter,
    includeStyles = true,
    customCss,
  } = options;

  // Create rendering context
  const context = createHtmlRenderContext();

  // Parse page layout from document
  const pageLayout = parsePageLayout(documentXml);
  context.pageLayout = pageLayout;

  // Load relationships for image resolution
  const relsXml = await loadRelationships(zip);
  const relsMap = parseRelationships(relsXml);

  // Load header/footer XML files
  const headers = await loadHeadersAndFooters(zip, documentXml, "header");
  const footers = await loadHeadersAndFooters(zip, documentXml, "footer");

  // Load footnotes and endnotes
  const footnotesXml = await loadFootnotes(zip);
  const endnotesXml = await loadEndnotes(zip);

  // Parse document relationships for images
  const docRelsXml = await getXmlEntry(zip, "word/_rels/document.xml.rels");
  const docRels = docRelsXml || "";

  // Build HTML document
  const lines: string[] = [];

  // DOCTYPE and html tag
  lines.push("<!DOCTYPE html>");
  lines.push('<html lang="en">');
  lines.push("<head>");
  lines.push('<meta charset="UTF-8">');
  lines.push('<meta name="viewport" content="width=device-width, initial-scale=1.0">');
  lines.push("<title>Word Document Preview</title>");

  // CSS
  lines.push("<style>");
  if (includeStyles) {
    lines.push(generateWordCss({
      widthPt: pageLayout.pageWidthTwips / 20,
      heightPt: pageLayout.pageHeightTwips / 20,
      marginTopPt: pageLayout.marginTopTwips / 20,
      marginBottomPt: pageLayout.marginBottomTwips / 20,
      marginLeftPt: pageLayout.marginLeftTwips / 20,
      marginRightPt: pageLayout.marginRightTwips / 20,
    }));
  }
  if (customCss) {
    lines.push(customCss);
  }
  lines.push("</style>");

  lines.push("</head>");
  lines.push("<body>");

  // Determine which pages to render
  const pagesToRender = parsePageFilter(pageFilter);

  // Get body content
  const bodyContent = getBodyContent(documentXml);
  context.renderingBody = true;

  // Determine if we're rendering all or filtered pages
  const allParagraphs = getParagraphs(documentXml);
  const allTables = getTables(documentXml);

  // Apply page filter if specified
  const filteredIndices = pagesToRender
    ? getPageFilteredIndices(allParagraphs, pagesToRender, pageLayout)
    : null;

  // Render each paragraph
  let paragraphIndex = 0;
  let tableIndex = 0;
  let pageNum = 1;

  // In single page mode, content flows naturally without page wrapper
  const usePageWrapper = !options.singlePage;

  // Calculate available page height for pagination
  const pageHeightPt = pageLayout.pageHeightTwips / 20; // Convert twips to points
  const marginTopPt = pageLayout.marginTopTwips / 20;
  const marginBottomPt = pageLayout.marginBottomTwips / 20;
  const headerHeightPt = headers.default ? 30 : 0; // Approximate header height
  const footerHeightPt = footers.default ? 30 : 0; // Approximate footer height
  const availableHeightPt = pageHeightPt - marginTopPt - marginBottomPt - headerHeightPt - footerHeightPt;

  // Current page accumulated height
  let currentPageHeightPt = 0;

  // Helper function to estimate paragraph height
  const estimateParaHeight = (paraXml: string): number => {
    // Extract text length
    const texts = paraXml.match(/<w:t[^>]*>[^<]*<\/w:t>/gi) || [];
    const totalChars = texts.join("").replace(/<[^>]*>/g, "").length;

    // Get line height from para properties
    const lineSpacingMatch = /<w:spacing[^>]*w:line="([^"]*)"[^>]*>/i.exec(paraXml);
    let lineHeight = 14; // Default 14pt per line
    if (lineSpacingMatch) {
      const lineVal = parseInt(lineSpacingMatch[1], 10);
      if (lineVal > 240) {
        // Exact or atLeast value in twips
        lineHeight = lineVal / 20;
      } else {
        // Auto line spacing (240 = single, 480 = double)
        lineHeight = lineVal / 240 * 14;
      }
    }

    // Get font size
    const fontSizeMatch = /<w:sz[^>]*w:val="([^"]*)"/i.exec(paraXml);
    const fontSize = fontSizeMatch ? parseInt(fontSizeMatch[1], 10) / 2 : 11;

    // Estimate lines based on character count (approx 12 chars per inch at 14pt)
    const charsPerLine = Math.floor(7 * 12); // ~7 inch content width
    const estimatedLines = Math.max(1, Math.ceil(totalChars / charsPerLine));

    return estimatedLines * Math.max(lineHeight, fontSize) + 6; // Add some padding
  };

  // Helper to start a new page
  const startNewPage = () => {
    if (pageNum > 1) {
      lines.push("</div>");
    }
    lines.push('<div class="page">');
    if (headers.default) {
      lines.push(renderHeaderHtml(headers.default, context));
    }
    currentPageHeightPt = 0;
    pageNum++;
  };

  if (usePageWrapper) {
    startNewPage();
  }

  for (const item of bodyContent) {
    if (item.type === "paragraph") {
      paragraphIndex++;

      // Skip if page filtering is active and this paragraph is not in selected pages
      if (filteredIndices && !filteredIndices.has(paragraphIndex)) {
        continue;
      }

      const paraHtml = renderParagraphHtml(item.xml, item.index, context);
      const paraHeight = estimateParaHeight(item.xml);

      // Check for explicit page break
      if (item.hasPageBreak && usePageWrapper) {
        lines.push(paraHtml);
        startNewPage();
        continue;
      }

      // Check if paragraph fits on current page
      if (usePageWrapper && currentPageHeightPt + paraHeight > availableHeightPt) {
        startNewPage();
      }

      lines.push(paraHtml);
      currentPageHeightPt += paraHeight;
    } else if (item.type === "table") {
      tableIndex++;

      // Skip if page filtering is active
      if (filteredIndices && !filteredIndices.has(-tableIndex)) {
        continue;
      }

      const tableHtml = renderTable(item.xml);

      // Estimate table height (rough: 20pt per row + padding)
      const rowCount = (item.xml.match(/<w:tr[>\s]/gi) || []).length;
      const tableHeight = rowCount * 25 + 20;

      // Check for page break before table
      if (item.hasPageBreakBefore && usePageWrapper) {
        startNewPage();
      }

      // Check if table fits on current page
      if (usePageWrapper && currentPageHeightPt + tableHeight > availableHeightPt) {
        startNewPage();
      }

      lines.push(tableHtml);
      currentPageHeightPt += tableHeight;
    } else if (item.type === "section") {
      // Handle section break
      if (item.properties && usePageWrapper) {
        const newPageLayout = parseSectionProperties(item.properties);
        if (newPageLayout) {
          startNewPage();
        }
      }
    }
  }

  // Flush remaining content
  // (content is already flushed incrementally)

  // Render footer
  if (footers.default) {
    lines.push(renderFooterHtml(footers.default, context));
  }

  if (usePageWrapper) {
    lines.push("</div>");
  }

  // Render footnotes
  if (footnotesXml) {
    lines.push(renderFootnotesHtml(footnotesXml, context));
  }

  // Render endnotes
  if (endnotesXml) {
    lines.push(renderEndnotesHtml(endnotesXml, context));
  }

  lines.push("</body>");
  lines.push("</html>");

  return lines.join("\n");
}

// ============================================================================
// Helper Functions
// ============================================================================

interface BodyContentItem {
  type: "paragraph" | "table" | "section";
  index: number;
  xml: string;
  hasPageBreak?: boolean;
  hasPageBreakBefore?: boolean;
  properties?: string;
}

/**
 * Get body content as paragraphs and tables.
 */
function getBodyContent(xml: string): BodyContentItem[] {
  const content: BodyContentItem[] = [];

  // Get body element
  const bodyMatch = /<w:body\b[^>]*>([\s\S]*?)<\/w:body>/i.exec(xml);
  if (!bodyMatch) return content;

  const bodyXml = bodyMatch[1];
  let cursor = 0;

  // Process paragraphs
  const paraRegex = /<w:p[\s\S]*?<\/w:p>/gi;
  let paraMatch;

  while ((paraMatch = paraRegex.exec(bodyXml)) !== null) {
    const paraXml = paraMatch[0];
    const hasPageBreak = /<w:pageBreakBefore[^>]*>/i.test(paraXml)
      || /<w:br[^>]*w:type="page"/i.test(paraXml);

    content.push({
      type: "paragraph",
      index: content.filter(c => c.type === "paragraph").length + 1,
      xml: paraXml,
      hasPageBreak,
    });
  }

  // Process tables
  const tableRegex = /<w:tbl[\s\S]*?<\/w:tbl>/gi;
  let tableMatch;

  while ((tableMatch = tableRegex.exec(bodyXml)) !== null) {
    const tableXml = tableMatch[0];

    // Check for page break before table
    const prevContent = bodyXml.substring(
      Math.max(0, tableMatch.index - 500),
      tableMatch.index
    );
    const hasPageBreakBefore = /<w:pageBreakBefore[^>]*>/i.test(prevContent);

    content.push({
      type: "table",
      index: content.filter(c => c.type === "table").length + 1,
      xml: tableXml,
      hasPageBreakBefore,
    });
  }

  // Sort by position in document
  content.sort((a, b) => {
    const aPos = bodyXml.indexOf(a.xml);
    const bPos = bodyXml.indexOf(b.xml);
    return aPos - bPos;
  });

  // Re-index after sorting
  let paraIdx = 0;
  let tableIdx = 0;
  for (const item of content) {
    if (item.type === "paragraph") {
      item.index = ++paraIdx;
    } else if (item.type === "table") {
      item.index = ++tableIdx;
    }
  }

  return content;
}

/**
 * Get all paragraphs from document.
 */
function getParagraphs(xml: string): string[] {
  const paras: string[] = [];
  const regex = /<w:p[\s\S]*?<\/w:p>/gi;
  let match;

  while ((match = regex.exec(xml)) !== null) {
    paras.push(match[0]);
  }

  return paras;
}

/**
 * Get all tables from document.
 */
function getTables(xml: string): string[] {
  const tables: string[] = [];
  const regex = /<w:tbl[\s\S]*?<\/w:tbl>/gi;
  let match;

  while ((match = regex.exec(xml)) !== null) {
    tables.push(match[0]);
  }

  return tables;
}

/**
 * Parse page layout from document sectPr.
 */
function parsePageLayout(xml: string): {
  pageWidthTwips: number;
  pageHeightTwips: number;
  marginTopTwips: number;
  marginBottomTwips: number;
  marginLeftTwips: number;
  marginRightTwips: number;
  orientation?: string;
  columns?: number;
  columnSpaceTwips?: number;
} {
  // Default US Letter size in twips (1 inch = 1440 twips)
  const defaults = {
    pageWidthTwips: 12240, // 8.5 inches
    pageHeightTwips: 15840, // 11 inches
    marginTopTwips: 1440, // 1 inch
    marginBottomTwips: 1440,
    marginLeftTwips: 1440,
    marginRightTwips: 1440,
  };

  // Find sectPr
  const sectPrMatch = /<w:sectPr[\s\S]*?<\/w:sectPr>/i.exec(xml);
  if (!sectPrMatch) return defaults;

  const sectPrContent = sectPrMatch[0];

  // Page size
  const pgSzMatch = /<w:pgSz[^>]*>/i.exec(sectPrContent);
  let pageWidth = defaults.pageWidthTwips;
  let pageHeight = defaults.pageHeightTwips;
  let orientation: string | undefined;

  if (pgSzMatch) {
    const wMatch = /w:w="([^"]*)"/i.exec(pgSzMatch[0]);
    const hMatch = /w:h="([^"]*)"/i.exec(pgSzMatch[0]);
    const orientMatch = /w:orient="([^"]*)"/i.exec(pgSzMatch[0]);

    if (wMatch) pageWidth = parseInt(wMatch[1], 10);
    if (hMatch) pageHeight = parseInt(hMatch[1], 10);
    if (orientMatch) orientation = orientMatch[1];

    // Swap dimensions for landscape
    if (orientation === "landscape") {
      [pageWidth, pageHeight] = [pageHeight, pageWidth];
    }
  }

  // Margins
  const pgMarMatch = /<w:pgMar[^>]*>/i.exec(sectPrContent);
  let marginTop = defaults.marginTopTwips;
  let marginBottom = defaults.marginBottomTwips;
  let marginLeft = defaults.marginLeftTwips;
  let marginRight = defaults.marginRightTwips;

  if (pgMarMatch) {
    const topMatch = /w:top="([^"]*)"/i.exec(pgMarMatch[0]);
    const bottomMatch = /w:bottom="([^"]*)"/i.exec(pgMarMatch[0]);
    const leftMatch = /w:left="([^"]*)"/i.exec(pgMarMatch[0]);
    const rightMatch = /w:right="([^"]*)"/i.exec(pgMarMatch[0]);

    if (topMatch) marginTop = parseInt(topMatch[1], 10);
    if (bottomMatch) marginBottom = parseInt(bottomMatch[1], 10);
    if (leftMatch) marginLeft = parseInt(leftMatch[1], 10);
    if (rightMatch) marginRight = parseInt(rightMatch[1], 10);
  }

  // Columns
  const colsMatch = /<w:cols[^>]*>/i.exec(sectPrContent);
  let columns: number | undefined;
  let columnSpace: number | undefined;

  if (colsMatch) {
    const numMatch = /w:num="([^"]*)"/i.exec(colsMatch[0]);
    const spaceMatch = /w:space="([^"]*)"/i.exec(colsMatch[0]);

    if (numMatch) columns = parseInt(numMatch[1], 10);
    if (spaceMatch) columnSpace = parseInt(spaceMatch[1], 10);
  }

  return {
    pageWidthTwips: pageWidth,
    pageHeightTwips: pageHeight,
    marginTopTwips: marginTop,
    marginBottomTwips: marginBottom,
    marginLeftTwips: marginLeft,
    marginRightTwips: marginRight,
    orientation,
    columns,
    columnSpaceTwips: columnSpace,
  };
}

/**
 * Parse section properties for page layout changes.
 */
function parseSectionProperties(sectPrContent: string): {
  pageWidthTwips: number;
  pageHeightTwips: number;
  marginTopTwips: number;
  marginBottomTwips: number;
  marginLeftTwips: number;
  marginRightTwips: number;
} | null {
  const pgSzMatch = /<w:pgSz[^>]*>/i.exec(sectPrContent);
  if (!pgSzMatch) return null;

  const result = {
    pageWidthTwips: 12240,
    pageHeightTwips: 15840,
    marginTopTwips: 1440,
    marginBottomTwips: 1440,
    marginLeftTwips: 1440,
    marginRightTwips: 1440,
  };

  const wMatch = /w:w="([^"]*)"/i.exec(pgSzMatch[0]);
  const hMatch = /w:h="([^"]*)"/i.exec(pgSzMatch[0]);

  if (wMatch) result.pageWidthTwips = parseInt(wMatch[1], 10);
  if (hMatch) result.pageHeightTwips = parseInt(hMatch[1], 10);

  const pgMarMatch = /<w:pgMar[^>]*>/i.exec(sectPrContent);
  if (pgMarMatch) {
    const topMatch = /w:top="([^"]*)"/i.exec(pgMarMatch[0]);
    const bottomMatch = /w:bottom="([^"]*)"/i.exec(pgMarMatch[0]);
    const leftMatch = /w:left="([^"]*)"/i.exec(pgMarMatch[0]);
    const rightMatch = /w:right="([^"]*)"/i.exec(pgMarMatch[0]);

    if (topMatch) result.marginTopTwips = parseInt(topMatch[1], 10);
    if (bottomMatch) result.marginBottomTwips = parseInt(bottomMatch[1], 10);
    if (leftMatch) result.marginLeftTwips = parseInt(leftMatch[1], 10);
    if (rightMatch) result.marginRightTwips = parseInt(rightMatch[1], 10);
  }

  return result;
}

/**
 * Load relationships from zip.
 */
async function loadRelationships(zip: JSZip): Promise<string> {
  return await getXmlEntry(zip, "word/_rels/document.xml.rels") || "";
}

/**
 * Parse relationships into a map.
 */
function parseRelationships(relsXml: string): Map<string, string> {
  const map = new Map<string, string>();
  const regex = /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*>/gi;
  let match;

  while ((match = regex.exec(relsXml)) !== null) {
    map.set(match[1], match[2]);
  }

  return map;
}

/**
 * Load headers and footers from zip.
 */
async function loadHeadersAndFooters(
  zip: JSZip,
  documentXml: string,
  type: "header" | "footer"
): Promise<Record<string, string>> {
  const result: Record<string, string> = {};

  // Find references in document
  const refRegex = type === "header"
    ? /<w:headerReference[^>]*w:type="([^"]*)"[^>]*r:id="([^"]*)"/gi
    : /<w:footerReference[^>]*w:type="([^"]*)"[^>]*r:id="([^"]*)"/gi;

  const relsMap = parseRelationships(await loadRelationships(zip));
  let match;

  while ((match = refRegex.exec(documentXml)) !== null) {
    const headerType = match[1];
    const relId = match[2];

    // Get target from relationships
    const target = relsMap.get(relId);
    if (!target) continue;

    // Load the header/footer file
    let filePath = target;
    if (!filePath.startsWith("/") && !filePath.includes(":")) {
      filePath = "word/" + filePath;
    }

    const content = await getXmlEntry(zip, filePath);
    if (content) {
      result[headerType] = content;
    }
  }

  return result;
}

/**
 * Load footnotes from zip.
 */
async function loadFootnotes(zip: JSZip): Promise<string | null> {
  return await getXmlEntry(zip, "word/footnotes.xml");
}

/**
 * Load endnotes from zip.
 */
async function loadEndnotes(zip: JSZip): Promise<string | null> {
  return await getXmlEntry(zip, "word/endnotes.xml");
}

/**
 * Get XML entry from zip.
 */
async function getXmlEntry(zip: JSZip, entryName: string): Promise<string | null> {
  const entry = zip.file(entryName);
  if (!entry) return null;
  return await entry.async("string");
}

/**
 * Parse page filter string.
 */
function parsePageFilter(filter?: string): Set<number> | null {
  if (!filter) return null;

  const pages = new Set<number>();
  const parts = filter.split(",").map(s => s.trim());

  for (const part of parts) {
    if (part.includes("-")) {
      // Range like "2-5"
      const [start, end] = part.split("-").map(s => parseInt(s.trim(), 10));
      if (!isNaN(start) && !isNaN(end)) {
        for (let i = start; i <= end; i++) {
          pages.add(i);
        }
      }
    } else {
      // Single page
      const pageNum = parseInt(part, 10);
      if (!isNaN(pageNum)) {
        pages.add(pageNum);
      }
    }
  }

  return pages.size > 0 ? pages : null;
}

/**
 * Get paragraph indices that belong to specific pages.
 */
function getPageFilteredIndices(
  paragraphs: string[],
  pages: Set<number>,
  _pageLayout: { pageWidthTwips: number; pageHeightTwips: number; marginTopTwips: number; marginBottomTwips: number }
): Set<number> | null {
  // Simple approximation: each page holds ~45-50 lines of text
  // This is a rough heuristic for page filtering
  const linesPerPage = 45;
  const indices = new Set<number>();

  let currentPage = 1;
  let lineCount = 0;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];

    // Estimate lines in this paragraph
    const textLength = (para.match(/<w:t[^>]*>[^<]*<\/w:t>/gi) || []).join("").length;
    const estimatedLines = Math.max(1, Math.ceil(textLength / 80));

    lineCount += estimatedLines;

    if (lineCount >= linesPerPage * currentPage) {
      currentPage++;
    }

    if (pages.has(currentPage)) {
      indices.add(i + 1);
    }
  }

  return indices;
}

/**
 * Render a table - wrapper for table module.
 */
function renderTable(tableXml: string): string {
  // Extract table style if present
  const styleMatch = /<w:tblStyle[^>]*w:val="([^"]*)"/i.exec(tableXml);
  const styleId = styleMatch ? styleMatch[1] : undefined;

  // Determine if borderless
  const bordersMatch = /<w:tblBorders>([\s\S]*?)<\/w:tblBorders>/i.exec(tableXml);
  let borderless = false;
  if (bordersMatch) {
    borderless = /<w:top[^>]*w:val="none"[^>]*\/>/i.test(bordersMatch[1])
      && /<w:left[^>]*w:val="none"[^>]*\/>/i.test(bordersMatch[1])
      && /<w:bottom[^>]*w:val="none"[^>]*\/>/i.test(bordersMatch[1])
      && /<w:right[^>]*w:val="none"[^>]*\/>/i.test(bordersMatch[1]);
  }

  return renderTableHtml(tableXml, {
    borderless,
    styleId,
  });
}

// ============================================================================
// Re-export types for external use
// ============================================================================

export type { HtmlRenderContext, HtmlPreviewOptions } from "./context.js";
