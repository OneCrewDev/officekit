/**
 * SVG Preview rendering for @officekit/ppt.
 *
 * Provides functions to render PowerPoint presentations as SVG:
 * - viewAsSvg: Renders entire presentation or specific slide as SVG
 */

import { readFile } from "node:fs/promises";
import path from "node:path";
import { readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput } from "./result.js";
import type { Result } from "./types.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Result from SVG preview - contains the SVG string.
 */
export interface ViewSvgResult {
  /** Total slide count in the preview */
  slideCount: number;
  /** The rendered SVG string */
  svg: string;
}

// ============================================================================
// Constants
// ============================================================================

/** EMU to Points conversion factor (914400 EMUs = 72 points) */
const EMU_TO_PT = 914400;

/** Default slide width in points */
const DEFAULT_SLIDE_WIDTH = 720;

/** Default slide height in points */
const DEFAULT_SLIDE_HEIGHT = 540;

// ============================================================================
// Helper Functions (duplicated for self-containment)
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
 * Converts EMUs to points.
 */
function emuToPt(emu: number): number {
  return emu / EMU_TO_PT * 72;
}

/**
 * Gets slide dimensions from presentation.xml.
 */
function getSlideDimensions(zip: Map<string, Buffer>): { width: number; height: number } {
  try {
    const presXml = requireEntry(zip, "ppt/presentation.xml");
    const sldSzMatch = presXml.match(/<p:sldSz\s[^>]*cx="(\d+)"[^>]*cy="(\d+)"[^>]*>/);
    if (sldSzMatch) {
      return {
        width: emuToPt(parseInt(sldSzMatch[1], 10)),
        height: emuToPt(parseInt(sldSzMatch[2], 10)),
      };
    }
  } catch {
    // Fall through to defaults
  }
  return { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT };
}

/**
 * Gets theme colors from the presentation.
 */
function getThemeColors(zip: Map<string, Buffer>): Map<string, string> {
  const colors = new Map<string, string>();

  try {
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);

    const themeRel = relationships.find(r =>
      r.type?.includes("theme") || r.target?.includes("theme")
    );

    if (themeRel) {
      const themePath = normalizeZipPath("ppt", themeRel.target);
      const themeXml = requireEntry(zip, themePath);

      const colorSchemeMatch = themeXml.match(/<a:clrScheme([^>]*)>([\s\S]*?)<\/a:clrScheme>/);
      if (colorSchemeMatch) {
        const schemeContent = colorSchemeMatch[2];

        const colorPatterns: [string, RegExp][] = [
          ["dk1", /<a:dk1>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:dk1>/],
          ["dk2", /<a:dk2>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:dk2>/],
          ["lt1", /<a:lt1>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:lt1>/],
          ["lt2", /<a:lt2>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:lt2>/],
          ["accent1", /<a:accent1>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:accent1>/],
          ["accent2", /<a:accent2>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:accent2>/],
          ["accent3", /<a:accent3>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:accent3>/],
          ["accent4", /<a:accent4>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:accent4>/],
          ["accent5", /<a:accent5>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:accent5>/],
          ["accent6", /<a:accent6>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:accent6>/],
        ];

        for (const [name, pattern] of colorPatterns) {
          const match = schemeContent.match(pattern);
          if (match) {
            colors.set(name, match[1]);
          }
        }
      }
    }
  } catch {
    // Use default colors if theme parsing fails
  }

  if (!colors.has("dk1")) colors.set("dk1", "000000");
  if (!colors.has("dk2")) colors.set("dk2", "333333");
  if (!colors.has("lt1")) colors.set("lt1", "FFFFFF");
  if (!colors.has("lt2")) colors.set("lt2", "EEEEEE");
  if (!colors.has("accent1")) colors.set("accent1", "4472C4");

  return colors;
}

/**
 * Escapes XML special characters.
 */
function xmlEncode(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Gets slide background as SVG rect.
 */
function getSlideBackgroundSvg(slideXml: string): string {
  const bgMatch = slideXml.match(/<p:bg([\s\S]*?)<\/p:bg>/);
  if (!bgMatch || !bgMatch[1]) return "";

  const bgContent = bgMatch[1];

  const bgPrMatch = bgContent.match(/<p:bgPr([\s\S]*?)<\/p:bgPr>/);
  if (bgPrMatch && bgPrMatch[1]) {
    const bgPrContent = bgPrMatch[1];
    const solidFill = bgPrContent.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
    if (solidFill && solidFill[1]) {
      const colorMatch = solidFill[1].match(/<a:srgbClr[^>]*val="([^"]*)"[^>]*>/);
      if (colorMatch && colorMatch[1]) {
        return `<rect width="100%" height="100%" fill="#${colorMatch[1]}"/>`;
      }
    }
  }

  return "";
}

/**
 * Extracts shape text content as SVG tspan elements.
 */
function extractTextRunsAsSvg(shapeXml: string, themeColors: Map<string, string>): { text: string; tspans: string } {
  const tspans: string[] = [];

  // Match a:r elements
  const runPattern = /<a:r([\s\S]*?)<\/a:r>/g;
  for (const runMatch of shapeXml.matchAll(runPattern)) {
    if (!runMatch[1]) continue;
    const runContent = runMatch[1];

    const textMatch = runContent.match(/<a:t>([^<]*)<\/a:t>/);
    if (!textMatch || !textMatch[1]) continue;

    const text = textMatch[1];
    if (!text) continue;

    let tspan = `<tspan>`;
    const lines = text.split(/\n/);
    for (let i = 0; i < lines.length; i++) {
      if (i > 0) {
        tspan += `</tspan><tspan x="0" dy="1.2em">`;
      }
      tspan += xmlEncode(lines[i]);
    }
    tspan += `</tspan>`;

    // Build full tspan with styles
    let fullTspan = `<tspan`;

    const rPrMatch = runContent.match(/<a:rPr([^>]*)>([\s\S]*?)<\/a:rPr>/);
    if (rPrMatch && rPrMatch[2]) {
      const rPrContent = rPrMatch[2];

      // Bold
      if (rPrContent.includes("<a:b/>") || rPrContent.match(/<a:rPr[^>]*b="1"/)) {
        fullTspan += ` font-weight="bold"`;
      }

      // Italic
      if (rPrContent.includes("<a:i/>") || rPrContent.match(/<a:rPr[^>]*i="1"/)) {
        fullTspan += ` font-style="italic"`;
      }

      // Font size
      const fontSizeMatch = rPrContent.match(/<a:defRPr[^>]*sz="(\d+)"/) || rPrContent.match(/<a:latin[^>]*sz="(\d+)"/);
      if (fontSizeMatch && fontSizeMatch[1]) {
        const size = parseInt(fontSizeMatch[1], 10) / 100;
        fullTspan += ` font-size="${size}"`;
      }

      // Font color
      const solidFillMatch = rPrContent.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
      if (solidFillMatch && solidFillMatch[1]) {
        const colorMatch = solidFillMatch[1].match(/<a:srgbClr[^>]*val="([^"]*)"[^>]*>/);
        if (colorMatch && colorMatch[1]) {
          fullTspan += ` fill="#${colorMatch[1]}"`;
        }
      } else {
        // Use default text color
        const defaultColor = themeColors.get("dk1") || "000000";
        fullTspan += ` fill="#${defaultColor}"`;
      }

      // Font face
      const fontFaceMatch = rPrContent.match(/<a:latin[^>]*typeface="([^"]*)"[^>]*>/);
      if (fontFaceMatch && fontFaceMatch[1]) {
        fullTspan += ` font-family="'${fontFaceMatch[1]}', sans-serif"`;
      }
    } else {
      // Default styling
      const defaultColor = themeColors.get("dk1") || "000000";
      fullTspan += ` fill="#${defaultColor}"`;
    }

    fullTspan += `>`;
    fullTspan += tspan.replace(/^<tspan>/, "").replace(/<\/tspan>$/, "");
    fullTspan += `</tspan>`;

    tspans.push(fullTspan);
  }

  const allText = tspans.join("");
  return { text: tspans.map(t => t.replace(/<[^>]+>/g, "")).join(""), tspans: allText };
}

/**
 * Extracts shape properties (position, size, fill, line) from shape XML.
 */
function extractShapeProps(shapeXml: string): {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  fill?: string;
  line?: string;
  lineWidth?: number;
  name?: string;
} {
  const props: ReturnType<typeof extractShapeProps> = {};

  const spPrMatch = shapeXml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);
  if (spPrMatch && spPrMatch[1]) {
    const spPrContent = spPrMatch[1];

    const xfrmMatch = spPrContent.match(/<a:xfrm(?:[^>]*)>([\s\S]*?)<\/a:xfrm>/);
    if (xfrmMatch && xfrmMatch[1]) {
      const xfrmContent = xfrmMatch[1];

      const offMatch = xfrmContent.match(/<a:off[^>]*x="([^"]*)"[^>]*y="([^"]*)"[^>]*>/);
      if (offMatch && offMatch[1] && offMatch[2]) {
        props.x = emuToPt(parseInt(offMatch[1], 10));
        props.y = emuToPt(parseInt(offMatch[2], 10));
      }

      const extMatch = xfrmContent.match(/<a:ext[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"[^>]*>/);
      if (extMatch && extMatch[1] && extMatch[2]) {
        props.width = emuToPt(parseInt(extMatch[1], 10));
        props.height = emuToPt(parseInt(extMatch[2], 10));
      }
    }

    const solidFillMatch = spPrContent.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
    if (solidFillMatch && solidFillMatch[1]) {
      const colorMatch = solidFillMatch[1].match(/<a:srgbClr[^>]*val="([^"]*)"[^>]*>/);
      if (colorMatch && colorMatch[1]) {
        props.fill = colorMatch[1];
      }
    }

    const lnMatch = spPrContent.match(/<a:ln(?:[^>]*)>([\s\S]*?)<\/a:ln>/);
    if (lnMatch && lnMatch[1]) {
      const lnContent = lnMatch[1];
      const lnColorMatch = lnContent.match(/<a:solidFill[^>]*>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:solidFill>/);
      if (lnColorMatch && lnColorMatch[1]) {
        props.line = lnColorMatch[1];
      }
      const lnWidthMatch = lnContent.match(/<a:ln[^>]*w="([^"]*)"[^>]*>/);
      if (lnWidthMatch && lnWidthMatch[1]) {
        props.lineWidth = parseInt(lnWidthMatch[1], 10) / 12700;
      }
    }
  }

  const nameMatch = shapeXml.match(/<p:cNvPr[^>]*name="([^"]*)"[^>]*>/);
  if (nameMatch && nameMatch[1]) {
    props.name = nameMatch[1];
  }

  return props;
}

/**
 * Renders a shape as SVG.
 */
function renderShapeAsSvg(shapeXml: string, slideIndex: number, shapeIndex: number, themeColors: Map<string, string>): string {
  const props = extractShapeProps(shapeXml);
  const { tspans } = extractTextRunsAsSvg(shapeXml, themeColors);

  const x = props.x ?? 0;
  const y = props.y ?? 0;
  const width = props.width ?? 100;
  const height = props.height ?? 50;

  let svg = `<g class="shape" data-slide="${slideIndex}" data-shape="${shapeIndex}">`;

  // Draw rectangle
  if (props.fill || props.line) {
    svg += `<rect x="${x}" y="${y}" width="${width}" height="${height}"`;
    if (props.fill) svg += ` fill="#${props.fill}"`;
    if (props.line) {
      svg += ` stroke="#${props.line}"`;
      svg += ` stroke-width="${props.lineWidth ?? 1}"`;
    } else {
      svg += ` stroke="none"`;
    }
    svg += `/>`;
  }

  // Draw text
  if (tspans) {
    svg += `<text x="${x}" y="${y}" width="${width}" height="${height}"`;
    svg += ` font-family="sans-serif" font-size="12"`;
    svg += `><tspan x="${x}" y="${y + 12}">${tspans}</tspan></text>`;
  }

  svg += `</g>`;

  return svg;
}

/**
 * Renders a table as SVG.
 */
function renderTableAsSvg(frameXml: string, slideIndex: number, tableIndex: number): string {
  const tblMatch = frameXml.match(/<a:tbl>([\s\S]*?)<\/a:tbl>/);
  if (!tblMatch || !tblMatch[1]) return "";

  const tblContent = tblMatch[1];

  // Extract rows
  const rows: string[][] = [];
  const rowPattern = /<a:tr([^>]*)>([\s\S]*?)<\/a:tr>/g;
  for (const rowMatch of tblContent.matchAll(rowPattern)) {
    if (!rowMatch[2]) continue;
    const cells: string[] = [];
    const cellPattern = /<a:tc([\s\S]*?)<\/a:tc>/g;
    for (const cellMatch of rowMatch[2].matchAll(cellPattern)) {
      if (!cellMatch[1]) continue;
      const cellContent = cellMatch[1];
      const textMatch = cellContent.match(/<a:t>([^<]*)<\/a:t>/);
      cells.push(textMatch && textMatch[1] ? xmlEncode(textMatch[1]) : "");
    }
    rows.push(cells);
  }

  if (rows.length === 0) return "";

  // Extract position and size
  let x = 0, y = 0, width = 100, height = 50;

  const xfrmMatch = frameXml.match(/<a:xfrm(?:[^>]*)>([\s\S]*?)<\/a:xfrm>/);
  if (xfrmMatch && xfrmMatch[1]) {
    const xfrmContent = xfrmMatch[1];
    const offMatch = xfrmContent.match(/<a:off[^>]*x="([^"]*)"[^>]*y="([^"]*)"[^>]*>/);
    if (offMatch && offMatch[1] && offMatch[2]) {
      x = emuToPt(parseInt(offMatch[1], 10));
      y = emuToPt(parseInt(offMatch[2], 10));
    }
    const extMatch = xfrmContent.match(/<a:ext[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"[^>]*>/);
    if (extMatch && extMatch[1] && extMatch[2]) {
      width = emuToPt(parseInt(extMatch[1], 10));
      height = emuToPt(parseInt(extMatch[2], 10));
    }
  }

  const colCount = rows[0]?.length || 1;
  const rowCount = rows.length;
  const cellWidth = width / colCount;
  const cellHeight = height / rowCount;

  let svg = `<g class="table" data-slide="${slideIndex}" data-table="${tableIndex}">`;

  // Draw cells
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const cx = x + c * cellWidth;
      const cy = y + r * cellHeight;

      svg += `<rect x="${cx}" y="${cy}" width="${cellWidth}" height="${cellHeight}" fill="none" stroke="#333" stroke-width="0.5"/>`;
      svg += `<text x="${cx + 2}" y="${cy + cellHeight / 2 + 4}" font-size="8" font-family="sans-serif">${rows[r][c]}</text>`;
    }
  }

  svg += `</g>`;

  return svg;
}

/**
 * Renders a picture as SVG (placeholder for embedded images).
 */
function renderPictureAsSvg(pictureXml: string, slideIndex: number, pictureIndex: number): string {
  let x = 0, y = 0, width = 100, height = 100;

  const xfrmMatch = pictureXml.match(/<p:spPr([\s\S]*?)<\/p:spPr>/);
  if (xfrmMatch && xfrmMatch[1]) {
    const xfrmContent = xfrmMatch[1].match(/<a:xfrm(?:[^>]*)>([\s\S]*?)<\/a:xfrm>/);
    if (xfrmContent && xfrmContent[1]) {
      const content = xfrmContent[1];
      const offMatch = content.match(/<a:off[^>]*x="([^"]*)"[^>]*y="([^"]*)"[^>]*>/);
      if (offMatch && offMatch[1] && offMatch[2]) {
        x = emuToPt(parseInt(offMatch[1], 10));
        y = emuToPt(parseInt(offMatch[2], 10));
      }
      const extMatch = content.match(/<a:ext[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"[^>]*>/);
      if (extMatch && extMatch[1] && extMatch[2]) {
        width = emuToPt(parseInt(extMatch[1], 10));
        height = emuToPt(parseInt(extMatch[2], 10));
      }
    }
  }

  const altMatch = pictureXml.match(/<p:cNvPr[^>]*descr="([^"]*)"[^>]*>/);
  const alt = (altMatch && altMatch[1]) ? xmlEncode(altMatch[1]) : "Image";

  let svg = `<g class="picture" data-slide="${slideIndex}" data-picture="${pictureIndex}">`;
  svg += `<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="#f0f0f0" stroke="#ccc" stroke-width="0.5"/>`;
  svg += `<text x="${x + width / 2}" y="${y + height / 2}" text-anchor="middle" font-size="10" fill="#666" font-family="sans-serif">[${alt}]</text>`;
  svg += `</g>`;

  return svg;
}

/**
 * Renders slide elements as SVG group.
 */
function renderSlideElementsAsSvg(slideXml: string, slideIndex: number, themeColors: Map<string, string>): string {
  let svg = "";

  // Render shapes
  const shapePattern = /<p:sp([\s\S]*?)<\/p:sp>/g;
  let shapeIndex = 0;
  for (const shapeMatch of slideXml.matchAll(shapePattern)) {
    if (!shapeMatch[1]) continue;
    shapeIndex++;
    svg += renderShapeAsSvg(shapeMatch[1], slideIndex, shapeIndex, themeColors);
  }

  // Render tables
  const framePattern = /<p:graphicFrame([\s\S]*?)<\/p:graphicFrame>/g;
  let tableIndex = 0;
  for (const frameMatch of slideXml.matchAll(framePattern)) {
    if (!frameMatch[1]) continue;
    const frameContent = frameMatch[1];
    if (frameContent.includes("<a:tbl>")) {
      tableIndex++;
      svg += renderTableAsSvg(frameContent, slideIndex, tableIndex);
    }
  }

  // Render pictures
  const picturePattern = /<p:pic([\s\S]*?)<\/p:pic>/g;
  let pictureIndex = 0;
  for (const picMatch of slideXml.matchAll(picturePattern)) {
    if (!picMatch[1]) continue;
    pictureIndex++;
    svg += renderPictureAsSvg(picMatch[1], slideIndex, pictureIndex);
  }

  return svg;
}

/**
 * Renders a single slide as SVG.
 */
function renderSlideSvg(slideXml: string, slideIndex: number, slideWidthPt: number, slideHeightPt: number, themeColors: Map<string, string>): string {
  const bgRect = getSlideBackgroundSvg(slideXml);

  let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${slideWidthPt} ${slideHeightPt}" width="${slideWidthPt}" height="${slideHeightPt}" class="slide" data-slide="${slideIndex}">`;

  // Background
  if (bgRect) {
    svg += bgRect;
  } else {
    svg += `<rect width="100%" height="100%" fill="#ffffff"/>`;
  }

  // Elements
  svg += renderSlideElementsAsSvg(slideXml, slideIndex, themeColors);

  svg += `</svg>`;

  return svg;
}

// ============================================================================
// Public API
// ============================================================================

/**
 * Renders a presentation as SVG.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - Optional 1-based slide index to render specific slide
 * @returns Result with SVG string
 *
 * @example
 * // Render entire presentation as SVG
 * const result = await viewAsSvg("/path/to/presentation.pptx");
 *
 * // Render specific slide as SVG
 * const result = await viewAsSvg("/path/to/presentation.pptx", 1);
 */
export async function viewAsSvg(
  filePath: string,
  slideIndex?: number,
): Promise<Result<ViewSvgResult>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    if (zipResult.error) {
      return err(zipResult.error.code, zipResult.error.message, zipResult.error.suggestion);
    }
    return err("operation_failed", "Failed to load presentation");
  }
  const zip = zipResult.data;
  if (!zip) {
    return err("operation_failed", "Failed to load presentation");
  }

  const slidesInfoResult = getAllSlideEntries(zip);
  if (!slidesInfoResult.ok) {
    if (slidesInfoResult.error) {
      return err(slidesInfoResult.error.code, slidesInfoResult.error.message, slidesInfoResult.error.suggestion);
    }
    return err("operation_failed", "Failed to get slide entries");
  }
  const slidesInfo = slidesInfoResult.data;
  if (!slidesInfo) {
    return err("operation_failed", "Failed to get slide entries");
  }

  // Filter to specific slide if requested
  const targetSlides = slideIndex
    ? slidesInfo.filter(s => s.index === slideIndex)
    : slidesInfo;

  if (slideIndex && targetSlides.length === 0) {
    return invalidInput(`Slide index ${slideIndex} is out of range (1-${slidesInfo.length})`);
  }

  // Get slide dimensions
  const { width: slideWidthPt, height: slideHeightPt } = getSlideDimensions(zip);

  // Get theme colors
  const themeColors = getThemeColors(zip);

  // Render slides
  const slideSvgParts: string[] = [];
  for (const slideInfo of targetSlides) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);
    slideSvgParts.push(renderSlideSvg(slideXml, slideInfo.index, slideWidthPt, slideHeightPt, themeColors));
  }

  // If single slide, return just that SVG
  if (targetSlides.length === 1 && slideSvgParts[0]) {
    return ok({
      slideCount: 1,
      svg: slideSvgParts[0],
    });
  }

  // Multiple slides: wrap in a container SVG
  const combinedHeight = slideSvgParts.length * slideHeightPt;
  const combinedSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${slideWidthPt} ${combinedHeight}" width="${slideWidthPt}" height="${combinedHeight}">
  <style>
    .slide { display: block; }
  </style>
  ${slideSvgParts.map((svg, i) => `<g transform="translate(0, ${i * slideHeightPt})">${svg.replace(/<\/?svg[^>]*>/g, "")}</g>`).join("\n")}
</svg>`;

  return ok({
    slideCount: targetSlides.length,
    svg: combinedSvg,
  });
}
