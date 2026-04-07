/**
 * HTML Preview rendering for @officekit/ppt.
 *
 * Provides functions to render PowerPoint presentations as self-contained HTML:
 * - viewAsHtml: Renders entire presentation or specific slide as HTML
 * - generatePreview: Generates preview with format and options
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
 * Options for generating a preview.
 */
export interface PreviewOptions {
  /** Output format: 'html' or 'svg' */
  format?: "html" | "svg";
  /** Specific slide indices to include (1-based) */
  slides?: number[];
  /** Width for scaling (optional) */
  width?: number;
}

/**
 * Result from HTML preview - contains the HTML string.
 */
export interface ViewHtmlResult {
  /** Total slide count in the preview */
  slideCount: number;
  /** The rendered HTML string */
  html: string;
}

/**
 * Result from generatePreview.
 */
export interface GeneratePreviewResult {
  /** Output format used */
  format: "html" | "svg";
  /** Total slide count */
  slideCount: number;
  /** The rendered output string */
  output: string;
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
// Helper Functions (from views.ts - duplicated for self-containment)
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

// ============================================================================
// Additional Helper Functions
// ============================================================================

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
    // Try to get theme colors from the theme file
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);

    // Find the theme relationship
    const themeRel = relationships.find(r =>
      r.type?.includes("theme") || r.target?.includes("theme")
    );

    if (themeRel) {
      const themePath = normalizeZipPath("ppt", themeRel.target);
      const themeXml = requireEntry(zip, themePath);

      // Extract color scheme
      const colorSchemeMatch = themeXml.match(/<a:clrScheme([^>]*)>([\s\S]*?)<\/a:clrScheme>/);
      if (colorSchemeMatch) {
        const schemeContent = colorSchemeMatch[2];

        // Extract individual colors
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

  // Set defaults if not found
  if (!colors.has("dk1")) colors.set("dk1", "000000");
  if (!colors.has("dk2")) colors.set("dk2", "333333");
  if (!colors.has("lt1")) colors.set("lt1", "FFFFFF");
  if (!colors.has("lt2")) colors.set("lt2", "EEEEEE");
  if (!colors.has("accent1")) colors.set("accent1", "4472C4");

  return colors;
}

/**
 * Gets slide background CSS.
 */
function getSlideBackgroundCss(slideXml: string, themeColors: Map<string, string>): string {
  // Check for solid fill background
  const bgMatch = slideXml.match(/<p:bg([\s\S]*?)<\/p:bg>/);
  if (!bgMatch || !bgMatch[1]) return "";

  const bgContent = bgMatch[1];

  // Solid fill
  const bgPrMatch = bgContent.match(/<p:bgPr([\s\S]*?)<\/p:bgPr>/);
  if (bgPrMatch && bgPrMatch[1]) {
    const bgPrContent = bgPrMatch[1];
    const solidFill = bgPrContent.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
    if (solidFill && solidFill[1]) {
      const colorMatch = solidFill[1].match(/<a:srgbClr[^>]*val="([^"]*)"[^>]*>/);
      if (colorMatch && colorMatch[1]) {
        return `background-color:#${colorMatch[1]};`;
      }
    }
  }

  return "";
}

/**
 * Escapes HTML special characters.
 */
function htmlEncode(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Extracts shape text content.
 */
function extractTextRuns(shapeXml: string): Array<{ text: string; bold?: boolean; italic?: boolean; fontSize?: number; fontColor?: string; fontFace?: string }> {
  const runs: Array<{ text: string; bold?: boolean; italic?: boolean; fontSize?: number; fontColor?: string; fontFace?: string }> = [];

  // Match a:r elements
  const runPattern = /<a:r([\s\S]*?)<\/a:r>/g;
  for (const runMatch of shapeXml.matchAll(runPattern)) {
    if (!runMatch[1]) continue;
    const runContent = runMatch[1];

    // Extract text
    const textMatch = runContent.match(/<a:t>([^<]*)<\/a:t>/);
    if (!textMatch || !textMatch[1]) continue;

    const text = textMatch[1];
    if (!text) continue;

    const runProps: typeof runs[0] = { text };

    // Extract run properties
    const rPrMatch = runContent.match(/<a:rPr([^>]*)>([\s\S]*?)<\/a:rPr>/);
    if (rPrMatch && rPrMatch[2]) {
      const rPrContent = rPrMatch[2];

      // Bold
      if (rPrContent.includes("<a:defRPr") || rPrContent.includes("<a:latin")) {
        const boldMatch = rPrContent.match(/<a:defRPr[^>]*b="1"/) || rPrContent.match(/<a:latin[^>]*b="1"/);
        if (boldMatch) runProps.bold = true;
      }
      if (rPrContent.includes("<a:b/>") || rPrContent.match(/<a:rPr[^>]*b="1"/)) {
        runProps.bold = true;
      }

      // Italic
      if (rPrContent.includes("<a:i/>") || rPrContent.match(/<a:rPr[^>]*i="1"/)) {
        runProps.italic = true;
      }

      // Font size (in hundredths of points)
      const fontSizeMatch = rPrContent.match(/<a:defRPr[^>]*sz="(\d+)"/) || rPrContent.match(/<a:latin[^>]*sz="(\d+)"/);
      if (fontSizeMatch && fontSizeMatch[1]) {
        runProps.fontSize = parseInt(fontSizeMatch[1], 10) / 100;
      }

      // Font color
      const solidFillMatch = rPrContent.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
      if (solidFillMatch && solidFillMatch[1]) {
        const colorMatch = solidFillMatch[1].match(/<a:srgbClr[^>]*val="([^"]*)"[^>]*>/);
        if (colorMatch && colorMatch[1]) {
          runProps.fontColor = colorMatch[1];
        }
      }

      // Font face
      const fontFaceMatch = rPrContent.match(/<a:latin[^>]*typeface="([^"]*)"[^>]*>/);
      if (fontFaceMatch && fontFaceMatch[1]) {
        runProps.fontFace = fontFaceMatch[1];
      }
    }

    runs.push(runProps);
  }

  return runs;
}

/**
 * Renders text runs as HTML.
 */
function renderTextRunsAsHtml(runs: Array<{ text: string; bold?: boolean; italic?: boolean; fontSize?: number; fontColor?: string; fontFace?: string }>): string {
  return runs.map(run => {
    let span = `<span style="`;
    if (run.bold) span += `font-weight:bold;`;
    if (run.italic) span += `font-style:italic;`;
    if (run.fontSize) span += `font-size:${run.fontSize}pt;`;
    if (run.fontColor) span += `color:#${run.fontColor};`;
    if (run.fontFace) span += `font-family:'${run.fontFace}',sans-serif;`;
    span += `">${htmlEncode(run.text)}</span>`;
    return span;
  }).join("");
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

  // Extract spPr
  const spPrMatch = shapeXml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);
  if (spPrMatch && spPrMatch[1]) {
    const spPrContent = spPrMatch[1];

    // Extract xfrm values
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

    // Extract fill color
    const solidFillMatch = spPrContent.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
    if (solidFillMatch && solidFillMatch[1]) {
      const colorMatch = solidFillMatch[1].match(/<a:srgbClr[^>]*val="([^"]*)"[^>]*>/);
      if (colorMatch && colorMatch[1]) {
        props.fill = colorMatch[1];
      }
    }

    // Extract line
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

  // Extract name
  const nameMatch = shapeXml.match(/<p:cNvPr[^>]*name="([^"]*)"[^>]*>/);
  if (nameMatch && nameMatch[1]) {
    props.name = nameMatch[1];
  }

  return props;
}

/**
 * Renders a single shape as HTML.
 */
function renderShapeAsHtml(shapeXml: string, slideIndex: number, shapeIndex: number, themeColors: Map<string, string>): string {
  const props = extractShapeProps(shapeXml);
  const textRuns = extractTextRuns(shapeXml);

  // Build inline styles
  const styles: string[] = [];
  styles.push(`position:absolute`);
  if (props.x !== undefined) styles.push(`left:${props.x}pt`);
  if (props.y !== undefined) styles.push(`top:${props.y}pt`);
  if (props.width !== undefined) styles.push(`width:${props.width}pt`);
  if (props.height !== undefined) styles.push(`height:${props.height}pt`);
  if (props.fill) styles.push(`background-color:#${props.fill}`);
  if (props.line) {
    styles.push(`border: ${props.lineWidth || 1}pt solid #${props.line}`);
  } else {
    styles.push(`border:none`);
  }

  const style = styles.join(";");

  // Build text content
  let textHtml = "";
  if (textRuns.length > 0) {
    // Get default text color from theme if not specified
    const defaultColor = themeColors.get("dk1") || "000000";
    textHtml = renderTextRunsAsHtml(textRuns);

    // Wrap in a text div if there's substantial text
    return `<div class="shape" data-slide="${slideIndex}" data-shape="${shapeIndex}" style="${style}">
  <div class="shape-content" style="position:relative;width:100%;height:100%;overflow:hidden;">
    <div class="shape-text" style="color:#${defaultColor};">${textHtml}</div>
  </div>
</div>`;
  }

  // Shape without text
  return `<div class="shape" data-slide="${slideIndex}" data-shape="${shapeIndex}" style="${style}"></div>`;
}

/**
 * Renders a table as HTML.
 */
function renderTableAsHtml(frameXml: string, slideIndex: number, tableIndex: number): string {
  // Extract table content
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
      // Extract cell text
      const textMatch = cellContent.match(/<a:t>([^<]*)<\/a:t>/);
      cells.push(textMatch && textMatch[1] ? htmlEncode(textMatch[1]) : "");
    }
    rows.push(cells);
  }

  if (rows.length === 0) return "";

  // Extract table position and size from graphicFrame
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

  // Build HTML table
  let html = `<div class="table" data-slide="${slideIndex}" data-table="${tableIndex}" style="position:absolute;left:${x}pt;top:${y}pt;width:${width}pt;height:${height}pt;overflow:hidden;">`;
  html += `<table style="width:100%;height:100%;border-collapse:collapse;">`;

  for (const row of rows) {
    html += `<tr>`;
    for (const cell of row) {
      html += `<td style="border:1pt solid #333;padding:2pt;">${cell}</td>`;
    }
    html += `</tr>`;
  }

  html += `</table></div>`;

  return html;
}

/**
 * Renders a picture as HTML (placeholder for embedded images).
 */
function renderPictureAsHtml(pictureXml: string, slideIndex: number, pictureIndex: number): string {
  // Extract picture position and size
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

  // Extract alt text
  const altMatch = pictureXml.match(/<p:cNvPr[^>]*descr="([^"]*)"[^>]*>/);
  const alt = (altMatch && altMatch[1]) ? htmlEncode(altMatch[1]) : "Image";

  return `<div class="picture" data-slide="${slideIndex}" data-picture="${pictureIndex}" style="position:absolute;left:${x}pt;top:${y}pt;width:${width}pt;height:${height}pt;overflow:hidden;background:#f0f0f0;display:flex;align-items:center;justify-content:center;">
  <span style="color:#666;">[${alt}]</span>
</div>`;
}

/**
 * Renders a slide's elements as HTML.
 */
function renderSlideElementsAsHtml(slideXml: string, slideIndex: number, themeColors: Map<string, string>): string {
  let html = "";

  // Render shapes
  const shapePattern = /<p:sp([\s\S]*?)<\/p:sp>/g;
  let shapeIndex = 0;
  for (const shapeMatch of slideXml.matchAll(shapePattern)) {
    if (!shapeMatch[1]) continue;
    shapeIndex++;
    html += renderShapeAsHtml(shapeMatch[1], slideIndex, shapeIndex, themeColors);
  }

  // Render tables (graphic frames containing tables)
  const framePattern = /<p:graphicFrame([\s\S]*?)<\/p:graphicFrame>/g;
  let tableIndex = 0;
  for (const frameMatch of slideXml.matchAll(framePattern)) {
    if (!frameMatch[1]) continue;
    const frameContent = frameMatch[1];
    if (frameContent.includes("<a:tbl>")) {
      tableIndex++;
      html += renderTableAsHtml(frameContent, slideIndex, tableIndex);
    }
  }

  // Render pictures
  const picturePattern = /<p:pic([\s\S]*?)<\/p:pic>/g;
  let pictureIndex = 0;
  for (const picMatch of slideXml.matchAll(picturePattern)) {
    if (!picMatch[1]) continue;
    pictureIndex++;
    html += renderPictureAsHtml(picMatch[1], slideIndex, pictureIndex);
  }

  return html;
}

/**
 * Generates the CSS for the HTML preview.
 */
function generateCss(slideWidthPt: number, slideHeightPt: number): string {
  const aspect = slideWidthPt / slideHeightPt;

  return `:root {
  --slide-design-w: ${slideWidthPt}pt;
  --slide-design-h: ${slideHeightPt}pt;
  --slide-aspect: ${aspect};
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: #1a1a2e;
  color: #e2e8f0;
  min-height: 100vh;
}
.slide-container {
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
  margin: 20px auto;
  overflow: hidden;
}
.slide {
  position: relative;
  width: ${slideWidthPt}pt;
  height: ${slideHeightPt}pt;
  transform-origin: top left;
}
.slide-label {
  background: #333;
  color: #fff;
  padding: 8px 16px;
  font-size: 12px;
  font-weight: 500;
}
.shape {
  box-sizing: border-box;
}
.shape-text {
  padding: 4pt;
  line-height: 1.2;
}
.table {
  overflow: hidden;
}
.table table {
  font-size: 10pt;
  text-align: left;
}
.picture {
  text-align: center;
}
`;
}

/**
 * Renders a single slide as HTML fragment.
 */
function renderSlideHtml(slideXml: string, slideIndex: number, slideWidthPt: number, slideHeightPt: number, themeColors: Map<string, string>): string {
  const bgStyle = getSlideBackgroundCss(slideXml, themeColors);

  let html = `<div class="slide-container" data-slide="${slideIndex}">`;
  html += `<div class="slide-label">Slide ${slideIndex}</div>`;
  html += `<div class="slide-wrapper">`;
  html += `<div class="slide"${bgStyle ? ` style="${bgStyle}"` : ""}>`;

  html += renderSlideElementsAsHtml(slideXml, slideIndex, themeColors);

  html += `</div>`; // slide
  html += `</div>`; // slide-wrapper
  html += `</div>`; // slide-container

  return html;
}

// ============================================================================
// Public API
// ============================================================================

/**
 * Renders a presentation as self-contained HTML.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - Optional 1-based slide index to render specific slide
 * @returns Result with HTML string
 *
 * @example
 * // Render entire presentation as HTML
 * const result = await viewAsHtml("/path/to/presentation.pptx");
 *
 * // Render specific slide as HTML
 * const result = await viewAsHtml("/path/to/presentation.pptx", 1);
 */
export async function viewAsHtml(
  filePath: string,
  slideIndex?: number,
): Promise<Result<ViewHtmlResult>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error!.code, zipResult.error!.message, zipResult.error!.suggestion);
  }
  const zip = zipResult.data!;

  const slidesInfoResult = getAllSlideEntries(zip);
  if (!slidesInfoResult.ok) {
    return err(slidesInfoResult.error!.code, slidesInfoResult.error!.message, slidesInfoResult.error!.suggestion);
  }
  const slidesInfo = slidesInfoResult.data!;

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
  const slideHtmlParts: string[] = [];
  for (const slideInfo of targetSlides) {
    const slideXml = requireEntry(zip, slideInfo.entryPath);
    slideHtmlParts.push(renderSlideHtml(slideXml, slideInfo.index, slideWidthPt, slideHeightPt, themeColors));
  }

  // Generate full HTML document
  const css = generateCss(slideWidthPt, slideHeightPt);
  const fileName = path.basename(filePath);

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${htmlEncode(fileName)}</title>
  <style>
${css}
  </style>
</head>
<body>
  <div class="main">
    <h1 class="file-title">${htmlEncode(fileName)}</h1>
    ${slideHtmlParts.join("\n")}
  </div>
</body>
</html>`;

  return ok({
    slideCount: targetSlides.length,
    html,
  });
}

/**
 * Generates a preview of the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param options - Preview options
 * @returns Result with generated preview
 *
 * @example
 * // Generate HTML preview
 * const result = await generatePreview("/path/to/presentation.pptx", { format: 'html' });
 *
 * // Generate SVG preview for specific slides
 * const result = await generatePreview("/path/to/presentation.pptx", { format: 'svg', slides: [1, 2] });
 */
export async function generatePreview(
  filePath: string,
  options?: PreviewOptions,
): Promise<Result<GeneratePreviewResult>> {
  const format = options?.format || "html";

  if (format === "svg") {
    // Dynamic import to avoid circular dependency
    const { viewAsSvg } = await import("./preview-svg.js");
    const svgResult = await viewAsSvg(filePath, options?.slides?.[0]);
    if (!svgResult.ok) {
      return err(svgResult.error!.code, svgResult.error!.message, svgResult.error!.suggestion);
    }
    return ok({
      format: "svg",
      slideCount: svgResult.data!.slideCount,
      output: svgResult.data!.svg,
    });
  }

  // Default to HTML
  const htmlResult = await viewAsHtml(filePath, options?.slides?.[0]);
  if (!htmlResult.ok) {
    return err(htmlResult.error!.code, htmlResult.error!.message, htmlResult.error!.suggestion);
  }
  return ok({
    format: "html",
    slideCount: htmlResult.data!.slideCount,
    output: htmlResult.data!.html,
  });
}
