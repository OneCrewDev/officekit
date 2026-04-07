/**
 * Text content operations for @officekit/ppt.
 *
 * Provides functions to manipulate text content within shapes:
 * - Get and set text runs
 * - Add paragraphs
 * - Set text formatting
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput, notFound } from "./result.js";
import type { Result, RunModel, ParagraphModel } from "./types.js";
import { getSlideIndex } from "./path.js";

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
 * Escapes special XML characters.
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
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
 * Extracts the shape index from a path.
 */
function extractShapeIndex(pptPath: string): number | null {
  const pattern = /\/shape\[(\d+)\]/i;
  const match = pptPath.match(pattern);
  return match ? parseInt(match[1], 10) : null;
}

// ============================================================================
// Text Run Operations
// ============================================================================

/**
 * Gets all text runs in a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 *
 * @example
 * const result = await getTextRuns("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 * if (result.ok) {
 *   console.log(result.data.runs);
 *   console.log(result.data.paragraphs);
 * }
 */
export async function getTextRuns(
  filePath: string,
  pptPath: string,
): Promise<Result<{ runs: RunModel[]; paragraphs: ParagraphModel[] }>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("getTextRuns requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message);
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
    }

    // Find the shape
    const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
    const matches = slideXml.match(pattern);

    if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
      return notFound("Shape", pptPath);
    }

    const shapeXml = matches[shapeIndex - 1];

    // Extract paragraphs and runs
    const paragraphs = extractParagraphs(shapeXml);
    const runs = extractRuns(shapeXml);

    return ok({ runs, paragraphs });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Extracts paragraphs from a shape's text body.
 */
function extractParagraphs(shapeXml: string): ParagraphModel[] {
  const paragraphs: ParagraphModel[] = [];

  // Find txBody
  const txBodyMatch = shapeXml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
  if (!txBodyMatch) {
    return paragraphs;
  }

  const txBody = txBodyMatch[1];

  // Match all paragraphs
  const paraPattern = /<a:p>[\s\S]*?<\/a:p>/g;
  const paraMatches = txBody.match(paraPattern) || [];

  for (let pIdx = 0; pIdx < paraMatches.length; pIdx++) {
    const paraXml = paraMatches[pIdx];
    const runs = extractRunsFromParagraph(paraXml, pIdx + 1);

    // Extract paragraph properties
    const alignment = extractParagraphAlignment(paraXml);
    const marginLeft = extractMargin(paraXml, "left");
    const marginRight = extractMargin(paraXml, "right");

    // Get concatenated text
    const text = runs.map(r => r.text).join("");

    paragraphs.push({
      index: pIdx + 1,
      text,
      alignment,
      marginLeft,
      marginRight,
      runs,
      childCount: runs.length,
    });
  }

  return paragraphs;
}

/**
 * Extracts runs from a shape's text body.
 */
function extractRuns(shapeXml: string): RunModel[] {
  const runs: RunModel[] = [];

  // Find txBody
  const txBodyMatch = shapeXml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
  if (!txBodyMatch) {
    return runs;
  }

  const txBody = txBodyMatch[1];

  // Match all paragraphs first
  const paraPattern = /<a:p>[\s\S]*?<\/a:p>/g;
  const paraMatches = txBody.match(paraPattern) || [];

  let runIndex = 0;
  for (const paraXml of paraMatches) {
    const paraRuns = extractRunsFromParagraph(paraXml, runIndex + 1);
    runs.push(...paraRuns);
    runIndex += paraRuns.length;
  }

  return runs;
}

/**
 * Extracts runs from a paragraph.
 */
function extractRunsFromParagraph(paraXml: string, paragraphIndex: number): RunModel[] {
  const runs: RunModel[] = [];

  // Match all runs in the paragraph
  const runPattern = /<a:r>[\s\S]*?<\/a:r>/g;
  const runMatches = paraXml.match(runPattern) || [];

  for (let rIdx = 0; rIdx < runMatches.length; rIdx++) {
    const runXml = runMatches[rIdx];

    // Extract text content
    const textMatch = runXml.match(/<a:t>([^<]*)<\/a:t>/);
    const text = textMatch ? textMatch[1] : "";

    // Extract run properties
    const rPrMatch = runXml.match(/<a:rPr([^>]*)\/?>([\s\S]*?)<\/a:rPr>|<a:rPr([^>]*)\/>/);
    const runProps = rPrMatch ? (rPrMatch[0] || "") : "";

    // Extract font
    const fontMatch = runProps.match(/<a:latin[^>]*typeface="([^"]*)"/);
    const font = fontMatch ? fontMatch[1] : undefined;

    // Extract size
    const sizeMatch = runProps.match(/sz="(\d+)"/);
    const size = sizeMatch ? sizeMatch[1] : undefined;

    // Extract bold
    const boldMatch = runProps.match(/b="([^"]*)"/);
    const bold = boldMatch ? boldMatch[1] === "1" || boldMatch[1] === "true" : undefined;

    // Extract italic
    const italicMatch = runProps.match(/i="([^"]*)"/);
    const italic = italicMatch ? italicMatch[1] === "1" || italicMatch[1] === "true" : undefined;

    // Extract underline
    const underlineMatch = runProps.match(/u="([^"]*)"/);
    const underline = underlineMatch ? underlineMatch[1] : undefined;

    // Extract strikethrough
    const strikeMatch = runProps.match(/strike="([^"]*)"/);
    const strike = strikeMatch ? strikeMatch[1] : undefined;

    // Extract color
    const colorMatch = runProps.match(/<a:solidFill>[\s\S]*?<a:srgbClr[^>]*val="([^"]*)"[^>]*>[\s\S]*?<\/a:solidFill>/);
    const color = colorMatch ? colorMatch[1] : undefined;

    runs.push({
      index: rIdx + 1,
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
 * Extracts paragraph alignment.
 */
function extractParagraphAlignment(paraXml: string): "left" | "center" | "right" | "justify" | undefined {
  const alignMatch = paraXml.match(/<a:pPr[^>]*algn="([^"]*)"[^>]*>/);
  if (alignMatch) {
    const alignment = alignMatch[1];
    switch (alignment) {
      case "ctr": return "center";
      case "r": return "right";
      case "just": return "justify";
      case "l": return "left";
      default: return "left";
    }
  }
  return undefined;
}

/**
 * Extracts paragraph margin.
 */
function extractMargin(paraXml: string, side: "left" | "right"): number | undefined {
  const marginAttr = side === "left" ? "marL" : "marR";
  const marginMatch = paraXml.match(new RegExp(`<a:pPr[^>]*${marginAttr}="([^"]*)"[^>]*>`));
  if (marginMatch) {
    return parseInt(marginMatch[1], 10);
  }
  return undefined;
}

// ============================================================================
// Set Text Runs
// ============================================================================

/**
 * Text run specification for setting text.
 */
export interface TextRunSpec {
  /** Text content */
  text: string;
  /** Font typeface */
  font?: string;
  /** Font size in points (e.g., "14" or "14pt") */
  size?: string;
  /** Bold */
  bold?: boolean;
  /** Italic */
  italic?: boolean;
  /** Underline style */
  underline?: string;
  /** Strikethrough style */
  strike?: string;
  /** Text color as hex */
  color?: string;
}

/**
 * Sets the text runs in a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param runs - Array of text run specifications
 *
 * @example
 * const result = await setTextRuns("/path/to/presentation.pptx", "/slide[1]/shape[1]", [
 *   { text: "Hello ", font: "Arial", size: "18pt", bold: true },
 *   { text: "World", font: "Arial", size: "18pt", italic: true }
 * ]);
 */
export async function setTextRuns(
  filePath: string,
  pptPath: string,
  runs: TextRunSpec[],
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setTextRuns requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message);
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
    }

    const updatedSlideXml = setTextRunsInShape(slideXml, shapeIndex, runs);

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));
    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Sets text runs in a shape by index.
 */
function setTextRunsInShape(slideXml: string, shapeIndex: number, runs: TextRunSpec[]): string {
  const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);

  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    throw new Error(`Shape index ${shapeIndex} out of range`);
  }

  const targetShapeXml = matches[shapeIndex - 1];
  const updatedShapeXml = updateShapeTextRuns(targetShapeXml, runs);

  return slideXml.replace(targetShapeXml, updatedShapeXml);
}

/**
 * Updates the text runs in a shape.
 */
function updateShapeTextRuns(shapeXml: string, runs: TextRunSpec[]): string {
  // Build new paragraph XML from runs
  const runsXml = runs.map(run => buildRunXml(run)).join("");

  const newParagraph = `        <a:p>
${runsXml}
        </a:p>`;

  // Find and replace the txBody content
  const txBodyMatch = shapeXml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);

  if (txBodyMatch) {
    const txBodyContent = txBodyMatch[1];
    // Keep bodyPr and lstStyle, replace paragraphs
    const bodyPrMatch = txBodyContent.match(/<a:bodyPr[^>]*\/?>[\s\S]*?<a:lstStyle\/>/);
    if (bodyPrMatch) {
      const newTxBody = `<p:txBody>${bodyPrMatch[0]}${newParagraph}
          </p:txBody>`;
      return shapeXml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTxBody);
    }
  }

  // Fallback: create minimal txBody structure
  const newTxBody = `        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
${newParagraph}
        </p:txBody>`;

  // Try to insert before </p:sp> or at end
  if (shapeXml.includes("</p:txBody>")) {
    return shapeXml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTxBody);
  }

  return shapeXml;
}

/**
 * Builds XML for a single text run.
 */
function buildRunXml(run: TextRunSpec): string {
  const text = escapeXml(run.text);
  const rPrAttrs: string[] = [];

  if (run.font) {
    rPrAttrs.push(`<a:latin typeface="${run.font}"/>`);
  }
  if (run.size) {
    // Convert points to half-points (sz value is in 100ths of points)
    const sizeValue = run.size.replace(/[^\d.]/g, "");
    const halfPoints = Math.round(parseFloat(sizeValue) * 100);
    rPrAttrs.push(`sz="${halfPoints}"`);
  }
  if (run.bold) {
    rPrAttrs.push('b="1"');
  }
  if (run.italic) {
    rPrAttrs.push('i="1"');
  }
  if (run.underline) {
    rPrAttrs.push(`u="${run.underline}"`);
  }
  if (run.strike) {
    rPrAttrs.push(`strike="${run.strike}"`);
  }
  if (run.color) {
    const color = run.color.replace("#", "");
    rPrAttrs.push(`<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`);
  }

  const rPrContent = rPrAttrs.length > 0 ? `\n          <a:rPr lang="en-US" ${rPrAttrs.join(" ")}/>` : '\n          <a:rPr lang="en-US"/>';

  return `          <a:r>${rPrContent}
            <a:t>${text}</a:t>
          </a:r>`;
}

// ============================================================================
// Add Text Paragraph
// ============================================================================

/**
 * Paragraph format specification.
 */
export interface ParagraphFormat {
  /** Text alignment */
  alignment?: "left" | "center" | "right" | "justify";
  /** Left margin in EMUs */
  marginLeft?: number;
  /** Right margin in EMUs */
  marginRight?: number;
  /** Line spacing (e.g., "1.5", "double", "240" for fixed) */
  lineSpacing?: string;
  /** Space before in points */
  spaceBefore?: string;
  /** Space after in points */
  spaceAfter?: string;
}

/**
 * Adds a new paragraph to a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param text - The paragraph text content
 * @param format - Optional paragraph formatting
 *
 * @example
 * const result = await addTextParagraph("/path/to/presentation.pptx", "/slide[1]/shape[1]", "New paragraph", { alignment: "center" });
 */
export async function addTextParagraph(
  filePath: string,
  pptPath: string,
  text: string,
  format?: ParagraphFormat,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("addTextParagraph requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message);
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
    }

    const updatedSlideXml = addParagraphInShape(slideXml, shapeIndex, text, format);

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));
    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Adds a paragraph to a shape.
 */
function addParagraphInShape(
  slideXml: string,
  shapeIndex: number,
  text: string,
  format?: ParagraphFormat,
): string {
  const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);

  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    throw new Error(`Shape index ${shapeIndex} out of range`);
  }

  const targetShapeXml = matches[shapeIndex - 1];
  const updatedShapeXml = insertParagraph(targetShapeXml, text, format);

  return slideXml.replace(targetShapeXml, updatedShapeXml);
}

/**
 * Inserts a paragraph into a shape's text body.
 */
function insertParagraph(shapeXml: string, text: string, format?: ParagraphFormat): string {
  // Build paragraph properties
  let pPrContent = "";
  if (format) {
    if (format.alignment) {
      const alignMap: Record<string, string> = {
        left: "l",
        center: "ctr",
        right: "r",
        justify: "just",
      };
      pPrContent += ` algn="${alignMap[format.alignment]}"`;
    }
    if (format.marginLeft !== undefined) {
      pPrContent += ` marL="${format.marginLeft}"`;
    }
    if (format.marginRight !== undefined) {
      pPrContent += ` marR="${format.marginRight}"`;
    }
    if (format.lineSpacing) {
      pPrContent += ` lnSpc="${format.lineSpacing}"`;
    }
    if (format.spaceBefore) {
      pPrContent += ` spcBef="${format.spaceBefore}"`;
    }
    if (format.spaceAfter) {
      pPrContent += ` spcAft="${format.spaceAfter}"`;
    }
  }

  const pPrXml = pPrContent ? `\n          <a:pPr${pPrContent}/>` : "";

  const newParagraph = `          <a:p>${pPrXml}
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>${escapeXml(text)}</a:t>
            </a:r>
          </a:p>`;

  // Find txBody and append paragraph before </a:lstStyle/> or at end
  const txBodyMatch = shapeXml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);

  if (txBodyMatch) {
    const txBodyContent = txBodyMatch[1];
    // Insert before </p:txBody> but after lstStyle
    const lstStyleMatch = txBodyContent.match(/<a:lstStyle\/>/);
    if (lstStyleMatch) {
      const insertIdx = txBodyContent.indexOf("</a:lstStyle/>") + "</a:lstStyle/>".length;
      const newTxBody = txBodyContent.slice(0, insertIdx) + "\n" + newParagraph + txBodyContent.slice(insertIdx);
      return shapeXml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, `<p:txBody>${newTxBody}</p:txBody>`);
    }
  }

  // Fallback: just add paragraph at end of txBody
  if (shapeXml.includes("</p:txBody>")) {
    return shapeXml.replace(
      /<\/p:txBody>/,
      `${newParagraph}
        </p:txBody>`
    );
  }

  return shapeXml;
}

// ============================================================================
// Set Text Format
// ============================================================================

/**
 * Text format specification.
 */
export interface TextFormat {
  /** Font typeface */
  font?: string;
  /** Font size in points */
  size?: number;
  /** Bold */
  bold?: boolean;
  /** Italic */
  italic?: boolean;
  /** Underline style */
  underline?: string;
  /** Strikethrough style */
  strike?: string;
  /** Text color as hex */
  color?: string;
  /** Vertical alignment within shape */
  valign?: "top" | "middle" | "bottom";
}

/**
 * Sets the text format for all runs in a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param format - The text format specification
 *
 * @example
 * const result = await setTextFormat("/path/to/presentation.pptx", "/slide[1]/shape[1]", {
 *   font: "Arial",
 *   size: 18,
 *   bold: true,
 *   color: "FF0000"
 * });
 */
export async function setTextFormat(
  filePath: string,
  pptPath: string,
  format: TextFormat,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setTextFormat requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message);
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
    }

    const updatedSlideXml = setTextFormatInShape(slideXml, shapeIndex, format);

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));
    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Sets text format in a shape by index.
 */
function setTextFormatInShape(slideXml: string, shapeIndex: number, format: TextFormat): string {
  const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);

  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    throw new Error(`Shape index ${shapeIndex} out of range`);
  }

  const targetShapeXml = matches[shapeIndex - 1];
  const updatedShapeXml = updateShapeTextFormat(targetShapeXml, format);

  return slideXml.replace(targetShapeXml, updatedShapeXml);
}

/**
 * Updates the text format in a shape.
 */
function updateShapeTextFormat(shapeXml: string, format: TextFormat): string {
  let result = shapeXml;

  // Handle vertical alignment in bodyPr
  if (format.valign) {
    const valignMap: Record<string, string> = {
      top: "t",
      middle: "ctr",
      bottom: "b",
    };
    const anchorVal = valignMap[format.valign];

    // Find or create bodyPr with anchor attribute
    if (/<a:bodyPr[^>]*>/.test(result)) {
      result = result.replace(
        /<a:bodyPr([^>]*)\/>/,
        `<a:bodyPr$1 anchor="${anchorVal}"/>`
      );
    }
  }

  // Get all rPr elements and update them
  const rPrPattern = /<a:rPr([^>]*)\/?>([\s\S]*?)<\/a:rPr>|<a:rPr([^>]*)\/>/g;

  result = result.replace(rPrPattern, (match, attrs1, content, attrs2) => {
    let rPrAttrs = attrs1 || attrs2 || "";
    let rPrContent = content || "";

    // Build new attributes
    const newAttrs: string[] = [];

    if (format.font) {
      newAttrs.push(`typeface="${format.font}"`);
    }
    if (format.size) {
      newAttrs.push(`sz="${format.size * 100}"`); // Convert points to half-points
    }
    if (format.bold) {
      newAttrs.push('b="1"');
    }
    if (format.italic) {
      newAttrs.push('i="1"');
    }
    if (format.underline) {
      newAttrs.push(`u="${format.underline}"`);
    }
    if (format.strike) {
      newAttrs.push(`strike="${format.strike}"`);
    }

    // Handle color - need to add solidFill element
    let colorFill = "";
    if (format.color) {
      const color = format.color.replace("#", "");
      colorFill = `\n          <a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
    }

    // Rebuild rPr
    if (match.endsWith("/>")) {
      // Self-closing rPr
      const combinedAttrs = [...newAttrs, ...rPrAttrs.split('"').filter((s: string) => s.trim())].join(" ");
      const attrStr = combinedAttrs ? ` ${combinedAttrs}` : "";
      const closeTag = colorFill ? `\n        </a:rPr>` : "";
      return `<a:rPr lang="en-US"${attrStr}${colorFill}${closeTag}`;
    } else {
      // rPr with content
      // Update or add attributes
      for (const attr of newAttrs) {
        const [name] = attr.split("=");
        // Remove existing attribute of same name
        rPrAttrs = rPrAttrs.replace(new RegExp(`${name}="[^"]*"\\s*`), "");
        rPrContent = rPrContent.replace(new RegExp(`<a:${name.slice(0, 1).toUpperCase() + name.slice(1)}[^>]*>\\s*<\\/a:${name.slice(0, 1).toUpperCase() + name.slice(1)}>`), "");
      }
      const combinedAttrs = rPrAttrs + (newAttrs.length > 0 ? " " + newAttrs.join(" ") : "");
      const attrStr = combinedAttrs ? ` ${combinedAttrs}` : "";
      return `<a:rPr lang="en-US"${attrStr}>${rPrContent}${colorFill}
        </a:rPr>`;
    }
  });

  return result;
}
