/**
 * Hyperlink operations for @officekit/ppt.
 *
 * Provides functions to manage hyperlinks in PowerPoint presentations:
 * - Get hyperlink from a shape/text
 * - Set hyperlink on a shape/text
 * - Remove hyperlink
 * - Set external hyperlink
 * - Set internal hyperlink to another slide
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, andThen, map, notFound, invalidInput } from "./result.js";
import type { Result } from "./types.js";
import { getSlideIndex } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Represents a hyperlink on a shape or text.
 */
export interface HyperlinkInfo {
  /** The URL of the hyperlink (null for internal hyperlinks) */
  url?: string;
  /** Whether this is an internal link to another slide */
  isInternal?: boolean;
  /** Target slide index for internal links (1-based) */
  targetSlideIndex?: number;
  /** Display text for the hyperlink */
  display?: string;
}

/**
 * Position for placing the link (used in setExternalHyperlink for positioning)
 */
export interface HyperlinkPosition {
  /** X position in EMUs */
  x?: number;
  /** Y position in EMUs */
  y?: number;
  /** Width in EMUs */
  width?: number;
  /** Height in EMUs */
  height?: number;
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
 * Throws an error if the entry is not found.
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
 * Gets the relationships entry name for a given entry.
 */
function getRelationshipsEntryName(entryName: string): string {
  const directory = path.posix.dirname(entryName);
  const basename = path.posix.basename(entryName);
  return path.posix.join(directory, "_rels", `${basename}.rels`);
}

/**
 * Generates a unique relationship ID.
 */
function generateRelId(existingRelIds: string[]): string {
  let id = 1;
  let relId = `rId${id}`;
  while (existingRelIds.includes(relId)) {
    id++;
    relId = `rId${id}`;
  }
  return relId;
}

/**
 * Extracts shape index from path.
 */
function extractShapeIndex(pptPath: string): number | null {
  const pattern = /\/shape\[(\d+)\]/i;
  const match = pptPath.match(pattern);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * Extracts paragraph index from path.
 */
function extractParagraphIndex(pptPath: string): number | null {
  const pattern = /\/paragraph\[(\d+)\]/i;
  const match = pptPath.match(pattern);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * Extracts run index from path.
 */
function extractRunIndex(pptPath: string): number | null {
  const pattern = /\/run\[(\d+)\]/i;
  const match = pptPath.match(pattern);
  return match ? parseInt(match[1], 10) : null;
}

// ============================================================================
// Hyperlink Operations
// ============================================================================

/**
 * Gets hyperlink information from a shape or text.
 *
 * @param filePath - Path to the PPTX file
 * @param path - Path to the shape or text (e.g., "/slide[1]/shape[1]" or "/slide[1]/shape[1]/paragraph[1]/run[1]")
 *
 * @example
 * const result = await getHyperlink("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 * if (result.ok) {
 *   console.log(result.data.url);
 * }
 */
export async function getHyperlink(
  filePath: string,
  path: string
): Promise<Result<HyperlinkInfo | null>> {
  try {
    const slideIndex = getSlideIndex(path);
    if (slideIndex === null) {
      return invalidInput("Invalid path - must include slide index");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    const relsXml = requireEntry(zip, relsEntry);
    const relationships = parseRelationshipEntries(relsXml);

    // Check for run-level hyperlink
    const paraIndex = extractParagraphIndex(path);
    const runIndex = extractRunIndex(path);

    if (paraIndex !== null && runIndex !== null) {
      // Get hyperlink from specific run
      return getHyperlinkFromRun(slideXml, relationships, paraIndex, runIndex);
    } else if (paraIndex !== null) {
      // Get hyperlink from paragraph (first run's hyperlink)
      return getHyperlinkFromParagraph(slideXml, relationships, paraIndex);
    } else {
      // Get hyperlink from shape
      return getHyperlinkFromShape(slideXml, relationships, path);
    }
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets hyperlink from a specific run.
 */
function getHyperlinkFromRun(
  slideXml: string,
  relationships: Array<{ id: string; target: string; type?: string }>,
  paraIndex: number,
  runIndex: number
): Result<HyperlinkInfo | null> {
  // Find paragraph
  const paraPattern = /<a:p(?:[^>]*)>([\s\S]*?)<\/a:p>/g;
  const paraMatches = slideXml.match(paraPattern) || [];

  if (paraIndex < 1 || paraIndex > paraMatches.length) {
    return invalidInput(`Paragraph index ${paraIndex} out of range`);
  }

  const paraContent = paraMatches[paraIndex - 1];

  // Find run
  const runPattern = /<a:r(?:[^>]*)>([\s\S]*?)<\/a:r>/g;
  const runMatches = paraContent.match(runPattern) || [];

  if (runIndex < 1 || runIndex > runMatches.length) {
    return invalidInput(`Run index ${runIndex} out of range`);
  }

  const runContent = runMatches[runIndex - 1];

  // Look for hlinkClick in rPr
  const rprMatch = /<a:rPr(?:[^>]*)>([\s\S]*?)<\/a:rPr>/.exec(runContent);
  if (!rprMatch) {
    return ok(null); // No run properties means no hyperlink
  }

  const rprContent = rprMatch[1];
  const hlinkMatch = /<a:hlinkClick(?:\s[^>]*)?\sr:id="([^"]*)"[^>]*>/.exec(rprContent);
  if (!hlinkMatch) {
    return ok(null);
  }

  const relId = hlinkMatch[1];
  const rel = relationships.find(r => r.id === relId);

  if (!rel) {
    return ok(null);
  }

  // Check if it's an internal link (relationship type contains "slide" or "internal")
  const isInternal = rel.type?.includes("slide") ||
    rel.target?.startsWith("slide") ||
    rel.target?.includes("action");

  if (isInternal) {
    // Try to parse target slide index
    const slideMatch = rel.target?.match(/slide(\d+)/i) ||
      rel.target?.match(/\[(\d+)\]/);
    const targetSlideIndex = slideMatch ? parseInt(slideMatch[1], 10) : undefined;

    return ok({
      isInternal: true,
      targetSlideIndex,
      display: undefined,
    });
  }

  return ok({
    url: rel.target,
    isInternal: false,
  });
}

/**
 * Gets hyperlink from a paragraph (first run with hyperlink).
 */
function getHyperlinkFromParagraph(
  slideXml: string,
  relationships: Array<{ id: string; target: string; type?: string }>,
  paraIndex: number
): Result<HyperlinkInfo | null> {
  // Find paragraph
  const paraPattern = /<a:p(?:[^>]*)>([\s\S]*?)<\/a:p>/g;
  const paraMatches = slideXml.match(paraPattern) || [];

  if (paraIndex < 1 || paraIndex > paraMatches.length) {
    return invalidInput(`Paragraph index ${paraIndex} out of range`);
  }

  const paraContent = paraMatches[paraIndex - 1];

  // Look for hlinkClick in any run's rPr
  const runPattern = /<a:r(?:[^>]*)>([\s\S]*?)<\/a:r>/g;
  const runMatches = paraContent.match(runPattern) || [];

  for (const runContent of runMatches) {
    const rprMatch = /<a:rPr(?:[^>]*)>([\s\S]*?)<\/a:rPr>/.exec(runContent);
    if (!rprMatch) continue;

    const rprContent = rprMatch[1];
    const hlinkMatch = /<a:hlinkClick(?:\s[^>]*)?\sr:id="([^"]*)"[^>]*>/.exec(rprContent);
    if (!hlinkMatch) continue;

    const relId = hlinkMatch[1];
    const rel = relationships.find(r => r.id === relId);

    if (!rel) continue;

    const isInternal = rel.type?.includes("slide") ||
      rel.target?.startsWith("slide") ||
      rel.target?.includes("action");

    if (isInternal) {
      const slideMatch = rel.target?.match(/slide(\d+)/i) ||
        rel.target?.match(/\[(\d+)\]/);
      const targetSlideIndex = slideMatch ? parseInt(slideMatch[1], 10) : undefined;

      return ok({
        isInternal: true,
        targetSlideIndex,
      });
    }

    return ok({
      url: rel.target,
      isInternal: false,
    });
  }

  return ok(null);
}

/**
 * Gets hyperlink from a shape.
 */
function getHyperlinkFromShape(
  slideXml: string,
  relationships: Array<{ id: string; target: string; type?: string }>,
  shapePath: string
): Result<HyperlinkInfo | null> {
  const shapeIndex = extractShapeIndex(shapePath);
  if (shapeIndex === null) {
    return invalidInput("Invalid shape path - must include shape[index]");
  }

  // Find all shapes (p:sp elements)
  const shapePattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  const shapeMatches = slideXml.match(shapePattern) || [];

  if (shapeIndex < 1 || shapeIndex > shapeMatches.length) {
    return notFound("Shape", String(shapeIndex));
  }

  const shapeXml = shapeMatches[shapeIndex - 1];

  // Look for hlinkClick in nvPr or cNvPr
  const hlinkMatch = /<a:hlinkClick(?:\s[^>]*)?\sr:id="([^"]*)"[^>]*>/.exec(shapeXml);
  if (!hlinkMatch) {
    return ok(null);
  }

  const relId = hlinkMatch[1];
  const rel = relationships.find(r => r.id === relId);

  if (!rel) {
    return ok(null);
  }

  const isInternal = rel.type?.includes("slide") ||
    rel.target?.startsWith("slide") ||
    rel.target?.includes("action");

  if (isInternal) {
    const slideMatch = rel.target?.match(/slide(\d+)/i) ||
      rel.target?.match(/\[(\d+)\]/);
    const targetSlideIndex = slideMatch ? parseInt(slideMatch[1], 10) : undefined;

    return ok({
      isInternal: true,
      targetSlideIndex,
    });
  }

  return ok({
    url: rel.target,
    isInternal: false,
  });
}

/**
 * Sets a hyperlink on a shape or text.
 *
 * @param filePath - Path to the PPTX file
 * @param path - Path to the shape or text (e.g., "/slide[1]/shape[1]" or "/slide[1]/shape[1]/paragraph[1]/run[1]")
 * @param url - The URL to link to
 * @param display - Optional display text for the link
 *
 * @example
 * const result = await setHyperlink(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/shape[1]",
 *   "https://example.com",
 *   "Click here"
 * );
 */
export async function setHyperlink(
  filePath: string,
  path: string,
  url: string,
  display?: string
): Promise<Result<void>> {
  try {
    // Validate URL
    if (!url || url === "none") {
      return invalidInput("URL is required for setting a hyperlink");
    }

    // Basic URL validation
    try {
      const parsedUrl = new URL(url);
      if (!parsedUrl.protocol || !parsedUrl.host) {
        return invalidInput("Invalid URL format");
      }
    } catch {
      return invalidInput("Invalid URL format");
    }

    const slideIndex = getSlideIndex(path);
    if (slideIndex === null) {
      return invalidInput("Invalid path - must include slide index");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    let relsXml = "";
    try {
      relsXml = requireEntry(zip, relsEntry);
    } catch {
      // Create empty rels if it doesn't exist
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    }
    const relationships = parseRelationshipEntries(relsXml);

    // Check for run-level or paragraph-level path
    const paraIndex = extractParagraphIndex(path);
    const runIndex = extractRunIndex(path);

    let updatedSlideXml: string;
    let updatedRelsXml: string;

    if (paraIndex !== null && runIndex !== null) {
      // Set hyperlink on specific run
      const result = setHyperlinkOnRun(slideXml, relationships, paraIndex, runIndex, url);
      if (!result.ok) {
        return err(result.error?.code ?? "operation_failed", result.error?.message ?? "Failed to set hyperlink on run");
      }
      if (!result.data) {
        return err("operation_failed", "No data returned");
      }
      updatedSlideXml = result.data.slideXml;
      updatedRelsXml = result.data.relsXml;
    } else if (paraIndex !== null) {
      // Set hyperlink on paragraph
      const result = setHyperlinkOnParagraph(slideXml, relationships, paraIndex, url);
      if (!result.ok) {
        return err(result.error?.code ?? "operation_failed", result.error?.message ?? "Failed to set hyperlink on paragraph");
      }
      if (!result.data) {
        return err("operation_failed", "No data returned");
      }
      updatedSlideXml = result.data.slideXml;
      updatedRelsXml = result.data.relsXml;
    } else {
      // Set hyperlink on shape
      const shapeIndex = extractShapeIndex(path);
      if (shapeIndex === null) {
        return invalidInput("Invalid shape path - must include shape[index]");
      }
      const result = setHyperlinkOnShape(slideXml, relationships, shapeIndex, url);
      if (!result.ok) {
        return err(result.error?.code ?? "operation_failed", result.error?.message ?? "Failed to set hyperlink on shape");
      }
      if (!result.data) {
        return err("operation_failed", "No data returned");
      }
      updatedSlideXml = result.data.slideXml;
      updatedRelsXml = result.data.relsXml;
    }

    // Build new zip with updated files
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else if (name === relsEntry) {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
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
 * Sets a hyperlink on a specific run.
 */
function setHyperlinkOnRun(
  slideXml: string,
  relationships: Array<{ id: string; target: string; type?: string }>,
  paraIndex: number,
  runIndex: number,
  url: string
): Result<{ slideXml: string; relsXml: string }> {
  // Find paragraph
  const paraPattern = /<a:p(?:[^>]*)>([\s\S]*?)<\/a:p>/g;
  const paraMatches = slideXml.match(paraPattern) || [];

  if (paraIndex < 1 || paraIndex > paraMatches.length) {
    return invalidInput(`Paragraph index ${paraIndex} out of range`);
  }

  const originalPara = paraMatches[paraIndex - 1];
  let paraContent = originalPara;

  // Find run
  const runPattern = /<a:r(?:[^>]*)>([\s\S]*?)<\/a:r>/g;
  const runMatches = paraContent.match(runPattern) || [];

  if (runIndex < 1 || runIndex > runMatches.length) {
    return invalidInput(`Run index ${runIndex} out of range`);
  }

  const originalRun = runMatches[runIndex - 1];

  // Generate new relationship ID
  const existingRelIds = relationships.map(r => r.id);
  const newRelId = generateRelId(existingRelIds);

  // Create relationship entry
  const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url}" TargetMode="External"/>`;

  // Get existing relsXml
  const relsPattern = /<Relationships[^>]*>([\s\S]*?)<\/Relationships>/;
  const relsMatch = slideXml.match(/<Relationships[^>]*>([\s\S]*?)<\/Relationships>/);
  let relsXml = relsMatch ? relsMatch[0] : "";

  // Check if run already has hlinkClick
  const rprMatch = /<a:rPr(?:[^>]*)>([\s\S]*?)<\/a:rPr>/.exec(originalRun);
  if (rprMatch) {
    // Modify existing rPr
    const rprContent = rprMatch[1];
    if (/<a:hlinkClick/.test(rprContent)) {
      // Replace existing hlinkClick
      const updatedRprContent = rprContent.replace(
        /<a:hlinkClick(?:\s[^>]*)?\sr:id="[^"]*"[^>]*>/,
        `<a:hlinkClick r:id="${newRelId}"/>`
      );
      const updatedRun = originalRun.replace(rprContent, updatedRprContent);
      paraContent = paraContent.replace(originalRun, updatedRun);
    } else {
      // Add hlinkClick to rPr
      const updatedRprContent = rprContent + `<a:hlinkClick r:id="${newRelId}"/>`;
      const updatedRun = originalRun.replace(rprContent, updatedRprContent);
      paraContent = paraContent.replace(originalRun, updatedRun);
    }
  } else {
    // Add rPr with hlinkClick
    const textMatch = /<a:t>([^<]*)<\/a:t>/.exec(originalRun);
    const text = textMatch ? textMatch[1] : "";
    const updatedRun = originalRun.replace(
      `<a:t>${text}</a:t>`,
      `<a:rPr lang="en-US"><a:hlinkClick r:id="${newRelId}"/></a:rPr><a:t>${text}</a:t>`
    );
    paraContent = paraContent.replace(originalRun, updatedRun);
  }

  // Update relationships
  relsXml = relsXml.replace("</Relationships>", `${newRelEntry}</Relationships>`);

  // Update slide XML
  const updatedSlideXml = slideXml.replace(originalPara, paraContent);

  return ok({ slideXml: updatedSlideXml, relsXml });
}

/**
 * Sets a hyperlink on a paragraph (all runs).
 */
function setHyperlinkOnParagraph(
  slideXml: string,
  relationships: Array<{ id: string; target: string; type?: string }>,
  paraIndex: number,
  url: string
): Result<{ slideXml: string; relsXml: string }> {
  // Find paragraph
  const paraPattern = /<a:p(?:[^>]*)>([\s\S]*?)<\/a:p>/g;
  const paraMatches = slideXml.match(paraPattern) || [];

  if (paraIndex < 1 || paraIndex > paraMatches.length) {
    return invalidInput(`Paragraph index ${paraIndex} out of range`);
  }

  const originalPara = paraMatches[paraIndex - 1];
  let paraContent = originalPara;

  // Generate new relationship ID
  const existingRelIds = relationships.map(r => r.id);
  const newRelId = generateRelId(existingRelIds);

  // Create relationship entry
  const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url}" TargetMode="External"/>`;

  // Get existing relsXml
  const relsMatch = slideXml.match(/<Relationships[^>]*>([\s\S]*?)<\/Relationships>/);
  let relsXml = relsMatch ? relsMatch[0] : "";

  // Process all runs in paragraph
  const runPattern = /<a:r(?:[^>]*)>([\s\S]*?)<\/a:r>/g;
  const runMatches = paraContent.match(runPattern) || [];

  for (const originalRun of runMatches) {
    const rprMatch = /<a:rPr(?:[^>]*)>([\s\S]*?)<\/a:rPr>/.exec(originalRun);
    if (rprMatch) {
      const rprContent = rprMatch[1];
      if (/<a:hlinkClick/.test(rprContent)) {
        // Replace existing hlinkClick
        const updatedRprContent = rprContent.replace(
          /<a:hlinkClick(?:\s[^>]*)?\sr:id="[^"]*"[^>]*>/,
          `<a:hlinkClick r:id="${newRelId}"/>`
        );
        const updatedRun = originalRun.replace(rprContent, updatedRprContent);
        paraContent = paraContent.replace(originalRun, updatedRun);
      } else {
        // Add hlinkClick to rPr
        const updatedRprContent = rprContent + `<a:hlinkClick r:id="${newRelId}"/>`;
        const updatedRun = originalRun.replace(rprContent, updatedRprContent);
        paraContent = paraContent.replace(originalRun, updatedRun);
      }
    } else {
      // Add rPr with hlinkClick
      const textMatch = /<a:t>([^<]*)<\/a:t>/.exec(originalRun);
      const text = textMatch ? textMatch[1] : "";
      const updatedRun = originalRun.replace(
        `<a:t>${text}</a:t>`,
        `<a:rPr lang="en-US"><a:hlinkClick r:id="${newRelId}"/></a:rPr><a:t>${text}</a:t>`
      );
      paraContent = paraContent.replace(originalRun, updatedRun);
    }
  }

  // Update relationships
  relsXml = relsXml.replace("</Relationships>", `${newRelEntry}</Relationships>`);

  // Update slide XML
  const updatedSlideXml = slideXml.replace(originalPara, paraContent);

  return ok({ slideXml: updatedSlideXml, relsXml });
}

/**
 * Sets a hyperlink on a shape.
 */
function setHyperlinkOnShape(
  slideXml: string,
  relationships: Array<{ id: string; target: string; type?: string }>,
  shapeIndex: number,
  url: string
): Result<{ slideXml: string; relsXml: string }> {
  // Find all shapes (p:sp elements)
  const shapePattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  const shapeMatches = slideXml.match(shapePattern) || [];

  if (shapeIndex < 1 || shapeIndex > shapeMatches.length) {
    return notFound("Shape", String(shapeIndex));
  }

  const originalShape = shapeMatches[shapeIndex - 1];

  // Generate new relationship ID
  const existingRelIds = relationships.map(r => r.id);
  const newRelId = generateRelId(existingRelIds);

  // Create relationship entry
  const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url}" TargetMode="External"/>`;

  // Get existing relsXml
  const relsMatch = slideXml.match(/<Relationships[^>]*>([\s\S]*?)<\/Relationships>/);
  let relsXml = relsMatch ? relsMatch[0] : "";

  let updatedShape = originalShape;

  // Check if shape already has hlinkClick in nvPr
  if (/<a:hlinkClick/.test(originalShape)) {
    // Replace existing hlinkClick
    updatedShape = originalShape.replace(
      /<a:hlinkClick(?:\s[^>]*)?\sr:id="[^"]*"[^>]*>/,
      `<a:hlinkClick r:id="${newRelId}"/>`
    );
  } else {
    // Add hlinkClick to cNvPr or nvPr
    // Find the cNvPr element and add hlinkClick as child
    const cnvpMatch = /<p:cNvPr([^>]*)>([\s\S]*?)<\/p:cNvPr>/.exec(updatedShape);
    if (cnvpMatch) {
      const attrs = cnvpMatch[1];
      const content = cnvpMatch[2];
      const updatedContent = content + `<a:hlinkClick r:id="${newRelId}"/>`;
      updatedShape = updatedShape.replace(
        `<p:cNvPr${attrs}>${content}</p:cNvPr>`,
        `<p:cNvPr${attrs}>${updatedContent}</p:cNvPr>`
      );
    }
  }

  // Update relationships
  relsXml = relsXml.replace("</Relationships>", `${newRelEntry}</Relationships>`);

  // Update slide XML
  const updatedSlideXml = slideXml.replace(originalShape, updatedShape);

  return ok({ slideXml: updatedSlideXml, relsXml });
}

/**
 * Removes a hyperlink from a shape or text.
 *
 * @param filePath - Path to the PPTX file
 * @param path - Path to the shape or text (e.g., "/slide[1]/shape[1]" or "/slide[1]/shape[1]/paragraph[1]/run[1]")
 *
 * @example
 * const result = await removeHyperlink("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 */
export async function removeHyperlink(
  filePath: string,
  path: string
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(path);
    if (slideIndex === null) {
      return invalidInput("Invalid path - must include slide index");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    const relsXml = requireEntry(zip, relsEntry);

    // Check for run-level or paragraph-level path
    const paraIndex = extractParagraphIndex(path);
    const runIndex = extractRunIndex(path);

    let updatedSlideXml: string;

    if (paraIndex !== null && runIndex !== null) {
      // Remove hyperlink from specific run
      updatedSlideXml = removeHyperlinkFromRun(slideXml, paraIndex, runIndex);
    } else if (paraIndex !== null) {
      // Remove hyperlink from paragraph
      updatedSlideXml = removeHyperlinkFromParagraph(slideXml, paraIndex);
    } else {
      // Remove hyperlink from shape
      const shapeIndex = extractShapeIndex(path);
      if (shapeIndex === null) {
        return invalidInput("Invalid shape path - must include shape[index]");
      }
      updatedSlideXml = removeHyperlinkFromShape(slideXml, shapeIndex);
    }

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
 * Removes hyperlink from a specific run.
 */
function removeHyperlinkFromRun(slideXml: string, paraIndex: number, runIndex: number): string {
  // Find paragraph
  const paraPattern = /<a:p(?:[^>]*)>([\s\S]*?)<\/a:p>/g;
  const paraMatches = slideXml.match(paraPattern) || [];

  if (paraIndex < 1 || paraIndex > paraMatches.length) {
    return slideXml; // Return unchanged
  }

  const originalPara = paraMatches[paraIndex - 1];
  let paraContent = originalPara;

  // Find run
  const runPattern = /<a:r(?:[^>]*)>([\s\S]*?)<\/a:r>/g;
  const runMatches = paraContent.match(runPattern) || [];

  if (runIndex < 1 || runIndex > runMatches.length) {
    return slideXml; // Return unchanged
  }

  const originalRun = runMatches[runIndex - 1];

  // Remove hlinkClick from rPr
  const rprMatch = /<a:rPr(?:[^>]*)>([\s\S]*?)<\/a:rPr>/.exec(originalRun);
  if (rprMatch) {
    const rprContent = rprMatch[1];
    if (/<a:hlinkClick/.test(rprContent)) {
      const updatedRprContent = rprContent.replace(/<a:hlinkClick(?:\s[^>]*)?\sr:id="[^"]*"[^>]*>/g, "");
      const updatedRun = originalRun.replace(rprContent, updatedRprContent);
      paraContent = paraContent.replace(originalRun, updatedRun);
    }
  }

  return slideXml.replace(originalPara, paraContent);
}

/**
 * Removes hyperlink from a paragraph.
 */
function removeHyperlinkFromParagraph(slideXml: string, paraIndex: number): string {
  // Find paragraph
  const paraPattern = /<a:p(?:[^>]*)>([\s\S]*?)<\/a:p>/g;
  const paraMatches = slideXml.match(paraPattern) || [];

  if (paraIndex < 1 || paraIndex > paraMatches.length) {
    return slideXml; // Return unchanged
  }

  const originalPara = paraMatches[paraIndex - 1];
  let paraContent = originalPara;

  // Remove hlinkClick from all runs
  paraContent = paraContent.replace(/<a:hlinkClick(?:\s[^>]*)?\sr:id="[^"]*"[^>]*>/g, "");

  return slideXml.replace(originalPara, paraContent);
}

/**
 * Removes hyperlink from a shape.
 */
function removeHyperlinkFromShape(slideXml: string, shapeIndex: number): string {
  // Find all shapes (p:sp elements)
  const shapePattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  const shapeMatches = slideXml.match(shapePattern) || [];

  if (shapeIndex < 1 || shapeIndex > shapeMatches.length) {
    return slideXml; // Return unchanged
  }

  const originalShape = shapeMatches[shapeIndex - 1];

  // Remove hlinkClick
  const updatedShape = originalShape.replace(/<a:hlinkClick(?:\s[^>]*)?\sr:id="[^"]*"[^>]*>/g, "");

  return slideXml.replace(originalShape, updatedShape);
}

/**
 * Sets an external hyperlink on a shape or text.
 * This is a convenience method that calls setHyperlink with an external URL.
 *
 * @param filePath - Path to the PPTX file
 * @param path - Path to the shape or text
 * @param url - The external URL to link to
 *
 * @example
 * const result = await setExternalHyperlink(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/shape[1]",
 *   "https://example.com"
 * );
 */
export async function setExternalHyperlink(
  filePath: string,
  path: string,
  url: string
): Promise<Result<void>> {
  return setHyperlink(filePath, path, url);
}

/**
 * Sets an internal hyperlink to another slide.
 *
 * @param filePath - Path to the PPTX file
 * @param path - Path to the shape or text to attach the link to
 * @param targetSlideIndex - 1-based index of the target slide
 *
 * @example
 * const result = await setInternalHyperlink(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/shape[1]",
 *   3  // Link to slide 3
 * );
 */
export async function setInternalHyperlink(
  filePath: string,
  path: string,
  targetSlideIndex: number
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(path);
    if (slideIndex === null) {
      return invalidInput("Invalid path - must include slide index");
    }

    // Validate target slide index
    if (targetSlideIndex < 1) {
      return invalidInput("Target slide index must be 1 or greater");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Check that target slide exists
    const targetPathResult = getSlideEntryPath(zip, targetSlideIndex);
    if (!targetPathResult.ok) {
      return err(targetPathResult.error?.code ?? "slide_not_found", targetPathResult.error?.message ?? "Failed to get target slide path");
    }

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    let relsXml = "";
    try {
      relsXml = requireEntry(zip, relsEntry);
    } catch {
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    }
    const relationships = parseRelationshipEntries(relsXml);

    // Generate new relationship ID
    const existingRelIds = relationships.map(r => r.id);
    const newRelId = generateRelId(existingRelIds);

    // Create internal link relationship (using ppaction for internal slide link)
    // The target format is "slide[index]" for internal links
    const internalTarget = `slide${targetSlideIndex}.xml`;
    const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="${internalTarget}"/>`;

    // Get existing relsXml
    const relsMatch = slideXml.match(/<Relationships[^>]*>([\s\S]*?)<\/Relationships>/);
    let currentRelsXml = relsMatch ? relsMatch[0] : "";

    // Check for run-level or paragraph-level path
    const paraIndex = extractParagraphIndex(path);
    const runIndex = extractRunIndex(path);

    let updatedSlideXml: string;

    if (paraIndex !== null && runIndex !== null) {
      // Set hyperlink on specific run
      const result = setHyperlinkOnRun(slideXml, relationships, paraIndex, runIndex, `#${internalTarget}`);
      if (!result.ok) {
        return err(result.error?.code ?? "operation_failed", result.error?.message ?? "Failed to set hyperlink on run");
      }
      if (!result.data) {
        return err("operation_failed", "No data returned");
      }
      updatedSlideXml = result.data.slideXml;
      currentRelsXml = result.data.relsXml;
    } else if (paraIndex !== null) {
      // Set hyperlink on paragraph
      const result = setHyperlinkOnParagraph(slideXml, relationships, paraIndex, `#${internalTarget}`);
      if (!result.ok) {
        return err(result.error?.code ?? "operation_failed", result.error?.message ?? "Failed to set hyperlink on paragraph");
      }
      if (!result.data) {
        return err("operation_failed", "No data returned");
      }
      updatedSlideXml = result.data.slideXml;
      currentRelsXml = result.data.relsXml;
    } else {
      // Set hyperlink on shape
      const shapeIndex = extractShapeIndex(path);
      if (shapeIndex === null) {
        return invalidInput("Invalid shape path - must include shape[index]");
      }
      const result = setHyperlinkOnShape(slideXml, relationships, shapeIndex, `#${internalTarget}`);
      if (!result.ok) {
        return err(result.error?.code ?? "operation_failed", result.error?.message ?? "Failed to set hyperlink on shape");
      }
      if (!result.data) {
        return err("operation_failed", "No data returned");
      }
      updatedSlideXml = result.data.slideXml;
      currentRelsXml = result.data.relsXml;
    }

    // Update relationships to include the slide relationship
    const updatedRelsXml = currentRelsXml.replace("</Relationships>", `${newRelEntry}</Relationships>`);

    // Build new zip with updated files
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else if (name === relsEntry) {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
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
