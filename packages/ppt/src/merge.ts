/**
 * Template merge operations for @officekit/ppt.
 *
 * Provides functions to merge data into PPTX templates with {{key}} placeholders:
 * - Simple: {{name}} - replaced with value
 * - Conditional: {{#if condition}}...{{/if}} - conditional block
 * - Loop: {{#each items}}...{{/each}} - repeat block
 * - Formatted: {{date:format}} - with formatting
 */

import { readFile, writeFile, copyFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput } from "./result.js";
import type { Result } from "./types.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Result of a merge operation.
 */
export interface MergedResult {
  /** Number of simple placeholders replaced */
  replacements: number;
  /** Number of conditional blocks processed */
  conditionals: number;
  /** Number of loop blocks processed */
  loops: number;
  /** Number of slides processed */
  slidesProcessed: number;
}

/**
 * Placeholder data for merge operations.
 * Can be simple key-value pairs or nested objects/arrays.
 */
export type MergeData = Record<string, unknown>;

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
 * Escapes special XML characters for output.
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
 * Unescapes XML characters for matching.
 */
function unescapeXml(text: string): string {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

// ============================================================================
// Placeholder Processing
// ============================================================================

/**
 * Gets a nested value from an object using dot notation.
 */
function getNestedValue(obj: Record<string, unknown>, key: string): unknown {
  const parts = key.split(".");
  let current: unknown = obj;
  for (const part of parts) {
    if (current === null || current === undefined) {
      return undefined;
    }
    if (typeof current === "object") {
      current = (current as Record<string, unknown>)[part];
    } else {
      return undefined;
    }
  }
  return current;
}

/**
 * Formats a date value according to a format string.
 */
function formatDate(value: string | Date, format: string): string {
  const date = typeof value === "string" ? new Date(value) : value;
  if (isNaN(date.getTime())) {
    return String(value);
  }

  // Simple format tokens: yyyy, mm, dd, HH, MM, SS
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");

  return format
    .replace(/yyyy/g, String(year))
    .replace(/yy/g, String(year).slice(-2))
    .replace(/mm/g, month)
    .replace(/dd/g, day)
    .replace(/HH/g, hours)
    .replace(/MM/g, minutes)
    .replace(/SS/g, seconds);
}

/**
 * Converts a value to a string for insertion.
 */
function valueToString(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "object") {
    return JSON.stringify(value);
  }
  return String(value);
}

/**
 * Processes a single placeholder expression and returns the replacement value.
 */
function processPlaceholder(expr: string, data: Record<string, unknown>): string {
  // Check for formatted placeholder: key:format
  const formatMatch = expr.match(/^([^:]+):(.+)$/);
  if (formatMatch) {
    const key = formatMatch[1];
    const format = formatMatch[2];
    const value = getNestedValue(data, key);
    if (value instanceof Date || typeof value === "string") {
      return formatDate(value, format);
    }
    return valueToString(value);
  }

  // Simple key lookup
  const value = getNestedValue(data, expr);
  return valueToString(value);
}

/**
 * Processes conditional blocks: {{#if condition}}...{{/if}}
 */
function processConditionals(text: string, data: Record<string, unknown>): { result: string; count: number } {
  let count = 0;
  let result = text;

  // Match {{#if key}}...{{/if}} blocks
  const conditionalPattern = /\{\{#if\s+([^}]+)\}\}([\s\S]*?)\{\{\/if\}\}/g;

  result = result.replace(conditionalPattern, (match, condition, content) => {
    count++;
    const value = getNestedValue(data, condition.trim());
    // Truthy values: non-empty string, non-zero number, true, object, array
    const isTruthy = value !== false && value !== null && value !== undefined && value !== "";
    return isTruthy ? content : "";
  });

  return { result, count };
}

/**
 * Processes loop blocks: {{#each items}}...{{/each}}
 */
function processLoops(text: string, data: Record<string, unknown>): { result: string; count: number } {
  let count = 0;
  let result = text;

  // Match {{#each key}}...{{/each}} blocks
  const loopPattern = /\{\{#each\s+([^}]+)\}\}([\s\S]*?)\{\{\/each\}\}/g;

  result = result.replace(loopPattern, (match, arrayKey, content) => {
    count++;
    const array = getNestedValue(data, arrayKey.trim());
    if (!Array.isArray(array)) {
      return "";
    }

    // Process each item in the array
    const items: string[] = [];
    for (const item of array) {
      let itemContent = content;

      // Replace placeholders within the loop content using the item as context
      // but also allow access to parent data using ../ prefix
      if (typeof item === "object" && item !== null) {
        const itemData = item as Record<string, unknown>;
        // Replace simple {{key}} patterns within the loop
        itemContent = itemContent.replace(/\{\{([^}]+)\}\}/g, (placeholder: string, expr: string) => {
          // Handle ../ prefix for parent context access
          if (expr.startsWith("../")) {
            const parentKey = expr.slice(3);
            return processPlaceholder(parentKey, data);
          }
          // Try item context first, then parent
          const itemValue = getNestedValue(itemData, expr);
          if (itemValue !== undefined) {
            return valueToString(itemValue);
          }
          // Fall back to parent context
          return processPlaceholder(expr, data);
        });
      } else {
        // Primitive array item - replace {{.}} with the value
        itemContent = itemContent.replace(/\{\{\.\}\}/g, String(item));
      }

      items.push(itemContent);
    }

    return items.join("");
  });

  return { result, count };
}

/**
 * Processes simple placeholders: {{key}}
 */
function processSimplePlaceholders(text: string, data: Record<string, unknown>): { result: string; count: number } {
  let count = 0;
  let result = text;

  // Match {{expression}} but not {{#...}} or {{/...}}
  const placeholderPattern = /\{\{(?!#|\/)([^}]+)\}\}/g;

  result = result.replace(placeholderPattern, (match, expr) => {
    count++;
    return processPlaceholder(expr.trim(), data);
  });

  return { result, count };
}

/**
 * Processes all placeholders in text.
 */
function processText(text: string, data: Record<string, unknown>): { result: string; stats: Partial<MergedResult> } {
  // First process conditionals
  let { result: afterConditionals, count: condCount } = processConditionals(text, data);

  // Then process loops
  let { result: afterLoops, count: loopCount } = processLoops(afterConditionals, data);

  // Finally process simple placeholders
  let { result: afterPlaceholders, count: placeholderCount } = processSimplePlaceholders(afterLoops, data);

  return {
    result: afterPlaceholders,
    stats: {
      replacements: placeholderCount,
      conditionals: condCount,
      loops: loopCount,
    },
  };
}

// ============================================================================
// XML Text Extraction and Replacement
// ============================================================================

/**
 * Extracts all text runs from an XML string.
 */
function extractTextRuns(xml: string): Array<{ match: string; text: string }> {
  const runs: Array<{ match: string; text: string }> = [];
  // Match text runs: <a:r>...<a:t>text</a:t>...</a:r>
  const runPattern = /<a:r>([\s\S]*?)<\/a:r>/g;
  for (const match of xml.matchAll(runPattern)) {
    const textMatch = match[1].match(/<a:t>([^<]*)<\/a:t>/);
    if (textMatch) {
      runs.push({ match: match[0], text: textMatch[1] });
    }
  }
  return runs;
}

/**
 * Replaces text in text runs while preserving formatting.
 */
function replaceTextInXml(xml: string, data: Record<string, unknown>): { result: string; stats: Partial<MergedResult> } {
  const stats: Partial<MergedResult> = {
    replacements: 0,
    conditionals: 0,
    loops: 0,
  };

  let result = xml;

  // Find all <a:t> tags and process their content
  const textPattern = /(<a:t>)([^<]*)(<\/a:t>)/g;

  result = result.replace(textPattern, (match, open, content, close) => {
    // Only process if content contains placeholders
    if (!content.includes("{{")) {
      return match;
    }

    const { result: processedText, stats: textStats } = processText(content, data);
    stats.replacements = (stats.replacements || 0) + (textStats.replacements || 0);
    stats.conditionals = (stats.conditionals || 0) + (textStats.conditionals || 0);
    stats.loops = (stats.loops || 0) + (textStats.loops || 0);

    return `${open}${processedText}${close}`;
  });

  return { result, stats };
}

/**
 * Processes all text in a slide's XML.
 */
function processSlideXml(slideXml: string, data: Record<string, unknown>): { result: string; stats: Partial<MergedResult> } {
  const stats: Partial<MergedResult> = {
    replacements: 0,
    conditionals: 0,
    loops: 0,
  };

  let result = slideXml;

  // Process text in shapes (p:sp)
  const shapePattern = /<p:sp>([\s\S]*?)<\/p:sp>/g;
  result = result.replace(shapePattern, (shapeMatch, shapeContent) => {
    const { result: processedShape, stats: shapeStats } = replaceTextInXml(shapeContent, data);
    stats.replacements = (stats.replacements || 0) + (shapeStats.replacements || 0);
    stats.conditionals = (stats.conditionals || 0) + (shapeStats.conditionals || 0);
    stats.loops = (stats.loops || 0) + (shapeStats.loops || 0);
    return shapeMatch.replace(shapeContent, processedShape);
  });

  // Process text in text boxes (p:sp)
  const txBodyPattern = /<p:txBody>([\s\S]*?)<\/p:txBody>/g;
  result = result.replace(txBodyPattern, (txBodyMatch, txBodyContent) => {
    const { result: processedTxBody, stats: txBodyStats } = replaceTextInXml(txBodyContent, data);
    stats.replacements = (stats.replacements || 0) + (txBodyStats.replacements || 0);
    stats.conditionals = (stats.conditionals || 0) + (txBodyStats.conditionals || 0);
    stats.loops = (stats.loops || 0) + (txBodyStats.loops || 0);
    return txBodyMatch.replace(txBodyContent, processedTxBody);
  });

  // Process tables (a:tbl)
  const tablePattern = /<a:tbl>([\s\S]*?)<\/a:tbl>/g;
  result = result.replace(tablePattern, (tableMatch, tableContent) => {
    const { result: processedTable, stats: tableStats } = replaceTextInXml(tableContent, data);
    stats.replacements = (stats.replacements || 0) + (tableStats.replacements || 0);
    stats.conditionals = (stats.conditionals || 0) + (tableStats.conditionals || 0);
    stats.loops = (stats.loops || 0) + (tableStats.loops || 0);
    return tableMatch.replace(tableContent, processedTable);
  });

  return { result, stats };
}

/**
 * Processes notes XML for placeholders.
 */
function processNotesXml(notesXml: string, data: Record<string, unknown>): { result: string; stats: Partial<MergedResult> } {
  return replaceTextInXml(notesXml, data);
}

// ============================================================================
// Main Merge Function
// ============================================================================

/**
 * Merges data into a PPTX template, replacing {{key}} placeholders.
 *
 * @param templatePath - Path to a PPTX template with {{key}} placeholders
 * @param data - Object containing data to merge (e.g., { name: "John", date: "2024-01-01" })
 * @param outputPath - Where to save the merged result
 *
 * @returns Result with stats about replacements made
 *
 * @example
 * const result = await merge("/path/to/template.pptx", { name: "John", date: "2024-01-01" }, "/path/to/output.pptx");
 * if (result.ok) {
 *   console.log(result.data.replacements);
 * }
 *
 * Placeholder patterns:
 * - Simple: {{name}} - replaced with value
 * - Conditional: {{#if condition}}...{{/if}} - conditional block
 * - Loop: {{#each items}}...{{/each}} - repeat block
 * - Formatted: {{date:yyyy-mm-dd}} - with date formatting
 */
export async function merge(
  templatePath: string,
  data: MergeData,
  outputPath: string,
): Promise<Result<MergedResult>> {
  try {
    // Copy template to output path first
    await copyFile(templatePath, outputPath);

    // Read the PPTX file
    const buffer = await readFile(outputPath);
    const zip = readStoredZip(buffer);

    const presentationXml = zip.get("ppt/presentation.xml")?.toString("utf8") ?? "";
    const relsXml = zip.get("ppt/_rels/presentation.xml.rels")?.toString("utf8") ?? "";

    const slideIds = getSlideIds(presentationXml);
    const relationships = parseRelationshipEntries(relsXml);

    // Track overall stats
    const overallStats: MergedResult = {
      replacements: 0,
      conditionals: 0,
      loops: 0,
      slidesProcessed: 0,
    };

    // Process modified entries
    const modifiedEntries = new Map<string, string>();

    // Process each slide
    for (let i = 0; i < slideIds.length; i++) {
      const slide = slideIds[i];
      const slideRel = relationships.find(r => r.id === slide.relId);
      const slidePath = normalizeZipPath("ppt", slideRel?.target ?? "");

      let slideXml = zip.get(slidePath)?.toString("utf8");
      if (!slideXml) {
        continue;
      }

      // Process slide content
      const { result: processedSlide, stats: slideStats } = processSlideXml(slideXml, data);
      slideXml = processedSlide;

      // Track stats
      overallStats.replacements += slideStats.replacements || 0;
      overallStats.conditionals += slideStats.conditionals || 0;
      overallStats.loops += slideStats.loops || 0;

      // Check for notes and process them
      const slideRelsPath = getRelationshipsEntryName(slidePath);
      const slideRelsXml = zip.get(slideRelsPath)?.toString("utf8") ?? "";
      const slideRels = parseRelationshipEntries(slideRelsXml);
      const notesRel = slideRels.find(r => r.type?.endsWith("/notesSlide"));

      if (notesRel) {
        const notesPath = normalizeZipPath(path.posix.dirname(slidePath), notesRel.target);
        const notesXml = zip.get(notesPath)?.toString("utf8");
        if (notesXml) {
          const { result: processedNotes, stats: notesStats } = processNotesXml(notesXml, data);
          overallStats.replacements += notesStats.replacements || 0;
          overallStats.conditionals += notesStats.conditionals || 0;
          overallStats.loops += notesStats.loops || 0;
          modifiedEntries.set(notesPath, processedNotes);
        }
      }

      modifiedEntries.set(slidePath, slideXml);
      overallStats.slidesProcessed++;
    }

    // Build new zip with all modifications
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (modifiedEntries.has(name)) {
        newEntries.push({ name, data: Buffer.from(modifiedEntries.get(name)!, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(outputPath, createStoredZip(newEntries));

    return ok(overallStats);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}
