/**
 * Core mutation operations for @officekit/ppt.
 *
 * Provides the foundational mutation capabilities:
 * - Set operations: Modify existing elements
 * - Remove operations: Delete elements
 * - Swap operations: Exchange elements
 * - CopyFrom operations: Copy elements
 * - Raw operations: Direct XML manipulation
 * - Batch operations: Multiple operations at once
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, andThen, map, invalidInput, notFound } from "./result.js";
import type { Result } from "./types.js";
import { parsePath, buildPath, getSlideIndex } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Represents a batch operation to be executed.
 */
export interface BatchOperation {
  /** Operation type */
  op: "set" | "remove" | "swap" | "copyFrom" | "rawSet" | "setShapeText";
  /** Operation parameters */
  params: Record<string, unknown>;
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
 * Gets the relationships entry name for a given entry.
 */
function getRelationshipsEntryName(entryName: string): string {
  const directory = path.posix.dirname(entryName);
  const basename = path.posix.basename(entryName);
  return path.posix.join(directory, "_rels", `${basename}.rels`);
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
 * Updates slide IDs in presentation.xml with new order.
 */
function reorderSlideIds(presentationXml: string, orderedIds: Array<{ id: string; relId: string }>): string {
  let result = presentationXml.replace(/<p:sldId\b[^>]*\/?>/g, "");

  const sldIdListMatch = result.match(/(<p:sldIdLst[^>]*>)([\s\S]*?)(<\/p:sldIdLst>)/);
  if (sldIdListMatch) {
    const newSlideIds = orderedIds.map(s => `<p:sldId id="${s.id}" r:id="${s.relId}"/>`).join("\n      ");
    result = result.replace(
      /<p:sldIdLst[^>]*>[\s\S]*?<\/p:sldIdLst>/,
      `<p:sldIdLst>\n      ${newSlideIds}\n    </p:sldIdLst>`
    );
  }

  return result;
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
 * Generates a unique slide ID.
 */
function generateSlideId(existingIds: number[]): string {
  let id = 256;
  while (existingIds.includes(id)) {
    id++;
  }
  return String(id);
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

// ============================================================================
// Slide Index Resolution
// ============================================================================

/**
 * Resolves a slide path or index to a 1-based slide index.
 */
async function resolveSlideIndex(filePath: string, slideRef: number | string): Promise<Result<number>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);
    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const slideIds = getSlideIds(presentationXml);

    if (typeof slideRef === "number") {
      if (slideRef < 1 || slideRef > slideIds.length) {
        return invalidInput(`Slide index ${slideRef} is out of range (1-${slideIds.length})`);
      }
      return ok(slideRef);
    }

    // slideRef is a path like "/slide[1]"
    const slideIndexMatch = slideRef.match(/^\/slide\[(\d+)\]/i);
    if (slideIndexMatch) {
      const index = parseInt(slideIndexMatch[1], 10);
      if (index < 1 || index > slideIds.length) {
        return invalidInput(`Slide index ${index} is out of range (1-${slideIds.length})`);
      }
      return ok(index);
    }

    return invalidInput(`Invalid slide reference: ${slideRef}`);
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
// Set Operations
// ============================================================================

/**
 * Sets the raw XML for an element at the given path.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the element (e.g., "/slide[1]/shape[2]")
 * @param xml - Raw XML to set
 *
 * @example
 * const result = await rawSet("/path/to/presentation.pptx", "/slide[1]/shape[1]", "<p:sp>...</p:sp>");
 */
export async function rawSet(filePath: string, pptPath: string, xml: string): Promise<Result<void>> {
  try {
    const parsed = parsePath(pptPath);
    if (!parsed.ok) {
      return err(parsed.error?.code ?? "invalid_path", parsed.error?.message ?? "Failed to parse path");
    }

    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("rawSet requires a slide path");
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
    const updatedSlideXml = rawSetElementInSlide(slideXml, pptPath, xml);

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
 * Sets an element's XML within a slide by matching the path.
 */
function rawSetElementInSlide(slideXml: string, pptPath: string, xml: string): string {
  // Extract the target element type and index from the path
  // e.g., "/slide[1]/shape[2]" -> type="shape", index=2
  const segments = pptPath.split("/").filter(s => s);

  // Find the last element type (the target element)
  let targetType = "";
  let targetIndex = 0;

  for (const segment of segments) {
    const match = segment.match(/^([a-zA-Z]+)\[(\d+)\]/);
    if (match) {
      targetType = match[1].toLowerCase();
      targetIndex = parseInt(match[2], 10);
    }
  }

  if (!targetType || targetIndex === 0) {
    throw new Error(`Invalid path format: ${pptPath}`);
  }

  // Map element types to XML element names
  const elementTypeMap: Record<string, string> = {
    shape: "p:sp",
    textbox: "p:sp",
    table: "a:tbl",
    picture: "p:pic",
    pic: "p:pic",
    chart: "c:chart",
    placeholder: "p:sp",
    connector: "p:cxnSp",
    group: "p:grpSp",
  };

  const xmlElementName = elementTypeMap[targetType] || targetType;

  // Find and replace the target element
  // This is a simplified approach - for complex shapes we need better matching
  let count = 0;
  const pattern = new RegExp(`<${xmlElementName}[\\s\\S]*?</${xmlElementName}>`, "g");
  return slideXml.replace(pattern, (match) => {
    count++;
    if (count === targetIndex) {
      return xml;
    }
    return match;
  });
}

// ============================================================================
// Raw Get Operations
// ============================================================================

/**
 * Gets the raw XML for an element at the given path.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the element (e.g., "/slide[1]/shape[2]")
 *
 * @example
 * const result = await rawGet("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 * if (result.ok) {
 *   console.log(result.data.xml);
 * }
 */
export async function rawGet(filePath: string, pptPath: string): Promise<Result<{ xml: string }>> {
  try {
    const parsed = parsePath(pptPath);
    if (!parsed.ok) {
      return err(parsed.error?.code ?? "invalid_path", parsed.error?.message ?? "Failed to parse path");
    }

    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("rawGet requires a slide path");
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

    // Extract element XML
    const elementXml = rawGetElementFromSlide(slideXml, pptPath);
    if (!elementXml) {
      return notFound("Element", pptPath);
    }

    return ok({ xml: elementXml });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets an element's XML from a slide by matching the path.
 */
function rawGetElementFromSlide(slideXml: string, pptPath: string): string | null {
  // Extract the target element type and index from the path
  const segments = pptPath.split("/").filter(s => s);

  let targetType = "";
  let targetIndex = 0;

  for (const segment of segments) {
    const match = segment.match(/^([a-zA-Z]+)\[(\d+)\]/);
    if (match) {
      targetType = match[1].toLowerCase();
      targetIndex = parseInt(match[2], 10);
    }
  }

  if (!targetType || targetIndex === 0) {
    return null;
  }

  // Map element types to XML element names
  const elementTypeMap: Record<string, string> = {
    shape: "p:sp",
    textbox: "p:sp",
    table: "a:tbl",
    picture: "p:pic",
    pic: "p:pic",
    chart: "c:chart",
    placeholder: "p:sp",
    connector: "p:cxnSp",
    group: "p:grpSp",
  };

  const xmlElementName = elementTypeMap[targetType] || targetType;

  // Find and return the target element
  let count = 0;
  const pattern = new RegExp(`<${xmlElementName}[\\s\\S]*?</${xmlElementName}>`, "g");
  let match;

  while ((match = pattern.exec(slideXml)) !== null) {
    count++;
    if (count === targetIndex) {
      return match[0];
    }
  }

  return null;
}

// ============================================================================
// Swap Operations
// ============================================================================

/**
 * Swaps two slides in the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param index1 - 1-based index of the first slide
 * @param index2 - 1-based index of the second slide
 *
 * @example
 * const result = await swapSlides("/path/to/presentation.pptx", 1, 3);
 */
export async function swapSlides(filePath: string, index1: number, index2: number): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const slideIds = getSlideIds(presentationXml);

    if (index1 < 1 || index1 > slideIds.length) {
      return invalidInput(`Slide index ${index1} is out of range (1-${slideIds.length})`);
    }
    if (index2 < 1 || index2 > slideIds.length) {
      return invalidInput(`Slide index ${index2} is out of range (1-${slideIds.length})`);
    }

    // Create new ordering with swapped slides
    const newSlideIds = [...slideIds];
    const temp = newSlideIds[index1 - 1];
    newSlideIds[index1 - 1] = newSlideIds[index2 - 1];
    newSlideIds[index2 - 1] = temp;

    const updatedPresentationXml = reorderSlideIds(presentationXml, newSlideIds);

    // Build new zip
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === "ppt/presentation.xml") {
        newEntries.push({ name, data: Buffer.from(updatedPresentationXml, "utf8") });
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
 * Swaps two shapes on the same slide.
 *
 * @param filePath - Path to the PPTX file
 * @param path1 - PPT path to the first shape (e.g., "/slide[1]/shape[1]")
 * @param path2 - PPT path to the second shape (e.g., "/slide[1]/shape[2]")
 *
 * @example
 * const result = await swapShapes("/path/to/presentation.pptx", "/slide[1]/shape[1]", "/slide[1]/shape[2]");
 */
export async function swapShapes(filePath: string, path1: string, path2: string): Promise<Result<void>> {
  try {
    const slideIndex1 = getSlideIndex(path1);
    const slideIndex2 = getSlideIndex(path2);

    if (slideIndex1 === null || slideIndex2 === null) {
      return invalidInput("Both paths must be slide paths");
    }

    if (slideIndex1 !== slideIndex2) {
      return invalidInput("Both shapes must be on the same slide to swap");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex1);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);

    // Extract shape indices
    const shape1Index = extractIndexFromPath(path1, "shape");
    const shape2Index = extractIndexFromPath(path2, "shape");

    if (shape1Index === null || shape2Index === null) {
      return invalidInput("Invalid shape paths");
    }

    const updatedSlideXml = swapElementsInSlide(slideXml, "p:sp", shape1Index, shape2Index);

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
 * Extracts the index of an element from a path.
 */
function extractIndexFromPath(pptPath: string, elementType: string): number | null {
  const pattern = new RegExp(`/${elementType}\\[(\\d+)\\]`, "i");
  const match = pptPath.match(pattern);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * Swaps two elements of a given type in a slide.
 */
function swapElementsInSlide(slideXml: string, xmlElementName: string, index1: number, index2: number): string {
  const pattern = new RegExp(`<${xmlElementName}[\\s\\S]*?</${xmlElementName}>`, "g");
  const elements: string[] = [];
  let match;

  while ((match = pattern.exec(slideXml)) !== null) {
    elements.push(match[0]);
  }

  if (index1 < 1 || index1 > elements.length || index2 < 1 || index2 > elements.length) {
    throw new Error("Element index out of range");
  }

  // Swap the elements
  const temp = elements[index1 - 1];
  elements[index1 - 1] = elements[index2 - 1];
  elements[index2 - 1] = temp;

  // Replace elements in the XML
  let result = slideXml;
  let count = 0;
  result = result.replace(pattern, () => elements[count++]);

  return result;
}

// ============================================================================
// CopyFrom Operations
// ============================================================================

/**
 * Copies a shape from one slide to another.
 *
 * @param filePath - Path to the PPTX file
 * @param sourcePath - PPT path to the source shape (e.g., "/slide[1]/shape[1]")
 * @param targetSlideIndex - 1-based index of the target slide
 *
 * @example
 * const result = await copyShape("/path/to/presentation.pptx", "/slide[1]/shape[1]", 2);
 * if (result.ok) {
 *   console.log(result.data.path); // Path to the new shape
 * }
 */
export async function copyShape(
  filePath: string,
  sourcePath: string,
  targetSlideIndex: number,
): Promise<Result<{ path: string }>> {
  try {
    const sourceSlideIndex = getSlideIndex(sourcePath);
    if (sourceSlideIndex === null) {
      return invalidInput("Source path must be a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Get source slide
    const sourceSlidePathResult = getSlideEntryPath(zip, sourceSlideIndex);
    if (!sourceSlidePathResult.ok) {
      return err(sourceSlidePathResult.error?.code ?? "slide_not_found", sourceSlidePathResult.error?.message ?? "Failed to get source slide path");
    }

    const sourceSlideEntry = sourceSlidePathResult.data;
    if (!sourceSlideEntry) {
      return err("slide_not_found", "Source slide entry not found");
    }
    const sourceSlideXml = requireEntry(zip, sourceSlideEntry);

    // Get target slide
    const targetSlidePathResult = getSlideEntryPath(zip, targetSlideIndex);
    if (!targetSlidePathResult.ok) {
      return err(targetSlidePathResult.error?.code ?? "slide_not_found", targetSlidePathResult.error?.message ?? "Failed to get target slide path");
    }

    const targetSlideEntry = targetSlidePathResult.data;
    if (!targetSlideEntry) {
      return err("slide_not_found", "Target slide entry not found");
    }
    const targetSlideXml = requireEntry(zip, targetSlideEntry);

    // Extract source shape XML
    const shapeIndex = extractIndexFromPath(sourcePath, "shape");
    if (shapeIndex === null) {
      return invalidInput("Invalid source shape path");
    }

    const shapeXml = extractElementFromSlide(sourceSlideXml, "p:sp", shapeIndex);
    if (!shapeXml) {
      return notFound("Shape", sourcePath);
    }

    // Find next shape index in target slide
    const targetShapeCount = countElementsInSlide(targetSlideXml, "p:sp");
    const newShapeIndex = targetShapeCount + 1;

    // Update shape IDs to avoid conflicts
    let newShapeXml = shapeXml;
    // Generate unique IDs for the copied shape
    const maxId = findMaxShapeId(targetSlideXml);
    const newId = maxId + 1;
    newShapeXml = replaceShapeId(newShapeXml, newId);

    // Insert shape into target slide
    const updatedTargetSlideXml = insertElementIntoSlide(targetSlideXml, newShapeXml, "p:sp");

    // Build new zip with updated slides
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === targetSlideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedTargetSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));

    return ok({ path: `/slide[${targetSlideIndex}]/shape[${newShapeIndex}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Counts elements of a given type in a slide.
 */
function countElementsInSlide(slideXml: string, xmlElementName: string): number {
  const pattern = new RegExp(`<${xmlElementName}[\\s\\S]*?</${xmlElementName}>`, "g");
  const matches = slideXml.match(pattern);
  return matches ? matches.length : 0;
}

/**
 * Extracts an element from a slide by index.
 */
function extractElementFromSlide(slideXml: string, xmlElementName: string, index: number): string | null {
  const pattern = new RegExp(`<${xmlElementName}[\\s\\S]*?</${xmlElementName}>`, "g");
  let count = 0;
  let match;

  while ((match = pattern.exec(slideXml)) !== null) {
    count++;
    if (count === index) {
      return match[0];
    }
  }

  return null;
}

/**
 * Finds the maximum shape ID in a slide.
 */
function findMaxShapeId(slideXml: string): number {
  const pattern = /<[pc]:(cNvPr|spPr|cNvSpPr)[^>]*\bid="(\d+)"/g;
  let maxId = 0;
  let match;

  while ((match = pattern.exec(slideXml)) !== null) {
    const id = parseInt(match[2], 10);
    if (id > maxId) {
      maxId = id;
    }
  }

  return maxId;
}

/**
 * Replaces the shape ID in an element's XML.
 */
function replaceShapeId(shapeXml: string, newId: number): string {
  // Replace id attributes in cNvPr elements
  return shapeXml.replace(/(<[pc]:cNvPr[^>]*\bid=")\d+(")/g, `$1${newId}$2`);
}

/**
 * Inserts an element into a slide's shape tree.
 */
function insertElementIntoSlide(slideXml: string, newElementXml: string, xmlElementName: string): string {
  // Find the closing </p:spTree> tag and insert before it
  const spTreeClosePattern = /<\/p:spTree>/;
  if (spTreeClosePattern.test(slideXml)) {
    return slideXml.replace(spTreeClosePattern, `${newElementXml}\n  </p:spTree>`);
  }

  throw new Error("Could not find shape tree in slide XML");
}

/**
 * Copies a slide within the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param sourceIndex - 1-based index of the slide to copy
 * @param targetIndex - 1-based index where the copy should be inserted (or use -1 for end)
 *
 * @example
 * const result = await copySlide("/path/to/presentation.pptx", 1, 3);
 * if (result.ok) {
 *   console.log(result.data.path); // Path to the new slide
 * }
 */
export async function copySlide(
  filePath: string,
  sourceIndex: number,
  targetIndex: number,
): Promise<Result<{ path: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const contentTypesXml = zip.get("[Content_Types].xml")?.toString("utf8") ?? "";

    const slideIds = getSlideIds(presentationXml);
    const relationships = parseRelationshipEntries(relsXml);

    if (sourceIndex < 1 || sourceIndex > slideIds.length) {
      return invalidInput(`Source slide index ${sourceIndex} is out of range (1-${slideIds.length})`);
    }

    const insertPosition = targetIndex === -1 ? slideIds.length : targetIndex;
    if (insertPosition < 1 || insertPosition > slideIds.length + 1) {
      return invalidInput(`Target slide index ${targetIndex} is out of range (1-${slideIds.length + 1})`);
    }

    // Get source slide data
    const sourceSlide = slideIds[sourceIndex - 1];
    const sourceRel = relationships.find(r => r.id === sourceSlide.relId);
    const sourceSlidePath = normalizeZipPath("ppt", sourceRel?.target ?? "");
    const sourceSlideXml = requireEntry(zip, sourceSlidePath);
    const sourceSlideRelsPath = getRelationshipsEntryName(sourceSlidePath);
    const sourceSlideRelsXml = zip.get(sourceSlideRelsPath)?.toString("utf8") ?? "";

    // Generate new IDs
    const existingIds = slideIds.map(s => parseInt(s.id, 10));
    const existingRelIds = relationships.map(r => r.id);

    const newSlideId = generateSlideId(existingIds);
    const newRelId = generateRelId(existingRelIds);

    // Determine new slide index
    const newSlideIndex = slideIds.length + 1;
    const newSlideEntry = `ppt/slides/slide${newSlideIndex}.xml`;
    const newSlideRelsEntry = `ppt/slides/_rels/slide${newSlideIndex}.xml.rels`;

    // Update slide ID in the duplicated slide
    let newSlideXml = sourceSlideXml;
    const idMatch = newSlideXml.match(/<p:cNvPr\b([^>]*)\bid="(\d+)"([^>]*)>/);
    if (idMatch) {
      newSlideXml = newSlideXml.replace(
        new RegExp(`(<p:cNvPr\\b[^>]*\\bid=")${idMatch[2]}("[^>]*>)`, "g"),
        `$1${newSlideId}$2`
      );
    }

    // Create new relationship for the duplicated slide (same layout as source)
    const newSlideRelsXml = sourceSlideRelsXml.replace(
      /Id="([^"]+)"\s+Type="[^"]*slideLayout"[^>]*Target="([^"]+)"/,
      `Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../${sourceRel?.target ?? ""}"`
    );

    // Add new slide ID to presentation.xml (insert at target position)
    const newSlideIds = [...slideIds];
    newSlideIds.splice(insertPosition - 1, 0, { id: newSlideId, relId: newRelId });
    const updatedPresentationXml = reorderSlideIds(presentationXml, newSlideIds);

    // Add new relationship to presentation.xml.rels
    const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideIndex}.xml"/>`;
    const updatedRelsXml = relsXml.replace(/<\/Relationships>/, `  ${newRelEntry}\n</Relationships>`);

    // Add Content_Type entry
    let updatedContentTypes = contentTypesXml;
    if (!contentTypesXml.includes(`slide${newSlideIndex}.xml`)) {
      const slideContentType = `<Override PartName="/ppt/slides/slide${newSlideIndex}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
      updatedContentTypes = updatedContentTypes.replace(/<\/Types>/, `  ${slideContentType}\n</Types>`);
    }

    // Build new zip
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === "ppt/presentation.xml") {
        newEntries.push({ name, data: Buffer.from(updatedPresentationXml, "utf8") });
      } else if (name === "ppt/_rels/presentation.xml.rels") {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
      } else if (name === "[Content_Types].xml") {
        newEntries.push({ name, data: Buffer.from(updatedContentTypes, "utf8") });
      } else if (name !== newSlideEntry && name !== newSlideRelsEntry) {
        newEntries.push({ name, data });
      }
    }

    // Add new slide and its relationships
    newEntries.push({ name: newSlideEntry, data: Buffer.from(newSlideXml, "utf8") });
    newEntries.push({ name: newSlideRelsEntry, data: Buffer.from(newSlideRelsXml, "utf8") });

    await writeFile(filePath, createStoredZip(newEntries));

    return ok({ path: `/slide[${insertPosition}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Batch Operations
// ============================================================================

/**
 * Executes multiple mutations in sequence.
 *
 * @param filePath - Path to the PPTX file
 * @param operations - Array of operations to execute
 *
 * @example
 * const result = await batch("/path/to/presentation.pptx", [
 *   { op: "setShapeText", params: { path: "/slide[1]/shape[1]", text: "Hello" } },
 *   { op: "setShapeText", params: { path: "/slide[1]/shape[2]", text: "World" } },
 * ]);
 */
export async function batch(filePath: string, operations: BatchOperation[]): Promise<Result<void>> {
  // For batch operations, we need to open the zip once and apply all changes
  // This is more efficient than opening/closing for each operation
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Track modified entries
    const modifiedEntries = new Map<string, string>();

    for (const operation of operations) {
      switch (operation.op) {
        case "rawSet": {
          const { path, xml } = operation.params as { path: string; xml: string };
          const slideIndex = getSlideIndex(path);
          if (slideIndex === null) {
            return invalidInput("rawSet requires a slide path");
          }

          const slidePathResult = getSlideEntryPathFromZip(zip, slideIndex);
          if (!slidePathResult.ok) {
            return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
          }

          const slideEntry = slidePathResult.data;
          if (!slideEntry) {
            return err("slide_not_found", "Slide entry not found");
          }

          let slideXml = modifiedEntries.get(slideEntry) || requireEntry(zip, slideEntry);
          slideXml = rawSetElementInSlide(slideXml, path, xml);
          modifiedEntries.set(slideEntry, slideXml);
          break;
        }

        default:
          return invalidInput(`Unknown batch operation: ${operation.op}`);
      }
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
 * Gets the slide entry path from a zip (doesn't require file read).
 */
function getSlideEntryPathFromZip(zip: Map<string, Buffer>, slideIndex: number): Result<string> {
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
