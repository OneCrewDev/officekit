/**
 * Shape mutation operations for @officekit/ppt.
 *
 * Provides functions to modify shapes on slides:
 * - Set shape text content
 * - Set shape properties (fill, line, etc.)
 * - Remove shapes from slides
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput, notFound, isOk } from "./result.js";
import type { Result } from "./types.js";
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

/**
 * Extracts the placeholder type from a path.
 */
function extractPlaceholderType(pptPath: string): string | null {
  const pattern = /\/placeholder\[([^\]]+)\]/i;
  const match = pptPath.match(pattern);
  return match ? match[1] : null;
}

// ============================================================================
// Text Operations
// ============================================================================

/**
 * Sets the text content of a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]" or "/slide[1]/placeholder[title]")
 * @param text - The new text content
 *
 * @example
 * const result = await setShapeText("/path/to/presentation.pptx", "/slide[1]/shape[1]", "Hello World");
 */
export async function setShapeText(
  filePath: string,
  pptPath: string,
  text: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShapeText requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<void>;
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    const placeholderType = extractPlaceholderType(pptPath);

    if (shapeIndex === null && placeholderType === null) {
      return invalidInput("Invalid shape path - must include shape[N] or placeholder[type]");
    }

    let updatedSlideXml: string;

    if (shapeIndex !== null) {
      updatedSlideXml = setTextInShapeByIndex(slideXml, shapeIndex, text);
    } else {
      updatedSlideXml = setTextInPlaceholderByType(slideXml, placeholderType!, text);
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
 * Sets text in a shape by its index.
 */
function setTextInShapeByIndex(slideXml: string, shapeIndex: number, text: string): string {
  // Find all p:sp elements
  const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);

  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    throw new Error(`Shape index ${shapeIndex} out of range`);
  }

  const targetShapeXml = matches[shapeIndex - 1];
  const updatedShapeXml = updateShapeText(targetShapeXml, text);

  return slideXml.replace(targetShapeXml, updatedShapeXml);
}

/**
 * Sets text in a placeholder by its type.
 */
function setTextInPlaceholderByType(slideXml: string, placeholderType: string, text: string): string {
  // Find placeholder with matching type
  const typePattern = new RegExp(`<p:sp[\\s\\S]*?<p:ph[^>]*type="${placeholderType}"[^>]*>[\\s\\S]*?</p:sp>`, "g");
  const matches = slideXml.match(typePattern);

  if (!matches || matches.length === 0) {
    throw new Error(`Placeholder type '${placeholderType}' not found`);
  }

  const targetShapeXml = matches[0];
  const updatedShapeXml = updateShapeText(targetShapeXml, text);

  return slideXml.replace(targetShapeXml, updatedShapeXml);
}

/**
 * Updates the text content of a shape's text body.
 */
function updateShapeText(shapeXml: string, text: string): string {
  // Create new text runs from the input text
  const paragraphs = text.split("\n");
  const newRuns = paragraphs.map((para, pIdx) =>
    `        <a:p>
          <a:r>
            <a:rPr lang="en-US"/>
            <a:t>${escapeXml(para)}</a:t>
          </a:r>
${pIdx < paragraphs.length - 1 ? "        </a:p>" : ""}`
  ).join("\n");

  // Handle case where text might be in a simple <a:t> element
  // or in structured <a:p> elements

  // First try to find and replace simple <a:t> content
  let result = shapeXml;

  // Find the text body - look for <p:txBody>
  const txBodyMatch = result.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
  if (txBodyMatch) {
    const oldContent = txBodyMatch[1];
    const newContent = createSimpleTextBody(text);
    result = result.replace(oldContent, newContent);
  } else {
    // No txBody found, need to create one
    // This is a simplified approach - actual implementation might need more context
    result = shapeXml;
  }

  return result;
}

/**
 * Creates a simple text body structure.
 */
function createSimpleTextBody(text: string): string {
  const paragraphs = text.split("\n");
  const runs = paragraphs.map(para =>
    `          <a:p>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>${escapeXml(para)}</a:t>
            </a:r>
          </a:p>`
  ).join("\n");

  return `
            <a:bodyPr/>
            <a:lstStyle/>
${runs}
          `;
}

/**
 * Gets text content from a shape.
 */
function getTextFromShape(shapeXml: string): string {
  const textRuns: string[] = [];

  // Match text runs
  for (const match of shapeXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) {
    textRuns.push(match[1]);
  }

  return textRuns.join("");
}

// ============================================================================
// Property Operations
// ============================================================================

/**
 * Shape property value types.
 */
export type ShapeProperty = "fill" | "line" | "lineWidth" | "fillColor" | "lineColor";

/**
 * Sets a property on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param property - The property to set
 * @param value - The new value
 *
 * @example
 * const result = await setShapeProperty("/path/to/presentation.pptx", "/slide[1]/shape[1]", "fillColor", "FF0000");
 */
export async function setShapeProperty(
  filePath: string,
  pptPath: string,
  property: ShapeProperty,
  value: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShapeProperty requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<void>;
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
    }

    const updatedSlideXml = setShapePropertyInSlide(slideXml, shapeIndex, property, value);

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
 * Sets a property on a shape in the slide XML.
 */
function setShapePropertyInSlide(
  slideXml: string,
  shapeIndex: number,
  property: ShapeProperty,
  value: string,
): string {
  // Find all p:sp elements
  const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);

  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    throw new Error(`Shape index ${shapeIndex} out of range`);
  }

  const targetShapeXml = matches[shapeIndex - 1];
  const updatedShapeXml = updateShapeProperty(targetShapeXml, property, value);

  return slideXml.replace(targetShapeXml, updatedShapeXml);
}

/**
 * Updates a property on a shape.
 */
function updateShapeProperty(shapeXml: string, property: ShapeProperty, value: string): string {
  let result = shapeXml;

  switch (property) {
    case "fillColor":
    case "fill":
      result = setShapeFillColor(shapeXml, value);
      break;
    case "lineColor":
    case "line":
      result = setShapeLineColor(shapeXml, value);
      break;
    case "lineWidth":
      result = setShapeLineWidth(shapeXml, value);
      break;
    default:
      throw new Error(`Unknown property: ${property}`);
  }

  return result;
}

/**
 * Sets the fill color of a shape.
 */
function setShapeFillColor(shapeXml: string, color: string): string {
  // Normalize color format (ensure it starts with # if needed)
  const normalizedColor = color.startsWith("#") ? color.slice(1) : color;

  // Find existing solid fill or create one
  const solidFillPattern = /<a:solidFill>([\s\S]*?)<\/a:solidFill>/;
  const hasSolidFill = solidFillPattern.test(shapeXml);

  if (hasSolidFill) {
    // Replace existing solid fill color
    return shapeXml.replace(
      /<a:solidFill>[\s\S]*?<\/a:solidFill>/,
      `<a:solidFill><a:srgbClr val="${normalizedColor}"/></a:solidFill>`
    );
  }

  // Need to add a solid fill - find spPr and insert after it
  const spPrMatch = shapeXml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);
  if (spPrMatch) {
    const newFill = `<a:solidFill><a:srgbClr val="${normalizedColor}"/></a:solidFill>`;
    // Insert after the opening spPr tag
    const newSpPr = shapeXml.replace(
      /<p:spPr>/,
      `<p:spPr>${newFill}`
    );
    return newSpPr;
  }

  // No spPr element - need to create one
  // This is a simplified approach
  return shapeXml;
}

/**
 * Sets the line color of a shape.
 */
function setShapeLineColor(shapeXml: string, color: string): string {
  const normalizedColor = color.startsWith("#") ? color.slice(1) : color;

  // Find existing ln element or create one
  const lnPattern = /<a:ln(?:[^>]*)>([\s\S]*?)<\/a:ln>/;
  const hasLn = lnPattern.test(shapeXml);

  if (hasLn) {
    // Check for existing solid fill in ln
    const hasSolidFill = /<a:solidFill>[\s\S]*?<\/a:solidFill>/.test(shapeXml);
    if (hasSolidFill) {
      return shapeXml.replace(
        /<a:solidFill>[\s\S]*?<\/a:solidFill>(?=[\s\S]*<\/a:ln>)/,
        `<a:solidFill><a:srgbClr val="${normalizedColor}"/></a:solidFill>`
      );
    }
    // Add solid fill to ln
    return shapeXml.replace(
      lnPattern,
      `<a:ln>$1<a:solidFill><a:srgbClr val="${normalizedColor}"/></a:solidFill></a:ln>`
    );
  }

  // No ln element - need to add one after spPr
  const spPrMatch = shapeXml.match(/<p:spPr>/);
  if (spPrMatch) {
    const newLn = `<a:ln><a:solidFill><a:srgbClr val="${normalizedColor}"/></a:solidFill></a:ln>`;
    return shapeXml.replace(/<p:spPr>/, `<p:spPr>${newLn}`);
  }

  return shapeXml;
}

/**
 * Sets the line width of a shape.
 */
function setShapeLineWidth(shapeXml: string, width: string): string {
  // Parse width value (in EMUs or points)
  const widthValue = parseFloat(width);
  const emuValue = Math.round(widthValue * 12700); // Convert points to EMUs

  // Find existing ln element
  const lnPattern = /<a:ln([^>]*)>/;
  const hasLn = lnPattern.test(shapeXml);

  if (hasLn) {
    // Add or update w attribute
    return shapeXml.replace(
      lnPattern,
      `<a:ln w="${emuValue}"$1>`
    );
  }

  // No ln element - need to add one after spPr
  const spPrMatch = shapeXml.match(/<p:spPr>/);
  if (spPrMatch) {
    const newLn = `<a:ln w="${emuValue}"></a:ln>`;
    return shapeXml.replace(/<p:spPr>/, `<p:spPr>${newLn}`);
  }

  return shapeXml;
}

// ============================================================================
// Add Shape Operations
// ============================================================================

/**
 * Known preset shape types.
 */
export type PresetShapeType =
  | "rectangle"
  | "roundRect"
  | "ellipse"
  | "diamond"
  | "triangle"
  | "rightTriangle"
  | "parallelogram"
  | "trapezoid"
  | "pentagon"
  | "hexagon"
  | "heptagon"
  | "octagon"
  | "star4"
  | "star5"
  | "star6"
  | "star8"
  | "star10"
  | "star12"
  | "star16"
  | "star24"
  | "star32"
  | "plus"
  | "circle"
  | "oval"
  | "arrowRight"
  | "arrowLeft"
  | "arrowUp"
  | "arrowDown"
  | "flowChartProcess"
  | "flowChartDecision"
  | "flowChartInputOutput"
  | "flowChartConnector";

/**
 * Position and size in EMUs.
 */
export interface ShapePosition {
  x: number;
  y: number;
}

export interface ShapeSize {
  width: number;
  height: number;
}

/**
 * Adds a new shape to a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param shapeType - The preset shape type (e.g., "rectangle", "ellipse")
 * @param position - The position (x, y) in EMUs
 * @param size - The size (width, height) in EMUs
 *
 * @example
 * const result = await addShape("/path/to/presentation.pptx", 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
 * if (result.ok) {
 *   console.log(result.data.path); // "/slide[1]/shape[5]"
 * }
 */
export async function addShape(
  filePath: string,
  slideIndex: number,
  shapeType: PresetShapeType,
  position: ShapePosition,
  size: ShapeSize,
): Promise<Result<{ path: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as unknown as Result<{ path: string }>;
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    // Count existing shapes to determine new shape index
    const shapePattern = /<p:sp[\s\S]*?<\/p:sp>/g;
    const existingShapes = slideXml.match(shapePattern) || [];
    const newShapeIndex = existingShapes.length + 1;

    // Generate unique shape ID
    const idPattern = /id="(\d+)"/g;
    let maxId = 0;
    let match;
    while ((match = idPattern.exec(slideXml)) !== null) {
      const id = parseInt(match[1], 10);
      if (id > maxId) maxId = id;
    }
    const newShapeId = maxId + 1;

    // Create new shape XML
    const newShapeXml = createShapeXml(newShapeId, newShapeIndex, shapeType, position, size);

    // Insert shape before </p:spTree>
    const updatedSlideXml = slideXml.replace(
      /<\/p:spTree>/,
      `${newShapeXml}\n  </p:spTree>`
    );

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
    return ok({ path: `/slide[${slideIndex}]/shape[${newShapeIndex}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Creates XML for a new shape.
 */
function createShapeXml(
  id: number,
  shapeIndex: number,
  shapeType: PresetShapeType,
  position: ShapePosition,
  size: ShapeSize,
): string {
  // Map shape type to OOXML preset geometry
  const presetMap: Record<PresetShapeType, string> = {
    rectangle: "rect",
    roundRect: "roundRect",
    ellipse: "ellipse",
    diamond: "diamond",
    triangle: "triangle",
    rightTriangle: "rightTriangle",
    parallelogram: "parallelogram",
    trapezoid: "trapezoid",
    pentagon: "pentagon",
    hexagon: "hexagon",
    heptagon: "heptagon",
    octagon: "octagon",
    star4: "star4",
    star5: "star5",
    star6: "star6",
    star8: "star8",
    star10: "star10",
    star12: "star12",
    star16: "star16",
    star24: "star24",
    star32: "star32",
    plus: "plus",
    circle: "ellipse",
    oval: "ellipse",
    arrowRight: "rightArrow",
    arrowLeft: "leftArrow",
    arrowUp: "upArrow",
    arrowDown: "downArrow",
    flowChartProcess: "flowChartProcess",
    flowChartDecision: "flowChartDecision",
    flowChartInputOutput: "flowChartInputOutput",
    flowChartConnector: "flowChartConnector",
  };

  const preset = presetMap[shapeType] || "rect";

  return `    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="${id}" name="Shape ${shapeIndex}"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="${position.x}" y="${position.y}"/>
          <a:ext cx="${size.width}" cy="${size.height}"/>
        </a:xfrm>
        <a:prstGeom prst="${preset}">
          <a:avLst/>
        </a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:endParaRPr/>
        </a:p>
      </p:txBody>
    </p:sp>`;
}

/**
 * Gets the type of a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 *
 * @example
 * const result = await getShapeType("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 * if (result.ok) {
 *   console.log(result.data.type); // "rectangle"
 * }
 */
export async function getShapeType(
  filePath: string,
  pptPath: string,
): Promise<Result<{ type: string }>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("getShapeType requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as unknown as Result<{ type: string }>;
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
    }

    // Find all p:sp elements
    const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
    const matches = slideXml.match(pattern);

    if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
      return notFound("Shape", pptPath);
    }

    const targetShapeXml = matches[shapeIndex - 1];

    // Extract shape type from preset geometry
    const presetMatch = targetShapeXml.match(/<a:prstGeom[^>]*prst="([^"]+)"/);
    if (presetMatch) {
      return ok({ type: presetMatch[1] });
    }

    // Check for placeholder type
    const phMatch = targetShapeXml.match(/<p:ph[^>]*type="([^"]+)"/);
    if (phMatch) {
      return ok({ type: `placeholder:${phMatch[1]}` });
    }

    return ok({ type: "shape" });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Fill Operations
// ============================================================================

/**
 * Fill type for shapes.
 */
export interface ShapeFill {
  type: "solid" | "gradient" | "picture" | "none";
  /** For solid fill: hex color (e.g., "FF0000") */
  color?: string;
  /** For gradient fill: gradient specification */
  gradient?: GradientFill;
  /** For picture fill: base64 encoded image data or relationship ID */
  picture?: string;
}

/**
 * Gradient fill specification.
 */
export interface GradientFill {
  type: "linear" | "radial";
  colors: Array<{ color: string; position: number }>; // position is 0-100000
  angle?: number; // for linear gradient, angle in degrees
}

/**
 * Sets the fill on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param fill - The fill specification
 *
 * @example
 * // Solid fill
 * const result = await setShapeFill("/path/to/presentation.pptx", "/slide[1]/shape[1]", { type: "solid", color: "FF0000" });
 * // Gradient fill
 * const result2 = await setShapeFill("/path/to/presentation.pptx", "/slide[1]/shape[1]", { type: "gradient", gradient: { type: "linear", colors: [{ color: "FF0000", position: 0 }, { color: "0000FF", position: 100000 }], angle: 90 } });
 */
export async function setShapeFill(
  filePath: string,
  pptPath: string,
  fill: ShapeFill,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShapeFill requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<void>;
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
    }

    const updatedSlideXml = setShapeFillInSlide(slideXml, shapeIndex, fill);

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
 * Sets fill on a shape in the slide XML.
 */
function setShapeFillInSlide(slideXml: string, shapeIndex: number, fill: ShapeFill): string {
  const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);

  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    throw new Error(`Shape index ${shapeIndex} out of range`);
  }

  const targetShapeXml = matches[shapeIndex - 1];
  const updatedShapeXml = updateShapeFill(targetShapeXml, fill);

  return slideXml.replace(targetShapeXml, updatedShapeXml);
}

/**
 * Updates the fill on a shape.
 */
function updateShapeFill(shapeXml: string, fill: ShapeFill): string {
  // Find or create spPr element
  let spPrMatch = shapeXml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);

  if (!spPrMatch) {
    // No spPr element, need to create one
    // Insert after nvSpPr
    const nvSpPrMatch = shapeXml.match(/<\/p:nvSpPr>/);
    if (nvSpPrMatch) {
      const newSpPr = `<p:spPr>${buildFillElement(fill)}</p:spPr>`;
      return shapeXml.replace(/<\/p:spPr>/, `${newSpPr}</p:spPr>`);
    }
    return shapeXml;
  }

  const existingSpPr = spPrMatch[0];
  const spPrContent = spPrMatch[1];

  // Remove existing fill elements (solidFill, gradFill, noFill, etc.)
  let newSpPrContent = spPrContent
    .replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/g, "")
    .replace(/<a:gradFill>[\s\S]*?<\/a:gradFill>/g, "")
    .replace(/<a:noFill>[\s\S]*?<\/a:noFill>/g, "")
    .replace(/<a:pattFill>[\s\S]*?<\/a:pattFill>/g, "");

  // Build new fill element
  const fillElement = buildFillElement(fill);

  // Insert fill at the beginning of spPr content
  newSpPrContent = fillElement + newSpPrContent;

  return shapeXml.replace(existingSpPr, `<p:spPr>${newSpPrContent}</p:spPr>`);
}

/**
 * Builds the XML for a fill element.
 */
function buildFillElement(fill: ShapeFill): string {
  switch (fill.type) {
    case "none":
      return "<a:noFill/>";

    case "solid":
      const color = (fill.color || "FFFFFF").replace("#", "");
      return `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;

    case "gradient":
      return buildGradientFillElement(fill.gradient!);

    case "picture":
      // Picture fill requires adding the image to the package and referencing it
      // This is a simplified version that just references by rel ID
      return `<a:blipFill><a:blip r:embed="${fill.picture}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;

    default:
      return "";
  }
}

/**
 * Builds the XML for a gradient fill element.
 */
function buildGradientFillElement(gradient: GradientFill): string {
  if (gradient.type === "linear") {
    const angle = gradient.angle || 0;
    const stops = gradient.colors
      .map(c => `<a:gradStop pos="${c.position}" type="rgb"><a:srgbClr val="${c.color.replace("#", "")}"/></a:gradStop>`)
      .join("");
    return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${angle * 60000}" scaled="1"/></a:gradFill>`;
  } else {
    // Radial gradient
    const stops = gradient.colors
      .map(c => `<a:gradStop pos="${c.position}" type="rgb"><a:srgbClr val="${c.color.replace("#", "")}"/></a:gradStop>`)
      .join("");
    return `<a:gradFill><a:gsLst>${stops}</a:gsLst><a:radialFill><a:srgbClr val="${gradient.colors[0]?.color.replace("#", "") || "FFFFFF"}"/></a:radialFill></a:gradFill>`;
  }
}

// ============================================================================
// Line Operations
// ============================================================================

/**
 * Line specification for shapes.
 */
export interface ShapeLine {
  /** Line color as hex (e.g., "FF0000") or "none" for no line */
  color?: string;
  /** Line width in points */
  width?: number;
  /** Line dash style */
  dash?: "solid" | "dot" | "dash" | "dashDot" | "lgDash" | "lgDashDot" | "lgDashDotDot";
}

/**
 * Sets the line properties on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param line - The line specification
 *
 * @example
 * const result = await setShapeLine("/path/to/presentation.pptx", "/slide[1]/shape[1]", { color: "FF0000", width: 2 });
 */
export async function setShapeLine(
  filePath: string,
  pptPath: string,
  line: ShapeLine,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShapeLine requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<void>;
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
    }

    const updatedSlideXml = setShapeLineInSlide(slideXml, shapeIndex, line);

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
 * Sets line on a shape in the slide XML.
 */
function setShapeLineInSlide(slideXml: string, shapeIndex: number, line: ShapeLine): string {
  const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);

  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    throw new Error(`Shape index ${shapeIndex} out of range`);
  }

  const targetShapeXml = matches[shapeIndex - 1];
  const updatedShapeXml = updateShapeLine(targetShapeXml, line);

  return slideXml.replace(targetShapeXml, updatedShapeXml);
}

/**
 * Updates the line on a shape.
 */
function updateShapeLine(shapeXml: string, line: ShapeLine): string {
  // Handle "none" line
  if (line.color === "none") {
    // Remove existing ln element or replace with noFill
    if (/<a:ln[^>]*>/.test(shapeXml)) {
      return shapeXml.replace(/<a:ln[\s\S]*?<\/a:ln>/, "<a:ln><a:noFill/></a:ln>");
    }
    return shapeXml;
  }

  // Find or create ln element
  let lnMatch = shapeXml.match(/<a:ln[^>]*>[\s\S]*?<\/a:ln>/);

  if (!lnMatch) {
    // No ln element, need to create one in spPr
    const spPrMatch = shapeXml.match(/<p:spPr>/);
    if (spPrMatch) {
      const newLn = buildLineElement(line);
      return shapeXml.replace(/<p:spPr>/, `<p:spPr>${newLn}`);
    }
    return shapeXml;
  }

  const existingLn = lnMatch[0];

  // Build updated line element
  const newLn = buildLineElement(line, existingLn);

  return shapeXml.replace(existingLn, newLn);
}

/**
 * Builds the XML for a line element.
 */
function buildLineElement(line: ShapeLine, existingLn?: string): string {
  const width = line.width ? Math.round(line.width * 12700) : 12700; // default 1pt
  const color = (line.color || "000000").replace("#", "");

  // Map dash styles to OOXML values
  const dashMap: Record<string, string> = {
    solid: "solid",
    dot: "dot",
    dash: "dash",
    dashDot: "dashDot",
    lgDash: "lgDash",
    lgDashDot: "lgDashDot",
    lgDashDotDot: "lgDashDotDot",
  };

  let dashAttr = "";
  if (line.dash && dashMap[line.dash]) {
    dashAttr = `<a:prstDash val="${dashMap[line.dash]}"/>`;
  }

  return `<a:ln w="${width}">${dashAttr}<a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln>`;
}

// ============================================================================
// Effect Operations
// ============================================================================

/**
 * Effect specification for shapes.
 */
export interface ShapeEffect {
  /** Shadow effect */
  shadow?: {
    type: "outer" | "inner" | "perspective";
    color?: string;
    blurRadius?: number;
    offsetX?: number;
    offsetY?: number;
    angle?: number;
    opacity?: number;
  };
  /** Glow effect */
  glow?: {
    color?: string;
    radius?: number;
  };
}

/**
 * Sets effects (shadow, glow, etc.) on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param effect - The effect specification
 *
 * @example
 * const result = await setShapeEffect("/path/to/presentation.pptx", "/slide[1]/shape[1]", {
 *   shadow: { type: "outer", color: "000000", blurRadius: 4, offsetX: 2, offsetY: 2, opacity: 0.5 }
 * });
 */
export async function setShapeEffect(
  filePath: string,
  pptPath: string,
  effect: ShapeEffect,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShapeEffect requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<void>;
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
    }

    const updatedSlideXml = setShapeEffectInSlide(slideXml, shapeIndex, effect);

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
 * Sets effect on a shape in the slide XML.
 */
function setShapeEffectInSlide(slideXml: string, shapeIndex: number, effect: ShapeEffect): string {
  const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);

  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    throw new Error(`Shape index ${shapeIndex} out of range`);
  }

  const targetShapeXml = matches[shapeIndex - 1];
  const updatedShapeXml = updateShapeEffect(targetShapeXml, effect);

  return slideXml.replace(targetShapeXml, updatedShapeXml);
}

/**
 * Updates the effect on a shape.
 */
function updateShapeEffect(shapeXml: string, effect: ShapeEffect): string {
  // Find or create spPr element
  let spPrMatch = shapeXml.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);

  if (!spPrMatch) {
    return shapeXml;
  }

  const existingSpPr = spPrMatch[0];
  let spPrContent = spPrMatch[1];

  // Remove existing effect elements
  spPrContent = spPrContent
    .replace(/<a:effectLst>[\s\S]*?<\/a:effectLst>/g, "")
    .replace(/<a:outerShdw>[\s\S]*?<\/a:outerShdw>/g, "")
    .replace(/<a:innerShdw>[\s\S]*?<\/a:innerShdw>/g, "")
    .replace(/<a:lstEffect>[\s\S]*?<\/a:lstEffect>/g, "");

  // Build new effect elements
  const effectElements: string[] = [];

  if (effect.shadow) {
    effectElements.push(buildShadowElement(effect.shadow));
  }

  if (effect.glow) {
    effectElements.push(buildGlowElement(effect.glow));
  }

  if (effectElements.length > 0) {
    const effectLst = `<a:effectLst>${effectElements.join("")}</a:effectLst>`;
    spPrContent += effectLst;
  }

  return shapeXml.replace(existingSpPr, `<p:spPr>${spPrContent}</p:spPr>`);
}

/**
 * Builds the XML for a shadow effect element.
 */
function buildShadowElement(shadow: NonNullable<ShapeEffect["shadow"]>): string {
  const color = (shadow.color || "000000").replace("#", "");
  const blurRadius = shadow.blurRadius ? shadow.blurRadius * 12700 : 4 * 12700;
  const offsetX = shadow.offsetX ? shadow.offsetX * 12700 : 2 * 12700;
  const offsetY = shadow.offsetY ? shadow.offsetY * 12700 : 2 * 12700;
  const angle = shadow.angle !== undefined ? shadow.angle * 60000 : 270 * 60000;
  const opacity = shadow.opacity !== undefined ? shadow.opacity * 100000 : 50000;

  return `<a:outerShdw blurRad="${blurRadius}" dist="${Math.sqrt(offsetX * offsetX + offsetY * offsetY)}" dir="${angle}" algn="tl" rotWithShape="0"><a:srgbClr val="${color}"><a:alpha val="${opacity}"/></a:srgbClr></a:outerShdw>`;
}

/**
 * Builds the XML for a glow effect element.
 */
function buildGlowElement(glow: NonNullable<ShapeEffect["glow"]>): string {
  const color = (glow.color || "000000").replace("#", "");
  const radius = glow.radius ? glow.radius * 12700 : 2 * 12700;

  return `<a:glow rad="${radius}"><a:srgbClr val="${color}"><a:alpha val="65000"/></a:srgbClr></a:glow>`;
}

// ============================================================================
// Remove Operations
// ============================================================================

/**
 * Removes a shape from a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 *
 * @example
 * const result = await removeShape("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 */
export async function removeShape(filePath: string, pptPath: string): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("removeShape requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<void>;
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeIndex = extractShapeIndex(pptPath);
    const placeholderType = extractPlaceholderType(pptPath);

    if (shapeIndex === null && placeholderType === null) {
      return invalidInput("Invalid shape path");
    }

    let updatedSlideXml: string;

    if (shapeIndex !== null) {
      updatedSlideXml = removeShapeByIndex(slideXml, shapeIndex);
    } else {
      updatedSlideXml = removePlaceholderByType(slideXml, placeholderType!);
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
 * Removes a shape by its index.
 */
function removeShapeByIndex(slideXml: string, shapeIndex: number): string {
  const pattern = /<p:sp[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);

  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    throw new Error(`Shape index ${shapeIndex} out of range`);
  }

  const targetShapeXml = matches[shapeIndex - 1];
  return slideXml.replace(targetShapeXml, "");
}

/**
 * Removes a placeholder by its type.
 */
function removePlaceholderByType(slideXml: string, placeholderType: string): string {
  const typePattern = new RegExp(`<p:sp[\\s\\S]*?<p:ph[^>]*type="${placeholderType}"[^>]*>[\\s\\S]*?</p:sp>`, "g");
  const matches = slideXml.match(typePattern);

  if (!matches || matches.length === 0) {
    throw new Error(`Placeholder type '${placeholderType}' not found`);
  }

  return slideXml.replace(matches[0], "");
}

// ============================================================================
// Slide Property Operations
// ============================================================================

/**
 * Sets a property on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param property - The property to set
 * @param value - The new value
 *
 * @example
 * const result = await setSlideProperty("/path/to/presentation.pptx", 1, "background", "FFCCCC");
 */
export async function setSlideProperty(
  filePath: string,
  slideIndex: number,
  property: string,
  value: string,
): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<void>;
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    let updatedSlideXml: string;

    switch (property) {
      case "background":
        updatedSlideXml = setSlideBackground(slideXml, value);
        break;
      default:
        return invalidInput(`Unknown slide property: ${property}`);
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
 * Sets the background of a slide.
 */
function setSlideBackground(slideXml: string, color: string): string {
  const normalizedColor = color.startsWith("#") ? color.slice(1) : color;

  // Find existing bg element or create one
  const bgPattern = /<p:bg>([\s\S]*?)<\/p:bg>/;
  const hasBg = bgPattern.test(slideXml);

  if (hasBg) {
    // Replace existing background
    return slideXml.replace(
      bgPattern,
      `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${normalizedColor}"/></a:solidFill></p:bgPr></p:bg>`
    );
  }

  // Find cSld and insert bg before it
  const cSldPattern = /(<p:cSld[^>]*>)/;
  if (cSldPattern.test(slideXml)) {
    const bgElement = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${normalizedColor}"/></a:solidFill></p:bgPr></p:bg>`;
    return slideXml.replace(cSldPattern, `${bgElement}\n  $1`);
  }

  return slideXml;
}
