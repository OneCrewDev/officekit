/**
 * Shape grouping operations for @officekit/ppt.
 *
 * Provides functions to group and ungroup shapes:
 * - Group multiple shapes together
 * - Ungroup shapes (dissolve the group, keep individual shapes)
 * - List shapes in a group
 * - Add/remove shapes from existing groups
 *
 * Groups use `<p:grpSp>` elements with children stored in `<p:grpSpPr>`.
 * Individual shapes inside a group use relative positioning.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput, notFound } from "./result.js";
import type { Result, GroupModel, ShapeModel } from "./types.js";
import { getSlideIndex, parsePath } from "./path.js";

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
 * Generates unique shape ID.
 */
function generateUniqueId(slideXml: string): number {
  const idPattern = /id="(\d+)"/g;
  let maxId = 0;
  let match;
  while ((match = idPattern.exec(slideXml)) !== null) {
    const id = parseInt(match[1], 10);
    if (id > maxId) maxId = id;
  }
  return maxId + 1;
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
 * Extracts the group index from a path.
 */
function extractGroupIndex(pptPath: string): number | null {
  const pattern = /\/group\[(\d+)\]/i;
  const match = pptPath.match(pattern);
  return match ? parseInt(match[1], 10) : null;
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
 * Gets all shapes from slide XML with their positions and sizes.
 */
interface ShapeInfo {
  index: number;
  xml: string;
  x: number;
  y: number;
  width: number;
  height: number;
  type: string;
}

function getShapesFromSlide(slideXml: string): ShapeInfo[] {
  const shapes: ShapeInfo[] = [];

  // Match all shape types: p:sp (shapes), p:pic (pictures), p:cxnSp (connectors), p:grpSp (groups)
  const shapePattern = /<(p:sp|p:pic|p:cxnSp|p:grpSp)([\s\S]*?)<\/\1>/g;
  let shapeIndex = 0;

  for (const match of slideXml.matchAll(shapePattern)) {
    const fullMatch = match[0];
    const shapeType = match[1];
    const content = match[2];
    shapeIndex++;

    // Extract position and size from xfrm
    const xfrmMatch = content.match(/<a:xfrm>[\s\S]*?<a:off x="(\d+)" y="(\d+)"\/>[\s\S]*?<a:ext cx="(\d+)" cy="(\d+)"\/><\/a:xfrm>/);
    let x = 0, y = 0, width = 0, height = 0;

    if (xfrmMatch) {
      x = parseInt(xfrmMatch[1], 10);
      y = parseInt(xfrmMatch[2], 10);
      width = parseInt(xfrmMatch[3], 10);
      height = parseInt(xfrmMatch[4], 10);
    }

    shapes.push({
      index: shapeIndex,
      xml: fullMatch,
      x,
      y,
      width,
      height,
      type: shapeType,
    });
  }

  return shapes;
}

/**
 * Calculates the bounding box of a set of shapes.
 */
function calculateBoundingBox(shapes: ShapeInfo[]): { x: number; y: number; width: number; height: number } {
  if (shapes.length === 0) {
    return { x: 0, y: 0, width: 0, height: 0 };
  }

  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;

  for (const shape of shapes) {
    minX = Math.min(minX, shape.x);
    minY = Math.min(minY, shape.y);
    maxX = Math.max(maxX, shape.x + shape.width);
    maxY = Math.max(maxY, shape.y + shape.height);
  }

  return {
    x: minX,
    y: minY,
    width: maxX - minX,
    height: maxY - minY,
  };
}

// ============================================================================
// Group Operations
// ============================================================================

/**
 * Groups multiple shapes together.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - Path to where the group should be created (slide or existing group)
 * @param shapePaths - Array of paths to shapes to group
 *
 * @example
 * // Group shapes together on slide 1
 * const result = await groupShapes(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]",
 *   ["/slide[1]/shape[1]", "/slide[1]/shape[2]"]
 * );
 * // Returns: { path: "/slide[1]/group[1]" }
 */
export async function groupShapes(
  filePath: string,
  pptPath: string,
  shapePaths: string[],
): Promise<Result<{ path: string }>> {
  try {
    // Validate inputs
    if (!shapePaths || shapePaths.length < 2) {
      return invalidInput("At least 2 shapes are required to create a group");
    }

    // Parse the parent path to determine slide index
    const parsedParent = parsePath(pptPath);
    if (!parsedParent.ok) {
      return err(parsedParent.error?.code ?? "invalid_path", parsedParent.error?.message ?? "Failed to parse path");
    }

    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("groupShapes requires a slide path");
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

    // Get all shapes and find the ones to group
    const allShapes = getShapesFromSlide(slideXml);
    const shapesToGroup: ShapeInfo[] = [];
    const shapesToGroupIndices: number[] = [];

    for (const shapePath of shapePaths) {
      const shapeIndex = extractShapeIndex(shapePath);
      if (shapeIndex === null) {
        return invalidInput(`Invalid shape path: ${shapePath}`);
      }

      // Validate that shape is on the same slide as the parent path
      const shapeSlideIndex = getSlideIndex(shapePath);
      if (shapeSlideIndex !== slideIndex) {
        return invalidInput(`Shape ${shapePath} is not on the same slide as the group destination`);
      }

      const shape = allShapes.find(s => s.index === shapeIndex && s.type === "p:sp");
      if (!shape) {
        return notFound("Shape", shapePath);
      }

      // Check if shape is already in a group (not at top level)
      const shapeXml = shape.xml;
      const parentMatch = slideXml.match(/<p:grpSp>[\s\S]*?<p:sp>[\s\S]*?<\/p:grpSp>/);
      if (parentMatch && parentMatch[0].includes(shapeXml)) {
        return invalidInput(`Shape ${shapePath} is already in a group`);
      }

      shapesToGroup.push(shape);
      shapesToGroupIndices.push(shapeIndex);
    }

    // Sort by index descending to remove from end first (avoid index shifting)
    const sortedIndices = [...shapesToGroupIndices].sort((a, b) => b - a);

    // Calculate bounding box for the group
    const bbox = calculateBoundingBox(shapesToGroup);

    // Generate unique ID for the group
    const newGroupId = generateUniqueId(slideXml);

    // Count existing groups
    const groupPattern = /<p:grpSp[\s\S]*?<\/p:grpSp>/g;
    const existingGroups = slideXml.match(groupPattern) || [];
    const newGroupIndex = existingGroups.length + 1;

    // Create relative positions for shapes within the group
    // Children positions are relative to the group's bounding box origin
    const childShapesXml = shapesToGroup.map((shape, i) => {
      const relX = shape.x - bbox.x;
      const relY = shape.y - bbox.y;

      // For shapes inside a group, we need to adjust the xfrm
      let adjustedShapeXml = shape.xml;

      // Replace the position in the shape's xfrm
      adjustedShapeXml = adjustedShapeXml.replace(
        /<a:xfrm>[\s\S]*?<\/a:xfrm>/,
        `<a:xfrm>
          <a:off x="${relX}" y="${relY}"/>
          <a:ext cx="${shape.width}" cy="${shape.height}"/>
        </a:xfrm>`
      );

      return adjustedShapeXml;
    }).join("\n");

    // Create group XML
    const groupXml = `    <p:grpSp>
      <p:nvGrpSpPr>
        <p:cNvPr id="${newGroupId}" name="Group ${newGroupIndex}"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="${bbox.x}" y="${bbox.y}"/>
          <a:ext cx="${bbox.width}" cy="${bbox.height}"/>
        </a:xfrm>
        <a:grpFill>
          <a:solidFill/>
        </a:grpFill>
      </p:grpSpPr>
${childShapesXml}
    </p:grpSp>`;

    // First, remove the shapes from their original positions
    let updatedSlideXml = slideXml;
    for (const idx of sortedIndices) {
      // Find and remove the shape at this index
      let count = 0;
      const shapeToRemove = /<p:sp[\s\S]*?<\/p:sp>/g;
      updatedSlideXml = updatedSlideXml.replace(shapeToRemove, (match) => {
        count++;
        if (count === idx) {
          return ""; // Remove the shape
        }
        return match;
      });
    }

    // Insert the group before </p:spTree>
    updatedSlideXml = updatedSlideXml.replace(
      /<\/p:spTree>/,
      `${groupXml}\n  </p:spTree>`
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
    return ok({ path: `/slide[${slideIndex}]/group[${newGroupIndex}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Ungroups shapes (dissolves the group, keeps the individual shapes).
 *
 * @param filePath - Path to the PPTX file
 * @param groupPath - Path to the group to ungroup (e.g., "/slide[1]/group[1]")
 *
 * @example
 * const result = await ungroupShapes("/path/to/presentation.pptx", "/slide[1]/group[1]");
 */
export async function ungroupShapes(
  filePath: string,
  pptPath: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("ungroupShapes requires a slide path");
    }

    const groupIndex = extractGroupIndex(pptPath);
    if (groupIndex === null) {
      return invalidInput("Invalid group path - must include group[N]");
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

    // Find the group
    const groupPattern = /<p:grpSp[\s\S]*?<\/p:grpSp>/g;
    const groups = slideXml.match(groupPattern) || [];

    if (groupIndex < 1 || groupIndex > groups.length) {
      return notFound("Group", pptPath);
    }

    const targetGroup = groups[groupIndex - 1];

    // Get the group's bounding box position
    const bboxMatch = targetGroup.match(/<a:xfrm>[\s\S]*?<a:off x="(\d+)" y="(\d+)"\/>[\s\S]*?<\/a:xfrm>/);
    const groupX = bboxMatch ? parseInt(bboxMatch[1], 10) : 0;
    const groupY = bboxMatch ? parseInt(bboxMatch[2], 10) : 0;

    // Extract the child shapes and adjust their positions back to absolute
    // Match all child shape elements within the group
    const childShapes: string[] = [];
    const childPattern = /<(p:sp|p:pic|p:grpSp)([\s\S]*?)<\/\1>/g;
    let childMatch;

    // We need to find shapes that are direct children of the group
    // The group content is between <p:grpSpPr>...</p:grpSpPr> and </p:grpSp>
    const grpSpPrEnd = targetGroup.indexOf("</p:grpSpPr>");
    const groupContent = targetGroup.slice(grpSpPrEnd + "</p:grpSpPr>".length);

    while ((childMatch = childPattern.exec(groupContent)) !== null) {
      let childXml = childMatch[0];

      // Adjust position: add group offset to make it absolute
      childXml = childXml.replace(
        /<a:xfrm>[\s\S]*?<\/a:xfrm>/,
        (xfrmMatch) => {
          const childPosMatch = xfrmMatch.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
          const childExtMatch = xfrmMatch.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);

          if (childPosMatch) {
            const relX = parseInt(childPosMatch[1], 10);
            const relY = parseInt(childPosMatch[1], 10);
            const absX = relX + groupX;
            const absY = relY + groupY;
            const cx = childExtMatch ? childExtMatch[1] : "0";
            const cy = childExtMatch ? childExtMatch[2] : "0";

            return `<a:xfrm>
          <a:off x="${absX}" y="${absY}"/>
          <a:ext cx="${cx}" cy="${cy}"/>
        </a:xfrm>`;
          }
          return xfrmMatch;
        }
      );

      childShapes.push(childXml);
    }

    // Replace the group with the ungrouped shapes
    let updatedSlideXml = slideXml.replace(targetGroup, childShapes.join("\n"));

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
 * Lists the shapes/children in a group.
 *
 * @param filePath - Path to the PPTX file
 * @param groupPath - Path to the group (e.g., "/slide[1]/group[1]")
 *
 * @example
 * const result = await getGroupChildren("/path/to/presentation.pptx", "/slide[1]/group[1]");
 * if (result.ok) {
 *   console.log(result.data.children);
 *   // [{ path: "/slide[1]/group[1]/shape[1]", type: "shape" }, ...]
 * }
 */
export async function getGroupChildren(
  filePath: string,
  pptPath: string,
): Promise<Result<{ children: ShapeModel[] }>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("getGroupChildren requires a slide path");
    }

    const groupIndex = extractGroupIndex(pptPath);
    if (groupIndex === null) {
      return invalidInput("Invalid group path - must include group[N]");
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

    // Find the group
    const groupPattern = /<p:grpSp[\s\S]*?<\/p:grpSp>/g;
    const groups = slideXml.match(groupPattern) || [];

    if (groupIndex < 1 || groupIndex > groups.length) {
      return notFound("Group", pptPath);
    }

    const targetGroup = groups[groupIndex - 1];

    // Extract children from the group
    const children: ShapeModel[] = [];

    // Match all child shape elements within the group
    const childPattern = /<(p:sp|p:pic|p:grpSp)([\s\S]*?)<\/\1>/g;
    let childMatch;
    let childIndex = 0;

    // Get content after </p:grpSpPr>
    const grpSpPrEnd = targetGroup.indexOf("</p:grpSpPr>");
    const groupContent = targetGroup.slice(grpSpPrEnd + "</p:grpSpPr>".length);

    while ((childMatch = childPattern.exec(groupContent)) !== null) {
      childIndex++;
      const childType = childMatch[1];
      const childContent = childMatch[2];

      // Extract common properties
      const nameMatch = childContent.match(/<p:cNvPr[^>]*name="([^"]+)"/);
      const name = nameMatch ? nameMatch[1] : undefined;

      // Determine element type
      let elementType = "shape";
      if (childType === "p:pic") elementType = "picture";
      if (childType === "p:grpSp") elementType = "group";

      children.push({
        path: `${pptPath}/${elementType}[${childIndex}]`,
        name,
        type: elementType,
      });
    }

    return ok({ children });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Adds a shape to an existing group.
 *
 * @param filePath - Path to the PPTX file
 * @param groupPath - Path to the group (e.g., "/slide[1]/group[1]")
 * @param shapePath - Path to the shape to add (e.g., "/slide[1]/shape[3]")
 *
 * @example
 * const result = await addShapeToGroup(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/group[1]",
 *   "/slide[1]/shape[3]"
 * );
 */
export async function addShapeToGroup(
  filePath: string,
  pptPath: string,
  shapePath: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("addShapeToGroup requires a group path");
    }

    const groupIndex = extractGroupIndex(pptPath);
    if (groupIndex === null) {
      return invalidInput("Invalid group path - must include group[N]");
    }

    const shapeIndex = extractShapeIndex(shapePath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path");
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

    // Get all shapes
    const allShapes = getShapesFromSlide(slideXml);

    // Find the shape to add
    const shapeToAdd = allShapes.find(s => s.index === shapeIndex && s.type === "p:sp");
    if (!shapeToAdd) {
      return notFound("Shape", shapePath);
    }

    // Find the group
    const groupPattern = /<p:grpSp[\s\S]*?<\/p:grpSp>/g;
    const groups = slideXml.match(groupPattern) || [];

    if (groupIndex < 1 || groupIndex > groups.length) {
      return notFound("Group", pptPath);
    }

    const targetGroup = groups[groupIndex - 1];

    // Get the group's bounding box position
    const bboxMatch = targetGroup.match(/<a:xfrm>[\s\S]*?<a:off x="(\d+)" y="(\d+)"\/>[\s\S]*?<\/a:xfrm>/);
    const groupX = bboxMatch ? parseInt(bboxMatch[1], 10) : 0;
    const groupY = bboxMatch ? parseInt(bboxMatch[2], 10) : 0;

    // Calculate relative position for the shape within the group
    const relX = shapeToAdd.x - groupX;
    const relY = shapeToAdd.y - groupY;

    // Adjust the shape's position
    let adjustedShapeXml = shapeToAdd.xml.replace(
      /<a:xfrm>[\s\S]*?<\/a:xfrm>/,
      `<a:xfrm>
          <a:off x="${relX}" y="${relY}"/>
          <a:ext cx="${shapeToAdd.width}" cy="${shapeToAdd.height}"/>
        </a:xfrm>`
    );

    // Remove the shape from its original position
    let updatedSlideXml = slideXml;
    let count = 0;
    const shapeToRemove = /<p:sp[\s\S]*?<\/p:sp>/g;
    updatedSlideXml = updatedSlideXml.replace(shapeToRemove, (match) => {
      count++;
      if (count === shapeIndex) {
        return ""; // Remove the shape
      }
      return match;
    });

    // Add the shape to the group (before </p:grpSp>)
    updatedSlideXml = updatedSlideXml.replace(
      targetGroup,
      targetGroup.replace(/<\/p:grpSp>/, `${adjustedShapeXml}\n    </p:grpSp>`)
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
    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Removes a shape from a group (shape becomes a standalone shape).
 *
 * @param filePath - Path to the PPTX file
 * @param groupPath - Path to the group (e.g., "/slide[1]/group[1]")
 * @param shapePath - Path to the shape to remove (e.g., "/slide[1]/group[1]/shape[1]")
 *
 * @example
 * const result = await removeShapeFromGroup(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/group[1]",
 *   "/slide[1]/group[1]/shape[1]"
 * );
 */
export async function removeShapeFromGroup(
  filePath: string,
  pptPath: string,
  shapePath: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("removeShapeFromGroup requires a group path");
    }

    const groupIndex = extractGroupIndex(pptPath);
    if (groupIndex === null) {
      return invalidInput("Invalid group path - must include group[N]");
    }

    const shapeIndex = extractShapeIndex(shapePath);
    if (shapeIndex === null) {
      return invalidInput("Invalid shape path - must include shape[N]");
    }

    // Validate that the shape path is a child of the group
    // Child paths look like "/slide[N]/group[M]/shape[K]" not "/slide[N]/shape[K]"
    if (!shapePath.startsWith(pptPath + "/")) {
      return invalidInput(`Shape ${shapePath} is not a child of group ${pptPath}`);
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

    // Find the group
    const groupPattern = /<p:grpSp[\s\S]*?<\/p:grpSp>/g;
    const groups = slideXml.match(groupPattern) || [];

    if (groupIndex < 1 || groupIndex > groups.length) {
      return notFound("Group", pptPath);
    }

    const targetGroup = groups[groupIndex - 1];

    // Get the group's bounding box position
    const bboxMatch = targetGroup.match(/<a:xfrm>[\s\S]*?<a:off x="(\d+)" y="(\d+)"\/>[\s\S]*?<\/a:xfrm>/);
    const groupX = bboxMatch ? parseInt(bboxMatch[1], 10) : 0;
    const groupY = bboxMatch ? parseInt(bboxMatch[2], 10) : 0;

    // Get content after </p:grpSpPr>
    const grpSpPrEnd = targetGroup.indexOf("</p:grpSpPr>");
    const groupContent = targetGroup.slice(grpSpPrEnd + "</p:grpSpPr>".length);

    // Find all child shapes in the group
    const childPattern = /<(p:sp|p:pic|p:grpSp)([\s\S]*?)<\/\1>/g;
    const childMatches = [...groupContent.matchAll(childPattern)];

    if (shapeIndex < 1 || shapeIndex > childMatches.length) {
      return notFound("Shape in group", shapePath);
    }

    const targetChild = childMatches[shapeIndex - 1][0];

    // Adjust the shape's position to be absolute (add group offset)
    let adjustedChildXml = targetChild.replace(
      /<a:xfrm>[\s\S]*?<\/a:xfrm>/,
      (xfrmMatch) => {
        const childPosMatch = xfrmMatch.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        const childExtMatch = xfrmMatch.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);

        if (childPosMatch) {
          const relX = parseInt(childPosMatch[1], 10);
          const relY = parseInt(childPosMatch[2], 10);
          const absX = relX + groupX;
          const absY = relY + groupY;
          const cx = childExtMatch ? childExtMatch[1] : "0";
          const cy = childExtMatch ? childExtMatch[2] : "0";

          return `<a:xfrm>
          <a:off x="${absX}" y="${absY}"/>
          <a:ext cx="${cx}" cy="${cy}"/>
        </a:xfrm>`;
        }
        return xfrmMatch;
      }
    );

    // Remove the shape from the group and add it to the slide
    let updatedSlideXml = slideXml.replace(targetGroup, targetGroup.replace(targetChild, ""));
    updatedSlideXml = updatedSlideXml.replace(/<\/p:spTree>/, `${adjustedChildXml}\n  </p:spTree>`);

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
 * Gets information about a group.
 *
 * @param filePath - Path to the PPTX file
 * @param groupPath - Path to the group (e.g., "/slide[1]/group[1]")
 *
 * @example
 * const result = await getGroup("/path/to/presentation.pptx", "/slide[1]/group[1]");
 * if (result.ok) {
 *   console.log(result.data.group);
 *   // { path: "/slide[1]/group[1]", name: "Group 1", childCount: 3 }
 * }
 */
export async function getGroup(
  filePath: string,
  pptPath: string,
): Promise<Result<{ group: GroupModel }>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("getGroup requires a slide path");
    }

    const groupIndex = extractGroupIndex(pptPath);
    if (groupIndex === null) {
      return invalidInput("Invalid group path - must include group[N]");
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

    // Find the group
    const groupPattern = /<p:grpSp[\s\S]*?<\/p:grpSp>/g;
    const groups = slideXml.match(groupPattern) || [];

    if (groupIndex < 1 || groupIndex > groups.length) {
      return notFound("Group", pptPath);
    }

    const targetGroup = groups[groupIndex - 1];

    // Extract group properties
    const nameMatch = targetGroup.match(/<p:cNvPr[^>]*name="([^"]+)"/);
    const name = nameMatch ? nameMatch[1] : undefined;

    // Count children by counting occurrences of child element opening tags after grpSpPr
    // Children appear between </p:grpSpPr> and </p:grpSp>
    const grpSpPrEnd = targetGroup.indexOf("</p:grpSpPr>");
    const afterGrpSpPr = targetGroup.slice(grpSpPrEnd + "</p:grpSpPr>".length);
    const grpSpStart = afterGrpSpPr.lastIndexOf("</p:grpSp>");
    const childrenContent = afterGrpSpPr.slice(0, grpSpStart);

    // Count child elements by matching their opening tags
    const childOpenPattern = /<(p:sp|p:pic|p:grpSp)[\s>]/g;
    const childMatches = childrenContent.match(childOpenPattern) || [];
    const childCount = childMatches.length;

    return ok({
      group: {
        path: pptPath,
        name,
        childCount,
      },
    });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}
