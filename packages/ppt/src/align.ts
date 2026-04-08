/**
 * Shape alignment and distribution operations for @officekit/ppt.
 *
 * Provides functions to align and distribute shapes on slides:
 * - Align shapes: left, center, right, top, middle, bottom
 * - Align relative to slide
 * - Distribute shapes: horizontal, vertical
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput } from "./result.js";
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
 * Gets the slide size from presentation.xml.
 * Returns dimensions in EMUs.
 */
function getSlideSize(zip: Map<string, Buffer>): { width: number; height: number } {
  const presXml = requireEntry(zip, "ppt/presentation.xml");
  const sizeMatch = /<p:sldSz\b[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"[^>]*>/.exec(presXml);
  if (sizeMatch) {
    return {
      width: parseInt(sizeMatch[1], 10),
      height: parseInt(sizeMatch[2], 10),
    };
  }
  // Default to 16:9 aspect ratio (9144000 x 5143500 EMUs = 10 inches x 5.625 inches)
  return { width: 9144000, height: 5143500 };
}

/**
 * Represents the transform data for a shape.
 */
interface TransformData {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Extracts transform data from a shape's XML.
 */
function extractTransformFromShape(shapeXml: string): TransformData | null {
  // Match the a:xfrm element inside p:spPr
  const xfrmMatch = /<a:xfrm(?:[^>]*)>([\s\S]*?)<\/a:xfrm>/.exec(shapeXml);
  if (!xfrmMatch) {
    return null;
  }

  const xfrmContent = xfrmMatch[1];

  // Extract offset (position)
  const offMatch = /<a:off[^>]*x="([^"]*)"[^>]*y="([^"]*)"[^>]*>/.exec(xfrmContent);
  if (!offMatch) {
    return null;
  }
  const x = parseInt(offMatch[1], 10);
  const y = parseInt(offMatch[2], 10);

  // Extract extents (size)
  const extMatch = /<a:ext[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"[^>]*>/.exec(xfrmContent);
  if (!extMatch) {
    return null;
  }
  const width = parseInt(extMatch[1], 10);
  const height = parseInt(extMatch[2], 10);

  return { x, y, width, height };
}

/**
 * Gets all shape elements from a slide's XML.
 */
function getAllShapes(slideXml: string): string[] {
  const pattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  return slideXml.match(pattern) || [];
}

/**
 * Resolves target shapes from a comma-separated list of shape indices.
 * Returns shapes specified by "shape[N]" or "N" format.
 */
function resolveTargets(slideXml: string, targets: string | null): string[] {
  const allShapes = getAllShapes(slideXml);

  if (!targets || targets.trim() === "") {
    return allShapes;
  }

  const result: string[] = [];
  const tokens = targets.split(",").map(t => t.trim()).filter(t => t.length > 0);

  for (const token of tokens) {
    // Accept "shape[N]" or just "N"
    const match = token.match(/^shape\[(\d+)\]$|^(\d+)$/i);
    if (match) {
      const idx = parseInt(match[1] || match[2], 10) - 1; // Convert to 0-based
      if (idx >= 0 && idx < allShapes.length) {
        result.push(allShapes[idx]);
      }
    }
  }

  return result;
}

// ============================================================================
// Alignment Types
// ============================================================================

/**
 * Alignment mode for horizontal alignment.
 */
export type HorizontalAlignMode = "left" | "center" | "right";

/**
 * Alignment mode for vertical alignment.
 */
export type VerticalAlignMode = "top" | "middle" | "bottom";

/**
 * Alignment mode that can include slide- prefix for relative-to-slide alignment.
 */
export type AlignMode = HorizontalAlignMode | VerticalAlignMode | `slide-${HorizontalAlignMode}` | `slide-${VerticalAlignMode}`;

/**
 * Distribution mode for distributing shapes.
 */
export type DistributeMode = "horizontal" | "vertical";

// ============================================================================
// Alignment Operations
// ============================================================================

/**
 * Aligns shapes on a slide along one axis.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the slide (e.g., "/slide[1]")
 * @param align - Alignment mode: "left", "center", "right", "top", "middle", "bottom", or "slide-left", "slide-center", etc.
 * @param targets - Optional comma-separated list of shape indices (e.g., "shape[1],shape[2],shape[3]" or "1,2,3")
 *
 * @example
 * // Align shapes 1, 2, 3 to the left edge of their bounding box
 * const result = await alignShapes("/path/to/presentation.pptx", "/slide[1]", "left", "1,2,3");
 *
 * // Align all shapes to the slide center
 * const result2 = await alignShapes("/path/to/presentation.pptx", "/slide[1]", "slide-center");
 */
export async function alignShapes(
  filePath: string,
  pptPath: string,
  align: AlignMode,
  targets?: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("alignShapes requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    // Get slide dimensions
    const { width: slideWidth, height: slideHeight } = getSlideSize(zip);

    // Resolve target shapes
    const shapes = resolveTargets(slideXml, targets ?? null);
    if (shapes.length < 1) {
      return invalidInput("No shapes found to align");
    }

    // Extract transforms from shapes
    const transforms = shapes.map(s => extractTransformFromShape(s));

    // Normalize alignment mode
    const alignLower = align.toLowerCase();
    const relative = alignLower.startsWith("slide-");
    const mode = relative ? alignLower.slice(6) : alignLower;

    // Calculate bounding box or use slide bounds
    let refLeft: number;
    let refTop: number;
    let refRight: number;
    let refBottom: number;

    if (relative) {
      refLeft = 0;
      refTop = 0;
      refRight = slideWidth;
      refBottom = slideHeight;
    } else {
      // Bounding box of selected shapes
      const validTransforms = transforms.filter((t): t is TransformData => t !== null);
      if (validTransforms.length === 0) {
        return invalidInput("No valid shapes found with transform data");
      }

      refLeft = Math.min(...validTransforms.map(t => t.x));
      refTop = Math.min(...validTransforms.map(t => t.y));
      refRight = Math.max(...validTransforms.map(t => t.x + t.width));
      refBottom = Math.max(...validTransforms.map(t => t.y + t.height));
    }

    const refCenterX = (refLeft + refRight) / 2;
    const refCenterY = (refTop + refBottom) / 2;

    // Apply alignment
    const updatedShapes = shapes.map((shapeXml, i) => {
      const xfrm = transforms[i];
      if (!xfrm) return shapeXml;

      const { x, y, width, height } = xfrm;
      let newX = x;
      let newY = y;

      switch (mode) {
        case "left":
          newX = refLeft;
          break;
        case "center":
        case "hcenter":
          newX = refCenterX - width / 2;
          break;
        case "right":
          newX = refRight - width;
          break;
        case "top":
          newY = refTop;
          break;
        case "middle":
        case "vcenter":
          newY = refCenterY - height / 2;
          break;
        case "bottom":
          newY = refBottom - height;
          break;
        default:
          throw new Error(`Invalid align value: '${align}'. Valid: left, center, right, top, middle, bottom, slide-left, slide-center, slide-right, slide-top, slide-middle, slide-bottom`);
      }

      // Replace the transform values
      let updatedShape = shapeXml;

      // Update offset (position)
      updatedShape = updatedShape.replace(
        /<a:off[^>]*x="[^"]*"[^>]*y="[^"]*"[^>]*>/,
        `<a:off x="${Math.round(newX)}" y="${Math.round(newY)}"/>`
      );

      return updatedShape;
    });

    // Reconstruct the slide XML with updated shapes
    let updatedSlideXml = slideXml;
    for (let i = 0; i < shapes.length; i++) {
      updatedSlideXml = updatedSlideXml.replace(shapes[i], updatedShapes[i]);
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

// ============================================================================
// Distribution Operations
// ============================================================================

/**
 * Distributes shapes evenly on a slide.
 * Shapes are distributed so that the gaps between them are equal.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the slide (e.g., "/slide[1]")
 * @param distribute - Distribution mode: "horizontal" or "vertical"
 * @param targets - Optional comma-separated list of shape indices
 *
 * @example
 * // Distribute shapes 1, 2, 3 horizontally
 * const result = await distributeShapes("/path/to/presentation.pptx", "/slide[1]", "horizontal", "1,2,3");
 */
export async function distributeShapes(
  filePath: string,
  pptPath: string,
  distribute: DistributeMode,
  targets?: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("distributeShapes requires a slide path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data!;
    let slideXml = requireEntry(zip, slideEntry);

    // Resolve target shapes
    const shapes = resolveTargets(slideXml, targets ?? null);
    if (shapes.length < 3) {
      return invalidInput("distributeShapes requires at least 3 shapes");
    }

    // Extract transforms from shapes
    const transforms = shapes.map(s => extractTransformFromShape(s));

    const mode = distribute.toLowerCase();

    if (mode === "horizontal" || mode === "h" || mode === "horiz") {
      // Sort shapes by their left edge
      const indexed = shapes.map((s, i) => ({ shape: s, xfrm: transforms[i] }))
        .filter(p => p.xfrm !== null)
        .sort((a, b) => a.xfrm!.x - b.xfrm!.x);

      if (indexed.length < 3) {
        return invalidInput("Not enough valid shapes with positions for horizontal distribution");
      }

      const first = indexed[0].xfrm!;
      const last = indexed[indexed.length - 1].xfrm!;
      const totalWidth = indexed.reduce((sum, p) => sum + p.xfrm!.width, 0);
      const span = (last.x + last.width) - first.x;
      const gap = (span - totalWidth) / (indexed.length - 1);

      let cursor = first.x;
      const updatedShapes = indexed.map(p => {
        const newShape = p.shape.replace(
          /<a:off[^>]*x="[^"]*"[^>]*y="[^"]*"[^>]*>/,
          `<a:off x="${Math.round(cursor)}" y="${p.xfrm!.y}"/>`
        );
        cursor += p.xfrm!.width + gap;
        return newShape;
      });

      // Update shapes in order
      for (let i = 0; i < indexed.length; i++) {
        slideXml = slideXml.replace(indexed[i].shape, updatedShapes[i]);
      }
    } else if (mode === "vertical" || mode === "v" || mode === "vert") {
      // Sort shapes by their top edge
      const indexed = shapes.map((s, i) => ({ shape: s, xfrm: transforms[i] }))
        .filter(p => p.xfrm !== null)
        .sort((a, b) => a.xfrm!.y - b.xfrm!.y);

      if (indexed.length < 3) {
        return invalidInput("Not enough valid shapes with positions for vertical distribution");
      }

      const first = indexed[0].xfrm!;
      const last = indexed[indexed.length - 1].xfrm!;
      const totalHeight = indexed.reduce((sum, p) => sum + p.xfrm!.height, 0);
      const span = (last.y + last.height) - first.y;
      const gap = (span - totalHeight) / (indexed.length - 1);

      let cursor = first.y;
      const updatedShapes = indexed.map(p => {
        const newShape = p.shape.replace(
          /<a:off[^>]*x="[^"]*"[^>]*y="[^"]*"[^>]*>/,
          `<a:off x="${p.xfrm!.x}" y="${Math.round(cursor)}"/>`
        );
        cursor += p.xfrm!.height + gap;
        return newShape;
      });

      // Update shapes in order
      for (let i = 0; i < indexed.length; i++) {
        slideXml = slideXml.replace(indexed[i].shape, updatedShapes[i]);
      }
    } else {
      return invalidInput(`Invalid distribute value: '${distribute}'. Valid: horizontal, vertical`);
    }

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(slideXml, "utf8") });
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
