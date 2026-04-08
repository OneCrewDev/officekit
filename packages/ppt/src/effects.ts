/**
 * Shape effects operations for @officekit/ppt.
 *
 * Provides functions to apply visual effects to shapes:
 * - Shadow (outer shadow)
 * - Glow (outer glow)
 * - Blur (soft edge)
 * - 3D effects (rotation, depth, bevel, material, light rig)
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
 * Gets a shape from the slide by its index.
 */
function getShapeByIndex(slideXml: string, shapeIndex: number): string | null {
  const pattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  const matches = slideXml.match(pattern);
  if (!matches || shapeIndex < 1 || shapeIndex > matches.length) {
    return null;
  }
  return matches[shapeIndex - 1];
}

/**
 * Normalizes a color value to 6-digit hex.
 */
function normalizeColor(color: string): string {
  // Remove # prefix if present
  let normalized = color.startsWith("#") ? color.slice(1) : color;

  // Handle short hex (e.g., "FFF" -> "FFFFFF")
  if (normalized.length === 3) {
    normalized = normalized.split("").map(c => c + c).join("");
  }

  // Validate hex color
  if (!/^[0-9A-Fa-f]{6}$/.test(normalized)) {
    throw new Error(`Invalid color value: '${color}'. Expected 6-digit hex color.`);
  }

  return normalized;
}

// ============================================================================
// Effect List Helpers
// ============================================================================

/**
 * Schema order for CT_EffectList children:
 * blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge
 */
const EFFECT_LIST_ORDER = [
  "blur",
  "fillOverlay",
  "glow",
  "innerShdw",
  "outerShdw",
  "prstShdw",
  "reflection",
  "softEdge",
] as const;

/**
 * Gets or creates an EffectList element in the correct schema position.
 * Schema order within p:spPr: fill → ln → effectLst → scene3d → sp3d → extLst
 */
function ensureEffectList(spPrContent: string): { hasEffectList: boolean; content: string } {
  // Check if effectLst already exists
  const effectListMatch = /<a:effectLst>([\s\S]*?)<\/a:effectLst>/.exec(spPrContent);

  if (effectListMatch) {
    return { hasEffectList: true, content: spPrContent };
  }

  // Need to insert effectLst before scene3d, sp3d, or extLst
  let newContent = spPrContent;

  // Check for scene3d
  if (/<a:scene3d>/.test(newContent)) {
    newContent = newContent.replace(/<a:scene3d>/, "<a:effectLst/>\n  <a:scene3d>");
  }
  // Check for sp3d
  else if (/<a:sp3d>/.test(newContent)) {
    newContent = newContent.replace(/<a:sp3d>/, "<a:effectLst/>\n  <a:sp3d>");
  }
  // Check for extLst
  else if (/<a:extLst>/.test(newContent)) {
    newContent = newContent.replace(/<a:extLst>/, "<a:effectLst/>\n  <a:extLst>");
  }
  // No special elements, append before closing </p:spPr>
  else {
    newContent = newContent.replace(/<\/p:spPr>/, "<a:effectLst/>\n</p:spPr>");
  }

  return { hasEffectList: true, content: newContent };
}

// ============================================================================
// Shadow Effect
// ============================================================================

/**
 * Shadow effect options.
 */
export interface ShadowOptions {
  /** Shadow color as hex (e.g., "000000"). Defaults to black. */
  color?: string;
  /** Blur radius in points. Defaults to 4. */
  blur?: number;
  /** Direction in degrees (0-360). 0/360 = up, 90 = right, 180 = down, 270 = left. Defaults to 45. */
  angle?: number;
  /** Distance in points. Defaults to 3. */
  distance?: number;
  /** Opacity as percentage (0-100). Defaults to 40. */
  opacity?: number;
}

/**
 * Builds the XML for an outer shadow effect.
 */
function buildOuterShadowXml(options: ShadowOptions): string {
  const color = normalizeColor(options.color ?? "000000");
  const blurPt = options.blur ?? 4;
  const angleDeg = options.angle ?? 45;
  const distPt = options.distance ?? 3;
  const opacity = options.opacity ?? 40;

  // Convert to OOXML units
  const blurRadius = Math.round(blurPt * 12700);
  const distance = Math.round(distPt * 12700);
  const direction = Math.round(angleDeg * 60000);
  const opacityVal = Math.round(opacity * 1000);

  return `<a:outerShdw blurRad="${blurRadius}" dist="${distance}" dir="${direction}" algn="tl" rotWithShape="0">
    <a:srgbClr val="${color}"><a:alpha val="${opacityVal}"/></a:srgbClr>
  </a:outerShdw>`;
}

/**
 * Sets or removes shadow effect on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param options - Shadow options, or null/"none" to remove shadow
 *
 * @example
 * // Add a shadow to shape 1
 * const result = await setShapeShadow("/path/to/presentation.pptx", "/slide[1]/shape[1]", { color: "000000", blur: 6, angle: 45, distance: 4 });
 *
 * // Remove shadow from shape 1
 * const result2 = await setShapeShadow("/path/to/presentation.pptx", "/slide[1]/shape[1]", null);
 */
export async function setShapeShadow(
  filePath: string,
  pptPath: string,
  options: ShadowOptions | null | "none" | "false",
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShapeShadow requires a slide path");
    }

    // Extract shape index from path
    const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid shape path - must include shape[index]");
    }
    const shapeIndex = parseInt(shapeIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeXml = getShapeByIndex(slideXml, shapeIndex);
    if (!shapeXml) {
      return invalidInput(`Shape index ${shapeIndex} out of range`);
    }

    let updatedShapeXml: string;

    // Handle "none", "false", or null to remove shadow
    if (options === null || options === "none" || options === "false") {
      updatedShapeXml = removeEffectFromShape(shapeXml, "outerShdw");
    } else {
      // Apply shadow effect
      updatedShapeXml = applyEffectToShape(shapeXml, "outerShdw", buildOuterShadowXml(options));
    }

    const updatedSlideXml = slideXml.replace(shapeXml, updatedShapeXml);

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
// Glow Effect
// ============================================================================

/**
 * Glow effect options.
 */
export interface GlowOptions {
  /** Glow color as hex (e.g., "0070FF"). Defaults to blue. */
  color?: string;
  /** Radius in points. Defaults to 8. */
  radius?: number;
  /** Opacity as percentage (0-100). Defaults to 75. */
  opacity?: number;
}

/**
 * Builds the XML for a glow effect.
 */
function buildGlowXml(options: GlowOptions): string {
  const color = normalizeColor(options.color ?? "0070FF");
  const radiusPt = options.radius ?? 8;
  const opacity = options.opacity ?? 75;

  // Convert to OOXML units
  const radius = Math.round(radiusPt * 12700);
  const opacityVal = Math.round(opacity * 1000);

  return `<a:glow rad="${radius}">
    <a:srgbClr val="${color}"><a:alpha val="${opacityVal}"/></a:srgbClr>
  </a:glow>`;
}

/**
 * Sets or removes glow effect on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param options - Glow options, or null/"none" to remove glow
 *
 * @example
 * // Add a blue glow to shape 1
 * const result = await setShapeGlow("/path/to/presentation.pptx", "/slide[1]/shape[1]", { color: "00B0F0", radius: 10 });
 *
 * // Remove glow from shape 1
 * const result2 = await setShapeGlow("/path/to/presentation.pptx", "/slide[1]/shape[1]", null);
 */
export async function setShapeGlow(
  filePath: string,
  pptPath: string,
  options: GlowOptions | null | "none" | "false",
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShapeGlow requires a slide path");
    }

    // Extract shape index from path
    const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid shape path - must include shape[index]");
    }
    const shapeIndex = parseInt(shapeIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeXml = getShapeByIndex(slideXml, shapeIndex);
    if (!shapeXml) {
      return invalidInput(`Shape index ${shapeIndex} out of range`);
    }

    let updatedShapeXml: string;

    // Handle "none", "false", or null to remove glow
    if (options === null || options === "none" || options === "false") {
      updatedShapeXml = removeEffectFromShape(shapeXml, "glow");
    } else {
      // Apply glow effect
      updatedShapeXml = applyEffectToShape(shapeXml, "glow", buildGlowXml(options));
    }

    const updatedSlideXml = slideXml.replace(shapeXml, updatedShapeXml);

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
// Blur (Soft Edge) Effect
// ============================================================================

/**
 * Sets or removes blur (soft edge) effect on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param radius - Blur radius in points, or null/"none" to remove blur
 *
 * @example
 * // Add blur to shape 1
 * const result = await setShapeBlur("/path/to/presentation.pptx", "/slide[1]/shape[1]", 5);
 *
 * // Remove blur from shape 1
 * const result2 = await setShapeBlur("/path/to/presentation.pptx", "/slide[1]/shape[1]", null);
 */
export async function setShapeBlur(
  filePath: string,
  pptPath: string,
  radius: number | null | "none" | "false",
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShapeBlur requires a slide path");
    }

    // Extract shape index from path
    const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid shape path - must include shape[index]");
    }
    const shapeIndex = parseInt(shapeIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeXml = getShapeByIndex(slideXml, shapeIndex);
    if (!shapeXml) {
      return invalidInput(`Shape index ${shapeIndex} out of range`);
    }

    let updatedShapeXml: string;

    // Handle "none", "false", or null to remove blur
    if (radius === null || radius === "none" || radius === "false") {
      updatedShapeXml = removeEffectFromShape(shapeXml, "softEdge");
    } else {
      // Validate radius
      if (typeof radius !== "number" || radius < 0 || !isFinite(radius)) {
        return invalidInput(`Invalid blur radius: '${radius}'. Expected a non-negative number.`);
      }

      // Build soft edge XML
      const radiusEmu = Math.round(radius * 12700);
      const softEdgeXml = `<a:softEdge rad="${radiusEmu}"/>`;

      updatedShapeXml = applyEffectToShape(shapeXml, "softEdge", softEdgeXml);
    }

    const updatedSlideXml = slideXml.replace(shapeXml, updatedShapeXml);

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
// Effect Application Helper
// ============================================================================

/**
 * Applies an effect to a shape by inserting or replacing in the effect list.
 */
function applyEffectToShape(shapeXml: string, effectType: string, effectXml: string): string {
  // Find the spPr element
  const spPrMatch = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(shapeXml);
  if (!spPrMatch) {
    throw new Error("Shape does not have spPr element");
  }

  const spPrContent = spPrMatch[1];
  const spPrFull = spPrMatch[0];

  // Ensure effect list exists
  const { content: contentWithEffectList } = ensureEffectList(spPrContent);

  // Remove existing effect of the same type
  const effectPattern = new RegExp(`<a:${effectType}(?:[^>]*)>([\\s\\S]*?)<\\/a:${effectType}>|<a:${effectType}(?:[^>]*)\/>`, "g");
  let contentWithoutEffect = contentWithEffectList.replace(effectPattern, "");

  // Parse the effect XML to determine where to insert it
  const effectTagMatch = effectXml.match(/<a:(\w+)/);
  if (!effectTagMatch) {
    throw new Error("Invalid effect XML");
  }
  const effectName = effectTagMatch[1];

  // Find insertion position based on schema order
  const effectIndex = EFFECT_LIST_ORDER.indexOf(effectName as typeof EFFECT_LIST_ORDER[number]);

  // Find the first effect that comes after this one in schema order
  let insertBefore: string | null = null;
  for (const laterEffect of EFFECT_LIST_ORDER.slice(effectIndex + 1)) {
    const laterPattern = new RegExp(`<a:${laterEffect}`);
    if (laterPattern.test(contentWithoutEffect)) {
      insertBefore = laterEffect;
      break;
    }
  }

  let newContent: string;
  if (insertBefore) {
    // Insert before the later effect
    newContent = contentWithoutEffect.replace(
      new RegExp(`<a:${insertBefore}`),
      `${effectXml}\n    <a:${insertBefore}`
    );
  } else {
    // Append at the end of effectLst
    // First ensure we have proper closing
    if (/<a:effectLst\/>/.test(contentWithoutEffect)) {
      newContent = contentWithoutEffect.replace("<a:effectLst/>", `<a:effectLst>${effectXml}</a:effectLst>`);
    } else {
      newContent = contentWithoutEffect.replace("</a:effectLst>", `${effectXml}</a:effectLst>`);
    }
  }

  return shapeXml.replace(spPrFull, `<p:spPr>${newContent}</p:spPr>`);
}

/**
 * Removes an effect from a shape.
 */
function removeEffectFromShape(shapeXml: string, effectType: string): string {
  // Find the spPr element
  const spPrMatch = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(shapeXml);
  if (!spPrMatch) {
    return shapeXml;
  }

  const spPrContent = spPrMatch[1];
  const spPrFull = spPrMatch[0];

  // Remove the effect
  const effectPattern = new RegExp(`<a:${effectType}(?:[^>]*)>([\\s\\S]*?)<\\/a:${effectType}>|<a:${effectType}(?:[^>]*)\/>`, "g");
  let newContent = spPrContent.replace(effectPattern, "");

  // If effectList is now empty, remove it
  newContent = newContent.replace(/<a:effectLst><\/a:effectLst>|<a:effectLst\/>|<\/a:effectLst>/g, "");

  return shapeXml.replace(spPrFull, `<p:spPr>${newContent}</p:spPr>`);
}

// ============================================================================
// 3D Effects
// ============================================================================

/**
 * 3D rotation options.
 */
export interface Rotation3DOptions {
  /** Rotation around X axis in degrees */
  x?: number;
  /** Rotation around Y axis in degrees */
  y?: number;
  /** Rotation around Z axis in degrees */
  z?: number;
}

/**
 * Converts degrees to OOXML 60000ths-of-a-degree units.
 * Accepts negative values (e.g., -20 degrees becomes 340 degrees).
 */
function degreesTo60k(degrees: number): number {
  const val = Math.round(degrees * 60000);
  const full = 360 * 60000; // 21600000
  const normalized = val % full;
  return normalized < 0 ? normalized + full : normalized;
}

/**
 * Sets or removes 3D rotation on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param rotation - 3D rotation options, or null/"none" to remove
 *
 * @example
 * // Add 3D rotation
 * const result = await setShape3DRotation("/path/to/presentation.pptx", "/slide[1]/shape[1]", { x: 45, y: 30, z: 0 });
 *
 * // Remove 3D rotation
 * const result2 = await setShape3DRotation("/path/to/presentation.pptx", "/slide[1]/shape[1]", null);
 */
export async function setShape3DRotation(
  filePath: string,
  pptPath: string,
  rotation: Rotation3DOptions | null | "none",
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShape3DRotation requires a slide path");
    }

    // Extract shape index from path
    const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid shape path - must include shape[index]");
    }
    const shapeIndex = parseInt(shapeIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeXml = getShapeByIndex(slideXml, shapeIndex);
    if (!shapeXml) {
      return invalidInput(`Shape index ${shapeIndex} out of range`);
    }

    let updatedShapeXml: string;

    // Handle "none" or null to remove rotation
    if (rotation === null || rotation === "none") {
      updatedShapeXml = remove3DRotationFromShape(shapeXml);
    } else {
      // Validate rotation values
      const rotX = rotation.x ?? 0;
      const rotY = rotation.y ?? 0;
      const rotZ = rotation.z ?? 0;

      if (!isFinite(rotX) || !isFinite(rotY) || !isFinite(rotZ)) {
        return invalidInput("3D rotation values must be finite numbers");
      }

      updatedShapeXml = apply3DRotationToShape(shapeXml, rotX, rotY, rotZ);
    }

    const updatedSlideXml = slideXml.replace(shapeXml, updatedShapeXml);

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
 * Applies 3D rotation to a shape.
 */
function apply3DRotationToShape(shapeXml: string, rotX: number, rotY: number, rotZ: number): string {
  // Find the spPr element
  const spPrMatch = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(shapeXml);
  if (!spPrMatch) {
    throw new Error("Shape does not have spPr element");
  }

  const spPrContent = spPrMatch[1];
  const spPrFull = spPrMatch[0];

  // Ensure scene3d exists
  let newContent = spPrContent;

  if (!/<a:scene3d>/.test(newContent)) {
    // Insert scene3d before sp3d, extLst, or at end
    if (/<a:sp3d>/.test(newContent)) {
      newContent = newContent.replace(/<a:sp3d>/, `<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>\n    <a:sp3d>`);
    } else if (/<a:extLst>/.test(newContent)) {
      newContent = newContent.replace(/<a:extLst>/, `<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>\n    <a:extLst>`);
    } else {
      newContent = newContent.replace(/<\/p:spPr>/, `<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>\n  </p:spPr>`);
    }
  }

  // Now update the camera rotation
  const lat60k = degreesTo60k(rotX);
  const lon60k = degreesTo60k(rotY);
  const rev60k = degreesTo60k(rotZ);

  // Check if rotation already exists
  if (/<a:rot/.test(newContent)) {
    // Replace existing rotation
    newContent = newContent.replace(
      /<a:rot[^>]*>[\s\S]*?<\/a:rot>/,
      `<a:rot lat="${lat60k}" lon="${lon60k}" rev="${rev60k}"/>`
    );
  } else {
    // Insert rotation after camera
    newContent = newContent.replace(
      /<a:camera[^>]*\/>/,
      `<a:camera prst="orthographicFront"/><a:rot lat="${lat60k}" lon="${lon60k}" rev="${rev60k}"/>`
    );
  }

  return shapeXml.replace(spPrFull, `<p:spPr>${newContent}</p:spPr>`);
}

/**
 * Removes 3D rotation from a shape.
 */
function remove3DRotationFromShape(shapeXml: string): string {
  // Find the spPr element
  const spPrMatch = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(shapeXml);
  if (!spPrMatch) {
    return shapeXml;
  }

  const spPrContent = spPrMatch[1];
  const spPrFull = spPrMatch[0];

  let newContent = spPrContent;

  // Remove rotation element
  newContent = newContent.replace(/<a:rot[^>]*>[\s\S]*?<\/a:rot>/g, "");

  // If scene3d is now empty (just has camera and lightRig with no rotation), remove it
  newContent = newContent.replace(/<a:scene3d>\s*<a:camera[^>]*\/>\s*<a:lightRig[^>]*\/>\s*<\/a:scene3d>/g, "");

  return shapeXml.replace(spPrFull, `<p:spPr>${newContent}</p:spPr>`);
}

// ============================================================================
// 3D Depth (Extrusion)
// ============================================================================

/**
 * Sets or removes 3D extrusion depth on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param depth - Depth in points, or null/"none" to remove
 *
 * @example
 * // Add 3D depth
 * const result = await setShape3DDepth("/path/to/presentation.pptx", "/slide[1]/shape[1]", 50);
 *
 * // Remove 3D depth
 * const result2 = await setShape3DDepth("/path/to/presentation.pptx", "/slide[1]/shape[1]", null);
 */
export async function setShape3DDepth(
  filePath: string,
  pptPath: string,
  depth: number | null | "none",
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShape3DDepth requires a slide path");
    }

    // Extract shape index from path
    const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid shape path - must include shape[index]");
    }
    const shapeIndex = parseInt(shapeIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeXml = getShapeByIndex(slideXml, shapeIndex);
    if (!shapeXml) {
      return invalidInput(`Shape index ${shapeIndex} out of range`);
    }

    let updatedShapeXml: string;

    // Handle "none" or null to remove depth
    if (depth === null || depth === "none" || depth === 0) {
      updatedShapeXml = remove3DDepthFromShape(shapeXml);
    } else {
      // Validate depth
      if (typeof depth !== "number" || !isFinite(depth)) {
        return invalidInput(`Invalid depth value: '${depth}'. Expected a finite number.`);
      }

      updatedShapeXml = apply3DDepthToShape(shapeXml, depth);
    }

    const updatedSlideXml = slideXml.replace(shapeXml, updatedShapeXml);

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
 * Applies 3D depth to a shape.
 */
function apply3DDepthToShape(shapeXml: string, depth: number): string {
  // Find the spPr element
  const spPrMatch = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(shapeXml);
  if (!spPrMatch) {
    throw new Error("Shape does not have spPr element");
  }

  const spPrContent = spPrMatch[1];
  const spPrFull = spPrMatch[0];

  let newContent = spPrContent;

  // Convert depth to EMUs (1 point = 12700 EMUs)
  const depthEmu = Math.round(depth * 12700);

  // Ensure sp3d exists
  if (!/<a:sp3d>/.test(newContent)) {
    // Insert sp3d at the end before closing
    newContent = newContent.replace(/<\/p:spPr>/, `<a:sp3d extrusionOk="1"><a:dimDepth>${depthEmu}</a:dimDepth></a:sp3d>\n  </p:spPr>`);
  } else {
    // Update or add dimDepth
    if (/<a:dimDepth>/.test(newContent)) {
      newContent = newContent.replace(/<a:dimDepth>[\s\S]*?<\/a:dimDepth>/, `<a:dimDepth>${depthEmu}</a:dimDepth>`);
    } else {
      // Insert dimDepth in sp3d
      newContent = newContent.replace(/<a:sp3d>/, `<a:sp3d extrusionOk="1"><a:dimDepth>${depthEmu}</a:dimDepth>`);
    }
  }

  return shapeXml.replace(spPrFull, `<p:spPr>${newContent}</p:spPr>`);
}

/**
 * Removes 3D depth from a shape.
 */
function remove3DDepthFromShape(shapeXml: string): string {
  // Find the spPr element
  const spPrMatch = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(shapeXml);
  if (!spPrMatch) {
    return shapeXml;
  }

  const spPrContent = spPrMatch[1];
  const spPrFull = spPrMatch[0];

  let newContent = spPrContent;

  // Remove dimDepth
  newContent = newContent.replace(/<a:dimDepth>[\s\S]*?<\/a:dimDepth>/g, "");

  // If sp3d is now empty, remove it
  newContent = newContent.replace(/<a:sp3d>\s*<\/a:sp3d>/g, "");

  return shapeXml.replace(spPrFull, `<p:spPr>${newContent}</p:spPr>`);
}

// ============================================================================
// 3D Bevel
// ============================================================================

/**
 * 3D bevel preset types.
 */
export type BevelPresetType =
  | "circle"
  | "relaxedInset"
  | "cross"
  | "coolSlant"
  | "angle"
  | "softRound"
  | "convex"
  | "slope"
  | "divot"
  | "riblet"
  | "hardEdge"
  | "artDeco";

/**
 * Bevel options.
 */
export interface BevelOptions {
  /** Bevel preset type */
  preset?: BevelPresetType;
  /** Width in points */
  width?: number;
  /** Height in points */
  height?: number;
}

/**
 * Applies bevel to a shape (top or bottom).
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param position - "top" or "bottom"
 * @param options - Bevel options, or null/"none" to remove
 *
 * @example
 * // Add top bevel
 * const result = await setShapeBevel("/path/to/presentation.pptx", "/slide[1]/shape[1]", "top", { preset: "circle", width: 6, height: 6 });
 *
 * // Remove top bevel
 * const result2 = await setShapeBevel("/path/to/presentation.pptx", "/slide[1]/shape[1]", "top", null);
 */
export async function setShapeBevel(
  filePath: string,
  pptPath: string,
  position: "top" | "bottom",
  options: BevelOptions | null | "none",
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShapeBevel requires a slide path");
    }

    // Extract shape index from path
    const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid shape path - must include shape[index]");
    }
    const shapeIndex = parseInt(shapeIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeXml = getShapeByIndex(slideXml, shapeIndex);
    if (!shapeXml) {
      return invalidInput(`Shape index ${shapeIndex} out of range`);
    }

    let updatedShapeXml: string;

    // Handle "none" or null to remove bevel
    if (options === null || options === "none") {
      updatedShapeXml = removeBevelFromShape(shapeXml, position);
    } else {
      const preset = options.preset ?? "circle";
      const widthPt = options.width ?? 6;
      const heightPt = options.height ?? 6;

      // Validate preset
      const validPresets: BevelPresetType[] = [
        "circle", "relaxedInset", "cross", "coolSlant", "angle", "softRound",
        "convex", "slope", "divot", "riblet", "hardEdge", "artDeco"
      ];
      if (!validPresets.includes(preset)) {
        return invalidInput(`Invalid bevel preset: '${preset}'. Valid: ${validPresets.join(", ")}`);
      }

      updatedShapeXml = applyBevelToShape(shapeXml, position, preset, widthPt, heightPt);
    }

    const updatedSlideXml = slideXml.replace(shapeXml, updatedShapeXml);

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
 * Applies bevel to a shape.
 */
function applyBevelToShape(
  shapeXml: string,
  position: "top" | "bottom",
  preset: string,
  widthPt: number,
  heightPt: number
): string {
  // Find the spPr element
  const spPrMatch = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(shapeXml);
  if (!spPrMatch) {
    throw new Error("Shape does not have spPr element");
  }

  const spPrContent = spPrMatch[1];
  const spPrFull = spPrMatch[0];

  let newContent = spPrContent;

  // Convert to EMUs
  const widthEmu = Math.round(widthPt * 12700);
  const heightEmu = Math.round(heightPt * 12700);

  // Ensure sp3d exists
  if (!/<a:sp3d>/.test(newContent)) {
    newContent = newContent.replace(/<\/p:spPr>/, `<a:sp3d extrusionOk="1"><a:bevelTop w="${widthEmu}" h="${heightEmu}" prst="${preset}"/></a:sp3d>\n  </p:spPr>`);
  } else {
    // Remove existing bevel of this position
    const bevelTag = position === "top" ? "bevelTop" : "bevelBottom";
    newContent = newContent.replace(new RegExp(`<a:${bevelTag}[^>]*>[\\s\\S]*?<\\/a:${bevelTag}>`, "g"), "");

    // Insert new bevel
    const bevelXml = `<a:bevelTop w="${widthEmu}" h="${heightEmu}" prst="${preset}"/>`;
    if (position === "top") {
      newContent = newContent.replace(/<a:sp3d>/, `<a:sp3d extrusionOk="1">${bevelXml}`);
    } else {
      // Insert before closing </a:sp3d>
      newContent = newContent.replace(/<\/a:sp3d>/, `${bevelXml}</a:sp3d>`);
    }
  }

  return shapeXml.replace(spPrFull, `<p:spPr>${newContent}</p:spPr>`);
}

/**
 * Removes bevel from a shape.
 */
function removeBevelFromShape(shapeXml: string, position: "top" | "bottom"): string {
  // Find the spPr element
  const spPrMatch = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(shapeXml);
  if (!spPrMatch) {
    return shapeXml;
  }

  const spPrContent = spPrMatch[1];
  const spPrFull = spPrMatch[0];

  let newContent = spPrContent;

  // Remove bevel of this position
  const bevelTag = position === "top" ? "bevelTop" : "bevelBottom";
  newContent = newContent.replace(new RegExp(`<a:${bevelTag}[^>]*>[\\s\\S]*?<\\/a:${bevelTag}>`, "g"), "");

  // If sp3d is now empty (no extrusion, no bevels), remove it
  newContent = newContent.replace(/<a:sp3d>\s*<\/a:sp3d>/g, "");

  return shapeXml.replace(spPrFull, `<p:spPr>${newContent}</p:spPr>`);
}

// ============================================================================
// 3D Material
// ============================================================================

/**
 * 3D material preset types.
 */
export type MaterialPresetType =
  | "warmMatte"
  | "plastic"
  | "metal"
  | "darkEdge"
  | "softEdge"
  | "flat"
  | "wire"
  | "powder"
  | "translucentPowder"
  | "clear"
  | "softMetal"
  | "matte";

/**
 * Sets 3D material on a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param material - Material preset
 *
 * @example
 * const result = await setShape3DMaterial("/path/to/presentation.pptx", "/slide[1]/shape[1]", "metal");
 */
export async function setShape3DMaterial(
  filePath: string,
  pptPath: string,
  material: MaterialPresetType,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setShape3DMaterial requires a slide path");
    }

    // Extract shape index from path
    const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid shape path - must include shape[index]");
    }
    const shapeIndex = parseInt(shapeIndexMatch[1], 10);

    // Validate material preset
    const validMaterials: MaterialPresetType[] = [
      "warmMatte", "plastic", "metal", "darkEdge", "softEdge", "flat",
      "wire", "powder", "translucentPowder", "clear", "softMetal", "matte"
    ];
    if (!validMaterials.includes(material)) {
      return invalidInput(`Invalid material: '${material}'. Valid: ${validMaterials.join(", ")}`);
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data!;
    const slideXml = requireEntry(zip, slideEntry);

    const shapeXml = getShapeByIndex(slideXml, shapeIndex);
    if (!shapeXml) {
      return invalidInput(`Shape index ${shapeIndex} out of range`);
    }

    const updatedShapeXml = apply3DMaterialToShape(shapeXml, material);
    const updatedSlideXml = slideXml.replace(shapeXml, updatedShapeXml);

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
 * Applies 3D material to a shape.
 */
function apply3DMaterialToShape(shapeXml: string, material: string): string {
  // Find the spPr element
  const spPrMatch = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(shapeXml);
  if (!spPrMatch) {
    throw new Error("Shape does not have spPr element");
  }

  const spPrContent = spPrMatch[1];
  const spPrFull = spPrMatch[0];

  let newContent = spPrContent;

  // Ensure sp3d exists
  if (!/<a:sp3d>/.test(newContent)) {
    newContent = newContent.replace(/<\/p:spPr>/, `<a:sp3d prstMaterial="${material}"/></a:sp3d>\n  </p:spPr>`);
  } else {
    // Update or add prstMaterial
    if (/<a:prstMaterial/.test(newContent)) {
      newContent = newContent.replace(/<a:prstMaterial[^>]*\/>/, `<a:prstMaterial val="${material}"/>`);
    } else {
      // Insert prstMaterial in sp3d
      newContent = newContent.replace(/<a:sp3d>/, `<a:sp3d prstMaterial="${material}">`);
    }
  }

  return shapeXml.replace(spPrFull, `<p:spPr>${newContent}</p:spPr>`);
}
