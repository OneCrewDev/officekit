/**
 * 3D Model operations for @officekit/ppt.
 *
 * Provides functions to manage 3D models (.glb files) in PowerPoint
 * presentations:
 * - List 3D models on a slide
 * - Add 3D models to slides
 * - Remove 3D models from slides
 * - Update 3D model rotation
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, notFound, invalidInput } from "./result.js";
import type { Result } from "./types.js";
import { getSlideIndex } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Represents a 3D model on a slide.
 */
export interface Model3DItem {
  /** Path to the 3D model (e.g., "/slide[1]/model3d[1]") */
  path: string;
  /** Name of the 3D model */
  name?: string;
  /** Position X in EMUs */
  x?: number;
  /** Position Y in EMUs */
  y?: number;
  /** Width in EMUs */
  width?: number;
  /** Height in EMUs */
  height?: number;
  /** Depth in EMUs */
  depth?: number;
  /** Rotation X in degrees */
  rotX?: number;
  /** Rotation Y in degrees */
  rotY?: number;
  /** Rotation Z in degrees */
  rotZ?: number;
  /** Content type (MIME type) */
  contentType?: string;
  /** Size in bytes */
  size?: number;
}

/**
 * Position for placing a 3D model on a slide.
 */
export interface Model3DPosition {
  /** X position in EMUs */
  x: number;
  /** Y position in EMUs */
  y: number;
  /** Width in EMUs */
  width: number;
  /** Height in EMUs */
  height: number;
  /** Depth in EMUs (optional, for 3D box sizing) */
  depth?: number;
}

/**
 * Rotation for a 3D model.
 */
export interface Model3DRotation {
  /** Rotation around X axis in degrees */
  x: number;
  /** Rotation around Y axis in degrees */
  y: number;
  /** Rotation around Z axis in degrees */
  z: number;
}

/**
 * Data for a 3D model file.
 */
export interface Model3DData {
  /** Path to the .glb file */
  filePath: string;
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
 * Generates a unique shape ID.
 */
function generateShapeId(existingIds: number[]): number {
  let id = 1;
  while (existingIds.includes(id)) {
    id++;
  }
  return id;
}

/**
 * Gets the slide size from presentation.xml.
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
 * Extracts 3D models from a slide's XML.
 */
function extractModels3DFromSlide(
  slideXml: string,
  slideIndex: number,
  slideRels: Array<{ id: string; target: string; type?: string }>,
  slideEntry: string,
  zip: Map<string, Buffer>
): Model3DItem[] {
  const models: Model3DItem[] = [];

  // Find all shape elements that contain scene3d (3D models)
  // 3D models are represented as p:sp elements with a:scene3d child
  const spPattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  const spMatches = slideXml.match(spPattern) || [];

  let modelIdx = 0;
  for (const spXml of spMatches) {
    // Check if this shape contains a 3D model (via cNvPr with a model-related name or via scene3d)
    const hasScene3d = /<a:scene3d>/.test(spXml);
    const hasModelRel = /<a:model3d>/.test(spXml);

    if (!hasScene3d && !hasModelRel) {
      continue;
    }

    modelIdx++;

    // Extract name
    const nameMatch = /<p:cNvPr[^>]*name="([^"]*)"[^>]*>/.exec(spXml);
    const name = nameMatch ? nameMatch[1] : `3D Model ${modelIdx}`;

    // Extract position and size from a:xfrm
    const xfrmMatch = /<a:xfrm(?:[^>]*)>([\s\S]*?)<\/a:xfrm>/.exec(spXml);
    let x: number | undefined;
    let y: number | undefined;
    let width: number | undefined;
    let height: number | undefined;
    let depth: number | undefined;

    if (xfrmMatch) {
      const xfrmContent = xfrmMatch[1];
      const offMatch = /<a:off[^>]*x="([^"]*)"[^>]*y="([^"]*)"[^>]*>/.exec(xfrmContent);
      if (offMatch) {
        x = parseInt(offMatch[1], 10);
        y = parseInt(offMatch[2], 10);
      }
      const extMatch = /<a:ext[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"[^>]*>/.exec(xfrmContent);
      if (extMatch) {
        width = parseInt(extMatch[1], 10);
        height = parseInt(extMatch[2], 10);
      }
      // Depth is typically in a:scene3d a:sp3d elements
      const depthMatch = /<a:sp3d[^>]*dimDepth="([^"]*)"[^>]*>/.exec(spXml);
      if (depthMatch) {
        depth = parseInt(depthMatch[1], 10);
      }
    }

    // Extract rotation from a:scene3d
    let rotX: number | undefined;
    let rotY: number | undefined;
    let rotZ: number | undefined;

    const rotMatch = /<a:rot(?:[^>]*)>([\s\S]*?)<\/a:rot>/.exec(spXml);
    if (rotMatch) {
      const rotContent = rotMatch[1];
      // Rotation is stored as fixed-point angle value (60,000 = 1 degree)
      const rotAttrMatch = /val="([^"]*)"/.exec(rotContent);
      if (rotAttrMatch) {
        const rotValue = parseInt(rotAttrMatch[1], 10);
        if (!isNaN(rotValue)) {
          rotZ = rotValue / 60000; // Convert from fixed-point to degrees
        }
      }
    }

    // Extract path to the 3D model
    const modelRelMatch = /<a:model3d\b[^>]*r:embed="([^"]*)"[^>]*>/.exec(spXml);
    let contentType = "model/gltf-binary";
    let size: number | undefined;

    if (modelRelMatch) {
      const relId = modelRelMatch[1];
      const rel = slideRels.find(r => r.id === relId);
      if (rel) {
        const slideDir = path.posix.dirname(slideEntry);
        const modelPath = normalizeZipPath(slideDir, rel.target);
        const modelData = zip.get(modelPath);
        if (modelData) {
          size = modelData.length;
        }
      }
    }

    models.push({
      path: `/slide[${slideIndex}]/model3d[${modelIdx}]`,
      name,
      x,
      y,
      width,
      height,
      depth,
      rotX,
      rotY,
      rotZ,
      contentType,
      size,
    });
  }

  return models;
}

// ============================================================================
// 3D Model Operations
// ============================================================================

/**
 * Lists all 3D models on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 *
 * @example
 * const result = await get3DModels("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.models);
 * }
 */
export async function get3DModels(
  filePath: string,
  slideIndex: number
): Promise<Result<{ models: Model3DItem[]; total: number }>> {
  try {
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

    const models = extractModels3DFromSlide(slideXml, slideIndex, relationships, slideEntry, zip);

    return ok({ models, total: models.length });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Adds a 3D model to a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param modelPath - Path to the .glb file
 * @param position - Position and size in EMUs
 * @param rotation - Optional rotation in degrees
 *
 * @example
 * const result = await add3DModel(
 *   "/path/to/presentation.pptx",
 *   1,
 *   "/path/to/model.glb",
 *   { x: 1000000, y: 1000000, width: 3000000, height: 3000000 },
 *   { x: 0, y: 0, z: 45 }
 * );
 */
export async function add3DModel(
  filePath: string,
  slideIndex: number,
  modelPath: string,
  position: Model3DPosition,
  rotation?: Model3DRotation
): Promise<Result<{ path: string }>> {
  try {
    // Validate model path
    if (!modelPath.toLowerCase().endsWith(".glb")) {
      return invalidInput("3D model file must have .glb extension");
    }

    // Read the GLB file
    const glbData = await readFile(modelPath);

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

    // Calculate position and size
    const modelWidth = position.width ?? 3000000; // Default ~3 inches
    const modelHeight = position.height ?? 3000000;
    const modelDepth = position.depth ?? 3000000;
    const modelX = position.x ?? 1000000;
    const modelY = position.y ?? 1000000;

    // Generate unique IDs
    const existingRelIds = relationships.map(r => r.id);
    const newRelId = generateRelId(existingRelIds);

    // Find existing shape IDs to generate unique shape ID
    const existingShapeIds: number[] = [];
    for (const match of slideXml.matchAll(/<p:cNvPr[^>]*id="(\d+)"[^>]*>/g)) {
      existingShapeIds.push(parseInt(match[1], 10));
    }
    const newShapeId = generateShapeId(existingShapeIds);

    // Count existing 3D models for naming
    const modelCount = (slideXml.match(/<a:scene3d>/g) || []).length;
    const modelName = `3D Model ${modelCount + 1}`;

    // Generate unique filename for the GLB in the archive
    const glbFilename = `model_${Date.now()}.glb`;
    const modelEntry = `ppt/media/${glbFilename}`;

    // Create the model part relationship
    const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships.model" Target="../media/${glbFilename}"/>`;

    // Build the rotation XML if provided
    let rotXml = "";
    if (rotation) {
      // Convert degrees to fixed-point (60000 units = 1 degree)
      const rotZFixed = Math.round((rotation.z ?? 0) * 60000);
      const rotXFixed = Math.round((rotation.x ?? 0) * 60000);
      const rotYFixed = Math.round((rotation.y ?? 0) * 60000);
      // Combined rotation in Z is typical for 3D model rotation
      rotXml = `<a:rot val="${rotZFixed}"/>`;
    }

    // Build the 3D model XML using a:scene3d structure
    // The 3D model is embedded in a p:sp (shape) element
    const modelXml = `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="${newShapeId}" name="${modelName}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${modelX}" y="${modelY}"/>
      <a:ext cx="${modelWidth}" cy="${modelHeight}"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
    <a:scene3d>
      <a:model3d r:embed="${newRelId}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
      <a:sp3d>
        <a: extrusionOk="1"/>
        <a:dimDepth>${modelDepth}</a:dimDepth>
        <a:bevelTop ang="2546309" h="19050" prst="circle"/>
        <a:bevelBottom ang="2546309" h="19050" prst="circle"/>
      </a:sp3d>
    </a:scene3d>
    ${rotXml}
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p/>
  </p:txBody>
</p:sp>`;

    // Insert model into slide XML before closing </p:spTree>
    const updatedSlideXml = slideXml.replace(
      "</p:spTree>",
      `${modelXml}</p:spTree>`
    );

    // Update relationships XML
    const updatedRelsXml = relsXml.replace(
      "</Relationships>",
      `${newRelEntry}</Relationships>`
    );

    // Ensure media directory exists and add GLB file
    const newEntries: Array<{ name: string; data: Buffer }> = [];

    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else if (name === relsEntry) {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
      } else if (name === modelEntry) {
        // Don't duplicate - skip since we're replacing
      } else {
        newEntries.push({ name, data });
      }
    }

    // Add the GLB data
    newEntries.push({ name: modelEntry, data: glbData });

    await writeFile(filePath, createStoredZip(newEntries));

    return ok({ path: `/slide[${slideIndex}]/model3d[${modelCount + 1}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Removes a 3D model from a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param modelPath - Path to the 3D model (e.g., "/slide[1]/model3d[1]")
 *
 * @example
 * const result = await remove3DModel("/path/to/presentation.pptx", "/slide[1]/model3d[1]");
 */
export async function remove3DModel(
  filePath: string,
  modelPath: string
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(modelPath);
    if (slideIndex === null) {
      return invalidInput("Invalid model path - must include slide index");
    }

    // Extract model index from path
    const modelIndexMatch = modelPath.match(/\/model3d\[(\d+)\]/i);
    if (!modelIndexMatch) {
      return invalidInput("Invalid model path - must include model3d[index]");
    }
    const modelIndex = parseInt(modelIndexMatch[1], 10);

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

    // Find all shapes with scene3d to locate the target model
    const spPattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
    const spMatches = slideXml.match(spPattern) || [];

    // Filter to only those with scene3d
    const modelShapes = spMatches.filter(sp => /<a:scene3d>/.test(sp));

    if (modelIndex < 1 || modelIndex > modelShapes.length) {
      return notFound("3D Model", String(modelIndex));
    }

    const targetSpXml = modelShapes[modelIndex - 1];

    // Extract the relationship ID for the model
    const modelRelMatch = /<a:model3d\b[^>]*r:embed="([^"]*)"[^>]*>/.exec(targetSpXml);
    const relId = modelRelMatch ? modelRelMatch[1] : null;

    // Remove the shape from slide XML
    const updatedSlideXml = slideXml.replace(targetSpXml, "");

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
 * Updates the rotation of a 3D model.
 *
 * @param filePath - Path to the PPTX file
 * @param modelPath - Path to the 3D model (e.g., "/slide[1]/model3d[1]")
 * @param rotation - New rotation values in degrees
 *
 * @example
 * const result = await set3DModelRotation(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/model3d[1]",
 *   { x: 45, y: 30, z: 60 }
 * );
 */
export async function set3DModelRotation(
  filePath: string,
  modelPath: string,
  rotation: Model3DRotation
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(modelPath);
    if (slideIndex === null) {
      return invalidInput("Invalid model path - must include slide index");
    }

    // Extract model index from path
    const modelIndexMatch = modelPath.match(/\/model3d\[(\d+)\]/i);
    if (!modelIndexMatch) {
      return invalidInput("Invalid model path - must include model3d[index]");
    }
    const modelIndex = parseInt(modelIndexMatch[1], 10);

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

    // Find all shapes with scene3d to locate the target model
    const spPattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
    const spMatches = slideXml.match(spPattern) || [];

    // Filter to only those with scene3d
    const modelShapes = spMatches.filter(sp => /<a:scene3d>/.test(sp));

    if (modelIndex < 1 || modelIndex > modelShapes.length) {
      return notFound("3D Model", String(modelIndex));
    }

    const targetSpXml = modelShapes[modelIndex - 1];

    // Build new rotation XML
    // Convert degrees to fixed-point (60000 units = 1 degree)
    const rotZFixed = Math.round((rotation.z ?? 0) * 60000);
    const newRotXml = `<a:rot val="${rotZFixed}"/>`;

    // Check if there's an existing a:rot element
    let updatedSpXml: string;
    if (/<a:rot/.test(targetSpXml)) {
      // Replace existing rotation
      updatedSpXml = targetSpXml.replace(/<a:rot[^>]*>[\s\S]*?<\/a:rot>/, newRotXml);
    } else {
      // Add new rotation after </a:scene3d> but before </p:spPr>
      updatedSpXml = targetSpXml.replace(
        /<\/p:spPr>/,
        `  ${newRotXml}\n  </p:spPr>`
      );
    }

    // Replace the old shape with the updated one
    const updatedSlideXml = slideXml.replace(targetSpXml, updatedSpXml);

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
