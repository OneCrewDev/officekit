/**
 * Media operations for @officekit/ppt.
 *
 * Provides functions to manage media (images, audio, video) in PowerPoint
 * presentations:
 * - List media on a slide
 * - Add pictures to slides
 * - Remove media from slides
 * - Replace pictures with new images
 * - Get binary data of media
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
 * Represents a media item on a slide.
 */
export interface MediaItem {
  /** Path to the media item (e.g., "/slide[1]/picture[1]") */
  path: string;
  /** Media type: "picture", "video", or "audio" */
  type: "picture" | "video" | "audio";
  /** Name of the media item */
  name: string;
  /** Alternative text description */
  alt?: string;
  /** Position X in EMUs */
  x?: number;
  /** Position Y in EMUs */
  y?: number;
  /** Width in EMUs */
  width?: number;
  /** Height in EMUs */
  height?: number;
  /** Content type (MIME type) */
  contentType?: string;
  /** Size in bytes */
  size?: number;
}

/**
 * Position for placing media on a slide.
 */
export interface MediaPosition {
  /** X position in EMUs (optional, defaults to center) */
  x?: number;
  /** Y position in EMUs (optional, defaults to center) */
  y?: number;
  /** Width in EMUs (optional, defaults to 6 inches / 5486400 EMUs) */
  width?: number;
  /** Height in EMUs (optional, defaults to 4 inches / 3657600 EMUs) */
  height?: number;
}

/**
 * Represents image data for adding/replacing pictures.
 */
export interface ImageData {
  /** Image data as a Buffer (raw image bytes) */
  data: Buffer;
  /** MIME type of the image (e.g., "image/png", "image/jpeg") */
  contentType: string;
  /** Optional filename for the image */
  filename?: string;
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
 * Reads an entry from the zip as a Buffer.
 * Returns null if the entry is not found.
 */
function getEntry(zip: Map<string, Buffer>, entryName: string): Buffer | null {
  return zip.get(entryName) ?? null;
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
 * Extracts media items from a slide's XML.
 */
function extractMediaFromSlide(
  slideXml: string,
  slideIndex: number,
  slideRels: Array<{ id: string; target: string; type?: string }>,
  slideEntry: string
): MediaItem[] {
  const mediaItems: MediaItem[] = [];

  // Find all picture elements
  const picturePattern = /<p:pic\b[\s\S]*?<\/p:pic>/g;
  const pictureMatches = slideXml.match(picturePattern) || [];

  let picIdx = 0;
  for (const picXml of pictureMatches) {
    picIdx++;

    // Check if it's a video or audio
    const isVideo = /<a:videoFile>/.test(picXml) || /<p:video>/.test(picXml);
    const isAudio = /<a:audioFile>/.test(picXml) || /<p:audio>/.test(picXml);

    // Skip video/audio for now - they are handled differently
    // For this implementation, we focus on pictures
    if (isVideo || isAudio) {
      continue;
    }

    // Extract name
    const nameMatch = /<p:cNvPr[^>]*name="([^"]*)"[^>]*>/.exec(picXml);
    const name = nameMatch ? nameMatch[1] : `Picture ${picIdx}`;

    // Extract alt text
    const altMatch = /<p:cNvPr[^>]*descr="([^"]*)"[^>]*>/.exec(picXml);
    const alt = altMatch ? altMatch[1] : undefined;

    // Extract position and size
    const xfrmMatch = /<a:xfrm(?:[^>]*)>([\s\S]*?)<\/a:xfrm>/.exec(picXml);
    let x: number | undefined;
    let y: number | undefined;
    let width: number | undefined;
    let height: number | undefined;

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
    }

    // Get media path
    const path = `/slide[${slideIndex}]/picture[${picIdx}]`;

    // Determine media type
    const mediaType = isVideo ? "video" : isAudio ? "audio" : "picture";

    mediaItems.push({
      path,
      type: mediaType,
      name,
      alt,
      x,
      y,
      width,
      height,
    });
  }

  return mediaItems;
}

// ============================================================================
// Media Operations
// ============================================================================

/**
 * Lists all media items on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 *
 * @example
 * const result = await getMedia("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.media);
 * }
 */
export async function getMedia(
  filePath: string,
  slideIndex: number
): Promise<Result<{ media: MediaItem[]; total: number }>> {
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

    const media = extractMediaFromSlide(slideXml, slideIndex, relationships, slideEntry);

    return ok({ media, total: media.length });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets binary data of a media item.
 *
 * @param filePath - Path to the PPTX file
 * @param mediaPath - Path to the media item (e.g., "/slide[1]/picture[1]")
 *
 * @example
 * const result = await getMediaData("/path/to/presentation.pptx", "/slide[1]/picture[1]");
 * if (result.ok) {
 *   console.log(result.data.data); // Buffer of image data
 *   console.log(result.data.contentType); // e.g., "image/png"
 * }
 */
export async function getMediaData(
  filePath: string,
  mediaPath: string
): Promise<Result<{ data: Buffer; contentType: string; filename?: string }>> {
  try {
    const slideIndex = getSlideIndex(mediaPath);
    if (slideIndex === null) {
      return invalidInput("Invalid media path - must include slide index");
    }

    // Extract picture index from path
    const picIndexMatch = mediaPath.match(/\/picture\[(\d+)\]/i);
    if (!picIndexMatch) {
      return invalidInput("Invalid media path - must include picture[index]");
    }
    const picIndex = parseInt(picIndexMatch[1], 10);

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

    // Find the picture element
    const picturePattern = /<p:pic\b[\s\S]*?<\/p:pic>/g;
    const pictureMatches = slideXml.match(picturePattern) || [];

    if (picIndex < 1 || picIndex > pictureMatches.length) {
      return notFound("Picture", String(picIndex));
    }

    const picXml = pictureMatches[picIndex - 1];

    // Extract the blip's embed attribute (relationship ID)
    const blipMatch = /<a:blip\b[^>]*r:embed="([^"]*)"[^>]*>/.exec(picXml);
    if (!blipMatch) {
      return invalidInput("Picture does not have an embedded image");
    }

    const relId = blipMatch[1];
    const rel = relationships.find(r => r.id === relId);
    if (!rel) {
      return invalidInput("Image relationship not found");
    }

    // Resolve the image path
    const slideDir = path.posix.dirname(slideEntry);
    const imagePath = normalizeZipPath(slideDir, rel.target);

    // Get the image data
    const imageData = getEntry(zip, imagePath);
    if (!imageData) {
      return invalidInput("Image data not found in archive");
    }

    // Determine content type from extension
    const ext = path.extname(rel.target).toLowerCase();
    const contentType = extToContentType(ext);

    return ok({
      data: imageData,
      contentType,
      filename: path.basename(rel.target),
    });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Converts file extension to content type.
 */
function extToContentType(ext: string): string {
  const contentTypes: Record<string, string> = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tiff": "image/tiff",
    ".tif": "image/tiff",
    ".webp": "image/webp",
    ".svg": "image/svg+xml",
    ".emf": "image/x-emf",
    ".wmf": "image/x-wmf",
    ".ico": "image/x-icon",
  };
  return contentTypes[ext] || "application/octet-stream";
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
 * Adds a picture to a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param imageData - Image data to add
 * @param position - Optional position and size
 *
 * @example
 * const imageBuffer = await readFile("image.png");
 * const result = await addPicture(
 *   "/path/to/presentation.pptx",
 *   1,
 *   { data: imageBuffer, contentType: "image/png" },
 *   { x: 1000000, y: 1000000, width: 3000000, height: 2000000 }
 * );
 */
export async function addPicture(
  filePath: string,
  slideIndex: number,
  imageData: ImageData,
  position?: MediaPosition
): Promise<Result<{ path: string }>> {
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
    let relsXml = "";
    try {
      relsXml = requireEntry(zip, relsEntry);
    } catch {
      // Create empty rels if it doesn't exist
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    }
    const relationships = parseRelationshipEntries(relsXml);

    // Get slide size for default positioning
    const slideSize = getSlideSize(zip);

    // Calculate position and size
    const picWidth = position?.width ?? 5486400; // Default 6 inches
    const picHeight = position?.height ?? 3657600; // Default 4 inches
    const picX = position?.x ?? Math.floor((slideSize.width - picWidth) / 2);
    const picY = position?.y ?? Math.floor((slideSize.height - picHeight) / 2);

    // Generate unique IDs
    const existingRelIds = relationships.map(r => r.id);
    const newRelId = generateRelId(existingRelIds);

    // Find existing picture IDs to generate unique shape ID
    const existingShapeIds: number[] = [];
    for (const match of slideXml.matchAll(/<p:cNvPr[^>]*id="(\d+)"[^>]*>/g)) {
      existingShapeIds.push(parseInt(match[1], 10));
    }
    const newShapeId = generateShapeId(existingShapeIds);

    // Count existing pictures for naming
    const picCount = (slideXml.match(/<p:pic\b/g) || []).length;
    const picName = `Picture ${picCount + 1}`;

    // Determine image extension from content type
    const ext = contentTypeToExt(imageData.contentType);
    const imageFilename = imageData.filename || `image${ext}`;
    const imageEntry = `ppt/media/${imageFilename}`;

    // Create image part relationship
    const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${imageFilename}"/>`;

    // Build the picture XML
    const pictureXml = `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${newShapeId}" name="${picName}"/>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="${newRelId}"/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${picX}" y="${picY}"/>
      <a:ext cx="${picWidth}" cy="${picHeight}"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
</p:pic>`;

    // Insert picture into slide XML before closing </p:spTree>
    const updatedSlideXml = slideXml.replace(
      "</p:spTree>",
      `${pictureXml}</p:spTree>`
    );

    // Update relationships XML
    const updatedRelsXml = relsXml.replace(
      "</Relationships>",
      `${newRelEntry}</Relationships>`
    );

    // Ensure media directory exists and add image
    const newEntries: Array<{ name: string; data: Buffer }> = [];

    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else if (name === relsEntry) {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
      } else if (name === imageEntry) {
        // Don't duplicate - skip since we're replacing
      } else {
        newEntries.push({ name, data });
      }
    }

    // Add the image data
    newEntries.push({ name: imageEntry, data: imageData.data });

    await writeFile(filePath, createStoredZip(newEntries));

    return ok({ path: `/slide[${slideIndex}]/picture[${picCount + 1}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Converts content type to file extension.
 */
function contentTypeToExt(contentType: string): string {
  const exts: Record<string, string> = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/tiff": ".tiff",
    "image/webp": ".webp",
    "image/svg+xml": ".svg",
    "image/x-emf": ".emf",
    "image/x-wmf": ".wmf",
    "image/x-icon": ".ico",
  };
  return exts[contentType] || ".png";
}

/**
 * Removes media from a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param mediaPath - Path to the media item (e.g., "/slide[1]/picture[1]")
 *
 * @example
 * const result = await removeMedia("/path/to/presentation.pptx", "/slide[1]/picture[1]");
 */
export async function removeMedia(
  filePath: string,
  mediaPath: string
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(mediaPath);
    if (slideIndex === null) {
      return invalidInput("Invalid media path - must include slide index");
    }

    // Extract picture index from path
    const picIndexMatch = mediaPath.match(/\/picture\[(\d+)\]/i);
    if (!picIndexMatch) {
      return invalidInput("Invalid media path - must include picture[index]");
    }
    const picIndex = parseInt(picIndexMatch[1], 10);

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

    // Find the picture element
    const picturePattern = /<p:pic\b[\s\S]*?<\/p:pic>/g;
    const pictureMatches = slideXml.match(picturePattern) || [];

    if (picIndex < 1 || picIndex > pictureMatches.length) {
      return notFound("Picture", String(picIndex));
    }

    const picXml = pictureMatches[picIndex - 1];

    // Remove the picture from slide XML
    const updatedSlideXml = slideXml.replace(picXml, "");

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
 * Replaces an existing picture with new image data.
 *
 * @param filePath - Path to the PPTX file
 * @param mediaPath - Path to the media item (e.g., "/slide[1]/picture[1]")
 * @param newImageData - New image data to replace with
 *
 * @example
 * const newImageBuffer = await readFile("new_image.png");
 * const result = await replacePicture(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/picture[1]",
 *   { data: newImageBuffer, contentType: "image/png" }
 * );
 */
export async function replacePicture(
  filePath: string,
  mediaPath: string,
  newImageData: ImageData
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(mediaPath);
    if (slideIndex === null) {
      return invalidInput("Invalid media path - must include slide index");
    }

    // Extract picture index from path
    const picIndexMatch = mediaPath.match(/\/picture\[(\d+)\]/i);
    if (!picIndexMatch) {
      return invalidInput("Invalid media path - must include picture[index]");
    }
    const picIndex = parseInt(picIndexMatch[1], 10);

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

    // Find the picture element
    const picturePattern = /<p:pic\b[\s\S]*?<\/p:pic>/g;
    const pictureMatches = slideXml.match(picturePattern) || [];

    if (picIndex < 1 || picIndex > pictureMatches.length) {
      return notFound("Picture", String(picIndex));
    }

    const picXml = pictureMatches[picIndex - 1];

    // Extract the existing relationship ID
    const blipMatch = /<a:blip\b[^>]*r:embed="([^"]*)"[^>]*>/.exec(picXml);
    if (!blipMatch) {
      return invalidInput("Picture does not have an embedded image");
    }

    const relId = blipMatch[1];
    const rel = relationships.find(r => r.id === relId);
    if (!rel) {
      return invalidInput("Image relationship not found");
    }

    // Resolve the existing image path
    const slideDir = path.posix.dirname(slideEntry);
    const oldImagePath = normalizeZipPath(slideDir, rel.target);

    // Determine new image extension and filename
    const ext = contentTypeToExt(newImageData.contentType);
    const imageFilename = `replaced_image${ext}`;
    const newImageEntry = `ppt/media/${imageFilename}`;

    // Update the relationship to point to the new image
    let updatedRelsXml = relsXml;
    if (rel.target !== `../media/${imageFilename}`) {
      updatedRelsXml = relsXml.replace(
        `Target="${rel.target}"`,
        `Target="../media/${imageFilename}"`
      );
    }

    // Build new zip with updated slide and image
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    let newImagePath = oldImagePath;

    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(slideXml, "utf8") });
      } else if (name === relsEntry) {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
      } else if (name === oldImagePath) {
        // Replace old image with new
        newEntries.push({ name: newImageEntry, data: newImageData.data });
        newImagePath = newImageEntry;
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
