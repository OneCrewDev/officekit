/**
 * Advanced media operations for @officekit/ppt.
 *
 * Provides functions to manage video and audio embedding in PowerPoint
 * presentations:
 * - Add video to slides
 * - Add audio to slides
 * - List video/audio on a slide
 * - Remove video/audio
 * - Update playback options
 */

import { readFile, writeFile, stat } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, andThen, invalidInput, notFound } from "./result.js";
import type { Result } from "./types.js";
import { getSlideIndex } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Position for placing media on a slide.
 */
export interface MediaPosition {
  /** X position in EMUs */
  x: number;
  /** Y position in EMUs */
  y: number;
  /** Width in EMUs */
  width: number;
  /** Height in EMUs */
  height: number;
}

/**
 * Options for video embedding.
 */
export interface VideoOptions {
  /** Whether to autoplay the video */
  autoplay?: boolean;
  /** Whether to loop the video */
  loop?: boolean;
  /** Whether to mute the video */
  mute?: boolean;
  /** Path to poster image file (preview image shown before playback) */
  posterImage?: string;
  /** Volume level (0-100) */
  volume?: number;
}

/**
 * Options for audio embedding.
 */
export interface AudioOptions {
  /** Whether to autoplay the audio */
  autoplay?: boolean;
  /** Whether to loop the audio */
  loop?: boolean;
  /** Volume level (0-100), defaults to 50 */
  volume?: number;
}

/**
 * Represents a media element (video or audio).
 */
export interface MediaElement {
  /** Path to the media element (e.g., "/slide[1]/media[1]") */
  path: string;
  /** Media type: "video" or "audio" */
  type: "video" | "audio";
  /** Name of the media element */
  name: string;
  /** Content type (MIME type) */
  contentType?: string;
  /** Size in bytes */
  size?: number;
  /** Position X in EMUs (for video) */
  x?: number;
  /** Position Y in EMUs (for video) */
  y?: number;
  /** Width in EMUs (for video) */
  width?: number;
  /** Height in EMUs (for video) */
  height?: number;
  /** Autoplay setting */
  autoplay?: boolean;
  /** Loop setting */
  loop?: boolean;
  /** Mute setting (video only) */
  mute?: boolean;
  /** Volume setting */
  volume?: number;
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
 * Converts file extension to content type.
 */
function extToContentType(ext: string): string {
  const contentTypes: Record<string, string> = {
    ".mp4": "video/mp4",
    ".avi": "video/avi",
    ".mov": "video/quicktime",
    ".wmv": "video/x-ms-wmv",
    ".mkv": "video/x-matroska",
    ".webm": "video/webm",
    ".m4v": "video/x-m4v",
    ".mp3": "audio/mpeg",
    ".wav": "audio/wav",
    ".aac": "audio/aac",
    ".ogg": "audio/ogg",
    ".wma": "audio/x-ms-wma",
    ".m4a": "audio/x-m4a",
    ".flac": "audio/flac",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
  };
  return contentTypes[ext.toLowerCase()] || "application/octet-stream";
}

/**
 * Converts content type to file extension.
 */
function contentTypeToExt(contentType: string): string {
  const exts: Record<string, string> = {
    "video/mp4": ".mp4",
    "video/avi": ".avi",
    "video/quicktime": ".mov",
    "video/x-ms-wmv": ".wmv",
    "video/x-matroska": ".mkv",
    "video/webm": ".webm",
    "video/x-m4v": ".m4v",
    "audio/mpeg": ".mp3",
    "audio/wav": ".wav",
    "audio/aac": ".aac",
    "audio/ogg": ".ogg",
    "audio/x-ms-wma": ".wma",
    "audio/x-m4a": ".m4a",
    "audio/flac": ".flac",
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
  };
  return exts[contentType] || "";
}

/**
 * Ensures the ContentTypes.xml includes the necessary content type for media.
 */
function ensureContentTypes(zip: Map<string, Buffer>, ext: string, contentType: string): Map<string, Buffer> {
  const contentTypesEntry = "[Content_Types].xml";
  let contentTypesXml: string;

  try {
    contentTypesXml = requireEntry(zip, contentTypesEntry);
  } catch {
    // ContentTypes.xml doesn't exist, create it
    contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`;
  }

  // Check if this content type is already registered
  const extPattern = new RegExp(`Extension="${ext}"`, "i");
  if (extPattern.test(contentTypesXml)) {
    return zip;
  }

  // Add the content type
  const newContentType = `<Default Extension="${ext}" ContentType="${contentType}"/>`;
  const updatedXml = contentTypesXml.replace(
    "</Types>",
    `${newContentType}</Types>`
  );

  const newEntries = new Map<string, Buffer>();
  for (const [name, data] of zip.entries()) {
    if (name === contentTypesEntry) {
      newEntries.set(name, Buffer.from(updatedXml, "utf8"));
    } else {
      newEntries.set(name, data);
    }
  }

  return newEntries;
}

// ============================================================================
// Video Operations
// ============================================================================

/**
 * Embeds a video in a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param videoPath - Path to the video file (.mp4, .avi, .mov, etc.)
 * @param position - Position and size of the video on the slide
 * @param options - Optional video options (autoplay, loop, mute, posterImage)
 *
 * @example
 * const result = await addVideo(
 *   "/path/to/presentation.pptx",
 *   1,
 *   "/path/to/video.mp4",
 *   { x: 1000000, y: 1000000, width: 3000000, height: 2000000 },
 *   { autoplay: false, loop: false, mute: true }
 * );
 */
export async function addVideo(
  filePath: string,
  slideIndex: number,
  videoPath: string,
  position: MediaPosition,
  options?: VideoOptions
): Promise<Result<{ path: string }>> {
  try {
    // Validate video file exists and get its data
    const videoStats = await stat(videoPath);
    if (!videoStats.isFile()) {
      return invalidInput(`Video path is not a file: ${videoPath}`);
    }

    const videoData = await readFile(videoPath);
    const videoExt = path.extname(videoPath).toLowerCase();
    const videoContentType = extToContentType(videoExt);

    // Read existing presentation
    const buffer = await readFile(filePath);
    let zip = readStoredZip(buffer);

    // Ensure content type is registered
    zip = ensureContentTypes(zip, videoExt, videoContentType);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    let relsXml = "";
    try {
      relsXml = requireEntry(zip, relsEntry);
    } catch {
      // Create empty rels if it doesn't exist
      relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    }
    const relationships = parseRelationshipEntries(relsXml);

    // Generate unique IDs
    const existingRelIds = relationships.map(r => r.id);
    const videoRelId = generateRelId(existingRelIds);
    const posterRelId = generateRelId([...existingRelIds, videoRelId]);

    // Find existing shape IDs
    const existingShapeIds: number[] = [];
    for (const match of slideXml.matchAll(/<p:cNvPr[^>]*id="(\d+)"[^>]*>/g)) {
      existingShapeIds.push(parseInt(match[1], 10));
    }
    const newShapeId = generateShapeId(existingShapeIds);

    // Count existing media elements for naming
    const mediaCount = (
      (slideXml.match(/<p:pic\b/g) || []).length +
      (slideXml.match(/<mc:AlternateContent\b/g) || []).length +
      1
    );
    const mediaName = `Video ${mediaCount}`;

    // Determine video filename
    const videoFilename = `video_${Date.now()}${videoExt}`;
    const videoEntry = `ppt/media/${videoFilename}`;

    // Handle poster image if provided
    let posterFilename: string | undefined;
    let posterEntry: string | undefined;
    let posterRelIdToUse: string | undefined;

    if (options?.posterImage) {
      const posterStats = await stat(options.posterImage);
      if (posterStats.isFile()) {
        const posterData = await readFile(options.posterImage);
        const posterExt = path.extname(options.posterImage).toLowerCase();
        const posterContentType = extToContentType(posterExt);

        // Ensure content type for poster
        zip = ensureContentTypes(zip, posterExt, posterContentType);

        posterFilename = `poster_${Date.now()}${posterExt}`;
        posterEntry = `ppt/media/${posterFilename}`;
        posterRelIdToUse = posterRelId;

        // Create video relationship with poster reference
        const videoRelEntry = `<Relationship Id="${videoRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/${videoFilename}"/>`;
        const posterRelEntry = `<Relationship Id="${posterRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${posterFilename}"/>`;

        // Build video XML with poster
        const autoplay = options.autoplay ?? false;
        const loop = options.loop ?? false;
        const mute = options.mute ?? true;
        const volume = options.volume ?? 50;

        const videoXml = `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${newShapeId}" name="${mediaName}"/>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${position.x}" y="${position.y}"/>
      <a:ext cx="${position.width}" cy="${position.height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
  <p:video>
    <p:videoFile name="${mediaName}" relId="${videoRelId}" posterImageRelId="${posterRelId}">
      <p:videoPr autoplay="${autoplay ? "1" : "0"}" loop="${loop ? "1" : "0"}" mute="${mute ? "1" : "0"}">
        <a:vol>${volume / 100}</a:vol>
      </p:videoPr>
    </p:videoFile>
  </p:video>
</p:pic>`;

        // Update relationships
        const updatedRelsXml = relsXml.replace(
          "</Relationships>",
          `${videoRelEntry}${posterRelEntry}</Relationships>`
        );

        // Build new zip
        const newEntries: Array<{ name: string; data: Buffer }> = [];
        for (const [name, data] of zip.entries()) {
          if (name === slideEntry) {
            newEntries.push({ name, data: Buffer.from(slideXml.replace("</p:spTree>", `${videoXml}</p:spTree>`), "utf8") });
          } else if (name === relsEntry) {
            newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
          } else if (name === posterEntry) {
            // Skip - we're adding a new one
          } else {
            newEntries.push({ name, data });
          }
        }

        // Add video and poster data
        newEntries.push({ name: videoEntry, data: videoData });
        newEntries.push({ name: posterEntry!, data: posterData });

        await writeFile(filePath, createStoredZip(newEntries));
        return ok({ path: `/slide[${slideIndex}]/media[${mediaCount}]` });
      }
    }

    // No poster image - create video without poster
    const autoplay = options?.autoplay ?? false;
    const loop = options?.loop ?? false;
    const mute = options?.mute ?? true;
    const volume = options?.volume ?? 50;

    // Create video relationship
    const videoRelEntry = `<Relationship Id="${videoRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/${videoFilename}"/>`;

    // Build video XML
    const videoXml = `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${newShapeId}" name="${mediaName}"/>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${position.x}" y="${position.y}"/>
      <a:ext cx="${position.width}" cy="${position.height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
  <p:video>
    <p:videoFile name="${mediaName}" relId="${videoRelId}">
      <p:videoPr autoplay="${autoplay ? "1" : "0"}" loop="${loop ? "1" : "0"}" mute="${mute ? "1" : "0"}">
        <a:vol>${volume / 100}</a:vol>
      </p:videoPr>
    </p:videoFile>
  </p:video>
</p:pic>`;

    // Update relationships
    const updatedRelsXml = relsXml.replace(
      "</Relationships>",
      `${videoRelEntry}</Relationships>`
    );

    // Build new zip
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(slideXml.replace("</p:spTree>", `${videoXml}</p:spTree>`), "utf8") });
      } else if (name === relsEntry) {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    // Add video data
    newEntries.push({ name: videoEntry, data: videoData });

    await writeFile(filePath, createStoredZip(newEntries));
    return ok({ path: `/slide[${slideIndex}]/media[${mediaCount}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Audio Operations
// ============================================================================

/**
 * Embeds an audio file in a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param audioPath - Path to the audio file (.mp3, .wav, etc.)
 * @param position - Optional position for the audio icon on the slide
 * @param options - Optional audio options (autoplay, loop, volume)
 *
 * @example
 * const result = await addAudio(
 *   "/path/to/presentation.pptx",
 *   1,
 *   "/path/to/audio.mp3",
 *   { x: 1000000, y: 1000000, width: 500000, height: 500000 },
 *   { autoplay: true, loop: false, volume: 75 }
 * );
 */
export async function addAudio(
  filePath: string,
  slideIndex: number,
  audioPath: string,
  position?: MediaPosition,
  options?: AudioOptions
): Promise<Result<{ path: string }>> {
  try {
    // Validate audio file exists and get its data
    const audioStats = await stat(audioPath);
    if (!audioStats.isFile()) {
      return invalidInput(`Audio path is not a file: ${audioPath}`);
    }

    const audioData = await readFile(audioPath);
    const audioExt = path.extname(audioPath).toLowerCase();
    const audioContentType = extToContentType(audioExt);

    // Read existing presentation
    const buffer = await readFile(filePath);
    let zip = readStoredZip(buffer);

    // Ensure content type is registered
    zip = ensureContentTypes(zip, audioExt, audioContentType);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    let relsXml = "";
    try {
      relsXml = requireEntry(zip, relsEntry);
    } catch {
      // Create empty rels if it doesn't exist
      relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    }
    const relationships = parseRelationshipEntries(relsXml);

    // Generate unique IDs
    const existingRelIds = relationships.map(r => r.id);
    const audioRelId = generateRelId(existingRelIds);

    // Find existing shape IDs
    const existingShapeIds: number[] = [];
    for (const match of slideXml.matchAll(/<p:cNvPr[^>]*id="(\d+)"[^>]*>/g)) {
      existingShapeIds.push(parseInt(match[1], 10));
    }
    const newShapeId = generateShapeId(existingShapeIds);

    // Count existing media elements for naming
    const mediaCount = (
      (slideXml.match(/<p:pic\b/g) || []).length +
      (slideXml.match(/<p:audio\b/g) || []).length +
      1
    );
    const mediaName = `Audio ${mediaCount}`;

    // Determine audio filename
    const audioFilename = `audio_${Date.now()}${audioExt}`;
    const audioEntry = `ppt/media/${audioFilename}`;

    // Create audio relationship
    const audioRelEntry = `<Relationship Id="${audioRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="../media/${audioFilename}"/>`;

    const autoplay = options?.autoplay ?? false;
    const loop = options?.loop ?? false;
    const volume = options?.volume ?? 50;

    // Default icon size and position if not provided (center of slide)
    const slideSize = getSlideSize(zip);
    const iconSize = 500000; // ~0.5 inch
    const iconX = position?.x ?? Math.floor((slideSize.width - iconSize) / 2);
    const iconY = position?.y ?? Math.floor((slideSize.height - iconSize) / 2);
    const iconWidth = position?.width ?? iconSize;
    const iconHeight = position?.height ?? iconSize;

    // Build audio XML - audio is typically just a cNvPr element with audioFile reference
    const audioXml = `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${newShapeId}" name="${mediaName}"/>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${iconX}" y="${iconY}"/>
      <a:ext cx="${iconWidth}" cy="${iconHeight}"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
  <p:audio>
    <p:audioFile name="${mediaName}" relId="${audioRelId}">
      <p:audioPr autoplay="${autoplay ? "1" : "0"}" loop="${loop ? "1" : "0"}">
        <a:vol>${volume / 100}</a:vol>
      </p:audioPr>
    </p:audioFile>
  </p:audio>
</p:pic>`;

    // Update relationships
    const updatedRelsXml = relsXml.replace(
      "</Relationships>",
      `${audioRelEntry}</Relationships>`
    );

    // Build new zip
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(slideXml.replace("</p:spTree>", `${audioXml}</p:spTree>`), "utf8") });
      } else if (name === relsEntry) {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    // Add audio data
    newEntries.push({ name: audioEntry, data: audioData });

    await writeFile(filePath, createStoredZip(newEntries));
    return ok({ path: `/slide[${slideIndex}]/media[${mediaCount}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Media Element Query Operations
// ============================================================================

/**
 * Lists all video and audio elements on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 *
 * @example
 * const result = await getMediaElements("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.media);
 * }
 */
export async function getMediaElements(
  filePath: string,
  slideIndex: number
): Promise<Result<{ media: MediaElement[]; total: number }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    const relsXml = requireEntry(zip, relsEntry);
    const relationships = parseRelationshipEntries(relsXml);

    const mediaElements: MediaElement[] = [];
    let mediaIdx = 0;

    // Find all video elements
    const videoPattern = /<p:video\b[\s\S]*?<\/p:video>/g;
    const videoMatches = slideXml.match(videoPattern) || [];

    for (const videoXml of videoMatches) {
      mediaIdx++;

      // Extract name
      const nameMatch = /<p:cNvPr[^>]*name="([^"]*)"[^>]*>/.exec(videoXml);
      const name = nameMatch ? nameMatch[1] : `Video ${mediaIdx}`;

      // Extract autoplay, loop, mute from videoPr
      const autoplayMatch = /<p:videoPr[^>]*autoplay="([^"]*)"[^>]*>/.exec(videoXml);
      const loopMatch = /<p:videoPr[^>]*loop="([^"]*)"[^>]*>/.exec(videoXml);
      const muteMatch = /<p:videoPr[^>]*mute="([^"]*)"[^>]*>/.exec(videoXml);

      // Extract volume
      const volMatch = /<a:vol>([^<]*)<\/a:vol>/.exec(videoXml);
      const volume = volMatch ? Math.round(parseFloat(volMatch[1]) * 100) : 50;

      // Get relId for the video
      const relIdMatch = /<p:videoFile[^>]*relId="([^"]*)"[^>]*>/.exec(videoXml);
      let contentType: string | undefined;

      if (relIdMatch) {
        const rel = relationships.find(r => r.id === relIdMatch[1]);
        if (rel) {
          const targetPath = normalizeZipPath(path.posix.dirname(slideEntry), rel.target);
          const ext = path.extname(targetPath).toLowerCase();
          contentType = extToContentType(ext);
        }
      }

      mediaElements.push({
        path: `/slide[${slideIndex}]/media[${mediaIdx}]`,
        type: "video",
        name,
        contentType,
        autoplay: autoplayMatch?.[1] === "1",
        loop: loopMatch?.[1] === "1",
        mute: muteMatch?.[1] === "1",
        volume,
      });
    }

    // Find all audio elements
    const audioPattern = /<p:audio\b[\s\S]*?<\/p:audio>/g;
    const audioMatches = slideXml.match(audioPattern) || [];

    for (const audioXml of audioMatches) {
      mediaIdx++;

      // Extract name
      const nameMatch = /<p:cNvPr[^>]*name="([^"]*)"[^>]*>/.exec(audioXml);
      const name = nameMatch ? nameMatch[1] : `Audio ${mediaIdx}`;

      // Extract autoplay, loop
      const autoplayMatch = /<p:audioPr[^>]*autoplay="([^"]*)"[^>]*>/.exec(audioXml);
      const loopMatch = /<p:audioPr[^>]*loop="([^"]*)"[^>]*>/.exec(audioXml);

      // Extract volume
      const volMatch = /<a:vol>([^<]*)<\/a:vol>/.exec(audioXml);
      const volume = volMatch ? Math.round(parseFloat(volMatch[1]) * 100) : 50;

      // Get relId for the audio
      const relIdMatch = /<p:audioFile[^>]*relId="([^"]*)"[^>]*>/.exec(audioXml);
      let contentType: string | undefined;

      if (relIdMatch) {
        const rel = relationships.find(r => r.id === relIdMatch[1]);
        if (rel) {
          const targetPath = normalizeZipPath(path.posix.dirname(slideEntry), rel.target);
          const ext = path.extname(targetPath).toLowerCase();
          contentType = extToContentType(ext);
        }
      }

      mediaElements.push({
        path: `/slide[${slideIndex}]/media[${mediaIdx}]`,
        type: "audio",
        name,
        contentType,
        autoplay: autoplayMatch?.[1] === "1",
        loop: loopMatch?.[1] === "1",
        volume,
      });
    }

    return ok({ media: mediaElements, total: mediaElements.length });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Remove Media Element
// ============================================================================

/**
 * Removes a video or audio element from a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param path - Path to the media element (e.g., "/slide[1]/media[1]")
 *
 * @example
 * const result = await removeMediaElement("/path/to/presentation.pptx", "/slide[1]/media[1]");
 */
export async function removeMediaElement(
  filePath: string,
  mediaPath: string
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(mediaPath);
    if (slideIndex === null) {
      return invalidInput("Invalid media path - must include slide index");
    }

    // Extract media index from path
    const mediaIndexMatch = mediaPath.match(/\/media\[(\d+)\]/i);
    if (!mediaIndexMatch) {
      return invalidInput("Invalid media path - must include media[index]");
    }
    const mediaIndex = parseInt(mediaIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);

    // Find all p:pic elements and identify which ones contain video/audio
    const picPattern = /<p:pic\b[\s\S]*?<\/p:pic>/g;
    const picMatches = slideXml.match(picPattern) || [];

    // Filter to only p:pic elements that contain video or audio
    const mediaPicElements: Array<{ xml: string; type: "video" | "audio" }> = [];
    for (const picXml of picMatches) {
      if (/<p:video\b[\s\S]*?<\/p:video>/.test(picXml)) {
        mediaPicElements.push({ xml: picXml, type: "video" });
      } else if (/<p:audio\b[\s\S]*?<\/p:audio>/.test(picXml)) {
        mediaPicElements.push({ xml: picXml, type: "audio" });
      }
    }

    if (mediaIndex < 1 || mediaIndex > mediaPicElements.length) {
      return notFound("Media", String(mediaIndex));
    }

    const targetPicXml = mediaPicElements[mediaIndex - 1].xml;

    // Remove the p:pic element from slide XML
    const updatedSlideXml = slideXml.replace(targetPicXml, "");

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
// Set Media Options
// ============================================================================

/**
 * Updates playback options for a video or audio element.
 *
 * @param filePath - Path to the PPTX file
 * @param mediaPath - Path to the media element (e.g., "/slide[1]/media[1]")
 * @param options - Options to update (autoplay, loop, volume, mute for video)
 *
 * @example
 * const result = await setMediaOptions(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/media[1]",
 *   { autoplay: true, loop: true, volume: 80 }
 * );
 */
export async function setMediaOptions(
  filePath: string,
  mediaPath: string,
  options: VideoOptions | AudioOptions
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(mediaPath);
    if (slideIndex === null) {
      return invalidInput("Invalid media path - must include slide index");
    }

    // Extract media index from path
    const mediaIndexMatch = mediaPath.match(/\/media\[(\d+)\]/i);
    if (!mediaIndexMatch) {
      return invalidInput("Invalid media path - must include media[index]");
    }
    const mediaIndex = parseInt(mediaIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);

    // Find all video and audio elements
    const videoPattern = /<p:video\b[\s\S]*?<\/p:video>/g;
    const audioPattern = /<p:audio\b[\s\S]*?<\/p:audio>/g;
    const videoMatches = slideXml.match(videoPattern) || [];
    const audioMatches = slideXml.match(audioPattern) || [];

    // Combine and index them
    const allMediaMatches: Array<{ xml: string; type: "video" | "audio" }> = [
      ...videoMatches.map(xml => ({ xml, type: "video" as const })),
      ...audioMatches.map(xml => ({ xml, type: "audio" as const })),
    ];

    if (mediaIndex < 1 || mediaIndex > allMediaMatches.length) {
      return notFound("Media", String(mediaIndex));
    }

    const targetMedia = allMediaMatches[mediaIndex - 1];
    let updatedXml = targetMedia.xml;

    // Update autoplay if provided
    if (options.autoplay !== undefined) {
      const autoplayValue = options.autoplay ? "1" : "0";
      updatedXml = updatedXml.replace(
        /autoplay="[^"]*"/,
        `autoplay="${autoplayValue}"`
      );
      // If autoplay attribute doesn't exist in videoPr, add it
      if (!updatedXml.includes('autoplay="')) {
        updatedXml = updatedXml.replace(
          /<p:(videoPr|audioPr)/,
          `<p:$1 autoplay="${autoplayValue}"`
        );
      }
    }

    // Update loop if provided
    if (options.loop !== undefined) {
      const loopValue = options.loop ? "1" : "0";
      updatedXml = updatedXml.replace(
        /loop="[^"]*"/,
        `loop="${loopValue}"`
      );
      // If loop attribute doesn't exist, add it
      if (!updatedXml.includes('loop="')) {
        updatedXml = updatedXml.replace(
          /<p:(videoPr|audioPr)/,
          `<p:$1 loop="${loopValue}"`
        );
      }
    }

    // Update mute only for video
    if (targetMedia.type === "video" && "mute" in options && options.mute !== undefined) {
      const muteValue = options.mute ? "1" : "0";
      updatedXml = updatedXml.replace(
        /mute="[^"]*"/,
        `mute="${muteValue}"`
      );
      // If mute attribute doesn't exist, add it
      if (!updatedXml.includes('mute="')) {
        updatedXml = updatedXml.replace(
          /<p:videoPr/,
          `<p:videoPr mute="${muteValue}"`
        );
      }
    }

    // Update volume if provided
    if (options.volume !== undefined) {
      const volumeValue = (options.volume / 100).toFixed(2);
      updatedXml = updatedXml.replace(
        /<a:vol>[^<]*<\/a:vol>/,
        `<a:vol>${volumeValue}</a:vol>`
      );
    }

    // Replace the old XML with updated XML in the slide
    let updatedSlideXml = slideXml.replace(targetMedia.xml, updatedXml);

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
