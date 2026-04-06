/**
 * Slide management operations for @officekit/ppt.
 *
 * Provides functions to add, remove, move, and duplicate slides within
 * a PowerPoint presentation.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, andThen, map, notFound, invalidInput } from "./result.js";
import type { Result } from "./types.js";
import { slidePath } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Represents a slide in the presentation.
 */
export interface SlideInfo {
  /** 1-based slide index */
  index: number;
  /** Path to the slide file */
  path: string;
  /** Relationship ID in presentation.xml.rels */
  relId: string;
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
  // Match <p:sldId id="256" r:id="rId2"/>
  for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*\bid="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g)) {
    slideIds.push({ id: match[1], relId: match[2] });
  }
  // Also handle case where r:id comes first
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
  // Remove all existing <p:sldId .../> elements
  let result = presentationXml.replace(/<p:sldId\b[^>]*\/?>/g, "");

  // Find the <p:sldIdLst> element and add new slide IDs
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

// ============================================================================
// Slide Management
// ============================================================================

/**
 * Gets all slides in the presentation.
 *
 * @example
 * const result = await getSlides("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(result.data.slides);
 * }
 */
export async function getSlides(filePath: string): Promise<Result<{ slides: SlideInfo[]; total: number }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);

    const slideIds = getSlideIds(presentationXml);
    const slides: SlideInfo[] = slideIds.map((s, index) => {
      const rel = relationships.find(r => r.id === s.relId);
      const target = rel?.target ?? "";
      const slidePath = normalizeZipPath("ppt", target);
      return {
        index: index + 1,
        path: slidePath,
        relId: s.relId,
      };
    });

    return ok({ slides, total: slides.length });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Adds a new slide to the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param layoutId - Optional 1-based layout index to use
 * @returns Result with the path of the new slide
 *
 * @example
 * const result = await addSlide("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.path); // "/slide[3]"
 * }
 */
export async function addSlide(filePath: string, layoutId?: number): Promise<Result<{ path: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Validate the file has required PPTX structure
    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const contentTypesXml = zip.get("[Content_Types].xml")?.toString("utf8") ?? "";

    // Get existing slide IDs and relationship IDs
    const slideIds = getSlideIds(presentationXml);
    const relationships = parseRelationshipEntries(relsXml);

    // Determine the next slide ID
    const existingIds = slideIds.map(s => parseInt(s.id, 10));
    const newSlideId = generateSlideId(existingIds);

    // Determine the next relationship ID
    const existingRelIds = relationships.map(r => r.id);
    const newRelId = generateRelId(existingRelIds);

    // Find slide layout if layoutId provided
    // Layouts are stored in ppt/slideLayouts/ folder, not in presentation.xml.rels
    let layoutTarget = "";
    // Find available layouts by looking at the zip entries
    const layoutEntries = [...zip.keys()].filter(name => name.startsWith("ppt/slideLayouts/slideLayout") && name.endsWith(".xml"));
    const availableLayouts = layoutEntries.map((name, idx) => ({
      index: idx + 1,
      name: name,
      target: name.replace("ppt/", "") // e.g., "slideLayouts/slideLayout1.xml"
    }));

    if (layoutId !== undefined) {
      if (layoutId < 1 || layoutId > availableLayouts.length) {
        return invalidInput(`Layout index ${layoutId} is out of range (1-${availableLayouts.length})`);
      }
      layoutTarget = availableLayouts[layoutId - 1].target;
    } else if (availableLayouts.length > 0) {
      // Use the first available layout if no layout specified
      layoutTarget = availableLayouts[0].target;
    }

    // Create new slide file
    const newSlideIndex = slideIds.length + 1;
    const newSlideEntry = `ppt/slides/slide${newSlideIndex}.xml`;
    const newSlideRelsEntry = `ppt/slides/_rels/slide${newSlideIndex}.xml.rels`;

    // Create minimal slide XML
    const newSlideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;

    // Create slide relationships XML
    const slideRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../${layoutTarget}"/>
</Relationships>`;

    // Add new slide to presentation.xml
    const maxId = Math.max(...existingIds);
    const newSlideIds = [...slideIds, { id: String(parseInt(newSlideId, 10) > maxId ? parseInt(newSlideId, 10) : maxId + 1), relId: newRelId }];
    const updatedPresentationXml = reorderSlideIds(presentationXml, newSlideIds);

    // Add new relationship to presentation.xml.rels
    const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideIndex}.xml"/>`;
    const updatedRelsXml = relsXml.replace(/<\/Relationships>/, `  ${newRelEntry}\n</Relationships>`);

    // Add Content_Type entry if needed
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
    newEntries.push({ name: newSlideRelsEntry, data: Buffer.from(slideRelsXml, "utf8") });

    await writeFile(filePath, createStoredZip(newEntries));

    return ok({ path: `/slide[${newSlideIndex}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Removes a slide from the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param index - 1-based slide index to remove
 *
 * @example
 * const result = await removeSlide("/path/to/presentation.pptx", 2);
 */
export async function removeSlide(filePath: string, index: number): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const contentTypesXml = zip.get("[Content_Types].xml")?.toString("utf8") ?? "";

    const slideIds = getSlideIds(presentationXml);
    const relationships = parseRelationshipEntries(relsXml);

    if (index < 1 || index > slideIds.length) {
      return invalidInput(`Slide index ${index} is out of range (1-${slideIds.length})`);
    }

    const removedSlide = slideIds[index - 1];
    const relToRemove = relationships.find(r => r.id === removedSlide.relId);
    const slideTarget = relToRemove?.target ?? "";

    // Remove slide ID from presentation.xml
    const newSlideIds = slideIds.filter((_, i) => i !== index - 1);
    let updatedPresentationXml = presentationXml;
    // Remove specific sldId
    updatedPresentationXml = updatedPresentationXml.replace(
      new RegExp(`<p:sldId\\b[^>]*\\bid="${removedSlide.id}"[^>]*r:id="${removedSlide.relId}"[^>]*\\/?>`, "g"),
      ""
    );
    // Also try the alternative order
    if (!updatedPresentationXml.includes(`id="${removedSlide.id}"`)) {
      // Already removed, try just the id
      updatedPresentationXml = updatedPresentationXml.replace(
        new RegExp(`<p:sldId\\b[^>]*r:id="${removedSlide.relId}"[^>]*\\/?>`, "g"),
        ""
      );
    }

    // Remove relationship from presentation.xml.rels
    const updatedRelsXml = relsXml.replace(
      new RegExp(`<Relationship\\b[^>]*\\bId="${removedSlide.relId}"[^>]*\\/?>`, "g"),
      ""
    );

    // Build list of entries to remove
    const slideEntryToRemove = normalizeZipPath("ppt", slideTarget);
    const slideRelsEntryToRemove = getRelationshipsEntryName(slideEntryToRemove);
    const slideContentTypeToRemove = `/ppt/slides/${path.posix.basename(slideEntryToRemove)}`;

    // Remove Content_Type entry
    let updatedContentTypes = contentTypesXml.replace(
      new RegExp(`<Override\\b[^>]*PartName="${slideContentTypeToRemove}"[^>]*\\/?>`, "g"),
      ""
    );

    // Build new zip without removed entries
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === "ppt/presentation.xml") {
        newEntries.push({ name, data: Buffer.from(updatedPresentationXml, "utf8") });
      } else if (name === "ppt/_rels/presentation.xml.rels") {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
      } else if (name === "[Content_Types].xml") {
        newEntries.push({ name, data: Buffer.from(updatedContentTypes, "utf8") });
      } else if (name === slideEntryToRemove || name === slideRelsEntryToRemove) {
        // Skip removed entries
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
 * Moves a slide from one position to another.
 *
 * @param filePath - Path to the PPTX file
 * @param fromIndex - Current 1-based position of the slide
 * @param toIndex - New 1-based position for the slide
 *
 * @example
 * // Move slide from position 3 to position 1
 * const result = await moveSlide("/path/to/presentation.pptx", 3, 1);
 */
export async function moveSlide(filePath: string, fromIndex: number, toIndex: number): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const presentationXml = requireEntry(zip, "ppt/presentation.xml");

    const slideIds = getSlideIds(presentationXml);

    if (fromIndex < 1 || fromIndex > slideIds.length) {
      return invalidInput(`Source index ${fromIndex} is out of range (1-${slideIds.length})`);
    }
    if (toIndex < 1 || toIndex > slideIds.length) {
      return invalidInput(`Target index ${toIndex} is out of range (1-${slideIds.length})`);
    }

    // Create new ordering
    const newSlideIds = [...slideIds];
    const [removed] = newSlideIds.splice(fromIndex - 1, 1);
    newSlideIds.splice(toIndex - 1, 0, removed);

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
 * Duplicates an existing slide.
 *
 * @param filePath - Path to the PPTX file
 * @param index - 1-based index of the slide to duplicate
 * @returns Result with the path of the new duplicated slide
 *
 * @example
 * const result = await duplicateSlide("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.path); // "/slide[5]" (or whatever the new position is)
 * }
 */
export async function duplicateSlide(filePath: string, index: number): Promise<Result<{ path: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const contentTypesXml = zip.get("[Content_Types].xml")?.toString("utf8") ?? "";

    const slideIds = getSlideIds(presentationXml);
    const relationships = parseRelationshipEntries(relsXml);

    if (index < 1 || index > slideIds.length) {
      return invalidInput(`Slide index ${index} is out of range (1-${slideIds.length})`);
    }

    const sourceSlide = slideIds[index - 1];
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

    // Determine new slide index (insert after the source slide)
    const newSlideIndex = slideIds.length + 1;
    const newSlideEntry = `ppt/slides/slide${newSlideIndex}.xml`;
    const newSlideRelsEntry = `ppt/slides/_rels/slide${newSlideIndex}.xml.rels`;

    // Update slide ID and rel ID in the duplicated slide
    let newSlideXml = sourceSlideXml;
    // Replace the slide's cNvPr id with the new id (but only for the main shape tree root)
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

    // Add new slide ID to presentation.xml (insert after source slide)
    const insertPosition = index; // Insert after source
    const newSlideIds = [...slideIds];
    newSlideIds.splice(insertPosition, 0, { id: newSlideId, relId: newRelId });
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

    return ok({ path: `/slide[${insertPosition + 1}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}
