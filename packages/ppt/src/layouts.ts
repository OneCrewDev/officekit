/**
 * Layout inheritance operations for @officekit/ppt.
 *
 * Slides inherit from slide layouts. This module provides functions to
 * get and set the layout a slide is using.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput } from "./result.js";
import type { Result } from "./types.js";

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
 * Extracts layout information from layout XML.
 */
function parseLayoutInfo(layoutXml: string): { name: string; type?: string } {
  const name = /<p:cSld\b[^>]*name="([^"]*)"/.exec(layoutXml)?.[1] ?? "";
  const type = /<p:sldLayout\b[^>]*type="([^"]+)"/.exec(layoutXml)?.[1];
  return { name, type };
}

// ============================================================================
// Layout Operations
// ============================================================================

/**
 * Gets the layout information for a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with layout information
 *
 * @example
 * const result = await getSlideLayout("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.layoutName); // "Title Slide"
 *   console.log(result.data.layoutType); // "title"
 * }
 */
export async function getSlideLayout(
  filePath: string,
  slideIndex: number,
): Promise<Result<{ layoutName: string; layoutType?: string; layoutIndex: number }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const presentationXml = zip.get("ppt/presentation.xml")?.toString("utf8") ?? "";
    const relsXml = zip.get("ppt/_rels/presentation.xml.rels")?.toString("utf8") ?? "";

    const slideIds = getSlideIds(presentationXml);
    const relationships = parseRelationshipEntries(relsXml);

    if (slideIndex < 1 || slideIndex > slideIds.length) {
      return invalidInput(`Slide index ${slideIndex} is out of range (1-${slideIds.length})`);
    }

    const slide = slideIds[slideIndex - 1];
    const slideRel = relationships.find(r => r.id === slide.relId);
    const slidePath = normalizeZipPath("ppt", slideRel?.target ?? "");
    const slideRelsPath = getRelationshipsEntryName(slidePath);
    const slideRelsXml = zip.get(slideRelsPath)?.toString("utf8") ?? "";

    // Find the layout relationship
    const slideRels = parseRelationshipEntries(slideRelsXml);
    const layoutRel = slideRels.find(r => r.type?.endsWith("/slideLayout"));
    if (!layoutRel) {
      return ok({ layoutName: "", layoutType: undefined, layoutIndex: 0 });
    }

    const layoutPath = normalizeZipPath(path.posix.dirname(slidePath), layoutRel.target);
    const layoutXml = zip.get(layoutPath)?.toString("utf8") ?? "";
    const { name, type } = parseLayoutInfo(layoutXml);

    // Find layout index
    const layoutRels = relationships.filter(r => r.type?.endsWith("/slideLayout"));
    const layoutIndex = layoutRels.findIndex(r => {
      const relPath = normalizeZipPath("ppt", r.target);
      return relPath === layoutPath || relPath.endsWith(layoutRel.target);
    }) + 1;

    return ok({
      layoutName: name || layoutRel.target,
      layoutType: type,
      layoutIndex: layoutIndex || 1,
    });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Sets the layout for a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param layoutIndex - 1-based layout index to set
 *
 * @example
 * const result = await setSlideLayout("/path/to/presentation.pptx", 1, 2);
 */
export async function setSlideLayout(
  filePath: string,
  slideIndex: number,
  layoutIndex: number,
): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const presentationXml = zip.get("ppt/presentation.xml")?.toString("utf8") ?? "";
    const relsXml = zip.get("ppt/_rels/presentation.xml.rels")?.toString("utf8") ?? "";

    const slideIds = getSlideIds(presentationXml);
    const relationships = parseRelationshipEntries(relsXml);

    if (slideIndex < 1 || slideIndex > slideIds.length) {
      return invalidInput(`Slide index ${slideIndex} is out of range (1-${slideIds.length})`);
    }

    // Find available layouts
    const layoutRels = relationships.filter(r => r.type?.endsWith("/slideLayout"));
    if (layoutIndex < 1 || layoutIndex > layoutRels.length) {
      return invalidInput(`Layout index ${layoutIndex} is out of range (1-${layoutRels.length})`);
    }

    const targetLayout = layoutRels[layoutIndex - 1];
    const targetLayoutPath = normalizeZipPath("ppt", targetLayout.target);

    const slide = slideIds[slideIndex - 1];
    const slideRel = relationships.find(r => r.id === slide.relId);
    const slidePath = normalizeZipPath("ppt", slideRel?.target ?? "");
    const slideRelsPath = getRelationshipsEntryName(slidePath);
    const slideRelsXml = zip.get(slideRelsPath)?.toString("utf8") ?? "";

    // Find the slide's slideLayout relationship and update it
    const slideRels = parseRelationshipEntries(slideRelsXml);
    const layoutRelIndex = slideRels.findIndex(r => r.type?.endsWith("/slideLayout"));

    let updatedSlideRelsXml: string;
    if (layoutRelIndex >= 0) {
      // Update existing layout relationship
      const oldLayoutRel = slideRels[layoutRelIndex];
      const newTargetRelative = path.posix.relative(path.posix.dirname(slidePath), targetLayoutPath);
      updatedSlideRelsXml = slideRelsXml.replace(
        new RegExp(`<Relationship\\b[^>]*\\bId="${oldLayoutRel.id}"[^>]*\\/?>`, "g"),
        `<Relationship Id="${oldLayoutRel.id}" Type="${oldLayoutRel.type}" Target="../${newTargetRelative}"/>`
      );
    } else {
      // Add new layout relationship
      const newRelId = `rId${slideRels.length + 1}`;
      const newTargetRelative = path.posix.relative(path.posix.dirname(slidePath), targetLayoutPath);
      const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../${newTargetRelative}"/>`;
      updatedSlideRelsXml = slideRelsXml.replace(/<\/Relationships>/, `  ${newRelEntry}\n</Relationships>`);
    }

    // Build new zip with updated slide rels
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideRelsPath) {
        newEntries.push({ name, data: Buffer.from(updatedSlideRelsXml, "utf8") });
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
 * Gets all available layouts in the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @returns Result with list of layouts
 *
 * @example
 * const result = await getLayouts("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(result.data.layouts);
 * }
 */
export async function getLayouts(
  filePath: string,
): Promise<Result<{ layouts: Array<{ index: number; name: string; type?: string }> }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find available layouts by looking at the zip entries
    const layoutEntries = [...zip.keys()].filter(name => name.startsWith("ppt/slideLayouts/slideLayout") && name.endsWith(".xml"));

    const layouts: Array<{ index: number; name: string; type?: string }> = [];

    for (let i = 0; i < layoutEntries.length; i++) {
      const layoutEntry = layoutEntries[i];
      const layoutXml = zip.get(layoutEntry)?.toString("utf8") ?? "";
      const { name, type } = parseLayoutInfo(layoutXml);
      layouts.push({
        index: i + 1,
        name: name || layoutEntry,
        type,
      });
    }

    return ok({ layouts });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}
