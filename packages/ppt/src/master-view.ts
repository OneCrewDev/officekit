/**
 * Master view and layout operations for @officekit/ppt.
 *
 * Provides functions to manage slide masters and their layouts:
 * - getMasters: Get all slide masters in the presentation
 * - getMaster: Get a specific slide master
 * - getMasterLayouts: Get layouts associated with a master
 * - setMasterProperty: Set properties on a master (background, header/footer, etc.)
 * - getLayouts: Get all layouts (independent)
 * - setLayoutForSlide: Assign a layout to a slide
 *
 * Reference: OOXML specification for slideMaster and slideLayout elements
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput } from "./result.js";
import type { Result, SlideMasterModel, SlideLayoutModel } from "./types.js";

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
 */
function requireEntry(zip: Map<string, Buffer>, entryName: string): string {
  const buffer = zip.get(entryName);
  if (!buffer) {
    throw new Error(`OOXML entry '${entryName}' is missing`);
  }
  return buffer.toString("utf8");
}

/**
 * Extracts slide master information from XML.
 */
function parseMasterInfo(masterXml: string, masterPath: string): SlideMasterModel {
  const name = /<p:cSld\b[^>]*name="([^"]*)"/.exec(masterXml)?.[1] ?? "";
  const themeMatch = masterXml.match(/<a:theme\b[^>]*name="([^"]*)"/);
  return {
    path: masterPath,
    name,
    theme: themeMatch?.[1],
  };
}

/**
 * Extracts slide layout information from XML.
 */
function parseLayoutInfo(layoutXml: string, layoutPath: string): SlideLayoutModel {
  const name = /<p:cSld\b[^>]*name="([^"]*)"/.exec(layoutXml)?.[1] ?? "";
  const type = /<p:sldLayout\b[^>]*type="([^"]+)"/.exec(layoutXml)?.[1];
  return {
    path: layoutPath,
    name,
    type,
  };
}

// ============================================================================
// Master Operations
// ============================================================================

/**
 * Gets all slide masters in the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @returns Result with list of slide masters
 *
 * @example
 * const result = await getMasters("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(`Found ${result.data.masters.length} masters`);
 *   for (const master of result.data.masters) {
 *     console.log(`  ${master.name} (${master.path})`);
 *   }
 * }
 */
export async function getMasters(
  filePath: string,
): Promise<Result<{ masters: SlideMasterModel[] }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find all slide masters by looking at presentation.xml.rels
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);

    // Filter to slide masters
    const masterRels = relationships.filter(r => r.type?.endsWith("/slideMaster"));
    const masters: SlideMasterModel[] = [];

    for (const masterRel of masterRels) {
      const masterPath = normalizeZipPath("ppt", masterRel.target);
      const masterXml = zip.get(masterPath)?.toString("utf8") ?? "";
      const masterInfo = parseMasterInfo(masterXml, masterPath);

      // Count associated layouts
      const masterRelsPath = getRelationshipsEntryName(masterPath);
      const masterRelsXml = zip.get(masterRelsPath)?.toString("utf8") ?? "";
      const masterRelationships = parseRelationshipEntries(masterRelsXml);
      const layoutCount = masterRelationships.filter(r => r.type?.endsWith("/slideLayout")).length;

      masters.push({
        ...masterInfo,
        layoutCount,
        shapeCount: (masterXml.match(/<p:sp\b/g) || []).length,
      });
    }

    return ok({ masters });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets a specific slide master by index.
 *
 * @param filePath - Path to the PPTX file
 * @param masterIndex - 1-based master index
 * @returns Result with master information
 *
 * @example
 * const result = await getMaster("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.master);
 * }
 */
export async function getMaster(
  filePath: string,
  masterIndex: number,
): Promise<Result<{ master: SlideMasterModel; layouts: SlideLayoutModel[] }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find all slide masters
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const masterRels = relationships.filter(r => r.type?.endsWith("/slideMaster"));

    if (masterIndex < 1 || masterIndex > masterRels.length) {
      return invalidInput(`Master index ${masterIndex} is out of range (1-${masterRels.length})`);
    }

    const masterRel = masterRels[masterIndex - 1];
    const masterPath = normalizeZipPath("ppt", masterRel.target);
    const masterXml = requireEntry(zip, masterPath);
    const masterInfo = parseMasterInfo(masterXml, masterPath);

    // Get associated layouts
    const masterRelsPath = getRelationshipsEntryName(masterPath);
    const masterRelsXml = zip.get(masterRelsPath)?.toString("utf8") ?? "";
    const masterRelationships = parseRelationshipEntries(masterRelsXml);
    const layoutRels = masterRelationships.filter(r => r.type?.endsWith("/slideLayout"));

    const layouts: SlideLayoutModel[] = [];
    for (const layoutRel of layoutRels) {
      const layoutPath = normalizeZipPath(path.posix.dirname(masterPath), layoutRel.target);
      const layoutXml = zip.get(layoutPath)?.toString("utf8") ?? "";
      layouts.push(parseLayoutInfo(layoutXml, layoutPath));
    }

    return ok({ master: masterInfo, layouts });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets the slide master used by a specific layout.
 *
 * @param filePath - Path to the PPTX file
 * @param layoutIndex - 1-based layout index
 * @returns Result with the master info
 *
 * @example
 * const result = await getLayoutMaster("/path/to/presentation.pptx", 2);
 * if (result.ok) {
 *   console.log(result.data.master);
 * }
 */
export async function getLayoutMaster(
  filePath: string,
  layoutIndex: number,
): Promise<Result<{ master: SlideMasterModel }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find all layouts from presentation.xml.rels
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const layoutRels = relationships.filter(r => r.type?.endsWith("/slideLayout"));

    if (layoutIndex < 1 || layoutIndex > layoutRels.length) {
      return invalidInput(`Layout index ${layoutIndex} is out of range (1-${layoutRels.length})`);
    }

    const layoutRel = layoutRels[layoutIndex - 1];
    const layoutPath = normalizeZipPath("ppt", layoutRel.target);
    const layoutRelsPath = getRelationshipsEntryName(layoutPath);
    const layoutRelsXml = zip.get(layoutRelsPath)?.toString("utf8") ?? "";

    // Find the slideMaster relationship
    const layoutRelationships = parseRelationshipEntries(layoutRelsXml);
    const masterRel = layoutRelationships.find(r => r.type?.endsWith("/slideMaster"));

    if (!masterRel) {
      return err("operation_failed", "Layout does not have an associated slide master");
    }

    const masterPath = normalizeZipPath(path.posix.dirname(layoutPath), masterRel.target);
    const masterXml = requireEntry(zip, masterPath);
    const masterInfo = parseMasterInfo(masterXml, masterPath);

    return ok({ master: masterInfo });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets the slide master used by a specific slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with the master info
 *
 * @example
 * const result = await getSlideMaster("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.master);
 * }
 */
export async function getSlideMaster(
  filePath: string,
  slideIndex: number,
): Promise<Result<{ master: SlideMasterModel }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Get slide IDs and relationships
    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);

    // Extract slide IDs
    const slideIds: Array<{ id: string; relId: string }> = [];
    for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*\bid="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g)) {
      slideIds.push({ id: match[1], relId: match[2] });
    }

    if (slideIndex < 1 || slideIndex > slideIds.length) {
      return invalidInput(`Slide index ${slideIndex} is out of range (1-${slideIds.length})`);
    }

    const slide = slideIds[slideIndex - 1];
    const slideRel = relationships.find(r => r.id === slide.relId);
    const slidePath = normalizeZipPath("ppt", slideRel?.target ?? "");
    const slideRelsPath = getRelationshipsEntryName(slidePath);
    const slideRelsXml = requireEntry(zip, slideRelsPath);
    const slideRelationships = parseRelationshipEntries(slideRelsXml);

    // Find the layout relationship
    const layoutRel = slideRelationships.find(r => r.type?.endsWith("/slideLayout"));
    if (!layoutRel) {
      return err("operation_failed", "Slide does not have an associated layout");
    }

    const layoutPath = normalizeZipPath(path.posix.dirname(slidePath), layoutRel.target);
    const layoutRelsPath = getRelationshipsEntryName(layoutPath);
    const layoutRelsXml = zip.get(layoutRelsPath)?.toString("utf8") ?? "";
    const layoutRelationships = parseRelationshipEntries(layoutRelsXml);

    // Find the slideMaster relationship
    const masterRel = layoutRelationships.find(r => r.type?.endsWith("/slideMaster"));
    if (!masterRel) {
      return err("operation_failed", "Layout does not have an associated slide master");
    }

    const masterPath = normalizeZipPath(path.posix.dirname(layoutPath), masterRel.target);
    const masterXml = requireEntry(zip, masterPath);
    const masterInfo = parseMasterInfo(masterXml, masterPath);

    return ok({ master: masterInfo });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Master Properties
// ============================================================================

/**
 * Master property specification.
 */
export interface MasterPropertySpec {
  /** Background fill - solid color hex (e.g., "FF0000" for red) */
  backgroundColor?: string;
  /** Background fill type: "solid", "gradient", "none" */
  backgroundFill?: "solid" | "gradient" | "none";
  /** Show header on this master */
  showHeader?: boolean;
  /** Show footer on this master */
  showFooter?: boolean;
  /** Show slide number on this master */
  showSlideNumber?: boolean;
  /** Show date/time on this master */
  showDateTime?: boolean;
  /** Footer text */
  footerText?: string;
  /** Header text */
  headerText?: string;
  /** Date/time format */
  dateTimeFormat?: string;
}

/**
 * Sets properties on a slide master.
 *
 * @param filePath - Path to the PPTX file
 * @param masterIndex - 1-based master index
 * @param props - Properties to set
 * @returns Result indicating success
 *
 * @example
 * // Set master background color
 * const result = await setMasterProperty("/path/to/presentation.pptx", 1, {
 *   backgroundColor: "4472C4",
 *   backgroundFill: "solid"
 * });
 *
 * // Enable header/footer elements
 * const result = await setMasterProperty("/path/to/presentation.pptx", 1, {
 *   showHeader: true,
 *   showFooter: true,
 *   showSlideNumber: true,
 *   footerText: "Confidential"
 * });
 */
export async function setMasterProperty(
  filePath: string,
  masterIndex: number,
  props: MasterPropertySpec,
): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find the master
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const masterRels = relationships.filter(r => r.type?.endsWith("/slideMaster"));

    if (masterIndex < 1 || masterIndex > masterRels.length) {
      return invalidInput(`Master index ${masterIndex} is out of range (1-${masterRels.length})`);
    }

    const masterRel = masterRels[masterIndex - 1];
    const masterPath = normalizeZipPath("ppt", masterRel.target);
    let masterXml = requireEntry(zip, masterPath);

    // Update background if specified
    if (props.backgroundFill !== undefined || props.backgroundColor !== undefined) {
      masterXml = updateMasterBackground(masterXml, props);
    }

    // Update header/footer if specified
    if (props.showHeader !== undefined || props.showFooter !== undefined ||
        props.showSlideNumber !== undefined || props.showDateTime !== undefined ||
        props.footerText !== undefined || props.headerText !== undefined ||
        props.dateTimeFormat !== undefined) {
      masterXml = updateMasterHeaderFooter(masterXml, props);
    }

    // Build new zip with updated master
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === masterPath) {
        newEntries.push({ name, data: Buffer.from(masterXml, "utf8") });
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
 * Updates the background in master XML.
 */
function updateMasterBackground(masterXml: string, props: MasterPropertySpec): string {
  // Remove existing background
  let result = masterXml.replace(/<p:bg\b[\s\S]*?<\/p:bg>/gi, "");

  if (props.backgroundFill === "none") {
    // Add empty background to indicate no fill
    result = result.replace(
      /(<p:cSld\b[^>]*>)/,
      `$1<p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg>`
    );
    return result;
  }

  // Build background properties
  let bgPrXml = "<p:bgPr>";
  if (props.backgroundColor) {
    bgPrXml += `<a:solidFill><a:srgbClr val="${props.backgroundColor}"/></a:solidFill>`;
  } else {
    bgPrXml += "<a:noFill/>";
  }
  bgPrXml += "</p:bgPr>";

  // Insert background before cSld content
  result = result.replace(
    /(<p:cSld\b[^>]*>)/,
    `$1<p:bg>${bgPrXml}</p:bg>`
  );

  return result;
}

/**
 * Updates header/footer properties in master XML.
 */
function updateMasterHeaderFooter(masterXml: string, props: MasterPropertySpec): string {
  // Check if hfPr exists
  let hasHfPr = /<p:hfPr\b/i.test(masterXml);

  if (!hasHfPr) {
    // Add hfPr after bgPr or cSld properties
    const insertPoint = masterXml.match(/<\/p:bgPr>/i) || masterXml.match(/<p:spTree>/i);
    if (insertPoint && insertPoint.index !== undefined) {
      const insertIndex = insertPoint.index + insertPoint[0].length;
      masterXml = masterXml.slice(0, insertIndex) + "<p:hfPr/>" + masterXml.slice(insertIndex);
    }
  }

  // Update individual properties
  let result = masterXml;

  // Show/hide flags
  if (props.showHeader !== undefined) {
    result = result.replace(
      /<p:hfPr\b[^>]*>/i,
      (match) => match.replace(/<p:showHeader\b[^>]*\/>/gi, "").replace(/>/, `><p:showHeader val="${props.showHeader ? "1" : "0"}"/>`)
    );
    if (!/<p:showHeader val="/.test(result)) {
      result = result.replace(/<p:hfPr\b[^>]*>/i, `$&<p:showHeader val="${props.showHeader ? "1" : "0"}"/>`);
    }
  }

  if (props.showFooter !== undefined) {
    result = result.replace(/<p:showFooter val="[^"]*"\/>/gi, "");
    result = result.replace(/<p:hfPr\b[^>]*>/i, `$&<p:showFooter val="${props.showFooter ? "1" : "0"}"/>`);
  }

  if (props.showSlideNumber !== undefined) {
    result = result.replace(/<p:showSlideNum val="[^"]*"\/>/gi, "");
    result = result.replace(/<p:hfPr\b[^>]*>/i, `$&<p:showSlideNum val="${props.showSlideNumber ? "1" : "0"}"/>`);
  }

  if (props.showDateTime !== undefined) {
    result = result.replace(/<p:showDateTime val="[^"]*"\/>/gi, "");
    result = result.replace(/<p:hfPr\b[^>]*>/i, `$&<p:showDateTime val="${props.showDateTime ? "1" : "0"}"/>`);
  }

  // Text values
  if (props.footerText !== undefined) {
    result = result.replace(/<p:dtFmt\b[^>]*\/>/gi, "");
    result = result.replace(/<p:hfPr\b[^>]*>/i, `$&<p:footerText val="${props.footerText.replace(/"/g, "&quot;")}"/>`);
  }

  if (props.headerText !== undefined) {
    result = result.replace(/<p:headerText\b[^>]*\/>/gi, "");
    result = result.replace(/<p:hfPr\b[^>]*>/i, `$&<p:headerText val="${props.headerText.replace(/"/g, "&quot;")}"/>`);
  }

  if (props.dateTimeFormat !== undefined) {
    result = result.replace(/<p:dtFmt\b[^>]*\/>/gi, "");
    result = result.replace(/<p:hfPr\b[^>]*>/i, `$&<p:dtFmt val="${props.dateTimeFormat.replace(/"/g, "&quot;")}"/>`);
  }

  return result;
}

// ============================================================================
// Layout Operations
// ============================================================================

/**
 * Gets all slide layouts in the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @returns Result with list of layouts
 *
 * @example
 * const result = await getAllLayouts("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(`Found ${result.data.layouts.length} layouts`);
 * }
 */
export async function getAllLayouts(
  filePath: string,
): Promise<Result<{ layouts: SlideLayoutModel[] }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find all layouts from presentation.xml.rels
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const layoutRels = relationships.filter(r => r.type?.endsWith("/slideLayout"));

    const layouts: SlideLayoutModel[] = [];
    for (const layoutRel of layoutRels) {
      const layoutPath = normalizeZipPath("ppt", layoutRel.target);
      const layoutXml = zip.get(layoutPath)?.toString("utf8") ?? "";
      layouts.push(parseLayoutInfo(layoutXml, layoutPath));
    }

    return ok({ layouts });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets layouts associated with a specific master.
 *
 * @param filePath - Path to the PPTX file
 * @param masterIndex - 1-based master index
 * @returns Result with list of layouts
 *
 * @example
 * const result = await getMasterLayouts("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(`Found ${result.data.layouts.length} layouts for this master`);
 * }
 */
export async function getMasterLayouts(
  filePath: string,
  masterIndex: number,
): Promise<Result<{ layouts: SlideLayoutModel[] }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find the master
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const masterRels = relationships.filter(r => r.type?.endsWith("/slideMaster"));

    if (masterIndex < 1 || masterIndex > masterRels.length) {
      return invalidInput(`Master index ${masterIndex} is out of range (1-${masterRels.length})`);
    }

    const masterRel = masterRels[masterIndex - 1];
    const masterPath = normalizeZipPath("ppt", masterRel.target);

    // Get master's relationships
    const masterRelsPath = getRelationshipsEntryName(masterPath);
    const masterRelsXml = zip.get(masterRelsPath)?.toString("utf8") ?? "";
    const masterRelationships = parseRelationshipEntries(masterRelsXml);
    const layoutRels = masterRelationships.filter(r => r.type?.endsWith("/slideLayout"));

    const layouts: SlideLayoutModel[] = [];
    for (const layoutRel of layoutRels) {
      const layoutPath = normalizeZipPath(path.posix.dirname(masterPath), layoutRel.target);
      const layoutXml = zip.get(layoutPath)?.toString("utf8") ?? "";
      layouts.push(parseLayoutInfo(layoutXml, layoutPath));
    }

    return ok({ layouts });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Sets the slide layout for a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param layoutIndex - 1-based layout index to set
 * @returns Result indicating success
 *
 * @example
 * // Set slide 1 to use layout 2
 * const result = await setSlideLayoutByIndex("/path/to/presentation.pptx", 1, 2);
 */
export async function setSlideLayoutByIndex(
  filePath: string,
  slideIndex: number,
  layoutIndex: number,
): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Get presentation relationships
    const presentationXml = requireEntry(zip, "ppt/presentation.xml");
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);

    // Extract slide IDs
    const slideIds: Array<{ id: string; relId: string }> = [];
    for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*\bid="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g)) {
      slideIds.push({ id: match[1], relId: match[2] });
    }

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

    // Get the slide
    const slide = slideIds[slideIndex - 1];
    const slideRel = relationships.find(r => r.id === slide.relId);
    const slidePath = normalizeZipPath("ppt", slideRel?.target ?? "");
    const slideRelsPath = getRelationshipsEntryName(slidePath);
    const slideRelsXml = requireEntry(zip, slideRelsPath);

    // Find the slideLayout relationship
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
 * Gets the name and type of a layout.
 *
 * @param filePath - Path to the PPTX file
 * @param layoutIndex - 1-based layout index
 * @returns Result with layout information
 *
 * @example
 * const result = await getLayoutInfo("/path/to/presentation.pptx", 2);
 * if (result.ok) {
 *   console.log(`Layout: ${result.data.name}, Type: ${result.data.type}`);
 * }
 */
export async function getLayoutInfo(
  filePath: string,
  layoutIndex: number,
): Promise<Result<SlideLayoutModel>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find all layouts
    const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
    const relationships = parseRelationshipEntries(relsXml);
    const layoutRels = relationships.filter(r => r.type?.endsWith("/slideLayout"));

    if (layoutIndex < 1 || layoutIndex > layoutRels.length) {
      return invalidInput(`Layout index ${layoutIndex} is out of range (1-${layoutRels.length})`);
    }

    const layoutRel = layoutRels[layoutIndex - 1];
    const layoutPath = normalizeZipPath("ppt", layoutRel.target);
    const layoutXml = requireEntry(zip, layoutPath);

    return ok(parseLayoutInfo(layoutXml, layoutPath));
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}
