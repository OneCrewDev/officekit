/**
 * Connector operations for @officekit/ppt.
 *
 * Connectors are lines that connect shapes and can auto-follow shape movements.
 * Connectors use `<p:cxnSp>` elements in PPTX with `<p:stCxn>` (start connection)
 * and `<p:endCxn>` (end connection) child elements.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput, notFound } from "./result.js";
import type { Result, ConnectorModel } from "./types.js";
import { getSlideIndex } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Connector types available in PowerPoint.
 */
export type ConnectorType = "straight" | "elbow" | "curved" | "arrow";

/**
 * Arrow style options for connector endpoints.
 */
export type ArrowStyle = "none" | "arrow" | "triangle" | "diamond" | "oval";

/**
 * Connector style options.
 */
export interface ConnectorStyle {
  /** Line color (hex format, e.g., "FF0000" for red) */
  color?: string;
  /** Line width in EMUs (1 inch = 914400 EMUs) */
  width?: number;
  /** Begin arrow style */
  beginArrow?: ArrowStyle;
  /** End arrow style */
  endArrow?: ArrowStyle;
  /** Optional label text */
  label?: string;
}

/**
 * Options for adding a connector.
 */
export interface AddConnectorOptions {
  /** Line color (hex format, e.g., "FF0000" for red) */
  color?: string;
  /** Line width in EMUs (default: 12700 = 1pt) */
  width?: number;
  /** Begin arrow style */
  beginArrow?: ArrowStyle;
  /** End arrow style */
  endArrow?: ArrowStyle;
  /** Optional label text */
  label?: string;
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
 * Maps connector type string to OOXML preset geometry.
 */
function connectorTypeToPreset(connectorType: ConnectorType): string {
  switch (connectorType) {
    case "straight":
    case "arrow":
      return "straightConnector1";
    case "elbow":
      return "elbowConnector1";
    case "curved":
      return "curvedConnector1";
    default:
      return "straightConnector1";
  }
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
 * Creates XML for a connector.
 */
function createConnectorXml(
  id: number,
  connectorIndex: number,
  connectorType: ConnectorType,
  startShape: string,
  endShape: string,
  options?: AddConnectorOptions,
): string {
  const preset = connectorTypeToPreset(connectorType);
  const color = options?.color || "000000";
  const width = options?.width || 12700;
  const beginArrow = options?.beginArrow || "none";
  const endArrow = options?.endArrow || (connectorType === "arrow" ? "arrow" : "none");

  // Build arrow XML if needed
  const beginArrowXml = beginArrow !== "none" ? `<a:tailEnd type="${beginArrow}"/>` : "";
  const endArrowXml = endArrow !== "none" ? `<a:tailEnd type="${endArrow}"/>` : "";

  // Build label XML if provided
  const labelXml = options?.label
    ? `<p:spPr>
        <a:prstGeom prst="${preset}">
          <a:avLst/>
        </a:prstGeom>
        <a:ln w="${width}">
          <a:solidFill>
            <a:srgbClr val="${color}"/>
          </a:solidFill>
          ${beginArrowXml}
          ${endArrowXml}
        </a:ln>
        <p:spCxn>
          <p:adj>0</p:adj>
        </p:spCxn>
        <p:stCxn id="${startShape}" idx="0"/>
        <p:endCxn id="${endShape}" idx="0"/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>${escapeXml(options.label)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:spPr>`
    : `<p:spPr>
        <a:prstGeom prst="${preset}">
          <a:avLst/>
        </a:prstGeom>
        <a:ln w="${width}">
          <a:solidFill>
            <a:srgbClr val="${color}"/>
          </a:solidFill>
          ${beginArrowXml}
          ${endArrowXml}
        </a:ln>
        <p:spCxn>
          <p:adj>0</p:adj>
        </p:spCxn>
        <p:stCxn id="${startShape}" idx="0"/>
        <p:endCxn id="${endShape}" idx="0"/>
      </p:spPr>`;

  return `    <p:cxnSp>
      <p:nvCxnSpPr>
        <p:cNvPr id="${id}" name="${capitalize(connectorType)} Connector ${connectorIndex}"/>
        <p:cNvCxnSpPr/>
        <p:nvPr/>
      </p:nvCxnSpPr>
      ${labelXml}
    </p:cxnSp>`;
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
 * Capitalizes first letter.
 */
function capitalize(str: string): string {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

// ============================================================================
// Connector Operations
// ============================================================================

/**
 * Adds a connector between two shapes on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param connectorType - Type of connector: "straight", "elbow", "curved", "arrow"
 * @param startShape - Path to starting shape (e.g., "/slide[1]/shape[1]")
 * @param endShape - Path to ending shape (e.g., "/slide[1]/shape[2]")
 * @param options - Optional connector style options (color, width, arrows, label)
 *
 * @example
 * // Add a straight connector between two shapes
 * const result = await addConnector(
 *   "/path/to/presentation.pptx",
 *   1,
 *   "straight",
 *   "/slide[1]/shape[1]",
 *   "/slide[1]/shape[2]"
 * );
 *
 * // Add an arrow connector with custom styling
 * const result = await addConnector(
 *   "/path/to/presentation.pptx",
 *   1,
 *   "arrow",
 *   "/slide[1]/shape[1]",
 *   "/slide[1]/shape[2]",
 *   { color: "FF0000", width: 25400, endArrow: "arrow", label: "Flow" }
 * );
 */
export async function addConnector(
  filePath: string,
  slideIndex: number,
  connectorType: ConnectorType,
  startShape: string,
  endShape: string,
  options?: AddConnectorOptions,
): Promise<Result<{ path: string }>> {
  try {
    // Validate connector type
    const validTypes: ConnectorType[] = ["straight", "elbow", "curved", "arrow"];
    if (!validTypes.includes(connectorType)) {
      return invalidInput(`Invalid connector type: ${connectorType}. Valid types: ${validTypes.join(", ")}`);
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);

    // Count existing connectors to determine new connector index
    const connectorPattern = /<p:cxnSp[\s\S]*?<\/p:cxnSp>/g;
    const existingConnectors = slideXml.match(connectorPattern) || [];
    const newConnectorIndex = existingConnectors.length + 1;

    // Generate unique shape ID
    const newId = generateUniqueId(slideXml);

    // Extract shape IDs from paths
    const startShapeId = extractShapeIdFromPath(startShape);
    const endShapeId = extractShapeIdFromPath(endShape);

    if (!startShapeId) {
      return invalidInput(`Invalid start shape path: ${startShape}`);
    }
    if (!endShapeId) {
      return invalidInput(`Invalid end shape path: ${endShape}`);
    }

    // Create connector XML
    const newConnectorXml = createConnectorXml(
      newId,
      newConnectorIndex,
      connectorType,
      startShapeId,
      endShapeId,
      options
    );

    // Insert connector before </p:spTree>
    const updatedSlideXml = slideXml.replace(
      /<\/p:spTree>/,
      `${newConnectorXml}\n  </p:spTree>`
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
    return ok({ path: `/slide[${slideIndex}]/connector[${newConnectorIndex}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Extracts the shape ID from a shape path.
 * For example: "/slide[1]/shape[1]" -> "1"
 */
function extractShapeIdFromPath(shapePath: string): string | null {
  // Extract shape index from path
  const match = shapePath.match(/\/shape\[(\d+)\]/i);
  return match ? match[1] : null;
}

/**
 * Gets all connectors on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 *
 * @example
 * const result = await getConnectors("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.connectors);
 *   // [{ path: "/slide[1]/connector[1]", name: "Straight Connector 1" }]
 * }
 */
export async function getConnectors(
  filePath: string,
  slideIndex: number,
): Promise<Result<{ connectors: ConnectorModel[] }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);

    // Find all connectors
    const connectorPattern = /<p:cxnSp[\s\S]*?<\/p:cxnSp>/g;
    const connectorMatches = slideXml.match(connectorPattern) || [];

    const connectors: ConnectorModel[] = connectorMatches.map((connectorXml, index) => {
      // Extract connector name
      const nameMatch = connectorXml.match(/<p:cNvPr[^>]*name="([^"]+)"/);
      const name = nameMatch ? nameMatch[1] : undefined;

      return {
        path: `/slide[${slideIndex}]/connector[${index + 1}]`,
        name,
      };
    });

    return ok({ connectors });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Sets the endpoints of a connector (updates which shapes it connects).
 *
 * @param filePath - Path to the PPTX file
 * @param path - PPT path to the connector (e.g., "/slide[1]/connector[1]")
 * @param startShape - New starting shape path (e.g., "/slide[1]/shape[1]")
 * @param endShape - New ending shape path (e.g., "/slide[1]/shape[3]")
 *
 * @example
 * const result = await setConnectorEndpoints(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/connector[1]",
 *   "/slide[1]/shape[1]",
 *   "/slide[1]/shape[3]"
 * );
 */
export async function setConnectorEndpoints(
  filePath: string,
  pptPath: string,
  startShape: string,
  endShape: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setConnectorEndpoints requires a slide path");
    }

    const connectorIndex = extractConnectorIndex(pptPath);
    if (connectorIndex === null) {
      return invalidInput("Invalid connector path");
    }

    const startShapeId = extractShapeIdFromPath(startShape);
    const endShapeId = extractShapeIdFromPath(endShape);

    if (!startShapeId) {
      return invalidInput(`Invalid start shape path: ${startShape}`);
    }
    if (!endShapeId) {
      return invalidInput(`Invalid end shape path: ${endShape}`);
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);

    // Find and update the connector
    const connectorPattern = /<p:cxnSp[\s\S]*?<\/p:cxnSp>/g;
    let updatedSlideXml = slideXml;
    let count = 0;

    updatedSlideXml = slideXml.replace(connectorPattern, (match) => {
      count++;
      if (count === connectorIndex) {
        // Update start connection
        let updated = match.replace(/<p:stCxn[^>]*\/>/, `<p:stCxn id="${startShapeId}" idx="0"/>`);
        // Update end connection
        updated = updated.replace(/<p:endCxn[^>]*\/>/, `<p:endCxn id="${endShapeId}" idx="0"/>`);
        return updated;
      }
      return match;
    });

    // Check if connector was found and updated
    const newConnectorPattern = /<p:cxnSp[\s\S]*?<\/p:cxnSp>/g;
    const newConnectors = updatedSlideXml.match(newConnectorPattern) || [];
    if (connectorIndex > newConnectors.length) {
      return notFound("Connector", pptPath);
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
 * Extracts the connector index from a path.
 */
function extractConnectorIndex(pptPath: string): number | null {
  const match = pptPath.match(/\/connector\[(\d+)\]/i);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * Removes a connector from a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param path - PPT path to the connector (e.g., "/slide[1]/connector[1]")
 *
 * @example
 * const result = await removeConnector("/path/to/presentation.pptx", "/slide[1]/connector[1]");
 */
export async function removeConnector(
  filePath: string,
  pptPath: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("removeConnector requires a slide path");
    }

    const connectorIndex = extractConnectorIndex(pptPath);
    if (connectorIndex === null) {
      return invalidInput("Invalid connector path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);

    // Find and remove the connector
    const connectorPattern = /<p:cxnSp[\s\S]*?<\/p:cxnSp>/g;
    let updatedSlideXml = slideXml;
    let count = 0;
    let found = false;

    updatedSlideXml = slideXml.replace(connectorPattern, (match) => {
      count++;
      if (count === connectorIndex) {
        found = true;
        return ""; // Remove the connector
      }
      return match;
    });

    if (!found) {
      return notFound("Connector", pptPath);
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
 * Sets the style of a connector (line color, width, arrows).
 *
 * @param filePath - Path to the PPTX file
 * @param path - PPT path to the connector (e.g., "/slide[1]/connector[1]")
 * @param style - Style options to apply
 *
 * @example
 * // Change connector color and add arrow
 * const result = await setConnectorStyle(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/connector[1]",
 *   { color: "00FF00", width: 19050, endArrow: "arrow" }
 * );
 */
export async function setConnectorStyle(
  filePath: string,
  pptPath: string,
  style: ConnectorStyle,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setConnectorStyle requires a slide path");
    }

    const connectorIndex = extractConnectorIndex(pptPath);
    if (connectorIndex === null) {
      return invalidInput("Invalid connector path");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult;
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry);

    // Find and update the connector
    const connectorPattern = /<p:cxnSp[\s\S]*?<\/p:cxnSp>/g;
    let updatedSlideXml = slideXml;
    let count = 0;
    let found = false;

    updatedSlideXml = slideXml.replace(connectorPattern, (match) => {
      count++;
      if (count === connectorIndex) {
        found = true;
        let updated = match;

        // Update line color
        if (style.color) {
          updated = updated.replace(
            /<a:srgbClr val="[^"]*"\/>/,
            `<a:srgbClr val="${style.color}"/>`
          );
          // If no solidFill exists, add one
          if (!updated.includes("<a:solidFill>")) {
            updated = updated.replace(
              /<a:ln w="[^"]*">/,
              `<a:ln w="${style.width || 12700}"><a:solidFill><a:srgbClr val="${style.color}"/></a:solidFill>`
            );
          }
        }

        // Update line width
        if (style.width) {
          updated = updated.replace(/<a:ln w="[^"]*"/, `<a:ln w="${style.width}"`);
        }

        // Update begin arrow
        if (style.beginArrow !== undefined) {
          if (style.beginArrow === "none") {
            updated = updated.replace(/<a:tailEnd[^>]*\/>/, "");
          } else {
            if (updated.includes("<a:tailEnd")) {
              updated = updated.replace(/<a:tailEnd[^>]*\/>/, `<a:tailEnd type="${style.beginArrow}"/>`);
            } else {
              updated = updated.replace(/<a:ln w="[^"]*">/, `<a:ln w="${style.width || 12700}"><a:tailEnd type="${style.beginArrow}"/>`);
            }
          }
        }

        // Update end arrow
        if (style.endArrow !== undefined) {
          if (style.endArrow === "none") {
            // Remove existing tailEnd
            updated = updated.replace(/<a:tailEnd[^>]*\/>/g, "");
          } else {
            if (updated.includes("<a:tailEnd")) {
              updated = updated.replace(/<a:tailEnd[^>]*\/>/, `<a:tailEnd type="${style.endArrow}"/>`);
            } else {
              // Add new tailEnd before closing </a:ln>
              updated = updated.replace(/<\/a:ln>/, `<a:tailEnd type="${style.endArrow}"/></a:ln>`);
            }
          }
        }

        // Update label
        if (style.label !== undefined) {
          if (updated.includes("<p:txBody>")) {
            updated = updated.replace(
              /<p:txBody>[\s\S]*?<\/p:txBody>/,
              `<p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>${escapeXml(style.label)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>`
            );
          }
        }

        return updated;
      }
      return match;
    });

    if (!found) {
      return notFound("Connector", pptPath);
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
