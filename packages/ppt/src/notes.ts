/**
 * Notes management operations for @officekit/ppt.
 *
 * Provides functions to get, set, and remove speaker notes from slides.
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
 * Extracts text from a notes slide.
 */
function extractNotesText(notesXml: string): string {
  // Find the body placeholder (idx=1) and extract text from runs
  const textRuns: string[] = [];
  // Match text runs in notes
  for (const match of notesXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) {
    textRuns.push(match[1]);
  }
  return textRuns.join("");
}

/**
 * Creates notes slide XML with the given text.
 */
function createNotesSlideXml(text: string): string {
  const paragraphs = text.split("\n").map(line =>
    `        <a:p>
          <a:r>
            <a:rPr lang="en-US"/>
            <a:t>${escapeXml(line)}</a:t>
          </a:r>
        </a:p>`
  ).join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Slide Image Placeholder 1"/>
          <p:cNvSpPr/>
          <p:nvPr>
            <p:ph type="slideImage" idx="0"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Notes Placeholder 2"/>
          <p:cNvSpPr/>
          <p:nvPr>
            <p:ph type="body" idx="1"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
${paragraphs}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:notes>`;
}

/**
 * Updates text in an existing notes slide XML.
 */
function updateNotesText(notesXml: string, text: string): string {
  // Find the body placeholder shape and update its text
  const paragraphs = text.split("\n").map(line =>
    `          <a:r>
            <a:rPr lang="en-US"/>
            <a:t>${escapeXml(line)}</a:t>
          </a:r>`
  ).join("\n");

  // Replace the content inside the txBody element within the body placeholder
  // This regex finds the txBody inside the shape with ph type="body"
  let result = notesXml;

  // Find and replace the paragraph content
  const txBodyPattern = /(<p:sp>[\s\S]*?<p:ph type="body"[\s\S]*?<\/p:nvSpPr>[\s\S]*?<p:txBody>)([\s\S]*?)(<\/p:txBody>[\s\S]*?<\/p:sp>)/;
  if (txBodyPattern.test(result)) {
    result = result.replace(txBodyPattern, (match, open, _oldContent, close) => {
      const newParagraphs = paragraphs.split("\n").map(p => `          <a:p>${p.replace(/<a:r>/g, "").replace(/<\/a:r>/g, "")}</a:p>`).join("\n");
      return `${open}\n            ${paragraphs}\n        ${close}`;
    });
  }

  return result;
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
// Notes Operations
// ============================================================================

/**
 * Gets the notes text for a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with notes text (empty string if no notes)
 *
 * @example
 * const result = await getNotes("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(result.data.text);
 * }
 */
export async function getNotes(filePath: string, slideIndex: number): Promise<Result<{ text: string }>> {
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

    // Find the notes relationship
    const slideRels = parseRelationshipEntries(slideRelsXml);
    const notesRel = slideRels.find(r => r.type?.endsWith("/notesSlide"));
    if (!notesRel) {
      return ok({ text: "" });
    }

    const notesPath = normalizeZipPath(path.posix.dirname(slidePath), notesRel.target);
    const notesXml = zip.get(notesPath)?.toString("utf8") ?? "";
    const text = extractNotesText(notesXml);

    return ok({ text });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Sets the notes text for a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param text - Notes text to set
 *
 * @example
 * const result = await setNotes("/path/to/presentation.pptx", 1, "Remember to explain the budget");
 */
export async function setNotes(filePath: string, slideIndex: number, text: string): Promise<Result<void>> {
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

    // Get existing relationship IDs to avoid conflicts
    const existingRelIds = relationships.map(r => r.id);

    // Find the notes relationship
    const slideRels = parseRelationshipEntries(slideRelsXml);
    const notesRelIndex = slideRels.findIndex(r => r.type?.endsWith("/notesSlide"));

    let notesPath: string;
    let updatedSlideRelsXml = slideRelsXml;

    if (notesRelIndex >= 0) {
      // Update existing notes
      notesPath = normalizeZipPath(path.posix.dirname(slidePath), slideRels[notesRelIndex].target);
    } else {
      // Create new notes slide
      const notesRelId = generateRelId(existingRelIds);
      const notesIndex = slideIndex; // Use slide index for naming consistency
      // notesPath is used for zip entries (with ppt/ prefix)
      notesPath = `ppt/notesSlides/notesSlide${notesIndex}.xml`;
      // But Target in relationship should be relative to the slide folder (without ppt/ prefix and leading /)
      const notesTarget = `../notesSlides/notesSlide${notesIndex}.xml`;
      const newRelEntry = `<Relationship Id="${notesRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="${notesTarget}"/>`;
      updatedSlideRelsXml = slideRelsXml.replace(/<\/Relationships>/, `  ${newRelEntry}\n</Relationships>`);

      // Add Content_Type entry (PartName is absolute from zip root)
      const contentTypesXml = zip.get("[Content_Types].xml")?.toString("utf8") ?? "";
      const notesContentType = `<Override PartName="/${notesPath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`;
      const updatedContentTypes = contentTypesXml.replace(/<\/Types>/, `  ${notesContentType}\n</Types>`);

      // Add notes slide relationships file
      const notesRelsPath = `ppt/notesSlides/_rels/notesSlide${notesIndex}.xml.rels`;
      const notesRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../${slideRel?.target ?? ""}"/>
</Relationships>`;

      // Create notes slide XML
      const notesXml = createNotesSlideXml(text);

      // Build new zip
      const newEntries: Array<{ name: string; data: Buffer }> = [];
      for (const [name, data] of zip.entries()) {
        if (name === slideRelsPath) {
          newEntries.push({ name, data: Buffer.from(updatedSlideRelsXml, "utf8") });
        } else if (name === "[Content_Types].xml") {
          newEntries.push({ name, data: Buffer.from(updatedContentTypes, "utf8") });
        } else if (name !== notesPath && name !== notesRelsPath) {
          newEntries.push({ name, data });
        }
      }
      newEntries.push({ name: notesPath, data: Buffer.from(notesXml, "utf8") });
      newEntries.push({ name: notesRelsPath, data: Buffer.from(notesRelsXml, "utf8") });

      await writeFile(filePath, createStoredZip(newEntries));
      return ok(void 0);
    }

    // Update existing notes
    const existingNotesXml = zip.get(notesPath)?.toString("utf8") ?? "";
    const updatedNotesXml = updateNotesText(existingNotesXml, text);

    // Build new zip with updated notes
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === notesPath) {
        newEntries.push({ name, data: Buffer.from(updatedNotesXml, "utf8") });
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
 * Removes notes from a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 *
 * @example
 * const result = await removeNotes("/path/to/presentation.pptx", 1);
 */
export async function removeNotes(filePath: string, slideIndex: number): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const presentationXml = zip.get("ppt/presentation.xml")?.toString("utf8") ?? "";
    const relsXml = zip.get("ppt/_rels/presentation.xml.rels")?.toString("utf8") ?? "";
    const contentTypesXml = zip.get("[Content_Types].xml")?.toString("utf8") ?? "";

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

    // Find the notes relationship
    const slideRels = parseRelationshipEntries(slideRelsXml);
    const notesRelIndex = slideRels.findIndex(r => r.type?.endsWith("/notesSlide"));
    if (notesRelIndex < 0) {
      // No notes to remove
      return ok(void 0);
    }

    const notesRel = slideRels[notesRelIndex];
    const notesPath = normalizeZipPath(path.posix.dirname(slidePath), notesRel.target);
    const notesRelsPath = getRelationshipsEntryName(notesPath);

    // Remove notes relationship from slide's rels
    const updatedSlideRelsXml = slideRelsXml.replace(
      new RegExp(`<Relationship\\b[^>]*\\bId="${notesRel.id}"[^>]*\\/?>`, "g"),
      ""
    );

    // Remove Content_Type entry
    const updatedContentTypes = contentTypesXml.replace(
      new RegExp(`<Override\\b[^>]*PartName="/${notesPath}"[^>]*\\/?>`, "g"),
      ""
    );

    // Build new zip without notes entries
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideRelsPath) {
        newEntries.push({ name, data: Buffer.from(updatedSlideRelsXml, "utf8") });
      } else if (name === "[Content_Types].xml") {
        newEntries.push({ name, data: Buffer.from(updatedContentTypes, "utf8") });
      } else if (name === notesPath || name === notesRelsPath) {
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
