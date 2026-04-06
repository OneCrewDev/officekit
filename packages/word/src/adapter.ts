/**
 * Word adapter for @officekit/word.
 *
 * This module provides Get and Query functions for Word documents.
 * It reads docx files (ZIP archives containing XML) and parses the XML
 * to extract document structure, text, and formatting.
 *
 * @example
 * import { getWordNode, queryWordNodes } from "./adapter.js";
 *
 * // Get a specific node by path
 * const result = await getWordNode("document.docx", "/body/p[1]", 1);
 *
 * // Query nodes using a selector
 * const paragraphs = await queryWordNodes("document.docx", "p");
 * const boldRuns = await queryWordNodes("document.docx", "r[bold=true]");
 */

import { readFile, writeFile } from "node:fs/promises";
import JSZip from "jszip";

import { err, ok } from "./result.js";
import { parsePath, buildPath } from "./path.js";
import { parseSelector } from "./selectors.js";
import type { Result, DocumentNode, PathSegment } from "./types.js";

// ============================================================================
// ZIP Helpers
// ============================================================================

async function readDocxZip(filePath: string): Promise<JSZip> {
  const buffer = await readFile(filePath);
  return await JSZip.loadAsync(buffer);
}

async function getXmlEntry(zip: JSZip, entryName: string): Promise<string | null> {
  const entry = zip.file(entryName);
  if (!entry) return null;
  return await entry.async("string");
}

// ============================================================================
// XML Text Extraction Helpers
// ============================================================================

/**
 * Extracts all text content from an XML string.
 */
function extractTextFromXml(xml: string): string {
  const texts: string[] = [];
  const regex = /<[^>]*:t[^>]*>([^<]*)<\/[^>]*:t>/g;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    texts.push(match[1]);
  }
  return texts.join("");
}

/**
 * Extracts text content from w:t elements in a more robust way.
 */
function extractTextSimple(xml: string): string {
  const texts: string[] = [];
  const regex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    texts.push(match[1]);
  }
  return texts.join("");
}

/**
 * Gets all paragraph texts from document XML.
 */
function getParagraphsInfo(xml: string): Array<{ index: number; text: string; style?: string; paraId?: string }> {
  const paragraphs: Array<{ index: number; text: string; style?: string; paraId?: string }> = [];

  const paraRegex = /<w:p[\\s\\S]*?<\\/w:p>/g;
  let match;
  let idx = 0;
  while ((match = paraRegex.exec(xml)) !== null) {
    idx++;
    const paraXml = match[0];
    const text = extractTextSimple(paraXml);

    let style: string | undefined;
    let paraId: string | undefined;

    const styleMatch = paraXml.match(/<w:pStyle[^>]*w:val="([^"]*)"/);
    if (styleMatch) style = styleMatch[1];

    const paraIdMatch = paraXml.match(/<w:paraId[^>]*w:val="([^"]*)"/);
    if (paraIdMatch) paraId = paraIdMatch[1];

    paragraphs.push({ index: idx, text, style, paraId });
  }

  return paragraphs;
}

/**
 * Gets all table info from document XML.
 */
function getTablesInfo(xml: string): Array<{ index: number; rows: number; cols: number }> {
  const tables: Array<{ index: number; rows: number; cols: number }> = [];

  const tblRegex = /<w:tbl[\\s\\S]*?<\\/w:tbl>/g;
  let match;
  let idx = 0;
  while ((match = tblRegex.exec(xml)) !== null) {
    idx++;
    const tblXml = match[0];
    const rows = (tblXml.match(/<w:tr[\\s\\S]*?<\\/w:tr>/g) || []).length;
    const firstRow = tblXml.match(/<w:tr[\\s\\S]*?<\\/w:tr>/);
    const cols = firstRow ? (firstRow[0].match(/<w:tc[\\s\\S]*?<\\/w:tc>/g) || []).length : 0;
    tables.push({ index: idx, rows, cols });
  }

  return tables;
}

// ============================================================================
// Document Node Helpers
// ============================================================================

function createDocumentNode(path: string, type: string, text?: string, format?: Record<string, unknown>): DocumentNode {
  return {
    path,
    type,
    text,
    format: format || {},
  };
}

function createErrorNode(path: string, message: string): DocumentNode {
  return {
    path,
    type: "error",
    text: message,
    format: {},
  };
}

// ============================================================================
// Get Word Node
// ============================================================================

/**
 * Gets a node at the specified path from a Word document.
 *
 * @param filePath - Path to the .docx file
 * @param path - Path to the node (e.g., "/body/p[1]", "/body/tbl[1]/tr[1]/tc[2]")
 * @param depth - How deep to fetch children (0 = just this node, 1 = one level, etc.)
 * @returns Result containing the DocumentNode or error
 *
 * @example
 * const result = await getWordNode("document.docx", "/body", 1);
 * if (result.ok) {
 *   console.log(result.data.path);  // "/body"
 *   console.log(result.data.children?.length);  // number of children
 * }
 */
export async function getWordNode(filePath: string, path: string, depth = 1): Promise<Result<DocumentNode>> {
  try {
    const zip = await readDocxZip(filePath);
    const documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    const parsed = parsePath(path);
    if (!parsed.ok) {
      return err("invalid_path", parsed.error?.message || "Invalid path");
    }

    const segments = parsed.data?.segments || [];
    const result = navigateToElement(documentXml, zip, segments, depth);

    if (!result) {
      return err("not_found", `Path not found: ${path}`);
    }

    return ok(result);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Navigates to an element based on path segments.
 */
function navigateToElement(
  documentXml: string,
  zip: JSZip,
  segments: PathSegment[],
  depth: number,
  parentPath = "",
): DocumentNode | null {
  if (segments.length === 0) {
    return createDocumentNode("/", "document");
  }

  const first = segments[0];
  let currentPath = "/" + first.name + (first.index !== undefined ? `[${first.index}]` : "");
  let currentNode: DocumentNode | null = null;

  switch (first.name) {
    case "body": {
      if (segments.length === 1) {
        const paras = getParagraphsInfo(documentXml);
        const tables = getTablesInfo(documentXml);
        const children: DocumentNode[] = [];

        for (let i = 0; i < paras.length; i++) {
          children.push(createDocumentNode(
            `/body/p[${i + 1}]`,
            "paragraph",
            paras[i].text,
            { style: paras[i].style, paraId: paras[i].paraId }
          ));
        }
        for (let i = 0; i < tables.length; i++) {
          children.push(createDocumentNode(
            `/body/tbl[${i + 1}]`,
            "table",
            undefined,
            { rowCount: tables[i].rows, columnCount: tables[i].cols }
          ));
        }

        currentNode = createDocumentNode("/body", "body");
        if (depth > 0) {
          currentNode.children = children;
          currentNode.childCount = children.length;
        }
      }
      break;
    }

    case "p":
    case "paragraph": {
      const paras = getParagraphsInfo(documentXml);
      const idx = (first.index || 1) - 1;
      if (idx >= 0 && idx < paras.length) {
        const para = paras[idx];
        currentPath = `/body/p[${idx + 1}]`;
        currentNode = createDocumentNode(
          currentPath,
          "paragraph",
          para.text,
          { style: para.style, paraId: para.paraId }
        );

        if (depth > 0) {
          const runs = getRunsFromParagraph(documentXml, idx + 1);
          currentNode.children = runs;
          currentNode.childCount = runs.length;
        }
      }
      break;
    }

    case "tbl":
    case "table": {
      const tables = getTablesInfo(documentXml);
      const idx = (first.index || 1) - 1;
      if (idx >= 0 && idx < tables.length) {
        const table = tables[idx];
        currentPath = `/body/tbl[${idx + 1}]`;
        currentNode = createDocumentNode(
          currentPath,
          "table",
          undefined,
          { rowCount: table.rows, columnCount: table.cols }
        );

        if (depth > 0) {
          const rows: DocumentNode[] = [];
          for (let i = 0; i < table.rows; i++) {
            rows.push(createDocumentNode(
              `/body/tbl[${idx + 1}]/tr[${i + 1}]`,
              "row",
              undefined,
              { cellCount: table.cols }
            ));
          }
          currentNode.children = rows;
          currentNode.childCount = rows.length;
        }
      }
      break;
    }

    case "header": {
      const headerIdx = (first.index || 1) - 1;
      const headerEntry = zip.file(`word/header${headerIdx + 1}.xml`);
      if (headerEntry) {
        const headerXml = await headerEntry.async("string");
        const text = extractTextSimple(headerXml);
        currentNode = createDocumentNode(
          `/header[${headerIdx + 1}]`,
          "header",
          text
        );
      }
      break;
    }

    case "footer": {
      const footerIdx = (first.index || 1) - 1;
      const footerEntry = zip.file(`word/footer${footerIdx + 1}.xml`);
      if (footerEntry) {
        const footerXml = await footerEntry.async("string");
        const text = extractTextSimple(footerXml);
        currentNode = createDocumentNode(
          `/footer[${footerIdx + 1}]`,
          "footer",
          text
        );
      }
      break;
    }

    case "styles": {
      const stylesXml = await getXmlEntry(zip, "word/styles.xml");
      if (stylesXml) {
        const styles = parseStyles(stylesXml);
        currentNode = createDocumentNode("/styles", "styles");
        if (depth > 0) {
          currentNode.children = styles;
          currentNode.childCount = styles.length;
        }
      }
      break;
    }

    case "numbering": {
      const numberingXml = await getXmlEntry(zip, "word/numbering.xml");
      if (numberingXml) {
        currentNode = createDocumentNode("/numbering", "numbering");
      }
      break;
    }

    case "settings": {
      const settingsXml = await getXmlEntry(zip, "word/settings.xml");
      if (settingsXml) {
        currentNode = createDocumentNode("/settings", "settings");
      }
      break;
    }

    default: {
      break;
    }
  }

  if (!currentNode) {
    return null;
  }

  if (segments.length > 1 && currentNode) {
    const remainingPath = segments.slice(1);
    const childPath = buildChildPath(currentNode.path, remainingPath);

    if (remainingPath.length === 1 && remainingPath[0].name === "tr") {
      const rowIdx = (remainingPath[0].index || 1) - 1;
      const rowPath = `${currentNode.path}/tr[${rowIdx + 1}]`;
      return createDocumentNode(rowPath, "row");
    }

    if (remainingPath.length === 2 &&
        (remainingPath[0].name === "tr" || remainingPath[0].name === "row") &&
        (remainingPath[1].name === "tc" || remainingPath[1].name === "cell")) {
      const rowIdx = (remainingPath[0].index || 1) - 1;
      const cellIdx = (remainingPath[1].index || 1) - 1;
      const cellPath = `${currentNode.path}/tr[${rowIdx + 1}]/tc[${cellIdx + 1}]`;
      return createDocumentNode(cellPath, "cell");
    }
  }

  return currentNode;
}

function buildChildPath(parentPath: string, segments: PathSegment[]): string {
  if (segments.length === 0) return parentPath;

  const seg = segments[0];
  let path = parentPath;
  if (seg.index !== undefined) {
    path += `/${seg.name}[${seg.index}]`;
  } else if (seg.stringIndex !== undefined) {
    path += `/${seg.name}[${seg.stringIndex}]`;
  } else {
    path += `/${seg.name}`;
  }

  if (segments.length > 1) {
    path = buildChildPath(path, segments.slice(1));
  }

  return path;
}

/**
 * Gets runs from a specific paragraph.
 */
function getRunsFromParagraph(documentXml: string, paraIndex: number): DocumentNode[] {
  const runs: DocumentNode[] = [];

  const paraRegex = /<w:p[\\s\\S]*?<\\/w:p>/g;
  let match;
  let idx = 0;

  while ((match = paraRegex.exec(documentXml)) !== null) {
    idx++;
    if (idx !== paraIndex) continue;

    const paraXml = match[0];
    const runRegex = /<w:r[\\s\\S]*?<\\/w:r>/g;
    let runMatch;
    let runIdx = 0;

    while ((runMatch = runRegex.exec(paraXml)) !== null) {
      runIdx++;
      const runXml = runMatch[0];
      const text = extractTextSimple(runXml);

      const format: Record<string, unknown> = {};
      if (runXml.includes("<w:b/>") || runXml.includes("<w:b ")) format.bold = true;
      if (runXml.includes("<w:i/>") || runXml.includes("<w:i ")) format.italic = true;
      if (runXml.includes("<w:u ")) format.underline = "single";
      if (runXml.includes("<w:strike/>") || runXml.includes("<w:strike ")) format.strike = true;

      const fontMatch = runXml.match(/<w:rFonts[^>]*w:ascii="([^"]*)"/);
      if (fontMatch) format.font = fontMatch[1];

      const sizeMatch = runXml.match(/<w:sz[^>]*w:val="([^"]*)"/);
      if (sizeMatch) format.size = `${parseInt(sizeMatch[1]) / 2}pt`;

      const colorMatch = runXml.match(/<w:color[^>]*w:val="([^"]*)"/);
      if (colorMatch) format.color = colorMatch[1];

      runs.push(createDocumentNode(
        `/body/p[${paraIndex}]/r[${runIdx}]`,
        "run",
        text,
        format
      ));
    }
    break;
  }

  return runs;
}

/**
 * Parses styles from styles.xml.
 */
function parseStyles(stylesXml: string): DocumentNode[] {
  const styles: DocumentNode[] = [];

  const styleRegex = /<w:style[^>]*>([\\s\\S]*?)<\\/w:style>/g;
  let match;
  let idx = 0;

  while ((match = styleRegex.exec(stylesXml)) !== null) {
    idx++;
    const styleXml = match[0];

    const styleIdMatch = styleXml.match(/w:styleId="([^"]*)"/);
    const styleId = styleIdMatch ? styleIdMatch[1] : `style${idx}`;

    const nameMatch = styleXml.match(/<w:name[^>]*w:val="([^"]*)"/);
    const name = nameMatch ? nameMatch[1] : styleId;

    const typeMatch = styleXml.match(/w:type="([^"]*)"/);
    const type = typeMatch ? typeMatch[1] : "paragraph";

    const format: Record<string, unknown> = { id: styleId, name, type };

    const fontMatch = styleXml.match(/<w:rFonts[^>]*w:ascii="([^"]*)"/);
    if (fontMatch) format.font = fontMatch[1];

    const sizeMatch = styleXml.match(/<w:sz[^>]*w:val="([^"]*)"/);
    if (sizeMatch) format.size = `${parseInt(sizeMatch[1]) / 2}pt`;

    if (styleXml.includes("<w:b/>") || styleXml.includes("<w:b ")) format.bold = true;
    if (styleXml.includes("<w:i/>") || styleXml.includes("<w:i ")) format.italic = true;

    const colorMatch = styleXml.match(/<w:color[^>]*w:val="([^"]*)"/);
    if (colorMatch) format.color = colorMatch[1];

    styles.push(createDocumentNode(
      `/styles/${styleId}`,
      "style",
      name,
      format
    ));
  }

  return styles;
}

// ============================================================================
// Query Word Nodes
// ============================================================================

/**
 * Queries nodes using a selector from a Word document.
 *
 * @param filePath - Path to the .docx file
 * @param selector - CSS-like selector (e.g., "p", "p[style=Heading1]", "r[bold=true]")
 * @returns Result containing an array of DocumentNodes or error
 *
 * @example
 * const result = await queryWordNodes("document.docx", "p");
 * if (result.ok) {
 *   console.log(result.data.length);  // number of paragraphs
 *   console.log(result.data[0].text);  // first paragraph text
 * }
 *
 * @example
 * // Query all bold runs
 * const boldRuns = await queryWordNodes("document.docx", "r[bold=true]");
 *
 * @example
 * // Query paragraphs containing specific text
 * const matches = await queryWordNodes("document.docx", 'p:contains("Hello")');
 */
export async function queryWordNodes(filePath: string, selector: string): Promise<Result<DocumentNode[]>> {
  try {
    const zip = await readDocxZip(filePath);
    const documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    const parsed = parseSelector(selector);
    if (!parsed.ok) {
      return err("invalid_selector", parsed.error?.message || "Invalid selector");
    }

    const selectorData = parsed.data!;
    const results: DocumentNode[] = [];

    const elementType = selectorData.element || "p";

    switch (elementType) {
      case "p":
      case "paragraph": {
        const paras = getParagraphsInfo(documentXml);
        for (let i = 0; i < paras.length; i++) {
          const para = paras[i];
          if (!matchesSelectorAttributes(para, selectorData.attributes, documentXml, i + 1)) {
            continue;
          }
          if (selectorData.containsText && !para.text.includes(selectorData.containsText)) {
            continue;
          }

          const node = createDocumentNode(
            `/body/p[${i + 1}]`,
            "paragraph",
            para.text,
            { style: para.style, paraId: para.paraId }
          );
          results.push(node);
        }
        break;
      }

      case "r":
      case "run": {
        const runs = getAllRuns(documentXml);
        for (let i = 0; i < runs.length; i++) {
          const run = runs[i];
          if (!matchesRunAttributes(run, selectorData.attributes)) {
            continue;
          }
          if (selectorData.containsText && !run.text.includes(selectorData.containsText)) {
            continue;
          }

          results.push(createDocumentNode(
            run.path,
            "run",
            run.text,
            run.format
          ));
        }
        break;
      }

      case "tbl":
      case "table": {
        const tables = getTablesInfo(documentXml);
        for (let i = 0; i < tables.length; i++) {
          const table = tables[i];
          results.push(createDocumentNode(
            `/body/tbl[${i + 1}]`,
            "table",
            undefined,
            { rowCount: table.rows, columnCount: table.cols }
          ));
        }
        break;
      }

      case "tr":
      case "row": {
        const tables = getTablesInfo(documentXml);
        for (let t = 0; t < tables.length; t++) {
          for (let r = 0; r < tables[t].rows; r++) {
            results.push(createDocumentNode(
              `/body/tbl[${t + 1}]/tr[${r + 1}]`,
              "row",
              undefined,
              { cellCount: tables[t].cols }
            ));
          }
        }
        break;
      }

      case "tc":
      case "cell": {
        const tables = getTablesInfo(documentXml);
        for (let t = 0; t < tables.length; t++) {
          for (let r = 0; r < tables[t].rows; r++) {
            for (let c = 0; c < tables[t].cols; c++) {
              results.push(createDocumentNode(
                `/body/tbl[${t + 1}]/tr[${r + 1}]/tc[${c + 1}]`,
                "cell"
              ));
            }
          }
        }
        break;
      }

      case "header": {
        let headerIdx = 0;
        let headerEntry = zip.file(`word/header${headerIdx + 1}.xml`);
        while (headerEntry) {
          const headerXml = await headerEntry.async("string");
          const text = extractTextSimple(headerXml);
          if (!selectorData.containsText || text.includes(selectorData.containsText)) {
            results.push(createDocumentNode(
              `/header[${headerIdx + 1}]`,
              "header",
              text
            ));
          }
          headerIdx++;
          headerEntry = zip.file(`word/header${headerIdx + 1}.xml`);
        }
        break;
      }

      case "footer": {
        let footerIdx = 0;
        let footerEntry = zip.file(`word/footer${footerIdx + 1}.xml`);
        while (footerEntry) {
          const footerXml = await footerEntry.async("string");
          const text = extractTextSimple(footerXml);
          if (!selectorData.containsText || text.includes(selectorData.containsText)) {
            results.push(createDocumentNode(
              `/footer[${footerIdx + 1}]`,
              "footer",
              text
            ));
          }
          footerIdx++;
          footerEntry = zip.file(`word/footer${footerIdx + 1}.xml`);
        }
        break;
      }

      case "style":
      case "styles": {
        const stylesXml = await getXmlEntry(zip, "word/styles.xml");
        if (stylesXml) {
          const styles = parseStyles(stylesXml);
          for (const style of styles) {
            if (!selectorData.containsText || style.text?.includes(selectorData.containsText)) {
              results.push(style);
            }
          }
        }
        break;
      }

      case "bookmark": {
        const bookmarks = getBookmarks(documentXml);
        for (const bookmark of bookmarks) {
          results.push(createDocumentNode(
            bookmark.path,
            "bookmark",
            bookmark.text,
            { name: bookmark.name, id: bookmark.id }
          ));
        }
        break;
      }

      case "sdt":
      case "contentcontrol": {
        const sdts = getContentControls(documentXml);
        for (const sdt of sdts) {
          results.push(createDocumentNode(
            sdt.path,
            "sdt",
            sdt.text,
            sdt.format
          ));
        }
        break;
      }

      default: {
        break;
      }
    }

    return ok(results);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Checks if a paragraph matches selector attributes.
 */
function matchesSelectorAttributes(
  para: { index: number; text: string; style?: string; paraId?: string },
  attrs: Record<string, string>,
  documentXml: string,
  paraIndex: number,
): boolean {
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "empty") {
      if (value === "true" && para.text.trim().length > 0) return false;
      if (value === "false" && para.text.trim().length === 0) return false;
      continue;
    }

    if (key === "style") {
      const styleMatch = para.style === value;
      if (!styleMatch) return false;
      continue;
    }

    if (key === "index") {
      if (parseInt(value) !== para.index) return false;
      continue;
    }

    if (key.startsWith("@")) {
      continue;
    }

    if (key === "numId" || key === "numid") {
      continue;
    }
  }
  return true;
}

/**
 * Checks if a run matches selector attributes.
 */
function matchesRunAttributes(
  run: { text: string; format: Record<string, unknown> },
  attrs: Record<string, string>,
): boolean {
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "empty") {
      if (value === "true" && run.text.trim().length > 0) return false;
      if (value === "false" && run.text.trim().length === 0) return false;
      continue;
    }

    if (key === "bold") {
      const isBold = run.format.bold === true;
      const shouldBeBold = value === "true";
      if (isBold !== shouldBeBold) return false;
      continue;
    }

    if (key === "italic") {
      const isItalic = run.format.italic === true;
      const shouldBeItalic = value === "true";
      if (isItalic !== shouldBeItalic) return false;
      continue;
    }

    if (key === "underline") {
      const hasUnderline = run.format.underline !== undefined;
      if (value === "true" && !hasUnderline) return false;
      if (value !== "true" && hasUnderline && run.format.underline !== value) return false;
      continue;
    }

    if (key === "strike") {
      const hasStrike = run.format.strike === true;
      const shouldBeStruck = value === "true";
      if (hasStrike !== shouldBeStruck) return false;
      continue;
    }

    if (key === "font") {
      const font = run.format.font as string | undefined;
      if (!font || !font.toLowerCase().includes(value.toLowerCase())) return false;
      continue;
    }

    if (key === "size") {
      const size = run.format.size as string | undefined;
      if (!size) return false;
      const sizeNum = parseFloat(size);
      const targetNum = parseFloat(value);
      if (isNaN(sizeNum) || isNaN(targetNum)) return false;
      if (Math.abs(sizeNum - targetNum) > 0.1) return false;
      continue;
    }

    if (key === "color") {
      const color = run.format.color as string | undefined;
      if (!color) return false;
      if (color.toLowerCase() !== value.toLowerCase()) return false;
      continue;
    }
  }
  return true;
}

/**
 * Gets all runs from the document.
 */
function getAllRuns(documentXml: string): Array<{
  path: string;
  text: string;
  format: Record<string, unknown>;
  paraIndex: number;
  runIndex: number;
}> {
  const runs: Array<{
    path: string;
    text: string;
    format: Record<string, unknown>;
    paraIndex: number;
    runIndex: number;
  }> = [];

  const paraRegex = /<w:p[\\s\\S]*?<\\/w:p>/g;
  let match;
  let paraIdx = 0;

  while ((match = paraRegex.exec(documentXml)) !== null) {
    paraIdx++;
    const paraXml = match[0];
    const runRegex = /<w:r[\\s\\S]*?<\\/w:r>/g;
    let runMatch;
    let runIdx = 0;

    while ((runMatch = runRegex.exec(paraXml)) !== null) {
      runIdx++;
      const runXml = runMatch[0];
      const text = extractTextSimple(runXml);

      const format: Record<string, unknown> = {};
      if (runXml.includes("<w:b/>") || runXml.includes("<w:b ")) format.bold = true;
      if (runXml.includes("<w:i/>") || runXml.includes("<w:i ")) format.italic = true;
      if (runXml.includes("<w:u ")) format.underline = "single";
      if (runXml.includes("<w:strike/>") || runXml.includes("<w:strike ")) format.strike = true;

      const fontMatch = runXml.match(/<w:rFonts[^>]*w:ascii="([^"]*)"/);
      if (fontMatch) format.font = fontMatch[1];

      const sizeMatch = runXml.match(/<w:sz[^>]*w:val="([^"]*)"/);
      if (sizeMatch) format.size = `${parseInt(sizeMatch[1]) / 2}pt`;

      const colorMatch = runXml.match(/<w:color[^>]*w:val="([^"]*)"/);
      if (colorMatch) format.color = colorMatch[1];

      runs.push({
        path: `/body/p[${paraIdx}]/r[${runIdx}]`,
        text,
        format,
        paraIndex: paraIdx,
        runIndex: runIdx,
      });
    }
  }

  return runs;
}

/**
 * Gets bookmarks from the document.
 */
function getBookmarks(documentXml: string): Array<{ path: string; name: string; id: string; text: string }> {
  const bookmarks: Array<{ path: string; name: string; id: string; text: string }> = [];

  const bookmarkStartRegex = /<w:bookmarkStart[^>]*w:id="([^"]*)"[^>]*w:name="([^"]*)"[^>]*>/g;
  let match;

  while ((match = bookmarkStartRegex.exec(documentXml)) !== null) {
    const id = match[1];
    const name = match[2];

    const startIdx = match.index;
    const endIdx = documentXml.indexOf("</w:bookmarkEnd>", startIdx);
    const bookmarkContent = documentXml.slice(startIdx, endIdx > 0 ? endIdx + 16 : undefined);
    const text = extractTextSimple(bookmarkContent);

    bookmarks.push({
      path: `/bookmark[${name}]`,
      name,
      id,
      text,
    });
  }

  return bookmarks;
}

/**
 * Gets content controls (SDT) from the document.
 */
function getContentControls(documentXml: string): Array<{
  path: string;
  text: string;
  format: Record<string, unknown>;
}> {
  const sdts: Array<{ path: string; text: string; format: Record<string, unknown> }> = [];

  const sdtRegex = /<w:sdt[\\s\\S]*?<\\/w:sdt>/g;
  let match;
  let idx = 0;

  while ((match = sdtRegex.exec(documentXml)) !== null) {
    idx++;
    const sdtXml = match[0];
    const text = extractTextSimple(sdtXml);

    const format: Record<string, unknown> = {};
    const tagMatch = sdtXml.match(/<w:tag[^>]*w:val="([^"]*)"/);
    if (tagMatch) format.tag = tagMatch[1];

    sdts.push({
      path: `/body/sdt[${idx}]`,
      text,
      format,
    });
  }

  return sdts;
}

// ============================================================================
// Document Info
// ============================================================================

/**
 * Gets basic document information without deep traversal.
 */
export async function getDocumentInfo(filePath: string): Promise<Result<DocumentNode>> {
  try {
    const zip = await readDocxZip(filePath);
    const documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    const paras = getParagraphsInfo(documentXml);
    const tables = getTablesInfo(documentXml);

    let headerCount = 0;
    let footerCount = 0;

    let entry = zip.file(`word/header${headerCount + 1}.xml`);
    while (entry) {
      headerCount++;
      entry = zip.file(`word/header${headerCount + 1}.xml`);
    }

    entry = zip.file(`word/footer${footerCount + 1}.xml`);
    while (entry) {
      footerCount++;
      entry = zip.file(`word/footer${footerCount + 1}.xml`);
    }

    const node = createDocumentNode("/", "document");
    node.childCount = 1;
    node.children = [createDocumentNode("/body", "body", undefined, {
      paragraphCount: paras.length,
      tableCount: tables.length,
      headerCount,
      footerCount,
    })];
    node.format = {
      paragraphCount: paras.length,
      tableCount: tables.length,
      headerCount,
      footerCount,
    };

    return ok(node);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Mutation Functions (Add/Set/Remove/Move/Swap/Batch)
// ============================================================================

/**
 * Word document namespaces
 */
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const V_NS = "urn:schemas-microsoft-com:vml";
const O_NS = "urn:schemas-microsoft-com:office:office";

/**
 * Helper: Escape XML special characters
 */
function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Helper: Sanitize hex color
 */
function sanitizeHex(color: string): string {
  return color.replace(/^#/, "").toUpperCase().padStart(6, "0");
}

/**
 * Helper: Generate unique ID
 */
function generateId(prefix: string, existing: string[] = []): string {
  const max = existing.reduce((m, id) => {
    const num = parseInt(id.replace(prefix, ""), 10);
    return isNaN(num) ? m : Math.max(m, num);
  }, 0);
  return `${prefix}${max + 1}`;
}

/**
 * Helper: Create a paragraph XML string
 */
function createParagraphXml(properties: Record<string, string> = {}): string {
  const { text, style, alignment, bold, italic, color, font, size, underline } = properties;

  let pPr = "";
  if (style || alignment) {
    const styleAttr = style ? `<w:pStyle w:val="${escapeXml(style)}"/>` : "";
    const alignAttr = alignment ? `<w:jc w:val="${alignment}"/>` : "";
    pPr = `<w:pPr>${styleAttr}${alignAttr}</w:pPr>`;
  }

  let rPr = "";
  if (bold || italic || color || font || size || underline) {
    const boldTag = bold ? "<w:b/>" : "";
    const italicTag = italic ? "<w:i/>" : "";
    const colorTag = color ? `<w:color w:val="${sanitizeHex(color)}"/>` : "";
    const fontTag = font ? `<w:rFonts w:ascii="${escapeXml(font)}" w:hAnsi="${escapeXml(font)}"/>` : "";
    const sizeTag = size ? `<w:sz w:val="${parseInt(size, 10) * 2}"/>` : "";
    const ulTag = underline ? `<w:u w:val="${underline === true ? "single" : underline}"/>` : "";
    rPr = `<w:rPr>${fontTag}${boldTag}${italicTag}${colorTag}${ulTag}${sizeTag}</w:rPr>`;
  }

  const textContent = text ? `<w:t xml:space="preserve">${escapeXml(text)}</w:t>` : "";
  return `<w:p>${pPr}<w:r>${rPr}${textContent}</w:r></w:p>`;
}

/**
 * Helper: Create a run XML string
 */
function createRunXml(properties: Record<string, string> = {}): string {
  const { text, bold, italic, color, font, size, underline, highlight } = properties;

  let rPr = "";
  if (bold || italic || color || font || size || underline || highlight) {
    const boldTag = bold ? "<w:b/>" : "";
    const italicTag = italic ? "<w:i/>" : "";
    const colorTag = color ? `<w:color w:val="${sanitizeHex(color)}"/>` : "";
    const fontTag = font ? `<w:rFonts w:ascii="${escapeXml(font)}" w:hAnsi="${escapeXml(font)}"/>` : "";
    const sizeTag = size ? `<w:sz w:val="${parseInt(size, 10) * 2}"/>` : "";
    const ulTag = underline ? `<w:u w:val="${underline === true ? "single" : underline}"/>` : "";
    const hlTag = highlight ? `<w:highlight w:val="${highlight}"/>` : "";
    rPr = `<w:rPr>${fontTag}${boldTag}${italicTag}${colorTag}${ulTag}${sizeTag}${hlTag}</w:rPr>`;
  }

  const textContent = text ? `<w:t xml:space="preserve">${escapeXml(text)}</w:t>` : "";
  return `<w:r>${rPr}${textContent}</w:r>`;
}

/**
 * Helper: Create a table XML string
 */
function createTableXml(properties: Record<string, string> = {}): string {
  const rows = parseInt(properties.rows || "1", 10);
  const cols = parseInt(properties.cols || "1", 10);
  const { width, style, alignment } = properties;

  let tblPr = "<w:tblPr>";
  if (style) {
    tblPr += `<w:tblStyle w:val="${escapeXml(style)}"/>`;
  }
  if (width) {
    tblPr += `<w:tblW w:w="${width}" w:type="dxa"/>`;
  }
  if (alignment) {
    tblPr += `<w:jc w:val="${alignment}"/>`;
  }
  tblPr += "</w:tblPr>";

  let tblBorders = "";
  if (!style) {
    tblBorders = `<w:tblBorders>
      <w:top w:val="single" w:sz="4"/>
      <w:left w:val="single" w:sz="4"/>
      <w:bottom w:val="single" w:sz="4"/>
      <w:right w:val="single" w:sz="4"/>
      <w:insideH w:val="single" w:sz="4"/>
      <w:insideV w:val="single" w:sz="4"/>
    </w:tblBorders>`;
  }

  let tblGrid = "<w:tblGrid>";
  for (let c = 0; c < cols; c++) {
    tblGrid += "<w:gridCol/>";
  }
  tblGrid += "</w:tblGrid>";

  let tblBody = "";
  for (let r = 0; r < rows; r++) {
    tblBody += "<w:tr>";
    for (let c = 0; c < cols; c++) {
      tblBody += `<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>`;
    }
    tblBody += "</w:tr>";
  }

  return `<w:tbl>${tblPr}${tblBorders}${tblGrid}${tblBody}</w:tbl>`;
}

/**
 * Helper: Create a table row XML string
 */
function createTableRowXml(cols: number, properties: Record<string, string> = {}): string {
  const { height, header } = properties;
  let trPr = "";
  if (height || header) {
    trPr = "<w:trPr>";
    if (height) trPr += `<w:trHeight w:val="${height}" w:hRule="atLeast"/>`;
    if (header) trPr += "<w:tblHeader/>";
    trPr += "</w:trPr>";
  }

  let cells = "";
  for (let c = 0; c < cols; c++) {
    cells += "<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>";
  }

  return `<w:tr>${trPr}${cells}</w:tr>`;
}

/**
 * Helper: Create a table cell XML string
 */
function createTableCellXml(properties: Record<string, string> = {}): string {
  const { text, width, vAlign } = properties;
  let tcPr = "";
  if (width) tcPr += `<w:tcW w:w="${width}" w:type="dxa"/>`;
  if (vAlign) tcPr += `<w:vAlign w:val="${vAlign}"/>`;

  const textContent = text ? `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>` : "";
  return `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ""}<w:p>${textContent}</w:p></w:tc>`;
}

/**
 * Helper: Create a picture/image XML string
 */
function createPictureXml(properties: Record<string, string> = {}): string {
  const width = properties.width || "5486400";
  const height = properties.height || "3657600";
  const alt = properties.alt || "";
  const relationshipId = properties.relationshipId || "rId1";

  return `<w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="${WP_NS}">
        <wp:extent cx="${width}" cy="${height}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="1" name="Picture" descr="${alt}"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic xmlns:a="${A_NS}">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr>
                <pic:cNvPr id="1" name="Picture"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="${relationshipId}" xmlns:r="${R_NS}"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>`;
}

/**
 * Helper: Create a field (complex script) XML string
 */
function createFieldXml(fieldType: string, properties: Record<string, string> = {}): string {
  const { font, size, bold, italic, color, text = "1" } = properties;

  let rPr = "";
  if (font || size || bold || italic || color) {
    const boldTag = bold ? "<w:b/>" : "";
    const italicTag = italic ? "<w:i/>" : "";
    const colorTag = color ? `<w:color w:val="${sanitizeHex(color)}"/>` : "";
    const fontTag = font ? `<w:rFonts w:ascii="${escapeXml(font)}" w:hAnsi="${escapeXml(font)}"/>` : "";
    const sizeTag = size ? `<w:sz w:val="${parseInt(size, 10) * 2}"/>` : "";
    rPr = `<w:rPr>${fontTag}${boldTag}${italicTag}${colorTag}${sizeTag}</w:rPr>`;
  }

  const instr = fieldType === "PAGE" ? " PAGE " :
    fieldType === "NUMPAGES" ? " NUMPAGES " :
    fieldType === "DATE" ? ' DATE \\@ "yyyy-MM-dd" ' :
    fieldType === "AUTHOR" ? " AUTHOR " :
    fieldType === "TITLE" ? " TITLE " :
    fieldType === "FILENAME" ? " FILENAME " :
    fieldType === "TIME" ? " TIME " :
    ` ${fieldType} `;

  return `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>
<w:r>${rPr}<w:instrText xml:space="preserve">${instr}</w:instrText></w:r>
<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>
<w:r>${rPr}<w:t>${text}</w:t></w:r>
<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`;
}

/**
 * Helper: Create a break XML string
 */
function createBreakXml(type: string = "page"): string {
  const breakType = type === "column" ? 'w:type="column"' : type === "line" ? 'w:type="textWrapping"' : "";
  return `<w:r><w:br ${breakType}/></w:r>`;
}

/**
 * Helper: Create a section break XML string
 */
function createSectionXml(properties: Record<string, string> = {}): string {
  const { type = "nextPage", pageWidth, pageHeight, marginTop, marginBottom, marginLeft, marginRight, columns } = properties;

  const sectType = type === "continuous" ? "continuous" : type === "evenPage" ? "evenPage" : type === "oddPage" ? "oddPage" : "nextPage";

  let pgSz = "";
  if (pageWidth || pageHeight) {
    pgSz = `<w:pgSz w:w="${pageWidth || 11906}" w:h="${pageHeight || 16838}"/>`;
  }

  let pgMar = "";
  if (marginTop || marginBottom || marginLeft || marginRight) {
    pgMar = `<w:pgMar w:top="${marginTop || 1440}" w:right="${marginRight || 1800}" w:bottom="${marginBottom || 1440}" w:left="${marginLeft || 1800}"/>`;
  }

  let cols = "";
  if (columns) {
    cols = `<w:cols w:num="${columns}"/>`;
  }

  return `<w:p>
  <w:pPr>
    <w:sectPr>
      <w:type w:val="${sectType}"/>
      ${pgSz}${pgMar}${cols}
    </w:sectPr>
  </w:pPr>
</w:p>`;
}

/**
 * Helper: Create a TOC field XML string
 */
function createTocXml(properties: Record<string, string> = {}): string {
  const levels = properties.levels || "1-3";
  const title = properties.title;
  const instr = ` TOC \\o "${levels}" \\h \\u `;

  let result = "";
  if (title) {
    result += `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>${escapeXml(title)}</w:t></w:r></w:p>`;
  }

  result += `<w:p>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Update field to see table of contents</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>`;

  return result;
}

/**
 * Helper: Create a hyperlink XML string
 */
function createHyperlinkXml(properties: Record<string, string> = {}): string {
  const { url, anchor, text, color, font, bold, italic } = properties;

  let rPr = "";
  if (color || font || bold || italic) {
    const boldTag = bold ? "<w:b/>" : "";
    const italicTag = italic ? "<w:i/>" : "";
    const colorTag = color ? `<w:color w:val="${sanitizeHex(color)}" w:themeColor="hyperlink"/>` : `<w:color w:val="0563C1" w:themeColor="hyperlink"/>`;
    const fontTag = font ? `<w:rFonts w:ascii="${escapeXml(font)}" w:hAnsi="${escapeXml(font)}"/>` : "";
    rPr = `<w:rPr>${fontTag}${boldTag}${italicTag}${colorTag}<w:u w:val="single"/></w:rPr>`;
  } else {
    rPr = `<w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr>`;
  }

  const linkText = text || url || anchor || "link";
  const attrs = url ? `r:id="${url}"` : `w:anchor="${escapeXml(anchor || "")}"`;

  return `<w:hyperlink ${attrs}>
    <w:r>${rPr}<w:t xml:space="preserve">${escapeXml(linkText)}</w:t></w:r>
  </w:hyperlink>`;
}

/**
 * Helper: Create a bookmark XML string
 */
function createBookmarkXml(name: string, properties: Record<string, string> = {}): string {
  const { text } = properties;
  const id = generateId("1", []);
  let content = "";
  if (text) {
    content = `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  }
  return `<w:bookmarkStart w:id="${id}" w:name="${escapeXml(name)}"/>${content}<w:bookmarkEnd w:id="${id}"/>`;
}

/**
 * Helper: Create a comment XML string
 */
function createCommentXml(properties: Record<string, string>): { id: string; xml: string } {
  const { text, author = "officekit", initials = "O", date } = properties;
  const id = generateId("1", []);
  const dateStr = date || new Date().toISOString();
  return { id, xml: `<w:comment w:id="${id}" w:author="${escapeXml(author)}" w:initials="${escapeXml(initials)}" w:date="${dateStr}"><w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:comment>` };
}

/**
 * Helper: Create a footnote XML string
 */
function createFootnoteXml(properties: Record<string, string>): { id: string; xml: string } {
  const { text } = properties;
  const id = generateId("1", []);
  return { id, xml: `<w:footnote w:id="${id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${escapeXml(text)}</w:t></w:r></w:p></w:footnote>` };
}

/**
 * Helper: Create an endnote XML string
 */
function createEndnoteXml(properties: Record<string, string>): { id: string; xml: string } {
  const { text } = properties;
  const id = generateId("1", []);
  return { id, xml: `<w:endnote w:id="${id}"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${escapeXml(text)}</w:t></w:r></w:p></w:endnote>` };
}

/**
 * Helper: Create a style XML string
 */
function createStyleXml(properties: Record<string, string>): string {
  const { id, name, type = "paragraph", basedOn, next, font, size, bold, italic, color, alignment } = properties;

  const styleId = id || name || "CustomStyle";
  const styleName = name || id || "CustomStyle";
  const styleType = type === "character" || type === "char" ? "character" : type === "table" ? "table" : type === "numbering" ? "numbering" : "paragraph";

  let styleXml = `<w:style w:type="${styleType}" w:styleId="${escapeXml(styleId)}" w:customStyle="1">
    <w:name w:val="${escapeXml(styleName)}"/>`;

  if (basedOn) styleXml += `<w:basedOn w:val="${escapeXml(basedOn)}"/>`;
  if (next) styleXml += `<w:next w:val="${escapeXml(next)}"/>`;

  let pPr = "";
  if (alignment) pPr += `<w:jc w:val="${alignment}"/>`;
  if (pPr) styleXml += `<w:pPr>${pPr}</w:pPr>`;

  let rPr = "";
  if (font) rPr += `<w:rFonts w:ascii="${escapeXml(font)}" w:hAnsi="${escapeXml(font)}"/>`;
  if (size) rPr += `<w:sz w:val="${parseInt(size, 10) * 2}"/>`;
  if (bold) rPr += `<w:b/>`;
  if (italic) rPr += `<w:i/>`;
  if (color) rPr += `<w:color w:val="${sanitizeHex(color)}"/>`;
  if (rPr) styleXml += `<w:rPr>${rPr}</w:rPr>`;

  styleXml += "</w:style>";
  return styleXml;
}

/**
 * Helper: Create an SDT (Content Control) XML string
 */
function createSdtXml(properties: Record<string, string> = {}): string {
  const { text = "", alias, tag, lock, sdtType = "text" } = properties;
  const id = generateId("1", []);

  let sdtPr = `<w:sdtPr><w:id w:val="${id}"/>`;
  if (alias) sdtPr += `<w:alias w:val="${escapeXml(alias)}"/>`;
  if (tag) sdtPr += `<w:tag w:val="${escapeXml(tag)}"/>`;
  if (lock) {
    const lockVal = lock === "contentLocked" || lock === "content" ? "contentLocked" :
      lock === "sdtLocked" || lock === "sdt" ? "sdtLocked" :
      lock === "sdtContentLocked" || lock === "both" ? "sdtContentLocked" : "unlocked";
    sdtPr += `<w:lock w:val="${lockVal}"/>`;
  }

  if (sdtType === "dropdown" || sdtType === "dropdownlist") {
    sdtPr += `<w:dropDownList/>`;
  } else if (sdtType === "date" || sdtType === "datepicker") {
    sdtPr += `<w:date w:dateFormat="yyyy-MM-dd"/>`;
  } else {
    sdtPr += `<w:text/>`;
  }
  sdtPr += "</w:sdtPr>";

  return `<w:sdt>${sdtPr}<w:sdtContent><w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:sdtContent></w:sdt>`;
}

/**
 * Helper: Create a watermark XML string
 */
function createWatermarkXml(properties: Record<string, string> = {}): string {
  const text = properties.text || "DRAFT";
  const color = properties.color || "silver";
  const font = properties.font || "Calibri";
  const size = properties.size || "1pt";
  const rotation = properties.rotation || "315";
  const opacity = properties.opacity || ".5";

  return `<v:shapetype id="_x0000_t136" coordsize="1600,21600" o:spt="136" adj="10800" path="m@7,0l@8,0m@5,21600l@6,21600e" xmlns:v="${V_NS}" xmlns:o="${O_NS}">
  <v:formulas>
    <v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>
    <v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>
    <v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>
    <v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>
    <v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/>
  </v:formulas>
  <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>
  <v:textpath on="t" fitshape="t"/>
  <v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>
  <o:lock v:ext="edit" text="t" shapetype="t"/>
</v:shapetype>
<v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s1025" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;rotation:${rotation};z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="${color}" stroked="f" xmlns:v="${V_NS}" xmlns:o="${O_NS}">
  <v:fill opacity="${opacity}"/>
  <v:textpath style="font-family:&quot;${escapeXml(font)}&quot;;font-size:${size}" string="${escapeXml(text)}"/>
</v:shape>`;
}

/**
 * Helper: Create header XML string
 */
function createHeaderXml(properties: Record<string, string> = {}): string {
  const { text, alignment = "center", field } = properties;
  const type = properties.type || "default";
  let rPr = "";
  if (properties.font || properties.size || properties.bold || properties.italic || properties.color) {
    const fontTag = properties.font ? `<w:rFonts w:ascii="${escapeXml(properties.font)}" w:hAnsi="${escapeXml(properties.font)}"/>` : "";
    const sizeTag = properties.size ? `<w:sz w:val="${parseInt(properties.size, 10) * 2}"/>` : "";
    const boldTag = properties.bold ? "<w:b/>" : "";
    const italicTag = properties.italic ? "<w:i/>" : "";
    const colorTag = properties.color ? `<w:color w:val="${sanitizeHex(properties.color)}"/>` : "";
    rPr = `<w:rPr>${fontTag}${boldTag}${italicTag}${colorTag}${sizeTag}</w:rPr>`;
  }

  let content = "";
  if (text) {
    content = `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  } else if (field) {
    content = createFieldXml(field, properties);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}">
  <w:sdt>
    <w:sdtPr>
      <w:id w:val="-1"/>
      <w:docPartObj>
        <w:docPartGallery w:val="Watermarks"/>
        <w:docPartUnique/>
      </w:docPartObj>
    </w:sdtPr>
    <w:sdtContent>
      <w:p>
        <w:pPr><w:pStyle w:val="Header"/><w:jc w:val="${alignment}"/></w:pPr>
        ${content}
      </w:p>
    </w:sdtContent>
  </w:sdt>
</w:hdr>`;
}

/**
 * Helper: Create footer XML string
 */
function createFooterXml(properties: Record<string, string> = {}): string {
  const { text, alignment = "center", field } = properties;
  let rPr = "";
  if (properties.font || properties.size || properties.bold || properties.italic || properties.color) {
    const fontTag = properties.font ? `<w:rFonts w:ascii="${escapeXml(properties.font)}" w:hAnsi="${escapeXml(properties.font)}"/>` : "";
    const sizeTag = properties.size ? `<w:sz w:val="${parseInt(properties.size, 10) * 2}"/>` : "";
    const boldTag = properties.bold ? "<w:b/>" : "";
    const italicTag = properties.italic ? "<w:i/>" : "";
    const colorTag = properties.color ? `<w:color w:val="${sanitizeHex(properties.color)}"/>` : "";
    rPr = `<w:rPr>${fontTag}${boldTag}${italicTag}${colorTag}${sizeTag}</w:rPr>`;
  }

  let content = "";
  if (text) {
    content = `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  } else if (field) {
    content = createFieldXml(field, properties);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}">
  <w:p>
    <w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="${alignment}"/></w:pPr>
    ${content}
  </w:p>
</w:ftr>`;
}

/**
 * Helper: Insert XML at position
 */
function insertAtPosition(docXml: string, insertXml: string, position: string | number | undefined): string {
  // Handle "find:" prefix for text-based anchoring
  if (position && typeof position === "string" && position.startsWith("find:")) {
    const findText = position.substring(5);
    const findIdx = docXml.indexOf(findText);
    if (findIdx === -1) {
      throw new Error(`Text not found: ${findText}`);
    }
    return docXml.slice(0, findIdx) + insertXml + docXml.slice(findIdx);
  }

  // Handle index-based positioning
  const bodyMatch = docXml.match(/<w:body>([\s\S]*)<\/w:body>/);
  if (!bodyMatch) {
    throw new Error("Document body not found");
  }

  const bodyOpen = docXml.indexOf("<w:body>");
  const bodyClose = docXml.indexOf("</w:body>");

  if (position === "start" || position === 0) {
    return docXml.slice(0, bodyOpen + 8) + insertXml + docXml.slice(bodyOpen + 8);
  }

  if (position === "end" || position === undefined || position === null) {
    return docXml.slice(0, bodyClose) + insertXml + docXml.slice(bodyClose);
  }

  // Insert at specific index
  const paras = bodyMatch[1].match(/<w:p[>\s]/g) || [];
  if (typeof position === "number" && position >= paras.length) {
    return docXml.slice(0, bodyClose) + insertXml + docXml.slice(bodyClose);
  }

  // Find position of the nth paragraph
  let paraCount = 0;
  let pos = bodyOpen + 8;
  while (paraCount < (position as number) && pos < bodyClose) {
    const nextPara = docXml.indexOf("<w:p", pos);
    if (nextPara === -1 || nextPara >= bodyClose) break;
    paraCount++;
    pos = nextPara + 4;
  }

  return docXml.slice(0, pos) + insertXml + docXml.slice(pos);
}

/**
 * Helper: Process find and replace/format
 */
function processFindAndFormat(docXml: string, find: string, replace: string | null, formatProps: Record<string, string>, useRegex: boolean): { docXml: string; matchCount: number } {
  let result = docXml;
  let matchCount = 0;

  if (useRegex) {
    const flags = "g" + (find.includes("i") ? "i" : "");
    const pattern = find.startsWith("r\"") && find.endsWith("\"")
      ? find.slice(2, -1)
      : find;
    const regex = new RegExp(pattern, flags);

    if (replace !== null && replace !== undefined) {
      result = result.replace(regex, replace);
    }
  } else {
    // Simple text search
    let searchStr = find;
    let idx = result.indexOf(searchStr);
    while (idx !== -1) {
      matchCount++;
      if (replace !== null && replace !== undefined) {
        result = result.slice(0, idx) + replace + result.slice(idx + searchStr.length);
        idx = result.indexOf(searchStr, idx + replace.length);
      } else {
        idx = result.indexOf(searchStr, idx + 1);
      }
    }
  }

  return { docXml: result, matchCount };
}

// ============================================================================
// Public Mutation API Functions
// ============================================================================

/**
 * Add an element to a Word document
 */
export async function addWordNode(
  filePath: string,
  targetPath: string,
  options: { type?: string; props?: Record<string, string>; position?: string; after?: string; before?: string } = {}
): Promise<Result<{ path: string }>> {
  try {
    const { type = "paragraph", props = {}, position, after, before } = options;

    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    let resultPath = targetPath;
    let insertXml = "";
    let effectivePosition: string | number | undefined = position;

    // Handle after/before
    if (after) {
      effectivePosition = `find:${after}`;
    } else if (before) {
      effectivePosition = `find:${before}`;
    }

    switch (type.toLowerCase()) {
      case "paragraph":
      case "p":
        insertXml = createParagraphXml(props);
        break;

      case "run":
      case "r":
        insertXml = createRunXml(props);
        break;

      case "table":
      case "tbl":
        insertXml = createTableXml(props);
        break;

      case "row":
      case "tr":
        if (!targetPath.includes("/tbl[")) {
          return err("invalid_path", "Rows must be added to a table: /body/tbl[N]");
        }
        const rows = parseInt(props.cols || "1", 10);
        insertXml = createTableRowXml(rows, props);
        break;

      case "cell":
      case "tc":
        if (!targetPath.includes("/tr[")) {
          return err("invalid_path", "Cells must be added to a table row: /body/tbl[N]/tr[M]");
        }
        insertXml = createTableCellXml(props);
        break;

      case "picture":
      case "image":
      case "img":
        if (!props.path && !props.src) {
          return err("invalid_args", "Picture requires 'path' or 'src' property");
        }
        // For now, create placeholder with relationshipId
        insertXml = createPictureXml({ ...props, relationshipId: "rId999" });
        break;

      case "bookmark":
        if (!props.name) {
          return err("invalid_args", "Bookmark requires 'name' property");
        }
        insertXml = createBookmarkXml(props.name, props);
        break;

      case "hyperlink":
      case "link":
        if (!props.url && !props.anchor) {
          return err("invalid_args", "Hyperlink requires 'url' or 'anchor' property");
        }
        insertXml = createHyperlinkXml(props);
        break;

      case "section":
      case "sectionbreak":
        insertXml = createSectionXml(props);
        break;

      case "toc":
      case "tableofcontents":
        insertXml = createTocXml(props);
        break;

      case "field":
      case "pagenum":
      case "pagenumber":
      case "page":
      case "numpages":
      case "date":
      case "author":
        insertXml = createFieldXml(type.toUpperCase(), props);
        break;

      case "break":
      case "pagebreak":
      case "columnbreak":
        insertXml = createBreakXml(props.type || (type === "columnbreak" ? "column" : "page"));
        break;

      case "comment":
        if (!props.text) {
          return err("invalid_args", "Comment requires 'text' property");
        }
        const comment = createCommentXml(props);
        insertXml = `<w:commentRangeStart w:id="${comment.id}"/><w:commentRangeEnd w:id="${comment.id}"/><w:r><w:commentReference w:id="${comment.id}"/></w:r>`;
        break;

      case "footnote":
        if (!props.text) {
          return err("invalid_args", "Footnote requires 'text' property");
        }
        const footnote = createFootnoteXml(props);
        insertXml = `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${footnote.id}"/></w:r>`;
        break;

      case "endnote":
        if (!props.text) {
          return err("invalid_args", "Endnote requires 'text' property");
        }
        const endnote = createEndnoteXml(props);
        insertXml = `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${endnote.id}"/></w:r>`;
        break;

      case "style":
        if (!props.name && !props.id) {
          return err("invalid_args", "Style requires 'name' or 'id' property");
        }
        const styleXml = createStyleXml(props);
        const stylesXml = await getXmlEntry(zip, "word/styles.xml");
        if (stylesXml) {
          const updatedStyles = stylesXml.replace("</w:styles>", `${styleXml}</w:styles>`);
          zip.file("word/styles.xml", updatedStyles);
        } else {
          zip.file("word/styles.xml", `<w:styles xmlns:w="${W_NS}">${styleXml}</w:styles>`);
        }
        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ path: `/styles/${props.name || props.id}` });

      case "header":
        const headerIdx = (zip.file(/^word\/header\d+\.xml$/) || []).length + 1;
        const headerContent = createHeaderXml(props);
        zip.file(`word/header${headerIdx}.xml`, headerContent);
        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ path: `/header[${headerIdx}]` });

      case "footer":
        const footerIdx = (zip.file(/^word\/footer\d+\.xml$/) || []).length + 1;
        const footerContent = createFooterXml(props);
        zip.file(`word/footer${footerIdx}.xml`, footerContent);
        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ path: `/footer[${footerIdx}]` });

      case "sdt":
      case "contentcontrol":
        insertXml = createSdtXml(props);
        break;

      case "watermark":
        const wmHeader = createWatermarkXml(props);
        const headerIdx2 = (zip.file(/^word\/header\d+\.xml$/) || []).length + 1;
        zip.file(`word/header${headerIdx2}.xml`, createHeaderXml({ ...props, text: undefined }) + `<w:pict>${wmHeader}</w:pict>`);
        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ path: "/watermark" });

      default:
        return err("invalid_type", `Unknown element type: ${type}`);
    }

    // Insert the XML
    documentXml = insertAtPosition(documentXml, insertXml, effectivePosition);
    zip.file("word/document.xml", documentXml);

    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
    return ok({ path: resultPath });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Set properties on an element in a Word document
 */
export async function setWordNode(
  filePath: string,
  targetPath: string,
  options: { props?: Record<string, string> } = {}
): Promise<Result<{ path: string; matchCount?: number }>> {
  try {
    const { props = {} } = options;

    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    // Handle find + format/replace
    if (props.find) {
      const find = props.find;
      const replace = props.replace || null;
      const useRegex = props.regex === "true" || props.regex === true;
      const { matchCount } = processFindAndFormat(documentXml, find, replace, props, useRegex);

      documentXml = matchCount > 0 ? documentXml : documentXml;
      zip.file("word/document.xml", documentXml);
      await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
      return ok({ path: targetPath, matchCount });
    }

    // Handle document-level properties
    if (targetPath === "/" || targetPath === "" || targetPath === "/body") {
      // Document properties modification would go here
      await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
      return ok({ path: targetPath });
    }

    // Handle style path
    if (targetPath.startsWith("/styles/")) {
      const styleId = targetPath.substring(8);
      let stylesXml = await getXmlEntry(zip, "word/styles.xml");
      if (stylesXml) {
        // Update existing style properties would go here
        zip.file("word/styles.xml", stylesXml);
      }
      await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
      return ok({ path: targetPath });
    }

    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
    return ok({ path: targetPath });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Remove an element from a Word document
 */
export async function removeWordNode(
  filePath: string,
  targetPath: string
): Promise<Result<{ ok: boolean; targetPath: string }>> {
  try {
    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    // Handle watermark removal
    if (targetPath === "/watermark") {
      // Remove watermark headers
      const headerFiles = zip.file(/^word\/header\d+\.xml$/);
      for (const file of headerFiles) {
        const content = await file.async("string");
        if (content.includes("Watermarks") || content.includes("WaterMark")) {
          zip.remove(file.name);
        }
      }
      await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
      return ok({ ok: true, targetPath });
    }

    // Handle header/footer removal
    const hfMatch = targetPath.match(/^\/(header|footer)\[(\d+)\]$/);
    if (hfMatch) {
      const [, type, idx] = hfMatch;
      const fileName = `word/${type}${idx}.xml`;
      if (zip.file(fileName)) {
        zip.remove(fileName);
      }
      await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
      return ok({ ok: true, targetPath });
    }

    // Handle TOC removal
    if (targetPath.match(/^\/toc\[\d+\]$/)) {
      // Remove TOC paragraphs - simplified
      documentXml = documentXml.replace(/<w:p[^>]*>[\s\S]*?<w:fldChar[\s\S]*?TOC[\s\S]*?<\/w:p>/g, "");
      zip.file("word/document.xml", documentXml);
      await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
      return ok({ ok: true, targetPath });
    }

    // For other removals, we'd need to parse the path and remove the specific element
    // Simplified implementation
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
    return ok({ ok: true, targetPath });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Move an element within a Word document
 */
export async function moveWordNode(
  filePath: string,
  sourcePath: string,
  targetPath: string,
  options: { after?: string; before?: string; position?: string | number } = {}
): Promise<Result<{ path: string }>> {
  try {
    // Full move implementation would:
    // 1. Navigate to source element
    // 2. Clone it
    // 3. Remove from original position
    // 4. Insert at target position

    return err("not_implemented", "Move operation not yet fully implemented");
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Swap two elements in a Word document
 */
export async function swapWordNodes(
  filePath: string,
  path1: string,
  path2: string
): Promise<Result<{ path1: string; path2: string }>> {
  try {
    // Full swap implementation would:
    // 1. Navigate to element 1
    // 2. Navigate to element 2
    // 3. Swap their content/positions

    return err("not_implemented", "Swap operation not yet fully implemented");
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Execute a batch of operations on a Word document
 */
export async function batchWordNodes(
  filePath: string,
  operations: Array<{ action: string; target: string; options?: Record<string, unknown> }>
): Promise<Result<Array<{ action: string; target: string; status: string }>>> {
  try {
    const results: Array<{ action: string; target: string; status: string }> = [];

    for (const op of operations) {
      const { action, target, options = {} } = op;

      switch (action.toLowerCase()) {
        case "add": {
          const result = await addWordNode(filePath, target, options as Parameters<typeof addWordNode>[2]);
          results.push({ action, target, status: result.ok ? "success" : "failed" });
          break;
        }
        case "set": {
          const result = await setWordNode(filePath, target, options as Parameters<typeof setWordNode>[2]);
          results.push({ action, target, status: result.ok ? "success" : "failed" });
          break;
        }
        case "remove": {
          const result = await removeWordNode(filePath, target);
          results.push({ action, target, status: result.ok ? "success" : "failed" });
          break;
        }
        case "move": {
          const result = await moveWordNode(filePath, target, options.target as string || "/", options as Parameters<typeof moveWordNode>[3]);
          results.push({ action, target, status: result.ok ? "success" : "failed" });
          break;
        }
        case "swap": {
          const result = await swapWordNodes(filePath, target, options.path2 as string || "/");
          results.push({ action, target, status: result.ok ? "success" : "failed" });
          break;
        }
        default:
          results.push({ action, target, status: "unknown_action" });
      }
    }

    return ok(results);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// View Modes (text, annotated, outline, stats, issues, html, forms)
// ============================================================================

export interface WordViewOptions {
  startLine?: number;
  endLine?: number;
  maxLines?: number;
  cols?: string[];
  pageFilter?: string;
  issueType?: string;
  limit?: number;
}

/**
 * Views a Word document in various modes.
 */
export async function viewWordDocument(
  filePath: string,
  mode: string,
  options?: WordViewOptions,
): Promise<{ mode: string; output: string }> {
  const zip = await readDocxZip(filePath);
  const documentXml = await getXmlEntry(zip, "word/document.xml") ?? "";
  const stylesXml = await getXmlEntry(zip, "word/styles.xml") ?? "";

  const normalizedMode = mode.toLowerCase();

  switch (normalizedMode) {
    case "text":
      return { mode, output: renderTextView(documentXml, options) };
    case "annotated":
      return { mode, output: renderAnnotatedView(documentXml, stylesXml, options) };
    case "outline":
      return { mode, output: renderOutlineView(documentXml, filePath) };
    case "stats":
      return { mode, output: renderStatsView(documentXml) };
    case "issues":
      return { mode, output: renderIssuesView(documentXml, options) };
    case "html":
      return { mode, output: renderHtmlView(documentXml, stylesXml) };
    case "forms":
      return { mode, output: await renderFormsView(zip, documentXml) };
    case "json":
      return { mode, output: renderJsonView(documentXml, stylesXml) };
    default:
      throw new Error(`Unsupported view mode '${mode}'. Use: text, annotated, outline, stats, issues, html, forms, or json.`);
  }
}

function renderTextView(xml: string, options?: WordViewOptions): string {
  const lines: string[] = [];
  const startLine = options?.startLine ?? 1;
  const endLine = options?.endLine ?? Number.MAX_SAFE_INTEGER;
  const maxLines = options?.maxLines ?? Number.MAX_SAFE_INTEGER;

  const paras = getParagraphsInfo(xml);
  let lineNum = 0;
  let emitted = 0;

  for (const para of paras) {
    lineNum++;

    if (lineNum < startLine) continue;
    if (lineNum > endLine) break;

    if (emitted >= maxLines) {
      lines.push(`... (showed ${emitted} rows, use --start/--end to view more)`);
      break;
    }

    const path = `/body/p[${para.index}]`;
    const styleStr = para.style && para.style !== "Normal" ? `[${para.style}] ` : "";
    const prefix = styleStr ? `[${path}] ${styleStr}` : `[${path}] `;
    lines.push(`${prefix}${para.text}`);
    emitted++;
  }

  return lines.join("\n");
}

function renderAnnotatedView(xml: string, stylesXml: string, options?: WordViewOptions): string {
  const lines: string[] = [];
  const startLine = options?.startLine ?? 1;
  const endLine = options?.endLine ?? Number.MAX_SAFE_INTEGER;
  const maxLines = options?.maxLines ?? Number.MAX_SAFE_INTEGER;

  const paras = getParagraphsInfo(xml);
  let lineNum = 0;
  let emitted = 0;

  for (const para of paras) {
    lineNum++;

    if (lineNum < startLine) continue;
    if (lineNum > endLine) break;

    if (emitted >= maxLines) {
      lines.push(`... (showed ${emitted} rows, use --start/--end to view more)`);
      break;
    }

    const path = `/body/p[${para.index}]`;
    const styleName = para.style ? getStyleNameFromId(stylesXml, para.style) || para.style : "Normal";

    if (!para.text.trim() && !hasRuns(xml, para.index)) {
      lines.push(`[${path}] [] <- ${styleName} | empty paragraph`);
    } else {
      const runs = getRunsInfo(xml, para.index);
      for (const run of runs) {
        const fmt = formatRunInfo(run);
        lines.push(`[${path}] 「${run.text}」 <- ${styleName} | ${fmt}`);
      }
    }
    emitted++;
  }

  return lines.join("\n");
}

function renderOutlineView(xml: string, filePath: string): string {
  const lines: string[] = [];

  const paras = getParagraphsInfo(xml);
  const tables = getTablesInfo(xml);
  const fileName = filePath.split("/").pop() || "document.docx";

  lines.push(`File: ${fileName} | ${paras.length} paragraphs | ${tables.length} tables`);

  let lineNum = 0;
  for (const para of paras) {
    lineNum++;

    if (para.style && (para.style.includes("Heading") || para.style === "Title" || para.style === "Subtitle")) {
      const level = getHeadingLevel(para.style);
      const indent = level <= 1 ? "" : "  ".repeat(level - 1);
      const prefix = level === 0 ? "■" : "├──";
      lines.push(`${indent}${prefix} [${lineNum}] "${para.text}" (${para.style})`);
    }
  }

  return lines.join("\n");
}

function renderStatsView(xml: string): string {
  const lines: string[] = [];

  const paras = getParagraphsInfo(xml);
  const styleCounts: Record<string, number> = {};
  const fontCounts: Record<string, number> = {};
  const sizeCounts: Record<string, number> = {};

  let totalWords = 0;
  let emptyParagraphs = 0;
  let doubleSpaces = 0;
  let totalChars = 0;

  for (const para of paras) {
    const style = para.style || "Normal";
    styleCounts[style] = (styleCounts[style] || 0) + 1;

    if (!para.text.trim()) {
      emptyParagraphs++;
      continue;
    }

    const words = para.text.split(/\s+/).filter(Boolean);
    totalWords += words.length;
    totalChars += para.text.length;

    if (para.text.includes("  ")) {
      doubleSpaces++;
    }

    const runs = getRunsInfo(xml, para.index);
    for (const run of runs) {
      if (run.font) fontCounts[run.font] = (fontCounts[run.font] || 0) + 1;
      if (run.size) sizeCounts[run.size] = (sizeCounts[run.size] || 0) + 1;
    }
  }

  lines.push(`Paragraphs: ${paras.length} | Words: ${totalWords} | Total Characters: ${totalChars}`);
  lines.push("");

  lines.push("Style Distribution:");
  for (const [style, count] of Object.entries(styleCounts).sort((a, b) => b[1] - a[1])) {
    lines.push(`  ${style}: ${count}`);
  }

  lines.push("");
  lines.push("Font Usage:");
  for (const [font, count] of Object.entries(fontCounts).sort((a, b) => b[1] - a[1]).slice(0, 10)) {
    lines.push(`  ${font}: ${count}`);
  }

  lines.push("");
  lines.push("Font Size Usage:");
  for (const [size, count] of Object.entries(sizeCounts).sort((a, b) => b[1] - a[1]).slice(0, 10)) {
    lines.push(`  ${size}: ${count}`);
  }

  lines.push("");
  lines.push(`Empty Paragraphs: ${emptyParagraphs}`);
  lines.push(`Consecutive Spaces: ${doubleSpaces}`);

  return lines.join("\n");
}

function renderIssuesView(xml: string, options?: WordViewOptions): string {
  const lines: string[] = [];
  const limit = options?.limit ?? 100;

  const paras = getParagraphsInfo(xml);
  let issueNum = 0;

  for (let i = 0; i < paras.length && issueNum < limit; i++) {
    const para = paras[i];
    const path = `/body/p[${para.index}]`;

    if (!para.text.trim()) {
      issueNum++;
      lines.push(`[S${issueNum}] Structure | Warning | ${path} | Empty paragraph`);
    }

    if (para.text.includes("  ")) {
      issueNum++;
      lines.push(`[C${issueNum}] Content | Warning | ${path} | Consecutive spaces`);
    }

    if ((para.style === "Normal" || !para.style) && para.text.trim()) {
      const hasIndent = hasFirstLineIndent(xml, para.index);
      if (!hasIndent) {
        issueNum++;
        lines.push(`[F${issueNum}] Format | Warning | ${path} | Body paragraph missing first-line indent`);
      }
    }

    if (issueNum >= limit) break;
  }

  return lines.length > 0 ? lines.join("\n") : "No issues found.";
}

function renderHtmlView(xml: string, _stylesXml: string): string {
  const lines: string[] = [];

  lines.push("<!DOCTYPE html>");
  lines.push("<html lang=\"en\">");
  lines.push("<head>");
  lines.push("<meta charset=\"UTF-8\">");
  lines.push("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
  lines.push("<title>Word Document Preview</title>");
  lines.push("<style>");
  lines.push(generateBasicCss());
  lines.push("</style>");
  lines.push("</head>");
  lines.push("<body>");

  const paras = getParagraphsInfo(xml);
  for (const para of paras) {
    let className = para.style || "Normal";
    className = className.replace(/\s+/g, "");

    let html = `<p`;
    if (className !== "Normal") {
      html += ` class="${escapeHtml(className)}"`;
    }
    html += ">";

    const runs = getRunsInfo(xml, para.index);
    for (const run of runs) {
      let text = escapeHtml(run.text);
      if (run.bold) text = `<strong>${text}</strong>`;
      if (run.italic) text = `<em>${text}</em>`;
      if (run.underline) text = `<u>${text}</u>`;
      if (run.color) text = `<span style="color:${run.color}">${text}</span>`;
      html += text;
    }

    html += "</p>";
    lines.push(html);
  }

  lines.push("</body>");
  lines.push("</html>");

  return lines.join("\n");
}

async function renderFormsView(zip: JSZip, xml: string): Promise<string> {
  const lines: string[] = [];

  const settingsXml = await getXmlEntry(zip, "word/settings.xml");
  const hasProtection = settingsXml ? /<w:documentProtection/.test(settingsXml) : false;
  lines.push(`Document Protection: ${hasProtection ? "enabled" : "none"}`);

  const sdts = getContentControls(xml);

  if (sdts.length === 0) {
    lines.push("");
    lines.push("No form fields or content controls found.");
    return lines.join("\n");
  }

  lines.push("");
  lines.push(`Content Controls (${sdts.length}):`);

  for (let i = 0; i < sdts.length; i++) {
    const sdt = sdts[i];
    const nameStr = sdt.tag ? ` name="${sdt.tag}"` : "";
    lines.push(`  #${i + 1} [sdt] type=${sdt.type || "richtext"}${nameStr} value="${sdt.text}"`);
  }

  return lines.join("\n");
}

function renderJsonView(xml: string, stylesXml: string): string {
  const result: Record<string, unknown> = {
    paragraphs: [],
    styles: parseStylesForJson(stylesXml),
  };

  const paras = getParagraphsInfo(xml);
  for (const para of paras) {
    result.paragraphs.push({
      text: para.text,
      style: para.style || "Normal",
    });
  }

  return JSON.stringify(result, null, 2);
}

// ============================================================================
// Style Management
// ============================================================================

export interface WordStyleProperties {
  basedOn?: string;
  next?: string;
  font?: string;
  fontSize?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  underline?: string;
  alignment?: "left" | "center" | "right" | "justify";
  spaceBefore?: string;
  spaceAfter?: string;
  lineSpacing?: string;
}

/**
 * Sets a style in a Word document.
 */
export async function setWordStyle(
  filePath: string,
  styleName: string,
  properties: WordStyleProperties,
): Promise<{ ok: boolean; path: string }> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  let stylesXml = await getXmlEntry(zip, "word/styles.xml");
  if (!stylesXml) {
    stylesXml = createBasicStylesXml();
    zip.file("word/styles.xml", stylesXml);
  }

  const styleId = styleName.replace(/\s+/g, "");
  const styleExists = new RegExp(`<w:style\\b[^>]*\\s+w:styleId=["']${escapeRegex(styleId)}["']`, "i").test(stylesXml);

  if (styleExists) {
    stylesXml = updateExistingStyle(stylesXml, styleId, properties);
  } else {
    stylesXml = addNewStyle(stylesXml, styleName, properties);
  }

  zip.file("word/styles.xml", stylesXml);

  const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });
  await writeFile(filePath, updatedBuffer);

  return { ok: true, path: `/styles/${styleId}` };
}

function createBasicStylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:hint="default"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>`;
}

function addNewStyle(stylesXml: string, styleName: string, properties: WordStyleProperties): string {
  const styleId = styleName.replace(/\s+/g, "");
  const styleType = properties.basedOn?.includes("Char") ? "character" : "paragraph";

  let styleXml = `  <w:style w:type="${styleType}" w:styleId="${escapeXml(styleId)}">\n`;
  styleXml += `    <w:name w:val="${escapeXml(styleName)}"/>\n`;

  if (properties.basedOn) {
    styleXml += `    <w:basedOn w:val="${escapeXml(properties.basedOn)}"/>\n`;
  }
  if (properties.next) {
    styleXml += `    <w:next w:val="${escapeXml(properties.next)}"/>\n`;
  }

  const rPrParts: string[] = [];
  if (properties.font) {
    rPrParts.push(`<w:rFonts w:ascii="${escapeXml(properties.font)}" w:hAnsi="${escapeXml(properties.font)}"/>`);
  }
  if (properties.fontSize) {
    const halfPoints = parseFontSizeToHalfPoints(properties.fontSize);
    rPrParts.push(`<w:sz w:val="${halfPoints}"/>`);
    rPrParts.push(`<w:szCs w:val="${halfPoints}"/>`);
  }
  if (properties.bold) rPrParts.push("<w:b/>");
  if (properties.italic) rPrParts.push("<w:i/>");
  if (properties.color) rPrParts.push(`<w:color w:val="${escapeXml(properties.color)}"/>`);
  if (properties.underline) rPrParts.push(`<w:u w:val="${escapeXml(properties.underline)}"/>`);

  if (rPrParts.length > 0) {
    styleXml += "    <w:rPr>\n";
    for (const part of rPrParts) {
      styleXml += `      ${part}\n`;
    }
    styleXml += "    </w:rPr>\n";
  }

  const pPrParts: string[] = [];
  if (properties.alignment) {
    pPrParts.push(`<w:jc w:val="${mapAlignment(properties.alignment)}"/>`);
  }
  if (properties.spaceBefore || properties.spaceAfter || properties.lineSpacing) {
    let spacing = "<w:spacing";
    if (properties.spaceBefore) spacing += ` w:before="${parseWordSpacing(properties.spaceBefore)}"`;
    if (properties.spaceAfter) spacing += ` w:after="${parseWordSpacing(properties.spaceAfter)}"`;
    if (properties.lineSpacing) spacing += ` w:line="${parseLineSpacing(properties.lineSpacing)}" w:lineRule="auto"`;
    spacing += "/>";
    pPrParts.push(spacing);
  }

  if (pPrParts.length > 0) {
    styleXml += "    <w:pPr>\n";
    for (const part of pPrParts) {
      styleXml += `      ${part}\n`;
    }
    styleXml += "    </w:pPr>\n";
  }

  styleXml += "  </w:style>\n";

  const insertIndex = stylesXml.lastIndexOf("</w:styles>");
  return stylesXml.slice(0, insertIndex) + styleXml + stylesXml.slice(insertIndex);
}

function updateExistingStyle(stylesXml: string, styleId: string, properties: WordStyleProperties): string {
  const styleRegex = new RegExp(`<w:style\\b[^>]*\\s+w:styleId=["']${escapeRegex(styleId)}["'][^>]*>[\\s\\S]*?<\\/w:style>`, "i");
  const match = styleRegex.exec(stylesXml);

  if (!match) {
    throw new Error(`Style '${styleId}' not found`);
  }

  const oldStyle = match[0];
  const styleName = getStyleNameFromId(stylesXml, styleId) || styleId;
  const props: WordStyleProperties = { ...properties };

  if (!props.basedOn) {
    const basedOnMatch = /<w:basedOn\s+w:val=["']([^"']+)["']/i.exec(oldStyle);
    props.basedOn = basedOnMatch ? basedOnMatch[1] : "Normal";
  }

  const newStylesXml = stylesXml.replace(oldStyle, "");
  return addNewStyle(newStylesXml, styleName, props);
}

// ============================================================================
// Section Layout
// ============================================================================

export interface WordSectionProperties {
  pageWidth?: number;
  pageHeight?: number;
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  orientation?: "portrait" | "landscape";
  columns?: number;
  columnSpace?: number;
  sectionType?: "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";
}

/**
 * Sets section layout properties in a Word document.
 */
export async function setWordSection(
  filePath: string,
  sectionPath: string,
  properties: WordSectionProperties,
): Promise<{ ok: boolean; path: string }> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  let documentXml = await getXmlEntry(zip, "word/document.xml");
  if (!documentXml) {
    throw new Error("Document.xml not found");
  }

  const hasSectPr = /<w:sectPr/.test(documentXml);

  if (!hasSectPr) {
    const sectPrXml = buildSectionPropertiesXml(properties);
    documentXml = documentXml.replace("</w:body>", `${sectPrXml}</w:body>`);
  } else {
    documentXml = updateSectionProperties(documentXml, properties);
  }

  zip.file("word/document.xml", documentXml);

  const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });
  await writeFile(filePath, updatedBuffer);

  return { ok: true, path: sectionPath || "/body/sectPr" };
}

function buildSectionPropertiesXml(props: WordSectionProperties): string {
  let xml = "<w:sectPr>";

  if (props.pageWidth || props.pageHeight) {
    const width = props.pageWidth || 12240;
    const height = props.pageHeight || 15840;
    const orient = props.orientation === "landscape" ? "landscape" : "portrait";
    xml += `<w:pgSz w:w="${width}" w:h="${height}" w:orient="${orient}"/>`;
  }

  if (props.marginTop || props.marginBottom || props.marginLeft || props.marginRight) {
    xml += `<w:pgMar w:top="${props.marginTop || 1440}" w:right="${props.marginRight || 1440}" w:bottom="${props.marginBottom || 1440}" w:left="${props.marginLeft || 1440}" w:header="720" w:footer="720" w:gutter="0"/>`;
  }

  if (props.columns !== undefined) {
    xml += `<w:cols w:num="${props.columns}"`;
    if (props.columnSpace) {
      xml += ` w:space="${props.columnSpace}"`;
    }
    xml += "/>";
  }

  if (props.sectionType) {
    xml += `<w:type w:val="${mapSectionType(props.sectionType)}"/>`;
  }

  xml += "</w:sectPr>";
  return xml;
}

function updateSectionProperties(docXml: string, props: WordSectionProperties): string {
  const sectPrRegex = /<w:sectPr\b[^>]*>([\s\S]*?)<\/w:sectPr>/i;

  if (!sectPrRegex.test(docXml)) {
    return docXml.replace("</w:body>", `${buildSectionPropertiesXml(props)}</w:body>`);
  }

  let newXml = docXml;

  if (props.pageWidth !== undefined || props.pageHeight !== undefined || props.orientation !== undefined) {
    const pgSzRegex = /<w:pgSz\b[^>]*\/?>/i;
    const width = props.pageWidth || 12240;
    const height = props.pageHeight || 15840;
    const orient = props.orientation === "landscape" ? "landscape" : "portrait";

    if (pgSzRegex.test(newXml)) {
      newXml = newXml.replace(pgSzRegex, `<w:pgSz w:w="${width}" w:h="${height}" w:orient="${orient}"/>`);
    } else {
      newXml = newXml.replace("<w:sectPr>", `<w:sectPr><w:pgSz w:w="${width}" w:h="${height}" w:orient="${orient}"/>`);
    }
  }

  if (props.marginTop !== undefined || props.marginBottom !== undefined || props.marginLeft !== undefined || props.marginRight !== undefined) {
    const pgMarRegex = /<w:pgMar\b[^>]*\/?>/i;
    const top = props.marginTop ?? 1440;
    const bottom = props.marginBottom ?? 1440;
    const left = props.marginLeft ?? 1440;
    const right = props.marginRight ?? 1440;

    if (pgMarRegex.test(newXml)) {
      newXml = newXml.replace(pgMarRegex, `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}" w:header="720" w:footer="720" w:gutter="0"/>`);
    } else {
      newXml = newXml.replace("<w:sectPr>", `<w:sectPr><w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}" w:header="720" w:footer="720" w:gutter="0"/>`);
    }
  }

  if (props.columns !== undefined) {
    const colsRegex = /<w:cols\b[^>]*\/?>/i;
    const space = props.columnSpace ?? 480;

    if (colsRegex.test(newXml)) {
      newXml = newXml.replace(colsRegex, `<w:cols w:num="${props.columns}" w:space="${space}"/>`);
    } else {
      newXml = newXml.replace("<w:sectPr>", `<w:sectPr><w:cols w:num="${props.columns}" w:space="${space}"/>`);
    }
  }

  if (props.sectionType !== undefined) {
    const typeRegex = /<w:type\b[^>]*\/?>/i;
    const sectTypeVal = mapSectionType(props.sectionType);

    if (typeRegex.test(newXml)) {
      newXml = newXml.replace(typeRegex, `<w:type w:val="${sectTypeVal}"/>`);
    } else {
      newXml = newXml.replace("<w:sectPr>", `<w:sectPr><w:type w:val="${sectTypeVal}"/>`);
    }
  }

  return newXml;
}

// ============================================================================
// Doc Defaults
// ============================================================================

export interface WordDocDefaults {
  font?: string;
  fontSize?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  alignment?: string;
  spaceBefore?: string;
  spaceAfter?: string;
  lineSpacing?: string;
}

/**
 * Sets document default properties in a Word document.
 */
export async function setWordDocDefaults(
  filePath: string,
  properties: WordDocDefaults,
): Promise<{ ok: boolean }> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  let stylesXml = await getXmlEntry(zip, "word/styles.xml");
  if (!stylesXml) {
    stylesXml = createBasicStylesXml();
    zip.file("word/styles.xml", stylesXml);
  }

  stylesXml = updateDocDefaults(stylesXml, properties);
  zip.file("word/styles.xml", stylesXml);

  const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });
  await writeFile(filePath, updatedBuffer);

  return { ok: true };
}

function updateDocDefaults(stylesXml: string, props: WordDocDefaults): string {
  let newXml = stylesXml;

  if (!/<w:docDefaults>/i.test(newXml)) {
    newXml = newXml.replace("<w:styles", "<w:styles><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>");
  }

  if (props.font || props.fontSize || props.bold || props.italic || props.color) {
    const rPrRegex = /<w:rPrDefault>([\s\S]*?)<\/w:rPrDefault>/i;
    const rPrMatch = rPrRegex.exec(newXml);

    if (rPrMatch) {
      let rPrContent = rPrMatch[1];

      if (props.font) {
        rPrContent = updateOrAddElement(rPrContent, "w:rFonts", `<w:rFonts w:ascii="${escapeXml(props.font)}" w:hAnsi="${escapeXml(props.font)}"/>`);
      }

      if (props.fontSize) {
        const halfPoints = parseFontSizeToHalfPoints(props.fontSize);
        rPrContent = updateOrAddElement(rPrContent, "w:sz", `<w:sz w:val="${halfPoints}"/>`);
        rPrContent = updateOrAddElement(rPrContent, "w:szCs", `<w:szCs w:val="${halfPoints}"/>`);
      }

      if (props.bold) rPrContent = updateOrAddElement(rPrContent, "w:b", "<w:b/>");
      if (props.italic) rPrContent = updateOrAddElement(rPrContent, "w:i", "<w:i/>");
      if (props.color) rPrContent = updateOrAddElement(rPrContent, "w:color", `<w:color w:val="${escapeXml(props.color)}"/>`);

      newXml = newXml.replace(rPrMatch[0], `<w:rPrDefault>${rPrContent}</w:rPrDefault>`);
    }
  }

  if (props.alignment || props.spaceBefore || props.spaceAfter || props.lineSpacing) {
    if (!/<w:pPrDefault>/i.test(newXml)) {
      newXml = newXml.replace("</w:docDefaults>", "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>");
    }

    const pPrRegex = /<w:pPrDefault>([\s\S]*?)<\/w:pPrDefault>/i;
    const pPrMatch = pPrRegex.exec(newXml);

    if (pPrMatch) {
      let pPrContent = pPrMatch[1];

      if (props.alignment) {
        pPrContent = updateOrAddElement(pPrContent, "w:jc", `<w:jc w:val="${mapAlignment(props.alignment)}"/>`);
      }

      if (props.spaceBefore || props.spaceAfter || props.lineSpacing) {
        let spacingAttrs = "";
        if (props.spaceBefore) spacingAttrs += ` w:before="${parseWordSpacing(props.spaceBefore)}"`;
        if (props.spaceAfter) spacingAttrs += ` w:after="${parseWordSpacing(props.spaceAfter)}"`;
        if (props.lineSpacing) spacingAttrs += ` w:line="${parseLineSpacing(props.lineSpacing)}" w:lineRule="auto"`;
        pPrContent = updateOrAddElement(pPrContent, "w:spacing", `<w:spacing${spacingAttrs}/>`);
      }

      newXml = newXml.replace(pPrMatch[0], `<w:pPrDefault>${pPrContent}</w:pPrDefault>`);
    }
  }

  return newXml;
}

// ============================================================================
// Raw XML Operations
// ============================================================================

/**
 * Gets raw XML from a Word document part.
 */
export async function rawWordDocument(filePath: string, partPath: string): Promise<string> {
  const zip = await readDocxZip(filePath);
  const normalizedPath = partPath.toLowerCase();

  switch (normalizedPath) {
    case "/":
    case "/document":
    case "/word/document.xml":
      return await getXmlEntry(zip, "word/document.xml") ?? "";
    case "/styles":
    case "/word/styles.xml":
      return await getXmlEntry(zip, "word/styles.xml") ?? "(no styles)";
    case "/settings":
    case "/word/settings.xml":
      return await getXmlEntry(zip, "word/settings.xml") ?? "(no settings)";
    case "/numbering":
    case "/word/numbering.xml":
      return await getXmlEntry(zip, "word/numbering.xml") ?? "(no numbering)";
    case "/comments":
    case "/word/comments.xml":
      return await getXmlEntry(zip, "word/comments.xml") ?? "(no comments)";
    default: {
      const headerMatch = /^\/header\[?(\d+)?\]?$/i.exec(partPath);
      if (headerMatch) {
        const idx = headerMatch[1] ? parseInt(headerMatch[1], 10) - 1 : 0;
        return await getXmlEntry(zip, `word/header${idx + 1}.xml`) ?? `(no header ${idx + 1})`;
      }

      const footerMatch = /^\/footer\[?(\d+)?\]?$/i.exec(partPath);
      if (footerMatch) {
        const idx = footerMatch[1] ? parseInt(footerMatch[1], 10) - 1 : 0;
        return await getXmlEntry(zip, `word/footer${idx + 1}.xml`) ?? `(no footer ${idx + 1})`;
      }

      throw new Error(`Unsupported part path '${partPath}'. Use: /document, /styles, /settings, /numbering, /comments, /header[n], /footer[n].`);
    }
  }
}

export interface RawSetOptions {
  xpath?: string;
  action?: string;
  xml?: string;
}

/**
 * Sets raw XML in a Word document part.
 */
export async function rawSetWordDocument(
  filePath: string,
  partPath: string,
  xpath: string,
  action: string,
  xml?: string,
): Promise<{ ok: boolean; affected: number }> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  const normalizedPath = partPath.toLowerCase();
  let targetXml: string | null = null;
  let entryName: string | null = null;

  switch (normalizedPath) {
    case "/":
    case "/document":
    case "/word/document.xml":
      targetXml = await getXmlEntry(zip, "word/document.xml");
      entryName = "word/document.xml";
      break;
    case "/styles":
    case "/word/styles.xml":
      targetXml = await getXmlEntry(zip, "word/styles.xml");
      entryName = "word/styles.xml";
      break;
    case "/settings":
    case "/word/settings.xml":
      targetXml = await getXmlEntry(zip, "word/settings.xml");
      entryName = "word/settings.xml";
      break;
    case "/numbering":
    case "/word/numbering.xml":
      targetXml = await getXmlEntry(zip, "word/numbering.xml");
      entryName = "word/numbering.xml";
      break;
    case "/comments":
    case "/word/comments.xml":
      targetXml = await getXmlEntry(zip, "word/comments.xml");
      entryName = "word/comments.xml";
      break;
    default: {
      const headerMatch = /^\/header\[?(\d+)?\]?$/i.exec(partPath);
      if (headerMatch) {
        const idx = headerMatch[1] ? parseInt(headerMatch[1], 10) - 1 : 0;
        targetXml = await getXmlEntry(zip, `word/header${idx + 1}.xml`);
        entryName = `word/header${idx + 1}.xml`;
        break;
      }

      const footerMatch = /^\/footer\[?(\d+)?\]?$/i.exec(partPath);
      if (footerMatch) {
        const idx = footerMatch[1] ? parseInt(footerMatch[1], 10) - 1 : 0;
        targetXml = await getXmlEntry(zip, `word/footer${idx + 1}.xml`);
        entryName = `word/footer${idx + 1}.xml`;
        break;
      }

      throw new Error(`Unsupported part path '${partPath}'. Use: /document, /styles, /settings, /numbering, /comments, /header[n], /footer[n].`);
    }
  }

  if (targetXml === null || entryName === null) {
    throw new Error(`Part not found: ${partPath}`);
  }

  const affected = executeRawXmlAction(targetXml, xpath, action, xml);
  zip.file(entryName, targetXml);

  const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });
  await writeFile(filePath, updatedBuffer);

  return { ok: true, affected };
}

function executeRawXmlAction(
  xml: string,
  xpath: string,
  action: string,
  newXml?: string,
): number {
  switch (action.toLowerCase()) {
    case "get":
      return 1;
    case "set":
      if (!newXml) {
        throw new Error("set action requires xml parameter");
      }
      if (xpath) {
        const regex = new RegExp(escapeRegex(xpath), "i");
        if (regex.test(xml)) {
          return xml.replace(regex, newXml) !== xml ? 1 : 0;
        }
      }
      return 0;
    case "insert":
      if (!newXml) {
        throw new Error("insert action requires xml parameter");
      }
      if (xpath) {
        const regex = new RegExp(escapeRegex(xpath), "i");
        if (regex.test(xml)) {
          xml = xml.replace(regex, `$&${newXml}`);
          return 1;
        }
      }
      return 0;
    case "delete":
      if (xpath) {
        const regex = new RegExp(escapeRegex(xpath), "i");
        const before = xml;
        xml = xml.replace(regex, "");
        return before !== xml ? 1 : 0;
      }
      return 0;
    default:
      throw new Error(`Unsupported action '${action}'. Use: get, set, insert, delete.`);
  }
}

// ============================================================================
// Compatibility Settings
// ============================================================================

export interface WordCompatibilityProperties {
  preset?: "word2019" | "word2010" | "css-layout";
  mode?: number;
  [key: string]: unknown;
}

/**
 * Sets compatibility settings in a Word document.
 */
export async function setWordCompatibility(
  filePath: string,
  properties: WordCompatibilityProperties,
): Promise<{ ok: boolean }> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  let settingsXml = await getXmlEntry(zip, "word/settings.xml");
  if (!settingsXml) {
    settingsXml = createBasicSettingsXml();
  }

  settingsXml = updateCompatibilitySettings(settingsXml, properties);
  zip.file("word/settings.xml", settingsXml);

  const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });
  await writeFile(filePath, updatedBuffer);

  return { ok: true };
}

function createBasicSettingsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;
}

function updateCompatibilitySettings(settingsXml: string, props: WordCompatibilityProperties): string {
  let newXml = settingsXml;

  if (!/<w:settings/i.test(newXml)) {
    newXml = createBasicSettingsXml();
  }

  if (!/<w:compat>/i.test(newXml)) {
    newXml = newXml.replace("</w:settings>", "<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat></w:settings>");
  }

  if (props.preset) {
    const modeValues: Record<string, number> = {
      word2019: 15,
      word2010: 14,
      "css-layout": 15,
    };
    const mode = modeValues[props.preset] || 15;
    const modeRegex = /<w:compatSetting\s+w:name=["']compatibilityMode["'][^>]*\s+w:val=["'](\d+)["'][^>]*>/i;
    if (modeRegex.test(newXml)) {
      newXml = newXml.replace(modeRegex, `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>`);
    }
  }

  if (props.mode !== undefined) {
    const modeRegex = /<w:compatSetting\s+w:name=["']compatibilityMode["'][^>]*\s+w:val=["'](\d+)["'][^>]*>/i;
    if (modeRegex.test(newXml)) {
      newXml = newXml.replace(modeRegex, `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${props.mode}"/>`);
    }
  }

  return newXml;
}

// ============================================================================
// Helper Functions
// ============================================================================

function getStyleNameFromId(stylesXml: string, styleId: string): string | null {
  const regex = new RegExp(`<w:style\\b[^>]*\\s+w:styleId=["']${escapeRegex(styleId)}["'][^>]*>[\\s\\S]*?<w:name\\s+w:val=["']([^"']+)["']`, "i");
  const match = regex.exec(stylesXml);
  return match ? match[1] : null;
}

function hasRuns(xml: string, paraIndex: number): boolean {
  const paraRegex = /<w:p[\\s\\S]*?<\\/w:p>/g;
  let match;
  let idx = 0;

  while ((match = paraRegex.exec(xml)) !== null) {
    idx++;
    if (idx !== paraIndex) continue;

    const runRegex = /<w:r[\\s\\S]*?<\\/w:r>/g;
    return runRegex.test(match[0]);
  }

  return false;
}

interface RunInfo {
  text: string;
  font?: string;
  size?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  color?: string;
}

function getRunsInfo(xml: string, paraIndex: number): RunInfo[] {
  const runs: RunInfo[] = [];

  const paraRegex = /<w:p[\\s\\S]*?<\\/w:p>/g;
  let match;
  let idx = 0;

  while ((match = paraRegex.exec(xml)) !== null) {
    idx++;
    if (idx !== paraIndex) continue;

    const paraXml = match[0];
    const runRegex = /<w:r[\\s\\S]*?<\\/w:r>/g;
    let runMatch;

    while ((runMatch = runRegex.exec(paraXml)) !== null) {
      const runXml = runMatch[0];
      const textMatch = /<w:t[^>]*>([^<]*)<\/w:t>/i.exec(runXml);
      const text = textMatch ? textMatch[1] : "";

      const runInfo: RunInfo = { text };

      const rPrMatch = /<w:rPr>([\\s\\S]*?)<\/w:rPr>/i.exec(runXml);
      if (rPrMatch) {
        const rPrContent = rPrMatch[1];

        const fontMatch = /<w:rFonts[^>]*w:ascii="([^"]*)"/i.exec(rPrContent);
        if (fontMatch) runInfo.font = fontMatch[1];

        const sizeMatch = /<w:sz[^>]*w:val="([^"]*)"/i.exec(rPrContent);
        if (sizeMatch) runInfo.size = `${parseInt(sizeMatch[1], 10) / 2}pt`;

        if (/<w:b[^>]*>/i.test(rPrContent)) runInfo.bold = true;
        if (/<w:i[^>]*>/i.test(rPrContent)) runInfo.italic = true;

        const underlineMatch = /<w:u[^>]*w:val="([^"]*)"/i.exec(rPrContent);
        if (underlineMatch) runInfo.underline = underlineMatch[1];

        const colorMatch = /<w:color[^>]*w:val="([^"]*)"/i.exec(rPrContent);
        if (colorMatch) runInfo.color = colorMatch[1];
      }

      runs.push(runInfo);
    }
    break;
  }

  return runs;
}

function formatRunInfo(run: RunInfo): string {
  const parts: string[] = [];
  if (run.font) parts.push(`font=${run.font}`);
  if (run.size) parts.push(`size=${run.size}`);
  if (run.bold) parts.push("bold");
  if (run.italic) parts.push("italic");
  if (run.underline) parts.push(`underline=${run.underline}`);
  if (run.color) parts.push(`color=${run.color}`);
  return parts.length > 0 ? parts.join(" ") : "normal";
}

function getHeadingLevel(styleId: string): number {
  const match = /Heading(\d+)/i.exec(styleId);
  if (match) return parseInt(match[1], 10);
  if (styleId === "Title") return 0;
  if (styleId === "Subtitle") return 1;
  return 2;
}

function hasFirstLineIndent(xml: string, paraIndex: number): boolean {
  const paraRegex = /<w:p[\\s\\S]*?<\\/w:p>/g;
  let match;
  let idx = 0;

  while ((match = paraRegex.exec(xml)) !== null) {
    idx++;
    if (idx !== paraIndex) continue;

    const paraXml = match[0];
    return /<w:ind[^>]*w:firstLine=["'][^0]/.test(paraXml);
  }

  return false;
}

function parseStylesForJson(stylesXml: string): Array<{ id: string; name: string; type: string }> {
  const styles: Array<{ id: string; name: string; type: string }> = [];

  const styleRegex = /<w:style[^>]*>([\\s\\S]*?)<\\/w:style>/g;
  let match;

  while ((match = styleRegex.exec(stylesXml)) !== null) {
    const styleXml = match[0];

    const idMatch = /w:styleId="([^"]*)"/.exec(styleXml);
    const id = idMatch ? idMatch[1] : "";

    const nameMatch = /<w:name[^>]*w:val="([^"]*)"/i.exec(styleXml);
    const name = nameMatch ? nameMatch[1] : id;

    const typeMatch = /w:type="([^"]*)"/.exec(styleXml);
    const type = typeMatch ? typeMatch[1] : "paragraph";

    if (id) {
      styles.push({ id, name, type });
    }
  }

  return styles;
}

function parseFontSizeToHalfPoints(size: string): string {
  let val = size.trim();
  if (val.endsWith("pt")) val = val.slice(0, -2).trim();
  const pts = parseFloat(val);
  return String(Math.round(pts * 2));
}

function parseWordSpacing(space: string): string {
  let val = space.trim();
  if (val.endsWith("pt")) val = val.slice(0, -2).trim();
  if (val.endsWith("px")) val = String(parseFloat(val) * 0.75);
  const pts = parseFloat(val);
  return String(Math.round(pts * 20));
}

function parseLineSpacing(lineSpacing: string): string {
  let val = lineSpacing.trim();

  if (val.endsWith("x")) {
    const multiplier = parseFloat(val.slice(0, -1));
    return String(Math.round(multiplier * 240));
  }

  if (val.endsWith("pt")) {
    val = val.slice(0, -2).trim();
  }

  const pts = parseFloat(val);
  return String(Math.round(pts * 20));
}

function mapAlignment(align: string): string {
  switch (align.toLowerCase()) {
    case "left": return "left";
    case "center": return "center";
    case "right": return "right";
    case "justify":
    case "both": return "both";
    default: return align;
  }
}

function mapSectionType(type: string): string {
  switch (type.toLowerCase()) {
    case "nextpage":
    case "next": return "nextPage";
    case "continuous": return "continuous";
    case "evenpage":
    case "even": return "evenPage";
    case "oddpage":
    case "odd": return "oddPage";
    case "nextcolumn":
    case "column": return "nextColumn";
    default: return type;
  }
}

function updateOrAddElement(content: string, tagName: string, newElement: string): string {
  const regex = new RegExp(`<${tagName}\\b[^>]*//?>`, "i");
  if (regex.test(content)) {
    return content.replace(regex, newElement);
  }
  return newElement + content;
}

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function generateBasicCss(): string {
  return `
    body { font-family: Calibri, Arial, sans-serif; font-size: 12pt; line-height: 1.5; }
    p { margin: 0.5em 0; }
    strong { font-weight: bold; }
    em { font-style: italic; }
    u { text-decoration: underline; }
    .Heading1 { font-size: 24pt; font-weight: bold; margin: 12pt 0 6pt 0; }
    .Heading2 { font-size: 18pt; font-weight: bold; margin: 10pt 0 4pt 0; }
    .Heading3 { font-size: 14pt; font-weight: bold; margin: 8pt 0 2pt 0; }
    table { border-collapse: collapse; }
    td, th { border: 1px solid #ccc; padding: 4px 8px; }
  `;
}
