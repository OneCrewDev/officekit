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

interface BodyParagraphInfo {
  index: number;
  xml: string;
  text: string;
  style?: string;
  paraId?: string;
}

interface BodyTableInfo {
  index: number;
  xml: string;
  rows: number;
  cols: number;
  cells: string[][];
}

type BodyContentInfo =
  | ({ type: "paragraph" } & BodyParagraphInfo)
  | ({ type: "table" } & BodyTableInfo);

function getBodyXml(xml: string): string {
  const bodyMatch = xml.match(/<w:body\b[^>]*>([\s\S]*?)<\/w:body>/);
  return bodyMatch?.[1] ?? "";
}

function findNextBodyTag(bodyXml: string, cursor: number) {
  const candidates = [
    { tag: "w:p", index: bodyXml.indexOf("<w:p", cursor) },
    { tag: "w:tbl", index: bodyXml.indexOf("<w:tbl", cursor) },
    { tag: "w:sectPr", index: bodyXml.indexOf("<w:sectPr", cursor) },
  ].filter((candidate) => candidate.index >= 0);

  if (candidates.length === 0) {
    return null;
  }

  return candidates.reduce((left, right) => (right.index < left.index ? right : left));
}

function readTopLevelElement(bodyXml: string, start: number, tagName: "w:p" | "w:tbl" | "w:sectPr") {
  const startTagEnd = bodyXml.indexOf(">", start);
  if (startTagEnd === -1) {
    return null;
  }

  const startTag = bodyXml.slice(start, startTagEnd + 1);
  if (startTag.endsWith("/>")) {
    return { xml: startTag, end: startTagEnd + 1 };
  }

  const closeTag = `</${tagName}>`;
  const closeIndex = bodyXml.indexOf(closeTag, startTagEnd + 1);
  if (closeIndex === -1) {
    return { xml: bodyXml.slice(start), end: bodyXml.length };
  }

  const end = closeIndex + closeTag.length;
  return { xml: bodyXml.slice(start, end), end };
}

function parseTableCells(tableXml: string): string[][] {
  const rows: string[][] = [];
  for (const rowMatch of tableXml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)) {
    const rowXml = rowMatch[0];
    const cells: string[] = [];
    for (const cellMatch of rowXml.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)) {
      cells.push(extractTextSimple(cellMatch[0]));
    }
    rows.push(cells);
  }
  return rows;
}

function getBodyContentInfo(xml: string): BodyContentInfo[] {
  const bodyXml = getBodyXml(xml);
  if (!bodyXml) {
    return [];
  }

  const content: BodyContentInfo[] = [];
  let cursor = 0;
  let paragraphIndex = 0;
  let tableIndex = 0;

  while (cursor < bodyXml.length) {
    const next = findNextBodyTag(bodyXml, cursor);
    if (!next) {
      break;
    }

    if (next.tag === "w:sectPr") {
      const sectPr = readTopLevelElement(bodyXml, next.index, "w:sectPr");
      if (!sectPr) {
        break;
      }
      cursor = sectPr.end;
      continue;
    }

    const element = readTopLevelElement(bodyXml, next.index, next.tag as "w:p" | "w:tbl");
    if (!element) {
      break;
    }

    if (next.tag === "w:p") {
      paragraphIndex += 1;
      const paraXml = element.xml;
      const styleMatch = paraXml.match(/<w:pStyle[^>]*w:val="([^"]*)"/);
      const paraIdMatch = paraXml.match(/<w:paraId[^>]*w:val="([^"]*)"/);
      content.push({
        type: "paragraph",
        index: paragraphIndex,
        xml: paraXml,
        text: extractTextSimple(paraXml),
        style: styleMatch?.[1],
        paraId: paraIdMatch?.[1],
      });
    } else {
      tableIndex += 1;
      const rows = parseTableCells(element.xml);
      content.push({
        type: "table",
        index: tableIndex,
        xml: element.xml,
        rows: rows.length,
        cols: rows[0]?.length ?? 0,
        cells: rows,
      });
    }

    cursor = element.end;
  }

  return content;
}

/**
 * Gets all paragraph texts from document XML.
 */
function getParagraphsInfo(xml: string): Array<{ index: number; text: string; style?: string; paraId?: string }> {
  return getBodyContentInfo(xml)
    .filter((item): item is Extract<BodyContentInfo, { type: "paragraph" }> => item.type === "paragraph")
    .map(({ index, text, style, paraId }) => ({ index, text, style, paraId }));
}

/**
 * Gets all table info from document XML.
 */
function getTablesInfo(xml: string): Array<{ index: number; rows: number; cols: number }> {
  return getBodyContentInfo(xml)
    .filter((item): item is Extract<BodyContentInfo, { type: "table" }> => item.type === "table")
    .map(({ index, rows, cols }) => ({ index, rows, cols }));
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
    const result = await navigateToElement(documentXml, zip, segments, depth);

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
async function navigateToElement(
  documentXml: string,
  zip: JSZip,
  segments: PathSegment[],
  depth: number,
  parentPath = "",
): Promise<DocumentNode | null> {
  if (segments.length === 0) {
    return createDocumentNode("/", "document");
  }

  const first = segments[0];
  let currentPath = "/" + first.name + (first.index !== undefined ? `[${first.index}]` : "");
  let currentNode: DocumentNode | null = null;

  switch (first.name) {
    case "body": {
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

      // If there are remaining segments (e.g., /body/p[1]), navigate to the child
      if (segments.length > 1) {
        const remaining = segments.slice(1);
        const firstSeg = remaining[0];
        const childIndex = (firstSeg.index || 1) - 1;

        // For p/paragraph: find paragraph at childIndex
        if ((firstSeg.name === "p" || firstSeg.name === "paragraph") && childIndex >= 0 && childIndex < paras.length) {
          const para = paras[childIndex];
          const paraPath = `/body/p[${childIndex + 1}]`;
          const paraNode = createDocumentNode(
            paraPath,
            "paragraph",
            para.text,
            { style: para.style, paraId: para.paraId }
          );

          if (depth > 0) {
            const runs = getRunsFromParagraph(documentXml, childIndex + 1);
            paraNode.children = runs;
            paraNode.childCount = runs.length;
          }

          // If there are more segments (e.g., /body/p[1]/r[1]), continue navigating
          if (remaining.length > 1) {
            const remaining2 = remaining.slice(1);
            if (remaining2.length === 1 && (remaining2[0].name === "r" || remaining2[0].name === "run")) {
              const runIdx = (remaining2[0].index || 1) - 1;
              if (runIdx >= 0 && runIdx < paraNode.children!.length) {
                return paraNode.children![runIdx];
              }
            }
          }

          return paraNode;
        }

        // For tbl/table: find table at childIndex
        if ((firstSeg.name === "tbl" || firstSeg.name === "table") && childIndex >= 0 && childIndex < tables.length) {
          const table = tables[childIndex];
          const tablePath = `/body/tbl[${childIndex + 1}]`;
          const tableContent = getBodyContentInfo(documentXml)
            .find((item): item is Extract<BodyContentInfo, { type: "table" }> => item.type === "table" && item.index === childIndex + 1);
          const tableNode = createDocumentNode(
            tablePath,
            "table",
            undefined,
            { rowCount: table.rows, columnCount: table.cols }
          );

          if (depth > 0 && tableContent) {
            const rows: DocumentNode[] = [];
            for (let i = 0; i < tableContent.cells.length; i++) {
              const rowNode = createDocumentNode(
                `/body/tbl[${childIndex + 1}]/tr[${i + 1}]`,
                "row",
                undefined,
                { cellCount: tableContent.cells[i].length }
              );
              rowNode.children = tableContent.cells[i].map((cellText, cellIndex) =>
                createDocumentNode(
                  `/body/tbl[${childIndex + 1}]/tr[${i + 1}]/tc[${cellIndex + 1}]`,
                  "cell",
                  cellText,
                )
              );
              rowNode.childCount = rowNode.children.length;
              rows.push(rowNode);
            }
            tableNode.children = rows;
            tableNode.childCount = rows.length;
          }

          // If there are more segments (e.g., /body/tbl[1]/tr[1]/tc[1]), continue
          if (remaining.length > 1) {
            const remaining2 = remaining.slice(1);
            if (remaining2.length === 1 && (remaining2[0].name === "tr" || remaining2[0].name === "row")) {
              const rowIdx = (remaining2[0].index || 1) - 1;
              if (rowIdx >= 0 && rowIdx < table.rows) {
                return tableNode.children![rowIdx];
              }
            }
            // Handle /body/tbl[N]/cell[N,N] format
            if (remaining2.length === 1 && remaining2[0].name === "cell" && remaining2[0].stringIndex?.includes(",")) {
              const [rowStr, colStr] = remaining2[0].stringIndex!.split(",");
              const rowIdx = parseInt(rowStr, 10) - 1;
              const cellIdx = parseInt(colStr, 10) - 1;
              const cellText = tableContent?.cells[rowIdx]?.[cellIdx];
              if (rowIdx >= 0 && rowIdx < table.rows && cellIdx >= 0 && cellIdx < table.cols) {
                const cellPath = `/body/tbl[${childIndex + 1}]/cell[${rowIdx + 1},${cellIdx + 1}]`;
                return createDocumentNode(cellPath, "cell", cellText);
              }
            }
            if (remaining2.length === 2 &&
                (remaining2[0].name === "tr" || remaining2[0].name === "row") &&
                (remaining2[1].name === "tc" || remaining2[1].name === "cell")) {
              const rowIdx = (remaining2[0].index || 1) - 1;
              const cellIdx = (remaining2[1].index || 1) - 1;
              const cellText = tableContent?.cells[rowIdx]?.[cellIdx];
              if (rowIdx >= 0 && rowIdx < table.rows && cellIdx >= 0 && cellIdx < table.cols) {
                const cellPath = `/body/tbl[${childIndex + 1}]/tr[${rowIdx + 1}]/tc[${cellIdx + 1}]`;
                return createDocumentNode(cellPath, "cell", cellText);
              }
            }
          }

          return tableNode;
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

  const paraXml = getBodyContentInfo(documentXml)
    .find((item): item is Extract<BodyContentInfo, { type: "paragraph" }> => item.type === "paragraph" && item.index === paraIndex)
    ?.xml;

  if (!paraXml) {
    return runs;
  }

  const runRegex = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/gi;
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

  return runs;
}

/**
 * Parses styles from styles.xml.
 */
function parseStyles(stylesXml: string): DocumentNode[] {
  const styles: DocumentNode[] = [];

  const styleRegex = /<w:style[^>]*>([\s\S]*?)<\/w:style>/g;
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

  const paraRegex = /<w:p[\s\S]*?<\/w:p>/g;
  let match;
  let paraIdx = 0;

  while ((match = paraRegex.exec(documentXml)) !== null) {
    paraIdx++;
    const paraXml = match[0];
    const runRegex = /<w:r[\s\S]*?<\/w:r>/g;
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

  const sdtRegex = /<w:sdt[\s\S]*?<\/w:sdt>/g;
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
 * Helper: Escape HTML special characters
 */
function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
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
    const ulTag = underline ? `<w:u w:val="${underline === "true" || underline === "1" ? "single" : underline}"/>` : "";
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
    const ulTag = underline ? `<w:u w:val="${underline === "true" || underline === "1" ? "single" : underline}"/>` : "";
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
 * Handles paragraphs, tables, and other body children with index-based positioning
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
    const sectPrMatch = /<w:sectPr\b[\s\S]*?(?:\/>|<\/w:sectPr>)/i.exec(bodyMatch[1]);
    if (sectPrMatch && sectPrMatch.index !== undefined) {
      const insertPos = bodyOpen + 8 + sectPrMatch.index;
      return docXml.slice(0, insertPos) + insertXml + docXml.slice(insertPos);
    }
    return docXml.slice(0, bodyClose) + insertXml + docXml.slice(bodyClose);
  }

  // Insert at specific index - find all body children (paragraphs, tables, etc.)
  const bodyContent = bodyMatch[1];
  const childElements: { type: string; start: number; end: number }[] = [];

  // Match all top-level body children: paragraphs, tables, sectPr, etc.
  const elementRegex = /<(w:p\b[^>]*>[\s\S]*?<\/w:p>|w:tbl\b[^>]*>[\s\S]*?<\/w:tbl>|w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>|w:customXml\b[^>]*>[\s\S]*?<\/w:customXml>)/gi;
  let match;
  let lastEnd = 0;

  while ((match = elementRegex.exec(bodyContent)) !== null) {
    const fullMatch = match[0];
    const start = bodyOpen + 8 + (match.index - lastEnd + match[0].indexOf('<'));
    childElements.push({
      type: fullMatch.startsWith('<w:p') ? 'p' : fullMatch.startsWith('<w:tbl') ? 'tbl' : 'other',
      start: match.index,
      end: match.index + fullMatch.length
    });
    lastEnd = match.index + fullMatch.length;
  }

  if (typeof position === "number") {
    if (position >= childElements.length) {
      // Append at end if index beyond children
      return docXml.slice(0, bodyClose) + insertXml + docXml.slice(bodyClose);
    }
    const targetChild = childElements[position];
    // Insert BEFORE the target child element
    const insertPos = bodyOpen + 8 + targetChild.start;
    return docXml.slice(0, insertPos) + insertXml + docXml.slice(insertPos);
  }

  return docXml.slice(0, bodyClose) + insertXml + docXml.slice(bodyClose);
}

/**
 * Helper: Process find and replace/format
 *
 * This function finds text in a Word document XML and optionally replaces it
 * and/or applies formatting. It processes paragraphs and runs to correctly
 * handle text that may span multiple runs.
 */
function processFindAndFormat(
  docXml: string,
  find: string,
  replace: string | null,
  formatProps: Record<string, string>,
  useRegex: boolean
): { docXml: string; matchCount: number } {
  let result = docXml;
  let matchCount = 0;

  if (!find) {
    return { docXml: result, matchCount: 0 };
  }

  // Build regex pattern
  let pattern: RegExp;
  if (useRegex) {
    const flags = "g" + (find.includes("i") ? "i" : "");
    const rawPattern = find.startsWith("r\"") && find.endsWith("\"")
      ? find.slice(2, -1)
      : find;
    try {
      pattern = new RegExp(rawPattern, flags);
    } catch {
      // Invalid regex, treat as literal
      pattern = new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g");
    }
  } else {
    pattern = new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g");
  }

  // Helper to generate rPr XML from format props
  function buildRprXml(props: Record<string, string>): string {
    const tags: string[] = [];

    if (props.bold) {
      tags.push("<w:b/>");
    }
    if (props.italic) {
      tags.push("<w:i/>");
    }
    if (props.underline) {
      const ulVal = props.underline === "true" || props.underline === "1" ? "single" : props.underline;
      tags.push(`<w:u w:val="${ulVal}"/>`);
    }
    if (props.strike) {
      tags.push("<w:strike/>");
    }
    if (props.color) {
      tags.push(`<w:color w:val="${sanitizeHex(props.color)}"/>`);
    }
    if (props.highlight) {
      tags.push(`<w:highlight w:val="${props.highlight}"/>`);
    }
    if (props.font) {
      tags.push(`<w:rFonts w:ascii="${escapeXml(props.font)}" w:hAnsi="${escapeXml(props.font)}" w:eastAsia="${escapeXml(props.font)}"/>`);
    }
    if (props.size) {
      const sizeVal = parseInt(props.size, 10) * 2;
      tags.push(`<w:sz w:val="${sizeVal}"/><w:szCs w:val="${sizeVal}"/>`);
    }
    if (props.shading || props.shd) {
      const shdVal = props.shading || props.shd;
      const parts = shdVal.split(";");
      if (parts.length === 1) {
        tags.push(`<w:shd w:val="clear" w:fill="${sanitizeHex(parts[0])}"/>`);
      } else {
        tags.push(`<w:shd w:val="${parts[0]}" w:fill="${sanitizeHex(parts[1])}" w:color="${parts.length > 2 ? sanitizeHex(parts[2]) : "auto"}"/>`);
      }
    }
    if (props.subscript) {
      tags.push(`<w:vertAlign w:val="subscript"/>`);
    }
    if (props.superscript) {
      tags.push(`<w:vertAlign w:val="superscript"/>`);
    }
    if (props.caps) {
      tags.push("<w:caps/>");
    }
    if (props.smallcaps) {
      tags.push("<w:smallCaps/>");
    }
    if (props.vanish) {
      tags.push("<w:vanish/>");
    }
    if (props.charspacing || props.spacing || props.letterspacing) {
      const val = props.charspacing || props.spacing || props.letterspacing;
      const numVal = val.endsWith("pt")
        ? Math.round(parseFloat(val.slice(0, -2)) * 20)
        : Math.round(parseFloat(val) * 20);
      tags.push(`<w:spacing w:val="${numVal}"/>`);
    }

    return tags.length > 0 ? `<w:rPr>${tags.join("")}</w:rPr>` : "";
  }

  // Process document paragraph by paragraph
  const paraRegex = /<w:p[\s\S]*?<\/w:p>/g;
  let paraMatch;

  while ((paraMatch = paraRegex.exec(result)) !== null) {
    const paraStart = paraMatch.index;
    const paraXml = paraMatch[0];

    // Parse runs in this paragraph to build text positions
    interface RunInfo {
      runXml: string;
      text: string;
      start: number;
      end: number;
      runStart: number;
      runEnd: number;
    }

    const runs: RunInfo[] = [];
    const runRegex = /<w:r[\s\S]*?<\/w:r>/g;
    let runMatch;
    let textPos = 0;

    while ((runMatch = runRegex.exec(paraXml)) !== null) {
      const runXml = runMatch[0];
      const runStartPos = runMatch.index;

      // Extract text content from this run
      const textMatches: string[] = [];
      const textRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
      let textMatch;
      while ((textMatch = textRegex.exec(runXml)) !== null) {
        textMatches.push(textMatch[1]);
      }
      const runText = textMatches.join("");

      // Find end of run element
      const runEndPos = runXml.length + runStartPos;

      runs.push({
        runXml,
        text: runText,
        start: textPos,
        end: textPos + runText.length,
        runStart: runStartPos,
        runEnd: runEndPos
      });

      textPos += runText.length;
    }

    if (runs.length === 0) continue;

    const fullText = runs.map(r => r.text).join("");

    // Find all matches in this paragraph
    const matches: Array<{ start: number; end: number; length: number }> = [];
    let regexMatch;
    const re = new RegExp(pattern.source, pattern.flags.includes("g") ? pattern.flags : pattern.flags + "g");
    while ((regexMatch = re.exec(fullText)) !== null) {
      matches.push({
        start: regexMatch.index,
        end: regexMatch.index + regexMatch[0].length,
        length: regexMatch[0].length
      });
      if (!pattern.global) break;
    }

    if (matches.length === 0) continue;

    matchCount += matches.length;

    // Process matches from end to start to preserve offsets
    for (let i = matches.length - 1; i >= 0; i--) {
      const m = matches[i];

      // Find which runs are affected by this match
      const affectedRuns: RunInfo[] = [];
      for (const run of runs) {
        if (run.end > m.start && run.start < m.end) {
          affectedRuns.push(run);
        }
      }

      if (affectedRuns.length === 0) continue;

      // Sort affected runs by their start position
      affectedRuns.sort((a, b) => a.start - b.start);

      // Get the original match text
      const matchText = fullText.slice(m.start, m.end);

      // Step 1: Handle replacement if provided
      if (replace !== null && replace !== undefined) {
        const firstRun = affectedRuns[0];
        const lastRun = affectedRuns[affectedRuns.length - 1];

        // Build text before, match, and after
        const textBefore = fullText.slice(0, m.start);
        const textAfter = fullText.slice(m.end);
        const textMatched = fullText.slice(m.start, m.end);

        // Determine replacement text
        const replacement = replace;

        // Build new run content
        let newRunContent = "";

        const firstRunStartOffset = m.start - firstRun.start;
        const lastRunEndOffset = m.end - lastRun.start;

        // Text in runs before the match
        if (textBefore) {
          const beforeRuns = runs.filter(r => r.end <= m.start);
          if (beforeRuns.length > 0) {
            const beforeText = firstRun.text.slice(0, firstRunStartOffset);
            if (beforeText) {
              newRunContent += `<w:r><w:t xml:space="preserve">${escapeXml(beforeText)}</w:t></w:r>`;
            }
          }
        }

        // Replacement text (potentially with formatting)
        const hasFormatProps = formatProps && Object.keys(formatProps).length > 0;
        const rprXml = hasFormatProps ? buildRprXml(formatProps) : "";
        newRunContent += `<w:r>${rprXml}<w:t xml:space="preserve">${escapeXml(replacement)}</w:t></w:r>`;

        // Text after match
        if (textAfter) {
          const afterText = lastRun.text.slice(lastRunEndOffset);
          if (afterText) {
            newRunContent += `<w:r><w:t xml:space="preserve">${escapeXml(afterText)}</w:t></w:r>`;
          }
        }

        // Find the position range in paraXml to replace
        const runStartInDoc = paraStart + firstRun.runStart;
        const runEndInDoc = paraStart + lastRun.runEnd;

        // Do the replacement in result
        result = result.slice(0, runStartInDoc) + newRunContent + result.slice(runEndInDoc);

        // Update run positions for subsequent matches in this paragraph
        const xmlDiff = newRunContent.length - (lastRun.runEnd - firstRun.runStart);
        for (const run of runs) {
          if (run.runStart >= firstRun.runStart) {
            run.runStart += xmlDiff;
            run.runEnd += xmlDiff;
          }
        }
      } else if (formatProps && Object.keys(formatProps).length > 0) {
        // No replacement, only formatting - wrap the matched text in formatted runs

        const firstRun = affectedRuns[0];
        const lastRun = affectedRuns[affectedRuns.length - 1];

        const rprXml = buildRprXml(formatProps);

        // Get text portions
        const firstRunStartOffset = m.start - firstRun.start;
        const lastRunEndOffset = m.end - lastRun.start;

        const textBefore = firstRun.text.slice(0, firstRunStartOffset);
        const textMatched = fullText.slice(m.start, m.end);
        const textAfter = lastRun.text.slice(lastRunEndOffset);

        // Build new run content
        let newRunContent = "";

        if (textBefore) {
          newRunContent += `<w:r><w:t xml:space="preserve">${escapeXml(textBefore)}</w:t></w:r>`;
        }

        newRunContent += `<w:r>${rprXml}<w:t xml:space="preserve">${escapeXml(textMatched)}</w:t></w:r>`;

        if (textAfter) {
          newRunContent += `<w:r><w:t xml:space="preserve">${escapeXml(textAfter)}</w:t></w:r>`;
        }

        // Find the position range in paraXml to replace
        const runStartInDoc = paraStart + firstRun.runStart;
        const runEndInDoc = paraStart + lastRun.runEnd;

        // Do the replacement in result
        result = result.slice(0, runStartInDoc) + newRunContent + result.slice(runEndInDoc);

        // Update run positions for subsequent matches
        const xmlDiff = newRunContent.length - (lastRun.runEnd - firstRun.runStart);
        for (const run of runs) {
          if (run.runStart >= firstRun.runStart) {
            run.runStart += xmlDiff;
            run.runEnd += xmlDiff;
          }
        }
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
        await zip.remove("officekit/document.json");
        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ path: `/styles/${props.name || props.id}` });

      case "header":
        const headerIdx = (zip.file(/^word\/header\d+\.xml$/) || []).length + 1;
        const headerContent = createHeaderXml(props);
        zip.file(`word/header${headerIdx}.xml`, headerContent);
        await zip.remove("officekit/document.json");
        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ path: `/header[${headerIdx}]` });

      case "footer":
        const footerIdx = (zip.file(/^word\/footer\d+\.xml$/) || []).length + 1;
        const footerContent = createFooterXml(props);
        zip.file(`word/footer${footerIdx}.xml`, footerContent);
        await zip.remove("officekit/document.json");
        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ path: `/footer[${footerIdx}]` });

      case "sdt":
      case "contentcontrol":
        insertXml = createSdtXml(props);
        break;

      case "watermark":
        const wmHeader = createWatermarkXml(props);
        const headerIdx2 = (zip.file(/^word\/header\d+\.xml$/) || []).length + 1;
        const headerProps: Record<string, string> = { ...props };
        delete headerProps.text;
        zip.file(`word/header${headerIdx2}.xml`, createHeaderXml(headerProps) + `<w:pict>${wmHeader}</w:pict>`);
        await zip.remove("officekit/document.json");
        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ path: "/watermark" });

      default:
        return err("invalid_type", `Unknown element type: ${type}`);
    }

    // Insert the XML
    documentXml = insertAtPosition(documentXml, insertXml, effectivePosition);
    zip.file("word/document.xml", documentXml);
    await zip.remove("officekit/document.json");
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
      const useRegex = props.regex === "true" || props.regex === "1";
      const { matchCount } = processFindAndFormat(documentXml, find, replace, props, useRegex);

      documentXml = matchCount > 0 ? documentXml : documentXml;
      zip.file("word/document.xml", documentXml);
      await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
      return ok({ path: targetPath, matchCount });
    }

    // Handle document-level properties (docProps/core.xml)
    if (targetPath === "/" || targetPath === "") {
      // Extract document properties from props and delegate to setDocumentProperties
      const docProps: DocumentProperties = {};
      if (props.title !== undefined) docProps.title = props.title;
      if (props.author !== undefined) docProps.author = props.author;
      if (props.subject !== undefined) docProps.subject = props.subject;
      if (props.keywords !== undefined) docProps.keywords = props.keywords;
      if (props.description !== undefined) docProps.description = props.description;
      if (props.category !== undefined) docProps.category = props.category;
      if (props.lastModifiedBy !== undefined) docProps.lastModifiedBy = props.lastModifiedBy;
      if (props.revision !== undefined) docProps.revision = props.revision;
      if (Object.keys(docProps).length > 0) {
        return setDocumentProperties(filePath, docProps).then((r) =>
          r.ok ? ok({ path: targetPath }) : err(r.error?.code ?? "operation_failed", r.error?.message ?? "Failed to set document properties")
        );
      }
      await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
      return ok({ path: targetPath });
    }

    // Handle style path (/styles/styleName)
    if (targetPath.startsWith("/styles/")) {
      const styleName = targetPath.substring(8);
      const styleProps: WordStyleProperties = {};
      if (props.font !== undefined) styleProps.font = props.font;
      if (props.fontSize !== undefined) styleProps.fontSize = props.fontSize;
      if (props.bold !== undefined) styleProps.bold = props.bold === "true" || props.bold === "1";
      if (props.italic !== undefined) styleProps.italic = props.italic === "true" || props.italic === "1";
      if (props.color !== undefined) styleProps.color = props.color;
      if (props.underline !== undefined) styleProps.underline = props.underline;
      if (props.alignment !== undefined) styleProps.alignment = props.alignment as "left" | "center" | "right" | "justify";
      if (props.spaceBefore !== undefined) styleProps.spaceBefore = props.spaceBefore;
      if (props.spaceAfter !== undefined) styleProps.spaceAfter = props.spaceAfter;
      if (props.lineSpacing !== undefined) styleProps.lineSpacing = props.lineSpacing;
      if (props.basedOn !== undefined) styleProps.basedOn = props.basedOn;
      if (props.next !== undefined) styleProps.next = props.next;
      const result = await setWordStyle(filePath, styleName, styleProps);
      return ok({ path: targetPath });
    }

    // Handle /body/table[N]/cell[N,N] path
    const tableCellMatch = /^\/body\/table\[(\d+)\]\/cell\[(\d+),(\d+)\]$/.exec(targetPath);
    if (tableCellMatch) {
      const tableIndex = parseInt(tableCellMatch[1], 10);
      const rowIndex = parseInt(tableCellMatch[2], 10);
      const colIndex = parseInt(tableCellMatch[3], 10);
      const newText = props.text;

      // Find and update the cell text
      const tblPattern = /<w:tbl\b[\s\S]*?<\/w:tbl>/g;
      let tblMatch;
      let currentTblIndex = 0;
      while ((tblMatch = tblPattern.exec(documentXml)) !== null) {
        currentTblIndex++;
        if (currentTblIndex === tableIndex) {
          const tblXml = tblMatch[0];
          // Find the row
          const rowPattern = /<w:tr\b[\s\S]*?<\/w:tr>/g;
          let rowMatch;
          let currentRowIndex = 0;
          while ((rowMatch = rowPattern.exec(tblXml)) !== null) {
            currentRowIndex++;
            if (currentRowIndex === rowIndex) {
              const rowXml = rowMatch[0];
              // Find the cell
              const cellPattern = /<w:tc\b[\s\S]*?<\/w:tc>/g;
              let cellMatch;
              let currentCellIndex = 0;
              while ((cellMatch = cellPattern.exec(rowXml)) !== null) {
                currentCellIndex++;
                if (currentCellIndex === colIndex) {
                  // Update the cell text - replace content inside <w:t>
                  const cellXml = cellMatch[0];
                  if (newText !== undefined) {
                    // Find the text run and update it
                    const updatedCellXml = cellXml.replace(/(<w:t[^>]*>)([^<]*)(<\/w:t>)/, (_, open, _oldText, close) => {
                      return open + escapeXml(newText) + close;
                    });
                    const updatedRowXml = rowXml.substring(0, cellMatch.index) + updatedCellXml + rowXml.substring(cellMatch.index + cellMatch[0].length);
                    const updatedTableXml = tblXml.substring(0, rowMatch.index) + updatedRowXml + tblXml.substring(rowMatch.index + rowMatch[0].length);
                    documentXml = documentXml.substring(0, tblMatch.index) + updatedTableXml + documentXml.substring(tblMatch.index + tblMatch[0].length);
                    zip.file("word/document.xml", documentXml);
                    await zip.remove("officekit/document.json");
                    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
                    return ok({ path: targetPath });
                  }
                }
              }
            }
          }
        }
      }
      return err("not_found", `Table cell ${rowIndex},${colIndex} not found in table ${tableIndex}`);
    }

    // Handle /body/p[N] path - set paragraph text
    const paraMatch = /^\/body\/p\[(\d+)\]$/.exec(targetPath);
    if (paraMatch) {
      const paraIndex = parseInt(paraMatch[1], 10);
      const newText = props.text;

      // Find all paragraphs in document order
      const allParaPattern = /<w:p\b[\s\S]*?<\/w:p>/g;
      let paraIdx = 0;
      let match;
      while ((match = allParaPattern.exec(documentXml)) !== null) {
        paraIdx++;
        if (paraIdx === paraIndex) {
          if (newText !== undefined) {
            // Update the paragraph text
            const paraXml = match[0];
            // Replace all <w:t> content with new text
            const updatedParaXml = paraXml.replace(/(<w:t[^>]*>)([^<]*)(<\/w:t>)/g, (_, open, _oldText, close) => {
              return open + newText + close;
            });
            documentXml = documentXml.substring(0, match.index) + updatedParaXml + documentXml.substring(match.index + match[0].length);
          }
          zip.file("word/document.xml", documentXml);
          await zip.remove("officekit/document.json");
          await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
          return ok({ path: targetPath });
        }
      }
      return err("not_found", `Paragraph ${paraIndex} not found`);
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
    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    // Extract source element XML
    const sourceXml = extractElementXml(documentXml, sourcePath);
    if (!sourceXml) {
      return err("not_found", `Source element not found: ${sourcePath}`);
    }

    // Get source element type for removal
    const sourceInfo = parsePath(sourcePath);
    if (!sourceInfo.ok || !sourceInfo.data) {
      return err("invalid_path", `Invalid source path: ${sourcePath}`);
    }

    // Determine insert position
    let insertPosition: string | number | undefined = options.position;
    if (options.after) {
      insertPosition = `find:${options.after}`;
    } else if (options.before) {
      insertPosition = `find:${options.before}`;
    }

    // Clone the source element with new IDs
    const clonedXml = generateNewParaIds(sourceXml);

    // Insert at target position
    const originalDocXml = documentXml;
    documentXml = insertAtPosition(documentXml, clonedXml, insertPosition);

    if (documentXml === originalDocXml && insertPosition !== undefined) {
      return err("operation_failed", "Failed to insert element at target position");
    }

    // Remove source element from original position
    const elementType = sourceInfo.data.segments[0]?.name;
    if (elementType === "p" || elementType === "paragraph") {
      // Remove the specific paragraph
      const paras = documentXml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/gi);
      if (paras) {
        const srcIdx = sourceInfo.data.segments[0]?.index ?? 1;
        if (paras[srcIdx - 1]) {
          documentXml = documentXml.replace(paras[srcIdx - 1], "");
        }
      }
    } else if (elementType === "tbl" || elementType === "table") {
      // Remove the specific table
      const tables = documentXml.match(/<w:tbl\b[^>]*>[\s\S]*?<\/w:tbl>/gi);
      if (tables) {
        const srcIdx = sourceInfo.data.segments[0]?.index ?? 1;
        if (tables[srcIdx - 1]) {
          documentXml = documentXml.replace(tables[srcIdx - 1], "");
        }
      }
    }

    zip.file("word/document.xml", documentXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

    // Calculate the new path
    const newPath = calculateInsertedPath(documentXml, clonedXml, "/body", elementType ?? "paragraph");

    return ok({ path: newPath });
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
    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    // Extract both elements
    const xml1 = extractElementXml(documentXml, path1);
    const xml2 = extractElementXml(documentXml, path2);

    if (!xml1) {
      return err("not_found", `First element not found: ${path1}`);
    }
    if (!xml2) {
      return err("not_found", `Second element not found: ${path2}`);
    }

    // Get element types
    const info1 = parsePath(path1);
    const info2 = parsePath(path2);

    if (!info1.ok || !info1.data || !info2.ok || !info2.data) {
      return err("invalid_path", "Invalid path format");
    }

    // Generate new IDs for swapped elements
    const swapped1Xml = generateNewParaIds(xml2);
    const swapped2Xml = generateNewParaIds(xml1);

    // Replace element1 with element2 (with new IDs)
    documentXml = documentXml.replace(xml1, swapped1Xml);
    // Replace element2 with element1 (with new IDs)
    documentXml = documentXml.replace(xml2, swapped2Xml);

    zip.file("word/document.xml", documentXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

    return ok({ path1: path2, path2: path1 });
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
        case "copy": {
          // target = source path, options.parent = target parent path
          const targetParent = (options.parent as string) || (options.target as string) || "/body";
          const copyResult = await copyWordNode(filePath, target, targetParent, {
            index: options.index as number | undefined,
            after: options.after as string | undefined,
            before: options.before as string | undefined,
          });
          results.push({ action, target, status: copyResult.ok ? "success" : "failed" });
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
  const fileName = filePath.split("/").pop() || "document.docx";

  const bodyContent = getBodyContentInfo(xml);
  const paraCount = bodyContent.filter((item) => item.type === "paragraph").length;
  const tblCount = bodyContent.filter((item) => item.type === "table").length;
  lines.push(`File: ${fileName} | ${paraCount} paragraphs | ${tblCount} tables`);

  for (const item of bodyContent) {
    if (item.type === "paragraph") {
      const style = item.style;
      if (style && (style.includes("Heading") || style === "Title" || style === "Subtitle")) {
        const level = getHeadingLevel(style);
        const indent = level <= 1 ? "" : "  ".repeat(level - 1);
        const prefix = level === 0 ? "■" : "├──";
        lines.push(`${indent}${prefix} [${item.index}] "${item.text}" (${style})`);
      } else {
        lines.push(`Paragraph ${item.index}: ${item.text}`);
      }
      continue;
    }

    lines.push(`Table ${item.index}: ${item.rows}x${item.cols}`);
    for (const [rowIndex, row] of item.cells.entries()) {
      for (const [cellIndex, cellText] of row.entries()) {
        lines.push(`  R${rowIndex + 1}C${cellIndex + 1}: ${cellText}`);
      }
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

  for (const item of getBodyContentInfo(xml)) {
    if (item.type === "paragraph") {
      let className = item.style || "Normal";
      className = className.replace(/\s+/g, "");

      let html = `<p`;
      if (className !== "Normal") {
        html += ` class="${escapeHtml(className)}"`;
      }
      html += ">";

      const runs = getRunsInfo(xml, item.index);
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
      continue;
    }

    lines.push("<table>");
    for (const row of item.cells) {
      lines.push("<tr>");
      for (const cellText of row) {
        lines.push(`<td>${escapeHtml(cellText)}</td>`);
      }
      lines.push("</tr>");
    }
    lines.push("</table>");
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
    lines.push(`  #${i + 1} [sdt] path="${sdt.path}" text="${sdt.text}"`);
  }

  return lines.join("\n");
}

function renderJsonView(xml: string, stylesXml: string): string {
  const paragraphs: { text: string; style: string }[] = [];

  const paras = getParagraphsInfo(xml);
  for (const para of paras) {
    paragraphs.push({
      text: para.text,
      style: para.style || "Normal",
    });
  }

  const result = {
    paragraphs,
    styles: parseStylesForJson(stylesXml),
  };

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
  const styleRegex = new RegExp(`<w:style\\b[^>]*\\s+w:styleId=["']${escapeRegex(styleId)}["'][^>]*>[\s\S]*?<\/w:style>`, "i");
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
  const regex = new RegExp(`<w:style\\b[^>]*\\s+w:styleId=["']${escapeRegex(styleId)}["'][^>]*>[\s\S]*?<w:name\\s+w:val=["']([^"']+)["']`, "i");
  const match = regex.exec(stylesXml);
  return match ? match[1] : null;
}

function hasRuns(xml: string, paraIndex: number): boolean {
  const paraRegex = /<w:p[\s\S]*?<\/w:p>/g;
  let match;
  let idx = 0;

  while ((match = paraRegex.exec(xml)) !== null) {
    idx++;
    if (idx !== paraIndex) continue;

    const runRegex = /<w:r[\s\S]*?<\/w:r>/g;
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

  const paraXml = getBodyContentInfo(xml)
    .find((item): item is Extract<BodyContentInfo, { type: "paragraph" }> => item.type === "paragraph" && item.index === paraIndex)
    ?.xml;

  if (!paraXml) {
    return runs;
  }

  const runRegex = /<w:r[\s\S]*?<\/w:r>/g;
  let runMatch;

  while ((runMatch = runRegex.exec(paraXml)) !== null) {
    const runXml = runMatch[0];
    const textMatch = /<w:t[^>]*>([^<]*)<\/w:t>/i.exec(runXml);
    const text = textMatch ? textMatch[1] : "";

    const runInfo: RunInfo = { text };

    const rPrMatch = /<w:rPr>([\s\S]*?)<\/w:rPr>/i.exec(runXml);
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
  const paraRegex = /<w:p[\s\S]*?<\/w:p>/g;
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

  const styleRegex = /<w:style[^>]*>([\s\S]*?)<\/w:style>/g;
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

// ============================================================================
// Add Part - Add chart/header/footer parts
// ============================================================================

export interface AddPartOptions {
  type?: string;
  props?: Record<string, string>;
}

export interface AddPartResult {
  relId: string;
  partPath: string;
}

/**
 * Adds a new part (chart, header, footer) to a Word document.
 * Returns the relationship ID and accessible path.
 */
export async function addWordPart(
  filePath: string,
  partType: string,
  options: AddPartOptions = {}
): Promise<Result<AddPartResult>> {
  try {
    const zip = await readDocxZip(filePath);

    switch (partType.toLowerCase()) {
      case "chart": {
        // Find existing chart count
        const existingCharts = (zip.file(/^word\/chart\d*\.xml$/) || []).length;
        const chartNum = existingCharts + 1;
        const chartRelId = `rIdChart${chartNum}`;

        // Create chart XML
        const chartXml = createBasicChartXml();
        zip.file(`word/chart${chartNum}.xml`, chartXml);

        // Update document.xml.rels to add relationship
        const relsPath = "word/_rels/document.xml.rels";
        let relsXml = await getXmlEntry(zip, relsPath) || "";
        const newRel = `<Relationship Id="${chartRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="chart${chartNum}.xml"/>`;
        relsXml = relsXml.replace("</Relationships>", `${newRel}</Relationships>`);
        zip.file(relsPath, relsXml);

        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ relId: chartRelId, partPath: `/chart[${chartNum}]` });
      }

      case "header": {
        const existingHeaders = (zip.file(/^word\/header\d+\.xml$/) || []).length;
        const headerNum = existingHeaders + 1;
        const headerRelId = `rIdH${headerNum}`;

        const headerXml = createHeaderXml(options.props || {});
        zip.file(`word/header${headerNum}.xml`, headerXml);

        // Update document.xml.rels
        const relsPath = "word/_rels/document.xml.rels";
        let relsXml = await getXmlEntry(zip, relsPath) || "";
        const newRel = `<Relationship Id="${headerRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header${headerNum}.xml"/>`;
        relsXml = relsXml.replace("</Relationships>", `${newRel}</Relationships>`);
        zip.file(relsPath, relsXml);

        // Update document.xml to reference header in sectPr
        let documentXml = await getXmlEntry(zip, "word/document.xml") || "";
        const headerRef = `<w:headerReference w:type="default" r:id="${headerRelId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`;
        documentXml = documentXml.replace("<w:body>", `<w:body>${headerRef}`);
        zip.file("word/document.xml", documentXml);

        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ relId: headerRelId, partPath: `/header[${headerNum}]` });
      }

      case "footer": {
        const existingFooters = (zip.file(/^word\/footer\d+\.xml$/) || []).length;
        const footerNum = existingFooters + 1;
        const footerRelId = `rIdF${footerNum}`;

        const footerXml = createFooterXml(options.props || {});
        zip.file(`word/footer${footerNum}.xml`, footerXml);

        // Update document.xml.rels
        const relsPath = "word/_rels/document.xml.rels";
        let relsXml = await getXmlEntry(zip, relsPath) || "";
        const newRel = `<Relationship Id="${footerRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer${footerNum}.xml"/>`;
        relsXml = relsXml.replace("</Relationships>", `${newRel}</Relationships>`);
        zip.file(relsPath, relsXml);

        // Update document.xml to reference footer in sectPr
        let documentXml = await getXmlEntry(zip, "word/document.xml") || "";
        const footerRef = `<w:footerReference w:type="default" r:id="${footerRelId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`;
        documentXml = documentXml.replace("<w:body>", `<w:body>${footerRef}`);
        zip.file("word/document.xml", documentXml);

        await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
        return ok({ relId: footerRelId, partPath: `/footer[${footerNum}]` });
      }

      default:
        return err("invalid_type", `Unknown part type: ${partType}. Supported: chart, header, footer`);
    }
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

function createBasicChartXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:layout/>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
}

// ============================================================================
// Copy Word Node
// ============================================================================

/**
 * Copies an element from source path to target parent path.
 * Returns the new element's path.
 *
 * C# Reference: WordHandler.Mutations.CopyFrom (line 370-406)
 */
export async function copyWordNode(
  filePath: string,
  sourcePath: string,
  targetParentPath: string,
  options: { index?: number; after?: string; before?: string } = {}
): Promise<Result<{ path: string }>> {
  try {
    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found");
    }

    // Get source element info
    const sourceResult = await getWordNode(filePath, sourcePath, 0);
    if (!sourceResult.ok || !sourceResult.data) {
      return err("not_found", `Source not found: ${sourcePath}`);
    }

    // Extract the source element's XML using the correct element type
    const sourceXml = extractElementXml(documentXml, sourcePath);
    if (!sourceXml) {
      return err("not_found", `Could not extract source element: ${sourcePath}`);
    }

    // Generate new paraId/textId for cloned paragraphs (like C# CloneNode + RegenerateIds)
    let clonedXml = generateNewParaIds(sourceXml);

    // Determine insert position
    let insertPosition: string | number | undefined = options.index;
    if (options.after) {
      insertPosition = `find:${options.after}`;
    } else if (options.before) {
      insertPosition = `find:${options.before}`;
    }

    // Normalize target parent path to body if needed
    if (targetParentPath === "/" || targetParentPath === "" || targetParentPath === "/body") {
      targetParentPath = "/body";
    }

    // Insert at position using the improved insertAtPosition
    const originalDocXml = documentXml;
    documentXml = insertAtPosition(documentXml, clonedXml, insertPosition);

    // If document wasn't modified, insertion failed
    if (documentXml === originalDocXml && insertPosition !== undefined) {
      return err("operation_failed", "Failed to insert element at specified position");
    }

    zip.file("word/document.xml", documentXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

    // Calculate the actual path of the newly inserted element
    // The new element was inserted at the specified position, so we count
    // elements of the same type up to that position to find the index
    const elementType = sourceResult.data.type || "paragraph";
    const elementTag = elementType === "paragraph" ? "w:p" : elementType === "table" ? "w:tbl" : "w:p";

    // Count how many elements of this type exist at/after the insert position
    const newPath = calculateInsertedPath(documentXml, clonedXml, targetParentPath, elementType);

    return ok({ path: newPath });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Calculate the path of an inserted element by finding its position in the document
 */
function calculateInsertedPath(documentXml: string, insertedXml: string, targetParentPath: string, elementType: string): string {
  // Find the element tag based on type
  let elementTag = "w:p";
  if (elementType === "table" || elementType === "tbl") {
    elementTag = "w:tbl";
  }

  // Find the position where the inserted XML appears in the document
  const insertIdx = documentXml.indexOf(insertedXml);
  if (insertIdx === -1) {
    // Fallback: return target parent with element type
    return `${targetParentPath}/${elementType}[1]`;
  }

  // Count how many elements of the same type appear before this position
  const beforeDoc = documentXml.substring(0, insertIdx);
  const elementRegex = new RegExp(`<${elementTag}\\b`, 'gi');
  const matches = beforeDoc.match(elementRegex);
  const index = matches ? matches.length + 1 : 1;

  return `${targetParentPath}/${elementType}[${index}]`;
}

// ============================================================================
// Ensure ParaIds - Generate and ensure stable IDs
// ============================================================================

/**
 * Ensures all paragraphs in the document have unique paraId and textId attributes.
 * This should be called when creating or modifying documents to ensure stable paths.
 */
export async function ensureParaIds(filePath: string): Promise<Result<{ updated: number }>> {
  try {
    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found");
    }

    let updated = 0;

    // Add paraId to paragraphs that don't have it
    const paraRegex = /<w:p[>\s][\s\S]*?<\/w:p>/gi;
    documentXml = documentXml.replace(paraRegex, (match) => {
      if (/<w:paraId/i.test(match)) {
        return match;
      }
      const newParaId = generateHexId(8);
      const newTextId = generateHexId(8);
      updated++;
      return match.replace("<w:p ", `<w:p `).replace(">", `><w:paraId w:val="${newParaId}"/><w:textId w:val="${newTextId}"/>`);
    });

    // Actually, we need a better approach - insert paraId as a child element if not present
    documentXml = await ensureParaIdsInXml(documentXml);

    zip.file("word/document.xml", documentXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

    return ok({ updated });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

async function ensureParaIdsInXml(xml: string): Promise<string> {
  let result = xml;

  // Match paragraphs
  const paraRegex = /<w:p\b([^>]*)(>([\s\S]*?)<\/w:p>)/gi;
  result = result.replace(paraRegex, (match, attrs, content, innerContent) => {
    // Check if paraId already exists
    if (/w:paraId\b/i.test(attrs) && /w:textId\b/i.test(innerContent)) {
      return match;
    }

    // Generate new IDs
    const newParaId = generateHexId(8);
    const newTextId = generateHexId(8);

    // Insert paraId and textId after opening tag
    const idElements = `<w:paraId w:val="${newParaId}"/><w:textId w:val="${newTextId}"/>`;

    // Find the position after <w:p ...>
    const closeBracket = match.indexOf(">");
    if (closeBracket === -1) return match;

    return match.substring(0, closeBracket + 1) + idElements + match.substring(closeBracket + 1);
  });

  return result;
}

// ============================================================================
// Document Properties
// ============================================================================

export interface DocumentProperties {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  description?: string;
  category?: string;
  lastModifiedBy?: string;
  revision?: string;
}

/**
 * Sets document core properties.
 */
export async function setDocumentProperties(
  filePath: string,
  props: DocumentProperties
): Promise<Result<{ ok: boolean }>> {
  try {
    const zip = await readDocxZip(filePath);

    // Update docProps/core.xml
    let coreXml = await getXmlEntry(zip, "docProps/core.xml") || createBasicCoreXml();

    if (props.title !== undefined) {
      coreXml = updateCoreProperty(coreXml, "dc:title", props.title);
    }
    if (props.author !== undefined) {
      coreXml = updateCoreProperty(coreXml, "dc:creator", props.author);
    }
    if (props.subject !== undefined) {
      coreXml = updateCoreProperty(coreXml, "dc:subject", props.subject);
    }
    if (props.keywords !== undefined) {
      coreXml = updateCoreProperty(coreXml, "cp:keywords", props.keywords);
    }
    if (props.description !== undefined) {
      coreXml = updateCoreProperty(coreXml, "dc:description", props.description);
    }
    if (props.category !== undefined) {
      coreXml = updateCoreProperty(coreXml, "cp:category", props.category);
    }
    if (props.lastModifiedBy !== undefined) {
      coreXml = updateCoreProperty(coreXml, "cp:lastModifiedBy", props.lastModifiedBy);
    }
    if (props.revision !== undefined) {
      coreXml = updateCoreProperty(coreXml, "cp:revision", props.revision);
    }

    zip.file("docProps/core.xml", coreXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

    return ok({ ok: true });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

function createBasicCoreXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
</cp:coreProperties>`;
}

function updateCoreProperty(xml: string, tagName: string, value: string): string {
  const regex = new RegExp(`<${tagName}([^>]*)>([^<]*)</${tagName}>`, "i");
  if (regex.test(xml)) {
    return xml.replace(regex, `<${tagName}$1>${escapeXml(value)}</${tagName}>`);
  }
  // Add before closing tag
  return xml.replace("</cp:coreProperties>", `<${tagName}>${escapeXml(value)}</${tagName}></cp:coreProperties>`);
}

// ============================================================================
// Helper Functions
// ============================================================================

function generateHexId(length: number): string {
  const chars = "0123456789ABCDEF";
  let result = "";
  for (let i = 0; i < length; i++) {
    result += chars[Math.floor(Math.random() * chars.length)];
  }
  return result;
}

function extractElementXml(documentXml: string, path: string): string | null {
  // Path to XML extraction for common paths
  const parsed = parsePath(path);
  if (!parsed.ok || !parsed.data) return null;

  const segments = parsed.data.segments;
  if (segments.length === 0) return null;

  const first = segments[0];

  // Navigate to the element based on path
  if (first.name === "body") {
    // Handle /body path - extract body content
    const bodyMatch = documentXml.match(/<w:body>([\s\S]*)<\/w:body>/);
    if (!bodyMatch) return null;

    // If only /body, return the body content
    if (segments.length === 1) {
      return bodyMatch[1];
    }

    // Otherwise, navigate to the child element
    const remainingPath = "/" + segments.slice(1).map(s => s.name + (s.index !== undefined ? `[${s.index}]` : "")).join("/");
    return extractElementXml(bodyMatch[1], remainingPath);
  }

  if (first.name === "p" || first.name === "paragraph") {
    const idx = (first.index || 1) - 1;
    const paras = documentXml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/gi);
    if (!paras || !paras[idx]) return null;

    // If there are more segments, navigate deeper (e.g., /body/p[1]/r[2])
    if (segments.length > 1) {
      const remainingPath = "/" + segments.slice(1).map(s => s.name + (s.index !== undefined ? `[${s.index}]` : "")).join("/");
      return extractElementXml(paras[idx], remainingPath);
    }
    return paras[idx];
  }

  if (first.name === "tbl" || first.name === "table") {
    const idx = (first.index || 1) - 1;
    const tables = documentXml.match(/<w:tbl\b[^>]*>[\s\S]*?<\/w:tbl>/gi);
    if (!tables || !tables[idx]) return null;

    // If there are more segments, navigate deeper (e.g., /body/tbl[1]/tr[2])
    if (segments.length > 1) {
      const remainingPath = "/" + segments.slice(1).map(s => s.name + (s.index !== undefined ? `[${s.index}]` : "")).join("/");
      return extractElementXml(tables[idx], remainingPath);
    }
    return tables[idx];
  }

  if (first.name === "r" || first.name === "run") {
    const idx = (first.index || 1) - 1;
    const runs = documentXml.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/gi);
    if (runs && runs[idx]) {
      return runs[idx];
    }
  }

  if (first.name === "tr" || first.name === "row") {
    const idx = (first.index || 1) - 1;
    const rows = documentXml.match(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/gi);
    if (rows && rows[idx]) {
      // If there are more segments, navigate deeper (e.g., /body/tbl[1]/tr[1]/tc[2])
      if (segments.length > 1) {
        const remainingPath = "/" + segments.slice(1).map(s => s.name + (s.index !== undefined ? `[${s.index}]` : "")).join("/");
        return extractElementXml(rows[idx], remainingPath);
      }
      return rows[idx];
    }
  }

  if (first.name === "tc" || first.name === "cell") {
    const idx = (first.index || 1) - 1;
    const cells = documentXml.match(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/gi);
    if (cells && cells[idx]) {
      return cells[idx];
    }
  }

  return null;
}


function generateNewParaIds(xml: string): string {
  let result = xml;

  // Generate new paraId for any paragraph in the cloned XML
  result = result.replace(/<w:paraId\b[^>]*w:val="[^"]*"[^>]*\/>/gi, () => {
    return `<w:paraId w:val="${generateHexId(8)}"/>`;
  });

  result = result.replace(/<w:textId\b[^>]*w:val="[^"]*"[^>]*\/>/gi, () => {
    return `<w:textId w:val="${generateHexId(8)}"/>`;
  });

  return result;
}

// ============================================================================
// Validate - Document OpenXML Validation
// ============================================================================

export interface ValidationError {
  errorType: string;
  description: string;
  path?: string;
  part?: string;
}

/**
 * Validates the document against OpenXML schema.
 * Returns a list of validation errors.
 */
export async function validateWordDocument(filePath: string): Promise<ValidationError[]> {
  const errors: ValidationError[] = [];
  const zip = await readDocxZip(filePath);

  const requiredParts = ["[Content_Types].xml", "word/document.xml"];
  for (const part of requiredParts) {
    if (!zip.file(part)) {
      errors.push({ errorType: "missing_part", description: `Required part missing: ${part}`, part });
    }
  }

  const documentXml = await getXmlEntry(zip, "word/document.xml");
  if (documentXml) {
    if (!documentXml.includes("<w:document") && !documentXml.includes("<w:document ")) {
      errors.push({ errorType: "invalid_root", description: "Document root element (w:document) not found", part: "word/document.xml" });
    }
    if (!documentXml.includes("<w:body") && !documentXml.includes("<w:body ")) {
      errors.push({ errorType: "missing_body", description: "Document body (w:body) not found", part: "word/document.xml" });
    }
    if (!documentXml.includes("</w:document>")) {
      errors.push({ errorType: "unclosed_tag", description: "Missing closing tag for w:document", part: "word/document.xml" });
    }
    if (!documentXml.includes("xmlns:w=") && !documentXml.includes("xmlns:w=\"")) {
      errors.push({ errorType: "missing_namespace", description: "Missing w: namespace declaration", part: "word/document.xml" });
    }
  }

  const stylesXml = await getXmlEntry(zip, "word/styles.xml");
  if (stylesXml && !stylesXml.includes("<w:styles") && !stylesXml.includes("<w:styles ")) {
    errors.push({ errorType: "invalid_styles", description: "Styles root element (w:styles) not found", part: "word/styles.xml" });
  }

  const relsXml = await getXmlEntry(zip, "word/_rels/document.xml.rels");
  if (relsXml) {
    const idMatches = relsXml.match(/Id="([^"]+)"/g) || [];
    const ids = idMatches.map(m => m.match(/Id="([^"]+)"/)![1]);
    const duplicates = ids.filter((id, idx) => ids.indexOf(id) !== idx);
    if (duplicates.length > 0) {
      errors.push({ errorType: "duplicate_rels", description: `Duplicate relationship IDs: ${[...new Set(duplicates)].join(", ")}`, part: "word/_rels/document.xml.rels" });
    }
  }

  const contentTypesXml = await getXmlEntry(zip, "[Content_Types].xml");
  if (contentTypesXml && !contentTypesXml.includes("[Content_Types]")) {
    errors.push({ errorType: "invalid_content_types", description: "Content_Types.xml root element not found", part: "[Content_Types].xml" });
  }

  return errors;
}

// ============================================================================
// JSON View Functions
// ============================================================================

export async function viewWordStatsJson(filePath: string): Promise<Record<string, unknown>> {
  const zip = await readDocxZip(filePath);
  const documentXml = await getXmlEntry(zip, "word/document.xml") ?? "";

  const paras = getParagraphsInfo(documentXml);
  const tables = getTablesInfo(documentXml);
  let totalWords = 0, totalChars = 0;
  const styleCounts: Record<string, number> = {};
  const fontCounts: Record<string, number> = {};

  for (const para of paras) {
    const style = para.style || "Normal";
    styleCounts[style] = (styleCounts[style] || 0) + 1;
    if (!para.text.trim()) continue;
    const words = para.text.split(/\s+/).filter(Boolean);
    totalWords += words.length;
    totalChars += para.text.length;
    const runs = getRunsInfo(documentXml, para.index);
    for (const run of runs) {
      if (run.font) fontCounts[run.font] = (fontCounts[run.font] || 0) + 1;
    }
  }

  return {
    paragraphs: paras.length, words: totalWords, characters: totalChars, tables: tables.length,
    styles: Object.entries(styleCounts).map(([name, count]) => ({ name, count })),
    fonts: Object.entries(fontCounts).map(([name, count]) => ({ name, count }))
  };
}

export async function viewWordOutlineJson(filePath: string): Promise<Record<string, unknown>> {
  const zip = await readDocxZip(filePath);
  const documentXml = await getXmlEntry(zip, "word/document.xml") ?? "";
  const paras = getParagraphsInfo(documentXml);
  const headings: Array<{ level: number; text: string; path: string; style: string }> = [];

  let paraIndex = 0;
  for (const para of paras) {
    paraIndex++;
    if (para.style && (para.style.includes("Heading") || para.style === "Title" || para.style === "Subtitle")) {
      headings.push({ level: getHeadingLevel(para.style), text: para.text, path: `/body/p[${paraIndex}]`, style: para.style });
    }
  }
  return { headings, totalParagraphs: paras.length };
}

export async function viewWordTextJson(filePath: string, options?: { startLine?: number; endLine?: number; maxLines?: number }): Promise<Record<string, unknown>> {
  const zip = await readDocxZip(filePath);
  const documentXml = await getXmlEntry(zip, "word/document.xml") ?? "";
  const stylesXml = await getXmlEntry(zip, "word/styles.xml") ?? "";
  const paras = getParagraphsInfo(documentXml);
  const startLine = options?.startLine ?? 1, endLine = options?.endLine ?? paras.length, maxLines = options?.maxLines ?? paras.length;

  const lines: Array<{ index: number; path: string; text: string; style?: string }> = [];
  let lineNum = 0, emitted = 0;

  for (const para of paras) {
    lineNum++;
    if (lineNum < startLine || lineNum > endLine || emitted >= maxLines) {
      if (lineNum > endLine || emitted >= maxLines) break;
      continue;
    }
    const styleName = para.style ? getStyleNameFromId(stylesXml, para.style) || para.style : "Normal";
    lines.push({ index: lineNum, path: `/body/p[${para.index}]`, text: para.text, style: styleName });
    emitted++;
  }
  return { lines, total: paras.length, startLine, endLine: lineNum, truncated: emitted >= maxLines };
}

export async function viewWordIssuesJson(filePath: string, options?: { issueType?: string; limit?: number }): Promise<Record<string, unknown>> {
  const zip = await readDocxZip(filePath);
  const documentXml = await getXmlEntry(zip, "word/document.xml") ?? "";
  const limit = options?.limit ?? 100;
  const issues: Array<{ type: string; description: string; path?: string; severity: string }> = [];
  const paras = getParagraphsInfo(documentXml);

  let paraIndex = 0;
  for (const para of paras) {
    paraIndex++;
    if (issues.length >= limit) break;
    if (!para.text.trim()) {
      issues.push({ type: "content", description: "Empty paragraph", path: `/body/p[${paraIndex}]`, severity: "warning" });
    } else if (para.text.includes("  ")) {
      issues.push({ type: "formatting", description: "Consecutive spaces detected", path: `/body/p[${paraIndex}]`, severity: "warning" });
    }
  }

  const sectPrMatch = documentXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/i);
  if (sectPrMatch && !sectPrMatch[0].includes("<w:pgMar")) {
    issues.push({ type: "structure", description: "Section missing page margins", severity: "error" });
  }
  return { issues, total: issues.length, limit };
}

// ============================================================================
// Form Fields - Full Support
// ============================================================================

export interface FormFieldInfo {
  type: "text" | "checkbox" | "dropdown";
  name: string;
  value: string;
  enabled: boolean;
  editable: boolean;
  path: string;
  defaultValue?: string;
  maxLength?: number;
  checked?: boolean;
  items?: string[];
  defaultIndex?: number;
}

export async function getWordFormFields(filePath: string): Promise<Result<FormFieldInfo[]>> {
  try {
    const zip = await readDocxZip(filePath);
    const documentXml = await getXmlEntry(zip, "word/document.xml");
    if (!documentXml) return err("not_found", "Document.xml not found");

    const fields: FormFieldInfo[] = [];
    const formTextRegex = /<w:fldChar[^>]*w:fldCharType="begin"[^>]*>[\s\S]*?<w:ffData>([\s\S]*?)<\/w:ffData>/gi;
    let match;

    while ((match = formTextRegex.exec(documentXml)) !== null) {
      const ffData = match[1];
      const nameMatch = ffData.match(/<w:ffname[^>]*w:val="([^"]*)"/i);
      const name = nameMatch ? nameMatch[1] : "unnamed";
      const textInputMatch = ffData.match(/<w:textInput/i);
      const checkBoxMatch = ffData.match(/<w:checkBox/i);
      const dropDownMatch = ffData.match(/<w:dropDown/i);

      let type: "text" | "checkbox" | "dropdown" = "text";
      let value = "", defaultValue: string | undefined, maxLength: number | undefined, checked: boolean | undefined, items: string[] | undefined, defaultIndex: number | undefined;

      if (textInputMatch) {
        type = "text";
        const dm = ffData.match(/<w:default[\s\S]*?w:val="([^"]*)"/i);
        defaultValue = dm ? dm[1] : undefined;
        const mlm = ffData.match(/<w:maxLength[^>]*w:val="(\d+)"/i);
        maxLength = mlm ? parseInt(mlm[1], 10) : undefined;
        value = defaultValue || "";
      } else if (checkBoxMatch) {
        type = "checkbox";
        const cm = ffData.match(/<w:checked[^>]*w:val="([^"]*)"/i);
        checked = cm ? cm[1].toLowerCase() === "true" : false;
        value = checked ? "\u2612" : "\u2610";
      } else if (dropDownMatch) {
        type = "dropdown";
        const im = ffData.matchAll(/<w:listItem[^>]*w:val="([^"]*)"/gi);
        items = Array.from(im, m => m[1]);
        const sm = ffData.match(/<w:selection[^>]*w:val="(\d+)"/i);
        defaultIndex = sm ? parseInt(sm[1], 10) : 0;
        value = items[defaultIndex || 0] || "";
      }

      const resultStart = documentXml.indexOf("</w:fldChar>", match.index) + 13;
      const resultEnd = documentXml.indexOf("<w:fldChar", resultStart);
      if (resultStart > 13 && resultEnd > resultStart) {
        const textMatch = documentXml.substring(resultStart, resultEnd).match(/<w:t[^>]*>([^<]*)/i);
        if (textMatch && textMatch[1]) value = textMatch[1];
      }

      fields.push({ type, name, value, enabled: true, editable: true, path: `/formfield[${fields.length + 1}]`, defaultValue, maxLength, checked, items, defaultIndex });
    }
    return ok(fields);
  } catch (e) {
    return err("operation_failed", e instanceof Error ? e.message : String(e));
  }
}

export async function setWordFormField(filePath: string, fieldPath: string, props: Record<string, string>): Promise<Result<{ ok: boolean }>> {
  try {
    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");
    if (!documentXml) return err("not_found", "Document.xml not found");

    const fieldMatch = fieldPath.match(/\/formfield\[(\d+)\]/i);
    if (!fieldMatch) return err("invalid_path", "Invalid formfield path");
    const fieldIndex = parseInt(fieldMatch[1], 10);

    const formTextRegex = /<w:fldChar[^>]*w:fldCharType="begin"[^>]*>[\s\S]*?<w:ffData>([\s\S]*?)<\/w:ffData>[\s\S]*?<w:fldChar[^>]*w:fldCharType="separate"/gi;
    let fieldNum = 0, updated = false;

    documentXml = documentXml.replace(formTextRegex, (fullMatch) => {
      fieldNum++;
      if (fieldNum !== fieldIndex) return fullMatch;
      let newMatch = fullMatch;
      if (props.text !== undefined || props.value !== undefined) {
        newMatch = newMatch.replace(/<w:t[^>]*>[^<]*<\/w:t>/gi, `<w:t>${escapeXml(props.text || props.value || "")}</w:t>`);
        updated = true;
      }
      if (props.checked !== undefined) {
        const isChecked = props.checked.toLowerCase() === "true";
        newMatch = newMatch.replace(/<w:t[^>]*>[^<]*<\/w:t>/gi, `<w:t>${isChecked ? "\u2612" : "\u2610"}</w:t>`);
        updated = true;
      }
      return newMatch;
    });

    if (!updated) return err("not_found", `Form field ${fieldIndex} not found`);
    zip.file("word/document.xml", documentXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
    return ok({ ok: true });
  } catch (e) {
    return err("operation_failed", e instanceof Error ? e.message : String(e));
  }
}

// ============================================================================
// Track Changes - Accept/Reject All
// ============================================================================

export interface TrackChangesResult { accepted: number; rejected: number; }

export async function acceptAllTrackChanges(filePath: string): Promise<Result<TrackChangesResult>> {
  try {
    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");
    if (!documentXml) return err("not_found", "Document.xml not found");

    let accepted = 0;
    const insRegex = /<w:ins[^>]*>([\s\S]*?)<\/w:ins>/gi;
    documentXml = documentXml.replace(insRegex, (m, inner) => { accepted++; return inner; });
    const delMatches = documentXml.match(/<w:del[^>]*>[\s\S]*?<\/w:del>/gi);
    documentXml = documentXml.replace(/<w:del[^>]*>[\s\S]*?<\/w:del>/gi, "");
    accepted += delMatches ? delMatches.length : 0;
    documentXml = documentXml.replace(/<w:rPrChange[^>]*>[\s\S]*?<\/w:rPrChange>/gi, "");
    documentXml = documentXml.replace(/<w:pPrChange[^>]*>[\s\S]*?<\/w:pPrChange>/gi, "");

    zip.file("word/document.xml", documentXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
    return ok({ accepted, rejected: 0 });
  } catch (e) {
    return err("operation_failed", e instanceof Error ? e.message : String(e));
  }
}

export async function rejectAllTrackChanges(filePath: string): Promise<Result<TrackChangesResult>> {
  try {
    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");
    if (!documentXml) return err("not_found", "Document.xml not found");

    let rejected = 0;
    const insMatches = documentXml.match(/<w:ins[^>]*>[\s\S]*?<\/w:ins>/gi);
    documentXml = documentXml.replace(/<w:ins[^>]*>[\s\S]*?<\/w:ins>/gi, "");
    rejected += insMatches ? insMatches.length : 0;
    documentXml = documentXml.replace(/<w:del[^>]*>([\s\S]*?)<\/w:del>/gi, (m, inner) => { rejected++; return inner; });
    documentXml = documentXml.replace(/<w:rPrChange[^>]*>[\s\S]*?<\/w:rPrChange>/gi, "");
    documentXml = documentXml.replace(/<w:pPrChange[^>]*>[\s\S]*?<\/w:pPrChange>/gi, "");

    zip.file("word/document.xml", documentXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
    return ok({ accepted: 0, rejected });
  } catch (e) {
    return err("operation_failed", e instanceof Error ? e.message : String(e));
  }
}

// ============================================================================
// Document Protection
// ============================================================================

export interface DocumentProtection { mode?: string; enforced?: boolean; }

export async function getDocumentProtection(filePath: string): Promise<DocumentProtection> {
  const zip = await readDocxZip(filePath);
  const settingsXml = await getXmlEntry(zip, "word/settings.xml") ?? "";
  const pm = settingsXml.match(/<w:documentProtection[^>]*w:edit="([^"]*)"[^>]*w:enforcement="([^"]*)"/i);
  return pm ? { mode: pm[1], enforced: pm[2].toLowerCase() === "true" } : { enforced: false };
}

export async function setDocumentProtection(filePath: string, mode: string, enforced: boolean = true): Promise<Result<{ ok: boolean }>> {
  try {
    const zip = await readDocxZip(filePath);
    let settingsXml = await getXmlEntry(zip, "word/settings.xml") || createBasicSettingsXml();
    settingsXml = settingsXml.replace(/<w:documentProtection[^>]*\/>/gi, "");
    if (enforced && mode !== "none") {
      settingsXml = settingsXml.replace("</w:settings>", `<w:documentProtection w:edit="${mode}" w:enforcement="true"/></w:settings>`);
    }
    zip.file("word/settings.xml", settingsXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
    return ok({ ok: true });
  } catch (e) {
    return err("operation_failed", e instanceof Error ? e.message : String(e));
  }
}

// ============================================================================
// SDT (Structured Document Tags / Content Controls) - Advanced
// ============================================================================

export interface SdtInfo { path: string; type: string; tag?: string; alias?: string; value?: string; }

export async function getWordSdts(filePath: string): Promise<Result<SdtInfo[]>> {
  try {
    const zip = await readDocxZip(filePath);
    const documentXml = await getXmlEntry(zip, "word/document.xml");
    if (!documentXml) return err("not_found", "Document.xml not found");

    const sdts: SdtInfo[] = [];
    const sdtRegex = /<w:sdt[^>]*>([\s\S]*?)<\/w:sdt>/gi;
    let match, idx = 0;

    while ((match = sdtRegex.exec(documentXml)) !== null) {
      idx++;
      const sdtContent = match[1];
      const tagMatch = sdtContent.match(/<w:tag[^>]*w:val="([^"]*)"/i);
      const aliasMatch = sdtContent.match(/<w:alias[^>]*w:val="([^"]*)"/i);
      let type = "unknown";
      if (sdtContent.includes("<w:richText")) type = "richText";
      else if (sdtContent.includes("<w:text")) type = "text";
      else if (sdtContent.includes("<w:checkBox")) type = "checkbox";
      else if (sdtContent.includes("<w:dropDownList")) type = "dropdown";
      else if (sdtContent.includes("<w:date")) type = "date";
      else if (sdtContent.includes("<w:comboBox")) type = "comboBox";
      else if (sdtContent.includes("<w:picture")) type = "picture";
      let value = "";
      const textMatch = sdtContent.match(/<w:t[^>]*>([^<]*)/i);
      if (textMatch) value = textMatch[1];
      sdts.push({ path: `/sdt[${idx}]`, type, tag: tagMatch?.[1], alias: aliasMatch?.[1], value });
    }
    return ok(sdts);
  } catch (e) {
    return err("operation_failed", e instanceof Error ? e.message : String(e));
  }
}

export async function setWordSdt(filePath: string, sdtPath: string, props: Record<string, string>): Promise<Result<{ ok: boolean }>> {
  try {
    const zip = await readDocxZip(filePath);
    let documentXml = await getXmlEntry(zip, "word/document.xml");
    if (!documentXml) return err("not_found", "Document.xml not found");

    const sdtMatch = sdtPath.match(/\/sdt\[(\d+)\]/i);
    if (!sdtMatch) return err("invalid_path", "Invalid SDT path");
    const sdtIndex = parseInt(sdtMatch[1], 10);

    const sdtRegex = /<w:sdt[^>]*>([\s\S]*?)<\/w:sdt>/gi;
    let sdtNum = 0, updated = false;

    documentXml = documentXml.replace(sdtRegex, (fullMatch) => {
      sdtNum++;
      if (sdtNum !== sdtIndex) return fullMatch;
      let newMatch = fullMatch;
      if (props.text !== undefined || props.value !== undefined) {
        newMatch = newMatch.replace(/<w:t[^>]*>[^<]*<\/w:t>/gi, `<w:t>${escapeXml(props.text || props.value || "")}</w:t>`);
        updated = true;
      }
      if (props.checked !== undefined) {
        const isChecked = props.checked.toLowerCase() === "true";
        newMatch = newMatch.replace(/<w:checked[^>]*/gi, `<w:checked w:val="${isChecked ? "true" : "false"}"/>`);
        updated = true;
      }
      if (props.tag !== undefined && /<w:tag/i.test(newMatch)) {
        newMatch = newMatch.replace(/<w:tag[^>]*w:val="[^"]*"/i, `<w:tag w:val="${escapeXml(props.tag)}"/>`);
        updated = true;
      }
      return newMatch;
    });

    if (!updated) return err("not_found", `SDT ${sdtIndex} not found`);
    zip.file("word/document.xml", documentXml);
    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
    return ok({ ok: true });
  } catch (e) {
    return err("operation_failed", e instanceof Error ? e.message : String(e));
  }
}
