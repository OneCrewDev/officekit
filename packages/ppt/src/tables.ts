/**
 * Table mutation operations for @officekit/ppt.
 *
 * Provides functions to modify tables on slides:
 * - Set table cell text
 * - Remove table rows
 * - Remove table columns
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput } from "./result.js";
import type { Result } from "./types.js";
import { getSlideIndex } from "./path.js";

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
 * Extracts indices from a table cell path.
 * Path format: /slide[N]/table[M]/tr[R]/tc[C]
 */
function extractTableCellIndices(pptPath: string): { tableIndex: number; rowIndex: number; cellIndex: number } | null {
  const pattern = /\/table\[(\d+)\]\/tr\[(\d+)\]\/tc\[(\d+)\]/i;
  const match = pptPath.match(pattern);
  if (!match) {
    return null;
  }
  return {
    tableIndex: parseInt(match[1], 10),
    rowIndex: parseInt(match[2], 10),
    cellIndex: parseInt(match[3], 10),
  };
}

/**
 * Extracts indices from a table row path.
 * Path format: /slide[N]/table[M]/tr[R]
 */
function extractTableRowIndices(pptPath: string): { tableIndex: number; rowIndex: number } | null {
  const pattern = /\/table\[(\d+)\]\/tr\[(\d+)\]/i;
  const match = pptPath.match(pattern);
  if (!match) {
    return null;
  }
  return {
    tableIndex: parseInt(match[1], 10),
    rowIndex: parseInt(match[2], 10),
  };
}

/**
 * Extracts the table index from a path.
 * Path format: /slide[N]/table[M]
 */
function extractTableIndex(pptPath: string): number | null {
  const pattern = /\/table\[(\d+)\]/i;
  const match = pptPath.match(pattern);
  return match ? parseInt(match[1], 10) : null;
}

// ============================================================================
// Table Cell Operations
// ============================================================================

/**
 * Sets the text content of a table cell.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the cell (e.g., "/slide[1]/table[1]/tr[1]/tc[1]")
 * @param text - The new text content
 *
 * @example
 * const result = await setTableCell("/path/to/presentation.pptx", "/slide[1]/table[1]/tr[1]/tc[1]", "Hello");
 */
export async function setTableCell(
  filePath: string,
  pptPath: string,
  text: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setTableCell requires a slide path");
    }

    const indices = extractTableCellIndices(pptPath);
    if (!indices) {
      return invalidInput("Invalid table cell path format. Expected: /slide[N]/table[M]/tr[R]/tc[C]");
    }

    const { tableIndex, rowIndex, cellIndex } = indices;

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message, slidePathResult.error!.suggestion);
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry!);

    const updatedSlideXml = setCellTextInTable(slideXml, tableIndex, rowIndex, cellIndex, text);

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
 * Sets the text in a table cell by finding and updating the cell.
 */
function setCellTextInTable(
  slideXml: string,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  text: string,
): string {
  // First, find the table by index
  const tablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  const tables = slideXml.match(tablePattern);

  if (!tables || tableIndex < 1 || tableIndex > tables.length) {
    throw new Error(`Table index ${tableIndex} out of range`);
  }

  const targetTableXml = tables[tableIndex - 1];
  const updatedTableXml = setTextInTableCell(targetTableXml, rowIndex, cellIndex, text);

  return slideXml.replace(targetTableXml, updatedTableXml);
}

/**
 * Sets text in a specific cell of a table.
 */
function setTextInTableCell(
  tableXml: string,
  rowIndex: number,
  cellIndex: number,
  text: string,
): string {
  // Find all rows (a:tr elements)
  const rowPattern = /<a:tr[^>]*>[\s\S]*?<\/a:tr>/g;
  const rows = tableXml.match(rowPattern);

  if (!rows || rowIndex < 1 || rowIndex > rows.length) {
    throw new Error(`Row index ${rowIndex} out of range`);
  }

  const targetRowXml = rows[rowIndex - 1];

  // Find all cells (a:tc elements) in the row
  const cellPattern = /<a:tc[\s\S]*?<\/a:tc>/g;
  const cells = targetRowXml.match(cellPattern);

  if (!cells || cellIndex < 1 || cellIndex > cells.length) {
    throw new Error(`Cell index ${cellIndex} out of range`);
  }

  const targetCellXml = cells[cellIndex - 1];
  const updatedCellXml = updateCellText(targetCellXml, text);

  const updatedRowXml = targetRowXml.replace(targetCellXml, updatedCellXml);
  return tableXml.replace(targetRowXml, updatedRowXml);
}

/**
 * Updates the text content of a table cell.
 */
function updateCellText(cellXml: string, text: string): string {
  // Find the text body or create text content
  // Table cells use <a:txBody> for rich text

  const txBodyPattern = /<a:txBody>([\s\S]*?)<\/a:txBody>/;
  const hasTxBody = txBodyPattern.test(cellXml);

  if (hasTxBody) {
    // Replace existing text content
    return cellXml.replace(txBodyPattern, () => {
      const newParagraph = `          <a:p>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>${escapeXml(text)}</a:t>
            </a:r>
          </a:p>`;
      return `<a:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
${newParagraph}
          </a:txBody>`;
    });
  }

  // No txBody - need to add one inside the cell
  // Find </a:tc> and insert before it
  const newContent = `
          <a:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US"/>
                <a:t>${escapeXml(text)}</a:t>
              </a:r>
            </a:p>
          </a:txBody>
        `;

  return cellXml.replace(/<\/a:tc>/, `${newContent}</a:tc>`);
}

// ============================================================================
// Table Row Operations
// ============================================================================

/**
 * Removes a row from a table.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the row (e.g., "/slide[1]/table[1]/tr[1]")
 *
 * @example
 * const result = await removeTableRow("/path/to/presentation.pptx", "/slide[1]/table[1]/tr[1]");
 */
export async function removeTableRow(
  filePath: string,
  pptPath: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("removeTableRow requires a slide path");
    }

    const indices = extractTableRowIndices(pptPath);
    if (!indices) {
      return invalidInput("Invalid table row path format. Expected: /slide[N]/table[M]/tr[R]");
    }

    const { tableIndex, rowIndex } = indices;

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message, slidePathResult.error!.suggestion);
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry!);

    const updatedSlideXml = removeRowFromTable(slideXml, tableIndex, rowIndex);

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
 * Removes a row from a table in the slide XML.
 */
function removeRowFromTable(slideXml: string, tableIndex: number, rowIndex: number): string {
  // First, find the table by index
  const tablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  const tables = slideXml.match(tablePattern);

  if (!tables || tableIndex < 1 || tableIndex > tables.length) {
    throw new Error(`Table index ${tableIndex} out of range`);
  }

  const targetTableXml = tables[tableIndex - 1];
  const updatedTableXml = removeRowFromTableXml(targetTableXml, rowIndex);

  return slideXml.replace(targetTableXml, updatedTableXml);
}

/**
 * Removes a row from the table XML.
 */
function removeRowFromTableXml(tableXml: string, rowIndex: number): string {
  // Find all rows (a:tr elements)
  const rowPattern = /<a:tr[^>]*>[\s\S]*?<\/a:tr>/g;
  const rows = tableXml.match(rowPattern);

  if (!rows || rowIndex < 1 || rowIndex > rows.length) {
    throw new Error(`Row index ${rowIndex} out of range`);
  }

  const targetRowXml = rows[rowIndex - 1];
  return tableXml.replace(targetRowXml, "");
}

// ============================================================================
// Table Column Operations
// ============================================================================

/**
 * Removes a column from a table.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the column (e.g., "/slide[1]/table[1]" with column index in params)
 * @param columnIndex - 1-based index of the column to remove
 *
 * @example
 * const result = await removeTableColumn("/path/to/presentation.pptx", "/slide[1]/table[1]", 2);
 */
export async function removeTableColumn(
  filePath: string,
  pptPath: string,
  columnIndex: number,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("removeTableColumn requires a slide path");
    }

    const tableIndex = extractTableIndex(pptPath);
    if (!tableIndex) {
      return invalidInput("Invalid table path format. Expected: /slide[N]/table[M]");
    }

    if (columnIndex < 1) {
      return invalidInput("Column index must be at least 1");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message, slidePathResult.error!.suggestion);
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry!);

    const updatedSlideXml = removeColumnFromTable(slideXml, tableIndex, columnIndex);

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
 * Removes a column from a table in the slide XML.
 */
function removeColumnFromTable(slideXml: string, tableIndex: number, columnIndex: number): string {
  // First, find the table by index
  const tablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  const tables = slideXml.match(tablePattern);

  if (!tables || tableIndex < 1 || tableIndex > tables.length) {
    throw new Error(`Table index ${tableIndex} out of range`);
  }

  const targetTableXml = tables[tableIndex - 1];
  const updatedTableXml = removeColumnFromTableXml(targetTableXml, columnIndex);

  return slideXml.replace(targetTableXml, updatedTableXml);
}

/**
 * Removes a column from the table XML.
 * This removes the gridCell from each row at the given column index.
 */
function removeColumnFromTableXml(tableXml: string, columnIndex: number): string {
  // Find all rows (a:tr elements)
  const rowPattern = /<a:tr[^>]*>[\s\S]*?<\/a:tr>/g;
  const rows = tableXml.match(rowPattern);

  if (!rows) {
    throw new Error("No rows found in table");
  }

  let result = tableXml;

  for (const rowXml of rows) {
    // Find all cells (a:tc elements) in this row
    const cellPattern = /<a:tc[\s\S]*?<\/a:tc>/g;
    const cells = rowXml.match(cellPattern);

    if (!cells || columnIndex < 1 || columnIndex > cells.length) {
      throw new Error(`Column index ${columnIndex} out of range for one or more rows`);
    }

    const targetCellXml = cells[columnIndex - 1];
    const updatedRowXml = rowXml.replace(targetCellXml, "");
    result = result.replace(rowXml, updatedRowXml);
  }

  // Also need to update the table grid to remove the gridCell
  const gridPattern = /<a:tblGrid>([\s\S]*?)<\/a:tblGrid>/;
  const gridMatch = result.match(gridPattern);

  if (gridMatch) {
    const gridCells = gridMatch[1].match(/<a:gridCol[^>]*\/>/g);
    if (gridCells && columnIndex <= gridCells.length) {
      const updatedGrid = result.replace(gridCells[columnIndex - 1], "");
      result = result.replace(gridPattern, updatedGrid);
    }
  }

  return result;
}

// ============================================================================
// Add Table Operations
// ============================================================================

/**
 * Position and size for table placement.
 */
export interface TablePosition {
  x: number;
  y: number;
}

export interface TableSize {
  width: number;
  height: number;
}

/**
 * Adds a new table to a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param rows - Number of rows
 * @param cols - Number of columns
 * @param position - The position (x, y) in EMUs
 * @param size - The size (width, height) in EMUs
 *
 * @example
 * const result = await addTable("/path/to/presentation.pptx", 1, 3, 4, { x: 1000000, y: 1000000 }, { width: 5000000, height: 3000000 });
 * if (result.ok) {
 *   console.log(result.data.path); // "/slide[1]/table[1]"
 * }
 */
export async function addTable(
  filePath: string,
  slideIndex: number,
  rows: number,
  cols: number,
  position: TablePosition,
  size: TableSize,
): Promise<Result<{ path: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message, slidePathResult.error!.suggestion);
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry!);

    // Count existing tables to determine new table index
    const tablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
    const existingTables = slideXml.match(tablePattern) || [];
    const newTableIndex = existingTables.length + 1;

    // Generate unique shape ID for the table
    const idPattern = /id="(\d+)"/g;
    let maxId = 0;
    let match;
    while ((match = idPattern.exec(slideXml)) !== null) {
      const id = parseInt(match[1], 10);
      if (id > maxId) maxId = id;
    }
    const newTableId = maxId + 1;

    // Calculate cell dimensions
    const cellWidth = Math.round(size.width / cols);
    const cellHeight = Math.round(size.height / rows);

    // Create table XML
    const newTableXml = createTableXml(newTableId, newTableIndex, rows, cols, position, size, cellWidth, cellHeight);

    // Insert table before </p:spTree>
    const updatedSlideXml = slideXml.replace(
      /<\/p:spTree>/,
      `${newTableXml}\n  </p:spTree>`
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
    return ok({ path: `/slide[${slideIndex}]/table[${newTableIndex}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Creates XML for a new table.
 */
function createTableXml(
  id: number,
  tableIndex: number,
  rows: number,
  cols: number,
  position: TablePosition,
  size: TableSize,
  cellWidth: number,
  cellHeight: number,
): string {
  // Build grid columns
  const gridCols = Array(cols).fill(0).map((_, i) =>
    `        <a:gridCol w="${cellWidth}"/>`
  ).join("\n");

  // Build table rows
  const tableRows = Array(rows).fill(0).map((_, rowIdx) => {
    const cells = Array(cols).fill(0).map((_, colIdx) =>
      `          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:endParaRPr/>
              </a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>`
    ).join("\n");

    return `        <a:tr h="${cellHeight}">
${cells}
        </a:tr>`;
  }).join("\n");

  return `    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="${id}" name="Table ${tableIndex}"/>
        <p:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="${position.x}" y="${position.y}"/>
        <a:ext cx="${size.width}" cy="${size.height}"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1" bandRow="1">
              <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
            </a:tblPr>
            <a:tblGrid>
${gridCols}
            </a:tblGrid>
${tableRows}
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>`;
}

/**
 * Sets the style of a table.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the table (e.g., "/slide[1]/table[1]")
 * @param styleId - The table style ID (e.g., "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
 *
 * @example
 * const result = await setTableStyle("/path/to/presentation.pptx", "/slide[1]/table[1]", "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
 */
export async function setTableStyle(
  filePath: string,
  pptPath: string,
  styleId: string,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("setTableStyle requires a slide path");
    }

    const tableIndex = extractTableIndex(pptPath);
    if (!tableIndex) {
      return invalidInput("Invalid table path format. Expected: /slide[N]/table[M]");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message, slidePathResult.error!.suggestion);
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry!);

    const updatedSlideXml = setTableStyleInSlide(slideXml, tableIndex, styleId);

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
 * Sets the style of a table in the slide XML.
 */
function setTableStyleInSlide(slideXml: string, tableIndex: number, styleId: string): string {
  // Find the table by index
  const tablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  const tables = slideXml.match(tablePattern);

  if (!tables || tableIndex < 1 || tableIndex > tables.length) {
    throw new Error(`Table index ${tableIndex} out of range`);
  }

  const targetTableXml = tables[tableIndex - 1];

  // Find and replace the tableStyleId
  const updatedTableXml = targetTableXml.replace(
    /<a:tableStyleId>[^<]*<\/a:tableStyleId>/,
    `<a:tableStyleId>${styleId}</a:tableStyleId>`
  );

  return slideXml.replace(targetTableXml, updatedTableXml);
}

// ============================================================================
// Insert Table Row Operations
// ============================================================================

/**
 * Inserts a new row into a table.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the table (e.g., "/slide[1]/table[1]")
 * @param beforeIndex - 1-based index where the new row should be inserted (before this row)
 *
 * @example
 * const result = await insertTableRow("/path/to/presentation.pptx", "/slide[1]/table[1]", 2);
 * if (result.ok) {
 *   console.log(result.data.path); // "/slide[1]/table[1]/tr[2]"
 * }
 */
export async function insertTableRow(
  filePath: string,
  pptPath: string,
  beforeIndex: number,
): Promise<Result<{ path: string }>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("insertTableRow requires a slide path");
    }

    const tableIndex = extractTableIndex(pptPath);
    if (!tableIndex) {
      return invalidInput("Invalid table path format. Expected: /slide[N]/table[M]");
    }

    if (beforeIndex < 1) {
      return invalidInput("beforeIndex must be at least 1");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message, slidePathResult.error!.suggestion);
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry!);

    const result = insertRowInTable(slideXml, tableIndex, beforeIndex);
    if (!result.ok) {
      return err(result.error!.code, result.error!.message, result.error!.suggestion);
    }

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(result.data!, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));
    return ok({ path: `/slide[${slideIndex}]/table[${tableIndex}]/tr[${beforeIndex}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Inserts a row into the table XML.
 */
function insertRowInTable(slideXml: string, tableIndex: number, beforeIndex: number): Result<string> {
  // Find the table by index
  const tablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  const tables = slideXml.match(tablePattern);

  if (!tables || tableIndex < 1 || tableIndex > tables.length) {
    return invalidInput(`Table index ${tableIndex} out of range`);
  }

  const targetTableXml = tables[tableIndex - 1];

  // Find the tblGrid to get column count and widths
  const gridMatch = targetTableXml.match(/<a:tblGrid>([\s\S]*?)<\/a:tblGrid>/);
  if (!gridMatch) {
    return invalidInput("Table has no grid definition");
  }

  const gridCols = gridMatch[1].match(/<a:gridCol[^>]*w="([^"]*)"[^>]*\/>/g);
  const colCount = gridCols ? gridCols.length : 1;
  const colWidth = gridCols && gridCols[0] ? parseInt(/w="(\d+)"/.exec(gridCols[0])?.[1] || "100000", 10) : 100000;

  // Get row height from existing rows
  const rowPattern = /<a:tr[^>]*h="([^"]*)"[^>]*>[\s\S]*?<\/a:tr>/g;
  const existingRows = targetTableXml.match(rowPattern);
  const rowHeight = existingRows && existingRows[0] ? parseInt(/h="(\d+)"/.exec(existingRows[0])?.[1] || "100000", 10) : 100000;

  // Create new row XML
  const newRowCells = Array(colCount).fill(0).map(() =>
    `          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:endParaRPr/>
              </a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>`
  ).join("\n");

  const newRowXml = `        <a:tr h="${rowHeight}">
${newRowCells}
        </a:tr>`;

  // Find all rows and insert at the correct position
  const fullRowPattern = /<a:tr[^>]*>[\s\S]*?<\/a:tr>/g;
  const rows = targetTableXml.match(fullRowPattern);

  if (!rows) {
    return invalidInput("Table has no rows");
  }

  if (beforeIndex > rows.length) {
    return invalidInput(`beforeIndex ${beforeIndex} is out of range (1-${rows.length + 1})`);
  }

  // Insert the new row
  let updatedTableXml: string;
  if (beforeIndex === 1) {
    // Insert at the beginning (after tblPr/tblGrid)
    const insertPoint = targetTableXml.indexOf("</a:tblGrid>") + "</a:tblGrid>".length;
    updatedTableXml = targetTableXml.slice(0, insertPoint) + "\n" + newRowXml + targetTableXml.slice(insertPoint);
  } else {
    // Insert after the (beforeIndex - 1)th row
    let insertAfterIdx = 0;
    let searchStart = 0;
    for (let i = 0; i < beforeIndex - 1; i++) {
      const rowMatch = targetTableXml.slice(searchStart).match(/<a:tr[^>]*>[\s\S]*?<\/a:tr>/);
      if (rowMatch) {
        insertAfterIdx = searchStart + rowMatch.index! + rowMatch[0].length;
        searchStart = insertAfterIdx;
      }
    }
    updatedTableXml = targetTableXml.slice(0, insertAfterIdx) + "\n" + newRowXml + targetTableXml.slice(insertAfterIdx);
  }

  return ok(slideXml.replace(targetTableXml, updatedTableXml));
}

// ============================================================================
// Insert Table Column Operations
// ============================================================================

/**
 * Inserts a new column into a table.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the table (e.g., "/slide[1]/table[1]")
 * @param beforeIndex - 1-based index where the new column should be inserted (before this column)
 *
 * @example
 * const result = await insertTableColumn("/path/to/presentation.pptx", "/slide[1]/table[1]", 2);
 */
export async function insertTableColumn(
  filePath: string,
  pptPath: string,
  beforeIndex: number,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("insertTableColumn requires a slide path");
    }

    const tableIndex = extractTableIndex(pptPath);
    if (!tableIndex) {
      return invalidInput("Invalid table path format. Expected: /slide[N]/table[M]");
    }

    if (beforeIndex < 1) {
      return invalidInput("beforeIndex must be at least 1");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message, slidePathResult.error!.suggestion);
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry!);

    const updatedSlideXml = insertColumnInTable(slideXml, tableIndex, beforeIndex);

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
 * Inserts a column into the table XML.
 */
function insertColumnInTable(slideXml: string, tableIndex: number, beforeIndex: number): string {
  // Find the table by index
  const tablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  const tables = slideXml.match(tablePattern);

  if (!tables || tableIndex < 1 || tableIndex > tables.length) {
    throw new Error(`Table index ${tableIndex} out of range`);
  }

  const targetTableXml = tables[tableIndex - 1];

  // Find the tblGrid to get column widths
  const gridMatch = targetTableXml.match(/<a:tblGrid>([\s\S]*?)<\/a:tblGrid>/);
  const gridCols = gridMatch?.[1].match(/<a:gridCol[^>]*w="([^"]*)"[^>]*\/>/g) || [];

  // Get column width (use existing or default)
  const colWidth = gridCols.length > 0 && gridCols[0] ? parseInt(/w="(\d+)"/.exec(gridCols[0])?.[1] || "100000", 10) : 100000;

  // Insert new gridCol at the correct position
  const newGridColXml = `<a:gridCol w="${colWidth}"/>`;
  let updatedTableXml = targetTableXml;

  if (gridCols.length > 0 && beforeIndex <= gridCols.length) {
    // Insert gridCol at position
    const gridColsList = targetTableXml.match(/<a:tblGrid>([\s\S]*?)<\/a:tblGrid>/);
    if (gridColsList) {
      const gridContent = gridColsList[1];
      const cols = gridContent.match(/<a:gridCol[^>]*\/>/g) || [];
      if (beforeIndex <= cols.length) {
        cols.splice(beforeIndex - 1, 0, newGridColXml);
        updatedTableXml = targetTableXml.replace(
          /<a:tblGrid>[\s\S]*?<\/a:tblGrid>/,
          `<a:tblGrid>${cols.join("")}</a:tblGrid>`
        );
      }
    }
  }

  // Insert new cell in each row at the correct position
  const rowPattern = /<a:tr[^>]*>[\s\S]*?<\/a:tr>/g;
  const rows = updatedTableXml.match(rowPattern);

  if (rows) {
    for (let i = 0; i < rows.length; i++) {
      const rowXml = rows[i];
      const cells = rowXml.match(/<a:tc[\s\S]*?<\/a:tc>/g) || [];

      const newCellXml = `          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:endParaRPr/>
              </a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>`;

      if (beforeIndex <= cells.length) {
        cells.splice(beforeIndex - 1, 0, newCellXml);

        // Reconstruct the row with updated cells
        const rowHeightMatch = rowXml.match(/<a:tr[^>]*h="([^"]*)"[^>]*>/);
        const rowHeight = rowHeightMatch ? rowHeightMatch[1] : "100000";

        const updatedRow = `        <a:tr h="${rowHeight}">
${cells.join("\n")}
        </a:tr>`;

        updatedTableXml = updatedTableXml.replace(rowXml, updatedRow);
      }
    }
  }

  return slideXml.replace(targetTableXml, updatedTableXml);
}

// ============================================================================
// Merge Table Cells Operations
// ============================================================================

/**
 * Merges cells in a table.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the table (e.g., "/slide[1]/table[1]")
 * @param startRow - 1-based starting row index
 * @param startCol - 1-based starting column index
 * @param endRow - 1-based ending row index
 * @param endCol - 1-based ending column index
 *
 * @example
 * // Merge cells from row 1, col 1 to row 2, col 3
 * const result = await mergeTableCells("/path/to/presentation.pptx", "/slide[1]/table[1]", 1, 1, 2, 3);
 */
export async function mergeTableCells(
  filePath: string,
  pptPath: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(pptPath);
    if (slideIndex === null) {
      return invalidInput("mergeTableCells requires a slide path");
    }

    const tableIndex = extractTableIndex(pptPath);
    if (!tableIndex) {
      return invalidInput("Invalid table path format. Expected: /slide[N]/table[M]");
    }

    if (startRow < 1 || startCol < 1 || endRow < startRow || endCol < startCol) {
      return invalidInput("Invalid cell range");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error!.code, slidePathResult.error!.message, slidePathResult.error!.suggestion);
    }

    const slideEntry = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntry!);

    const updatedSlideXml = mergeCellsInTable(slideXml, tableIndex, startRow, startCol, endRow, endCol);

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
 * Merges cells in the table XML.
 */
function mergeCellsInTable(
  slideXml: string,
  tableIndex: number,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): string {
  // Find the table by index
  const tablePattern = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  const tables = slideXml.match(tablePattern);

  if (!tables || tableIndex < 1 || tableIndex > tables.length) {
    throw new Error(`Table index ${tableIndex} out of range`);
  }

  const targetTableXml = tables[tableIndex - 1];

  // Get all rows
  const rowPattern = /<a:tr[^>]*>[\s\S]*?<\/a:tr>/g;
  const rows = targetTableXml.match(rowPattern);

  if (!rows || startRow > rows.length || endRow > rows.length) {
    throw new Error("Row index out of range");
  }

  let updatedTableXml = targetTableXml;

  // Process each row in the merge range
  for (let r = startRow - 1; r < endRow; r++) {
    const rowXml = rows[r];
    const cells = rowXml.match(/<a:tc[\s\S]*?<\/a:tc>/g) || [];

    if (startCol > cells.length || endCol > cells.length) {
      throw new Error("Column index out of range");
    }

    // For the starting cell, add gridSpan and rowSpan attributes
    if (r === startRow - 1) {
      const startCellXml = cells[startCol - 1];
      const gridSpan = endCol - startCol + 1;
      const rowSpan = endRow - startRow + 1;

      // Update start cell with merge attributes
      let updatedStartCell = startCellXml;

      // Add gridSpan to tcPr
      if (gridSpan > 1) {
        updatedStartCell = updatedStartCell.replace(
          /<a:tcPr([^>]*)\/>/,
          `<a:tcPr$1 gridSpan="${gridSpan}"/>`
        );
        // If no tcPr exists, create one
        if (!updatedStartCell.includes("gridSpan")) {
          updatedStartCell = updatedStartCell.replace(
            /<a:tc>/,
            `<a:tc><a:tcPr gridSpan="${gridSpan}"/>`
          );
        }
      }

      // Add rowSpan to tcPr
      if (rowSpan > 1) {
        updatedStartCell = updatedStartCell.replace(
          /<a:tcPr([^>]*)\/>/,
          `<a:tcPr$1 rowSpan="${rowSpan}"/>`
        );
        if (!updatedStartCell.includes("rowSpan")) {
          updatedStartCell = updatedStartCell.replace(
            /<a:tc>/,
            `<a:tc><a:tcPr rowSpan="${rowSpan}"/>`
          );
        }
      }

      // Replace cells in range with the merged cell (for start row, keep start cell with span)
      const mergedRowCells = cells.slice(0, startCol - 1);
      mergedRowCells.push(updatedStartCell);

      // For the remaining cells in the start row that are being merged, mark them as hmerged
      for (let c = startCol; c < endCol; c++) {
        const cellXml = cells[c];
        const hmergedCell = cellXml.replace(
          /<a:tc>/,
          `<a:tc><a:tcPr hMerge="1"/>`
        );
        mergedRowCells.push(hmergedCell);
      }

      // Add remaining cells after the merge
      for (let c = endCol; c < cells.length; c++) {
        mergedRowCells.push(cells[c]);
      }

      // Reconstruct the row
      const rowHeightMatch = rowXml.match(/<a:tr[^>]*h="([^"]*)"[^>]*>/);
      const rowHeight = rowHeightMatch ? rowHeightMatch[1] : "100000";

      const updatedRow = `        <a:tr h="${rowHeight}">
${mergedRowCells.join("\n")}
        </a:tr>`;

      updatedTableXml = updatedTableXml.replace(rowXml, updatedRow);
    } else {
      // For rows in the merge range (but not the start row), mark cells as vmerged
      const mergedRowCells = cells.slice(0, startCol - 1);

      for (let c = startCol - 1; c < endCol; c++) {
        const cellXml = cells[c];
        const vmergedCell = cellXml.replace(
          /<a:tc>/,
          `<a:tc><a:tcPr vMerge="1"/>`
        );
        mergedRowCells.push(vmergedCell);
      }

      // Add remaining cells after the merge
      for (let c = endCol; c < cells.length; c++) {
        mergedRowCells.push(cells[c]);
      }

      // Reconstruct the row
      const rowHeightMatch = rowXml.match(/<a:tr[^>]*h="([^"]*)"[^>]*>/);
      const rowHeight = rowHeightMatch ? rowHeightMatch[1] : "100000";

      const updatedRow = `        <a:tr h="${rowHeight}">
${mergedRowCells.join("\n")}
        </a:tr>`;

      updatedTableXml = updatedTableXml.replace(rowXml, updatedRow);
    }
  }

  return slideXml.replace(targetTableXml, updatedTableXml);
}
