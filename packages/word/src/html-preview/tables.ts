/**
 * Table rendering for Word HTML preview.
 * Handles tables, rows, cells, colspan, rowspan, and borders.
 */

import { generateTableCellCss } from "./css.js";

export interface TableRenderOptions {
  /** Whether table has no visible borders */
  borderless?: boolean;
  /** Table style ID for conditional formatting */
  styleId?: string;
  /** Table grid for column widths */
  gridCols?: number[];
}

/**
 * Render a table to HTML.
 */
export function renderTableHtml(tableXml: string, options: TableRenderOptions = {}): string {
  const { borderless = false, styleId, gridCols = [] } = options;

  // Parse table properties
  const tblPrMatch = /<w:tblPr>([\s\S]*?)<\/w:tblPr>/i.exec(tableXml);
  const tblProps = parseTableProperties(tblPrMatch?.[1] || "");

  // Determine if table is actually borderless
  const isBorderless = borderless || tblProps.bordersNone || false;

  // Parse table look for conditional formatting
  const tblLook = parseTableLook(tblPrMatch?.[1] || "");

  // Get table justification/alignment
  const tableJustify = tblProps.justification;

  // Build table styles
  const tableStyles: string[] = [];

  if (tblProps.width) {
    tableStyles.push(`width: ${tblProps.width}pt`);
  }

  if (tableJustify === "center") {
    tableStyles.push("margin-left: auto");
    tableStyles.push("margin-right: auto");
  } else if (tableJustify === "right") {
    tableStyles.push("margin-left: auto");
    tableStyles.push("margin-right: 0");
  }

  if (isBorderless) {
    tableStyles.push("border: none");
  }

  // Table-level float for floating tables
  if (tblProps.float) {
    if (tblProps.float.horizontalAnchor === "page" && (tblProps.float.positionX ?? 0) > 5000) {
      tableStyles.push("float: right");
    } else {
      tableStyles.push("float: left");
    }
    if (tblProps.float.marginLeft) {
      tableStyles.push(`margin-left: ${tblProps.float.marginLeft}pt`);
    }
    if (tblProps.float.marginRight) {
      tableStyles.push(`margin-right: ${tblProps.float.marginRight}pt`);
    }
  }

  const tableStyleAttr = tableStyles.length > 0 ? ` style="${tableStyles.join("; ")}"` : "";
  const tableClass = isBorderless ? " borderless" : "";

  let html = `<table class="${tableClass}"${tableStyleAttr}>`;

  // Render column group for column widths
  if (gridCols.length > 0) {
    html += "\n<colgroup>";
    for (const colWidth of gridCols) {
      if (colWidth > 0) {
        html += `<col style="width: ${colWidth}pt">`;
      } else {
        html += "<col>";
      }
    }
    html += "</colgroup>";
  }

  // Parse table grid for column widths if not provided
  const tableGrid = /<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/i.exec(tableXml);
  if (tableGrid && gridCols.length === 0) {
    const colWidths: number[] = [];
    const colRegex = /<w:gridCol[^>]*><w:w[^>]*w:w="([^"]*)"[^>]*\/><\/w:gridCol>/gi;
    let colMatch;
    while ((colMatch = colRegex.exec(tableGrid[1])) !== null) {
      colWidths.push(parseInt(colMatch[1], 10) / 20); // Convert twips to pt
    }
    if (colWidths.length > 0) {
      html += "\n<colgroup>";
      for (const colWidth of colWidths) {
        html += `<col style="width: ${colWidth}pt">`;
      }
      html += "</colgroup>";
    }
  }

  // Extract rows
  const rows: string[] = [];
  const rowRegex = /<w:tr[\s\S]*?<\/w:tr>/gi;
  let rowMatch;

  while ((rowMatch = rowRegex.exec(tableXml)) !== null) {
    rows.push(rowMatch[0]);
  }

  const totalRows = rows.length;
  let totalCols = gridCols.length || (totalRows > 0 ? estimateColumnCount(rows[0]) : 0);

  // Render each row
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const rowXml = rows[rowIdx];
    const rowHtml = renderTableRow(
      rowXml,
      rowIdx,
      totalRows,
      totalCols,
      isBorderless,
      tblProps.borders,
      tblLook,
      styleId
    );
    html += "\n" + rowHtml;
  }

  html += "\n</table>";

  return html;
}

/**
 * Render a table row to HTML.
 */
function renderTableRow(
  rowXml: string,
  rowIdx: number,
  totalRows: number,
  totalCols: number,
  borderless: boolean,
  tableBorders?: TableBorders,
  tblLook?: TableLook,
  styleId?: string
): string {
  // Check if this is a header row
  const isHeaderRow = /<w:tblHeader[^>]*\/>/i.test(rowXml)
    || (tblLook?.firstRow === true && rowIdx === 0);

  // Row height
  const trHeightMatch = /<w:trHeight[^>]*w:val="([^"]*)"/i.exec(rowXml);
  let trStyle = "";
  if (trHeightMatch) {
    const heightTwips = parseInt(trHeightMatch[1], 10);
    if (heightTwips > 0) {
      trStyle = ` style="height: ${heightTwips / 20}pt"`;
    }
  }

  const rowClass = isHeaderRow ? "header-row" : "";
  const rowClassAttr = rowClass ? ` class="${rowClass}"` : "";

  let html = `<tr${rowClassAttr}${trStyle}>`;

  // Extract cells
  const cells: string[] = [];
  const cellRegex = /<w:tc>[\s\S]*?<\/w:tc>/gi;
  let cellMatch;

  while ((cellMatch = cellRegex.exec(rowXml)) !== null) {
    cells.push(cellMatch[0]);
  }

  let colIdx = 0;
  for (const cellXml of cells) {
    // Check for grid span (colspan)
    const gridSpanMatch = /<w:gridSpan[^>]*w:val="([^"]*)"/i.exec(cellXml);
    const gridSpan = gridSpanMatch ? parseInt(gridSpanMatch[1], 10) : 1;

    // Check for vertical merge (rowspan)
    const vMergeMatch = /<w:vMerge[^>]*w:val="([^"]*)"/i.exec(cellXml);
    const vMergeStart = !vMergeMatch || vMergeMatch[1] !== "continue";

    // Calculate rowspan
    let rowSpan = 1;
    if (vMergeStart) {
      rowSpan = countRowSpan(cells, rowIdx, colIdx, totalRows);
    } else if (vMergeMatch) {
      // This cell is continued from previous row, skip it
      colIdx += gridSpan;
      continue;
    }

    // Determine cell type
    const cellTag = isHeaderRow ? "th" : "td";

    // Build cell attributes
    const cellAttrs: string[] = [];

    if (gridSpan > 1) {
      cellAttrs.push(`colspan="${gridSpan}"`);
    }

    if (rowSpan > 1) {
      cellAttrs.push(`rowspan="${rowSpan}"`);
    }

    // Determine conditional formatting type
    const condTypes = getConditionalTypes(tblLook, rowIdx, colIdx, totalRows, totalCols);

    // Generate cell CSS
    const cellCss = generateTableCellCss({
      fill: extractCellFill(cellXml),
      valign: extractCellValign(cellXml),
      width: extractCellWidth(cellXml),
      borders: extractCellBorders(cellXml, borderless, tableBorders),
    });

    if (cellCss) {
      cellAttrs.push(`style="${cellCss}"`);
    }

    // Check if this is an odd/even row for conditional formatting
    if (condTypes.isFirstColumn && tblLook?.firstColumn === false) {
      // Don't apply first column formatting
    } else if (condTypes.isLastColumn && tblLook?.lastColumn === false) {
      // Don't apply last column formatting
    }

    const cellAttrsStr = cellAttrs.length > 0 ? " " + cellAttrs.join(" ") : "";

    // Render cell content
    const cellContent = renderTableCellContent(cellXml);

    html += `<${cellTag}${cellAttrsStr}>${cellContent}</${cellTag}>`;

    colIdx += gridSpan;
  }

  html += "</tr>";

  return html;
}

/**
 * Render table cell content.
 */
function renderTableCellContent(cellXml: string): string {
  // Extract paragraphs from cell
  const paras: string[] = [];
  const paraRegex = /<w:p[\s\S]*?<\/w:p>/gi;
  let paraMatch;

  while ((paraMatch = paraRegex.exec(cellXml)) !== null) {
    const paraXml = paraMatch[0];
    const paraContent = extractTextFromCellPara(paraXml);
    paras.push(`<p>${paraContent}</p>`);
  }

  if (paras.length === 0) {
    return "<p>&nbsp;</p>";
  }

  return paras.join("");
}

/**
 * Extract text from a cell paragraph including runs.
 */
function extractTextFromCellPara(paraXml: string): string {
  let html = "";

  // Extract runs
  const runRegex = /<w:r[\s\S]*?<\/w:r>/gi;
  let runMatch;

  while ((runMatch = runRegex.exec(paraXml)) !== null) {
    const runXml = runMatch[0];

    // Extract text
    const textMatch = /<w:t[^>]*>([^<]*)<\/w:t>/i.exec(runXml);
    const text = textMatch ? textMatch[1] : "";

    // Check for formatting
    let runHtml = escapeHtml(text);

    // Bold
    if (/<w:b(?![a-z])[^>]*>/i.test(runXml)) {
      runHtml = `<strong>${runHtml}</strong>`;
    }

    // Italic
    if (/<w:i(?![a-z])[^>]*>/i.test(runXml)) {
      runHtml = `<em>${runHtml}</em>`;
    }

    // Underline
    if (/<w:u[^>]*>/i.test(runXml)) {
      runHtml = `<u>${runHtml}</u>`;
    }

    // Strike
    if (/<w:strike[^>]*>/i.test(runXml) || /<w:dstrike[^>]*>/i.test(runXml)) {
      runHtml = `<s>${runHtml}</s>`;
    }

    // Color
    const colorMatch = /<w:color[^>]*w:val="([^"]*)"/i.exec(runXml);
    if (colorMatch) {
      runHtml = `<span style="color: #${colorMatch[1]}">${runHtml}</span>`;
    }

    // Size
    const sizeMatch = /<w:sz[^>]*w:val="([^"]*)"/i.exec(runXml);
    if (sizeMatch) {
      const halfPt = parseInt(sizeMatch[1], 10);
      runHtml = `<span style="font-size: ${halfPt / 2}pt">${runHtml}</span>`;
    }

    html += runHtml;
  }

  // Handle tab in cell
  if (paraXml.includes("<w:tab/>")) {
    html += "\t";
  }

  // Handle line break
  if (paraXml.includes("<w:br/>")) {
    html += "<br>";
  }

  return html;
}

interface TableProperties {
  width?: number;
  justification?: string;
  bordersNone?: boolean;
  borders?: TableBorders;
  float?: {
    horizontalAnchor?: string;
    positionX?: number;
    marginLeft?: number;
    marginRight?: number;
    marginTop?: number;
    marginBottom?: number;
  };
}

interface TableBorders {
  top?: BorderInfo;
  bottom?: BorderInfo;
  left?: BorderInfo;
  right?: BorderInfo;
  insideH?: BorderInfo;
  insideV?: BorderInfo;
}

interface BorderInfo {
  style?: string;
  size?: number;
  color?: string;
}

interface TableLook {
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  noHBand?: boolean;
  noVBand?: boolean;
}

function parseTableProperties(tblPrContent: string): TableProperties {
  const props: TableProperties = {};

  // Table width
  const widthMatch = /<w:tblW[^>]*w:w="([^"]*)"/i.exec(tblPrContent);
  if (widthMatch) {
    props.width = parseInt(widthMatch[1], 10) / 20;
  }

  // Justification
  const jcMatch = /<w:jc[^>]*w:val="([^"]*)"/i.exec(tblPrContent);
  if (jcMatch) props.justification = jcMatch[1];

  // Table borders
  const bordersMatch = /<w:tblBorders>([\s\S]*?)<\/w:tblBorders>/i.exec(tblPrContent);
  if (bordersMatch) {
    props.borders = parseTableBorders(bordersMatch[1]);
    props.bordersNone = isAllBordersNone(props.borders);
  }

  // Floating table position
  const tblpPrMatch = /<w:tblpPr>([\s\S]*?)<\/w:tblpPr>/i.exec(tblPrContent);
  if (tblpPrMatch) {
    const tblpContent = tblpPrMatch[1];
    const hAnchorMatch = /<w:horzAnchor[^>]*w:val="([^"]*)"/i.exec(tblpContent);
    const tblpXMatch = /<w:tblpX[^>]*w:val="([^"]*)"/i.exec(tblpContent);
    const leftDistMatch = /<w:leftFromText[^>]*w:val="([^"]*)"/i.exec(tblpContent);
    const rightDistMatch = /<w:rightFromText[^>]*w:val="([^"]*)"/i.exec(tblpContent);
    const topDistMatch = /<w:topFromText[^>]*w:val="([^"]*)"/i.exec(tblpContent);
    const bottomDistMatch = /<w:bottomFromText[^>]*w:val="([^"]*)"/i.exec(tblpContent);

    props.float = {
      horizontalAnchor: hAnchorMatch ? hAnchorMatch[1] : undefined,
      positionX: tblpXMatch ? parseInt(tblpXMatch[1], 10) : undefined,
      marginLeft: leftDistMatch ? parseInt(leftDistMatch[1], 10) / 20 : undefined,
      marginRight: rightDistMatch ? parseInt(rightDistMatch[1], 10) / 20 : undefined,
      marginTop: topDistMatch ? parseInt(topDistMatch[1], 10) / 20 : undefined,
      marginBottom: bottomDistMatch ? parseInt(bottomDistMatch[1], 10) / 20 : undefined,
    };
  }

  return props;
}

function parseTableBorders(bordersContent: string): TableBorders {
  const borders: TableBorders = {};

  const topMatch = /<w:top[^>]*>/i.exec(bordersContent);
  if (topMatch) borders.top = parseBorder(topMatch[0]);

  const bottomMatch = /<w:bottom[^>]*>/i.exec(bordersContent);
  if (bottomMatch) borders.bottom = parseBorder(bottomMatch[0]);

  const leftMatch = /<w:left[^>]*>/i.exec(bordersContent);
  if (leftMatch) borders.left = parseBorder(leftMatch[0]);

  const rightMatch = /<w:right[^>]*>/i.exec(bordersContent);
  if (rightMatch) borders.right = parseBorder(rightMatch[0]);

  const insideHMatch = /<w:insideH[^>]*>/i.exec(bordersContent);
  if (insideHMatch) borders.insideH = parseBorder(insideHMatch[0]);

  const insideVMatch = /<w:insideV[^>]*>/i.exec(bordersContent);
  if (insideVMatch) borders.insideV = parseBorder(insideVMatch[0]);

  return borders;
}

function parseBorder(borderXml: string): BorderInfo {
  const info: BorderInfo = {};

  const valMatch = /w:val="([^"]*)"/i.exec(borderXml);
  if (valMatch) info.style = valMatch[1];

  const szMatch = /w:sz="([^"]*)"/i.exec(borderXml);
  if (szMatch) info.size = parseInt(szMatch[1], 10);

  const colorMatch = /w:color="([^"]*)"/i.exec(borderXml);
  if (colorMatch) info.color = colorMatch[1];

  return info;
}

function isAllBordersNone(borders: TableBorders): boolean {
  const borderVals = [
    borders.top?.style,
    borders.bottom?.style,
    borders.left?.style,
    borders.right?.style,
    borders.insideH?.style,
    borders.insideV?.style,
  ];
  return borderVals.every((v) => v === undefined || v === "none" || v === "nil");
}

function parseTableLook(tblPrContent: string): TableLook | undefined {
  const tblLookMatch = /<w:tblLook[^>]*>/i.exec(tblPrContent);
  if (!tblLookMatch) return undefined;

  const content = tblLookMatch[0];

  // tblLook is a 4-character hex bitmask
  const valMatch = /w:val="([^"]*)"/i.exec(content);
  if (!valMatch) return undefined;

  const val = parseInt(valMatch[1], 16);

  return {
    firstRow: (val & 0x020) !== 0,
    lastRow: (val & 0x040) !== 0,
    firstColumn: (val & 0x080) !== 0,
    lastColumn: (val & 0x100) !== 0,
    noHBand: (val & 0x200) !== 0,
    noVBand: (val & 0x400) !== 0,
  };
}

interface ConditionalTypes {
  isFirstRow: boolean;
  isLastRow: boolean;
  isFirstColumn: boolean;
  isLastColumn: boolean;
}

function getConditionalTypes(
  tblLook: TableLook | undefined,
  rowIdx: number,
  colIdx: number,
  totalRows: number,
  totalCols: number
): ConditionalTypes {
  return {
    isFirstRow: rowIdx === 0,
    isLastRow: rowIdx === totalRows - 1,
    isFirstColumn: colIdx === 0,
    isLastColumn: colIdx === totalCols - 1,
  };
}

function countRowSpan(cells: string[], startRowIdx: number, colIdx: number, totalRows: number): number {
  let rowSpan = 1;

  for (let r = startRowIdx + 1; r < totalRows; r++) {
    const rowXml = cells[r];
    if (!rowXml) break;

    // Find the cell at the same column index
    const rowCells: string[] = [];
    const cellRegex = /<w:tc>[\s\S]*?<\/w:tc>/gi;
    let cellMatch;

    while ((cellMatch = cellRegex.exec(rowXml)) !== null) {
      rowCells.push(cellMatch[0]);
    }

    const cellAtCol = rowCells[colIdx];
    if (!cellAtCol) break;

    // Check if this cell has vMerge continue
    const vMergeMatch = /<w:vMerge[^>]*w:val="([^"]*)"/i.exec(cellAtCol);
    if (vMergeMatch && vMergeMatch[1] === "continue") {
      rowSpan++;
    } else {
      break;
    }
  }

  return rowSpan;
}

function extractCellFill(cellXml: string): string | undefined {
  // Try cell properties first
  const tcPrMatch = /<w:tcPr>([\s\S]*?)<\/w:tcPr>/i.exec(cellXml);
  if (tcPrMatch) {
    const shadingMatch = /<w:shd[^>]*w:fill="([^"]*)"/i.exec(tcPrMatch[1]);
    if (shadingMatch) return shadingMatch[1];
  }

  // Check for w:tcBorders (sometimes fill is there)
  const tcBordersMatch = /<w:tcBorders>([\s\S]*?)<\/w:tcBorders>/i.exec(cellXml);
  if (tcBordersMatch) {
    // Cell borders don't have fill, but we check for completeness
  }

  return undefined;
}

function extractCellValign(cellXml: string): string | undefined {
  const tcPrMatch = /<w:tcPr>([\s\S]*?)<\/w:tcPr>/i.exec(cellXml);
  if (tcPrMatch) {
    const vAlignMatch = /<w:vAlign[^>]*w:val="([^"]*)"/i.exec(tcPrMatch[1]);
    if (vAlignMatch) return vAlignMatch[1];
  }
  return undefined;
}

function extractCellWidth(cellXml: string): number | undefined {
  const tcPrMatch = /<w:tcPr>([\s\S]*?)<\/w:tcPr>/i.exec(cellXml);
  if (tcPrMatch) {
    const widthMatch = /<w:tcW[^>]*w:w="([^"]*)"/i.exec(tcPrMatch[1]);
    if (widthMatch) {
      return parseInt(widthMatch[1], 10) / 20;
    }
  }
  return undefined;
}

function extractCellBorders(
  cellXml: string,
  borderless: boolean,
  tableBorders?: TableBorders
): { top?: BorderInfo; bottom?: BorderInfo; left?: BorderInfo; right?: BorderInfo } | undefined {
  const tcPrMatch = /<w:tcPr>([\s\S]*?)<\/w:tcPr>/i.exec(cellXml);
  if (!tcPrMatch) {
    // Use table-level borders if no cell-level borders
    return borderless ? undefined : tableBorders;
  }

  const tcBordersMatch = /<w:tcBorders>([\s\S]*?)<\/w:tcBorders>/i.exec(tcPrMatch[1]);
  if (!tcBordersMatch) {
    return borderless ? undefined : tableBorders;
  }

  const bordersContent = tcBordersMatch[1];

  // Check for "none" or "nil" borders
  const noBorders = /<w:top[^>]*w:val="none"[^>]*\/?>/i.test(bordersContent)
    || /<w:top[^>]*w:val="nil"[^>]*\/?>/i.test(bordersContent);

  if (noBorders) {
    return undefined; // No borders
  }

  const borders: { top?: BorderInfo; bottom?: BorderInfo; left?: BorderInfo; right?: BorderInfo } = {};

  const topMatch = /<w:top[^>]*>/i.exec(bordersContent);
  if (topMatch) borders.top = parseBorder(topMatch[0]);

  const bottomMatch = /<w:bottom[^>]*>/i.exec(bordersContent);
  if (bottomMatch) borders.bottom = parseBorder(bottomMatch[0]);

  const leftMatch = /<w:left[^>]*>/i.exec(bordersContent);
  if (leftMatch) borders.left = parseBorder(leftMatch[0]);

  const rightMatch = /<w:right[^>]*>/i.exec(bordersContent);
  if (rightMatch) borders.right = parseBorder(rightMatch[0]);

  return borders;
}

function estimateColumnCount(rowXml: string): number {
  let count = 0;
  const cellRegex = /<w:tc>[\s\S]*?<\/w:tc>/gi;
  let match;

  while ((match = cellRegex.exec(rowXml)) !== null) {
    const cellXml = match[0];
    const gridSpanMatch = /<w:gridSpan[^>]*w:val="([^"]*)"/i.exec(cellXml);
    count += gridSpanMatch ? parseInt(gridSpanMatch[1], 10) : 1;
  }

  return count;
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
