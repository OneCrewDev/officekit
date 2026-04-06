import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { OfficekitError, UsageError } from "./errors.js";
import { assertFormat, type SupportedFormat } from "./formats.js";
import { createStoredZip, readStoredZip } from "./zip.js";
import {
  addExcelNode,
  createExcelDocument,
  getExcelNode,
  importExcelDelimitedData,
  queryExcelNodes as queryExcelNodesFromAdapter,
  rawExcelDocument,
  removeExcelNode,
  renderExcelHtmlFromRoot,
  setExcelNode,
  summarizeExcelCheck,
  viewExcelDocument,
} from "../../excel/src/adapter.js";
import type {
  ExcelCellModel as ExcelCell,
  ExcelNamedRangeModel as ExcelNamedRange,
  ExcelSheetModel as ExcelSheet,
  ExcelWorkbookSettings,
} from "../../excel/src/adapter.js";

export interface WordParagraph {
  text: string;
}

export interface WordParagraphNode extends WordParagraph {
  type: "paragraph";
}

export interface WordTableCell {
  text: string;
}

export interface WordTableRow {
  cells: WordTableCell[];
}

export interface WordTable {
  rows: WordTableRow[];
}

export interface WordTableNode extends WordTable {
  type: "table";
}

export type WordBodyNode = WordParagraphNode | WordTableNode;

export interface PptShape {
  text: string;
  kind?: string;
  name?: string;
}

export interface PptSlide {
  title: string;
  layoutName?: string;
  layoutType?: string;
  themeName?: string;
  shapes: PptShape[];
}

export interface OfficekitDocument {
  product: "officekit";
  lineage: string;
  format: SupportedFormat;
  version: 1;
  updatedAt: string;
  word?: {
    body: WordBodyNode[];
    paragraphs?: WordParagraph[];
    tables?: WordTable[];
  };
  excel?: {
    path?: string;
    type?: string;
    sheets: Array<{
      name: string;
      cells: Record<string, ExcelCell>;
      autoFilter?: string;
      freezeTopLeftCell?: string;
      zoom?: number;
      showGridLines?: boolean;
      showHeadings?: boolean;
      tabColor?: string;
    }>;
    settings?: ExcelWorkbookSettings;
    styleSheetXml?: string;
    namedRanges?: ExcelNamedRange[];
    metadata?: Record<string, string>;
  };
  powerpoint?: { slides: PptSlide[] };
}

const METADATA_PATH = "officekit/document.json";
const LINEAGE = "officekit is migrated from OfficeCLI and currently persists metadata-backed OOXML vertical slices.";

export interface CommandOptions {
  type?: string;
  props: Record<string, string>;
  json?: boolean;
  mode?: string;
}

export interface ImportOptions {
  delimiter: string;
  hasHeader: boolean;
  startCell: string;
}

export interface RawDocumentOptions {
  partPath?: string;
  startRow?: number;
  endRow?: number;
  cols?: string[];
}

export async function createDocument(filePath: string) {
  const format = assertFormat(filePath);
  if (format === "excel") {
    return createExcelDocument(filePath);
  }
  const document = createBlankDocument(format);
  await persistDocument(filePath, document);
  return { format, filePath, document };
}

export async function addDocumentNode(filePath: string, targetPath: string, options: CommandOptions) {
  if (assertFormat(filePath) === "excel") {
    return addExcelNode(filePath, targetPath, options);
  }
  const document = await loadDocument(filePath);
  switch (document.format) {
    case "word": {
      if (targetPath !== "/body") {
        throw new UsageError("Word add currently supports only /body.", "Use /body with --type paragraph or --type table.");
      }
      if (options.type === "paragraph") {
        document.word!.body.push(createWordParagraph(options.props.text ?? ""));
        break;
      }
      if (options.type === "table") {
        const rows = Math.max(1, Number(options.props.rows ?? "2"));
        const cols = Math.max(1, Number(options.props.cols ?? "2"));
        document.word!.body.push(createWordTable(rows, cols));
        break;
      }
      throw new UsageError(
        "Word add currently supports: add <file.docx> /body --type paragraph|table ...",
        "Use /body with --type paragraph or --type table.",
      );
    }
    case "excel":
      throw new UsageError("Excel operations are handled by the Excel adapter.");
    case "powerpoint": {
      if (targetPath === "/" && options.type === "slide") {
        document.powerpoint!.slides.push({ title: options.props.title ?? "Untitled slide", shapes: [] });
        break;
      }
      const slide = resolveSlide(document, targetPath);
      if (options.type !== "shape") {
        throw new UsageError("PowerPoint add currently supports slide creation at / and shape insertion under /slide[n].", "Use / with --type slide or /slide[n] with --type shape.");
      }
      slide.shapes.push({ text: options.props.text ?? options.props.title ?? "" });
      break;
    }
  }

  stampDocument(document);
  await persistDocument(filePath, document);
  return materializePath(document, targetPath);
}

export async function importDelimitedData(
  filePath: string,
  parentPath: string,
  content: string,
  options: ImportOptions,
) {
  if (assertFormat(filePath) === "excel") {
    return importExcelDelimitedData(filePath, parentPath, content, options);
  }
  throw new UsageError("import currently supports only .xlsx files.");
}

export async function setDocumentNode(filePath: string, targetPath: string, options: CommandOptions) {
  if (assertFormat(filePath) === "excel") {
    return setExcelNode(filePath, targetPath, options);
  }
  const document = await loadDocument(filePath);
  if (document.format === "word") {
    const match = /^\/body\/p\[(\d+)\]$/.exec(targetPath);
    const tableMatch = /^\/body\/table\[(\d+)\]\/cell\[(\d+),(\d+)\]$/.exec(targetPath);
    if (match) {
      const paragraph = resolveWordParagraph(document, Number(match[1]));
      if (!paragraph) throw new OfficekitError(`Paragraph ${match[1]} does not exist.`, "not_found");
      paragraph.text = options.props.text ?? paragraph.text;
    } else if (tableMatch) {
      const table = resolveWordTable(document, Number(tableMatch[1]));
      const row = table?.rows[Number(tableMatch[2]) - 1];
      const cell = row?.cells[Number(tableMatch[3]) - 1];
      if (!cell) {
        throw new OfficekitError(
          `Table cell ${tableMatch[2]},${tableMatch[3]} does not exist in table ${tableMatch[1]}.`,
          "not_found",
        );
      }
      cell.text = options.props.text ?? cell.text;
    } else {
      throw new UsageError(
        "Word set currently supports /body/p[n] or /body/table[n]/cell[row,col].",
        "Example: officekit set demo.docx /body/table[1]/cell[1,1] --prop text=Updated",
      );
    }
  } else if (document.format === "excel") {
    throw new UsageError("Excel operations are handled by the Excel adapter.");
  } else {
    const shapeMatch = /^\/slide\[(\d+)\]\/shape\[(\d+)\]$/.exec(targetPath);
    const slideMatch = /^\/slide\[(\d+)\]$/.exec(targetPath);
    if (shapeMatch) {
      const slide = document.powerpoint!.slides[Number(shapeMatch[1]) - 1];
      const shape = slide?.shapes[Number(shapeMatch[2]) - 1];
      if (!shape) throw new OfficekitError(`Shape ${shapeMatch[2]} does not exist.`, "not_found");
      shape.text = options.props.text ?? shape.text;
    } else if (slideMatch) {
      const slide = document.powerpoint!.slides[Number(slideMatch[1]) - 1];
      if (!slide) throw new OfficekitError(`Slide ${slideMatch[1]} does not exist.`, "not_found");
      slide.title = options.props.title ?? options.props.text ?? slide.title;
    } else {
      throw new UsageError("PowerPoint set currently supports /slide[n] or /slide[n]/shape[n].");
    }
  }

  stampDocument(document);
  await persistDocument(filePath, document);
  return materializePath(document, targetPath);
}

export async function removeDocumentNode(filePath: string, targetPath: string) {
  if (assertFormat(filePath) === "excel") {
    return removeExcelNode(filePath, targetPath);
  }
  const document = await loadDocument(filePath);
  if (document.format === "word") {
    const match = /^\/body\/p\[(\d+)\]$/.exec(targetPath);
    const tableMatch = /^\/body\/table\[(\d+)\]$/.exec(targetPath);
    if (match) {
      removeWordBodyNode(document, "paragraph", Number(match[1]));
    } else if (tableMatch) {
      removeWordBodyNode(document, "table", Number(tableMatch[1]));
    } else {
      throw new UsageError("Word remove currently supports /body/p[n] or /body/table[n].");
    }
  } else if (document.format === "excel") {
    removeExcelNode(document, targetPath);
  } else {
    const shapeMatch = /^\/slide\[(\d+)\]\/shape\[(\d+)\]$/.exec(targetPath);
    const slideMatch = /^\/slide\[(\d+)\]$/.exec(targetPath);
    if (shapeMatch) {
      const slide = document.powerpoint!.slides[Number(shapeMatch[1]) - 1];
      slide?.shapes.splice(Number(shapeMatch[2]) - 1, 1);
    } else if (slideMatch) {
      document.powerpoint!.slides.splice(Number(slideMatch[1]) - 1, 1);
    } else {
      throw new UsageError("PowerPoint remove currently supports /slide[n] or /slide[n]/shape[n].");
    }
  }
  stampDocument(document);
  await persistDocument(filePath, document);
  return { ok: true, targetPath };
}

export async function getDocumentNode(filePath: string, targetPath: string) {
  if (assertFormat(filePath) === "excel") {
    return getExcelNode(filePath, targetPath);
  }
  const document = await loadDocument(filePath);
  return materializePath(document, targetPath);
}

export async function queryDocumentNodes(filePath: string, selector: string) {
  if (assertFormat(filePath) === "excel") {
    return queryExcelNodesFromAdapter(filePath, selector);
  }
  const document = await loadDocument(filePath);
  return [materializePath(document, selector)];
}

export async function viewDocument(filePath: string, mode: string) {
  if (assertFormat(filePath) === "excel") {
    return viewExcelDocument(filePath, mode);
  }
  const document = await loadDocument(filePath);
  if (mode === "html") {
    return {
      mode,
      output: renderDocumentHtml(document),
    };
  }

  if (mode === "outline") {
    return {
      mode,
      output: renderDocumentOutline(document),
    };
  }

  if (mode === "json") {
    return {
      mode,
      output: JSON.stringify(document, null, 2),
    };
  }

  throw new UsageError(`Unsupported view mode '${mode}'.`, "Use outline, html, or json.");
}

export async function checkDocument(filePath: string) {
  if (assertFormat(filePath) === "excel") {
    return summarizeExcelCheck(filePath);
  }
  const document = await loadDocument(filePath);
  return {
    ok: true,
    format: document.format,
    summary: renderDocumentOutline(document),
  };
}

export async function rawDocument(filePath: string, options: RawDocumentOptions = {}) {
  if (assertFormat(filePath) === "excel") {
    return rawExcelDocument(filePath, options.partPath ?? "/", {
      startRow: options.startRow,
      endRow: options.endRow,
      cols: options.cols,
    });
  }
  const document = await loadDocument(filePath);
  return JSON.stringify(document, null, 2);
}

export function renderDocumentHtml(document: OfficekitDocument): string {
  if (document.format === "word") {
    const body = document.word!.body
      .map((node) => (node.type === "paragraph" ? `<p>${escapeHtml(node.text)}</p>` : renderWordTableHtml(node)))
      .join("\n") || "<p><em>Empty document</em></p>";
    return `<article data-format="word">${body}</article>`;
  }

  if (document.format === "excel") {
    return renderExcelHtmlFromRoot(document.excel);
  }

  const slides = document.powerpoint!.slides.map((slide, index) => `<section class="slide"><h2>Slide ${index + 1}: ${escapeHtml(slide.title)}</h2>${slide.shapes.map((shape) => `<p>${escapeHtml(shape.text)}</p>`).join("")}</section>`);
  return `<main data-format="powerpoint">${slides.join("") || '<section class="slide"><em>Empty deck</em></section>'}</main>`;
}

export function renderDocumentOutline(document: OfficekitDocument): string {
  if (document.format === "word") {
    const lines: string[] = [];
    let paragraphIndex = 0;
    let tableIndex = 0;
    for (const node of document.word!.body) {
      if (node.type === "paragraph") {
        paragraphIndex += 1;
        lines.push(`Paragraph ${paragraphIndex}: ${node.text}`);
        continue;
      }

      tableIndex += 1;
      const rowCount = node.rows.length;
      const colCount = node.rows[0]?.cells.length ?? 0;
      lines.push(`Table ${tableIndex}: ${rowCount}x${colCount}`);
      for (const [rowIndex, row] of node.rows.entries()) {
        for (const [cellIndex, cell] of row.cells.entries()) {
          lines.push(`  R${rowIndex + 1}C${cellIndex + 1}: ${cell.text}`);
        }
      }
    }
    return lines.join("\n") || "Word document is empty.";
  }

  if (document.format === "excel") {
    const lines: string[] = [];
    for (const sheet of document.excel!.sheets) {
      lines.push(`Sheet ${sheet.name}`);
      const refs = Object.keys(sheet.cells).sort();
      for (const ref of refs) {
        const cell = sheet.cells[ref];
        lines.push(`  ${ref}: ${cell.value}${cell.formula ? ` (formula=${cell.formula})` : ""}`);
      }
    }
    return lines.join("\n") || "Workbook is empty.";
  }

  return document.powerpoint!.slides.map((slide, index) => {
    const shapeLines = slide.shapes.map((shape, shapeIndex) => `  Shape ${shapeIndex + 1}: ${shape.text}`).join("\n");
    return [`Slide ${index + 1}: ${slide.title}`, shapeLines].filter(Boolean).join("\n");
  }).join("\n") || "Presentation is empty.";
}

function renderExcelView(document: OfficekitDocument, mode: string) {
  switch (mode) {
    case "html":
      return renderDocumentHtml(document);
    case "outline":
      return renderDocumentOutline(document);
    case "json":
      return JSON.stringify(document.excel, null, 2);
    case "text":
      return renderExcelTextView(document);
    case "annotated":
      return renderExcelAnnotatedView(document);
    case "stats":
      return renderExcelStatsView(document);
    case "issues":
      return renderExcelIssuesView(document);
    default:
      throw new UsageError(`Unsupported Excel view mode '${mode}'.`, "Use text, annotated, outline, stats, issues, html, or json.");
  }
}

function renderExcelTextView(document: OfficekitDocument) {
  const lines: string[] = [];
  for (const sheet of document.excel!.sheets) {
    lines.push(`=== Sheet: ${sheet.name} ===`);
    const rows = groupCellsByRow(sheet);
    for (const [rowNumber, rowCells] of rows) {
      lines.push(`[/${sheet.name}/row[${rowNumber}]] ${rowCells.map(([, cell]) => cell.value).join("\t")}`);
    }
  }
  return lines.join("\n").trimEnd() || "(empty workbook)";
}

function renderExcelAnnotatedView(document: OfficekitDocument) {
  const lines: string[] = [];
  for (const sheet of document.excel!.sheets) {
    lines.push(`=== Sheet: ${sheet.name} ===`);
    for (const [ref, cell] of Object.entries(sheet.cells).sort(([left], [right]) => compareCellRefs(left, right))) {
      const annotation = cell.formula ? `=${cell.formula}` : cell.type ?? "number";
      const warnings = [
        cell.value === "" && !cell.formula ? "empty" : "",
        cell.formula && cell.value === "" ? "unevaluated-formula" : "",
      ].filter(Boolean);
      lines.push(`  ${ref}: [${cell.value}] <- ${annotation}${warnings.length > 0 ? ` (${warnings.join(", ")})` : ""}`);
    }
  }
  return lines.join("\n").trimEnd() || "(empty workbook)";
}

function renderExcelStatsView(document: OfficekitDocument) {
  let totalCells = 0;
  let emptyCells = 0;
  let formulaCells = 0;
  const typeCounts = new Map<string, number>();
  for (const sheet of document.excel!.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      totalCells += 1;
      if (cell.value === "") emptyCells += 1;
      if (cell.formula) formulaCells += 1;
      const type = cell.type ?? (cell.formula ? "formula" : "number");
      typeCounts.set(type, (typeCounts.get(type) ?? 0) + 1);
    }
  }
  return [
    `Sheets: ${document.excel!.sheets.length}`,
    `Total Cells: ${totalCells}`,
    `Empty Cells: ${emptyCells}`,
    `Formula Cells: ${formulaCells}`,
    "",
    "Data Type Distribution:",
    ...[...typeCounts.entries()].sort((left, right) => right[1] - left[1]).map(([type, count]) => `  ${type}: ${count}`),
  ].join("\n").trimEnd();
}

function renderExcelIssuesView(document: OfficekitDocument) {
  const issues: string[] = [];
  for (const sheet of document.excel!.sheets) {
    if (sheet.autoFilter && !/^[A-Z]+\d+:[A-Z]+\d+$/.test(sheet.autoFilter)) {
      issues.push(`${sheet.name}: invalid autoFilter range '${sheet.autoFilter}'`);
    }
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      if (cell.formula && cell.value === "") {
        issues.push(`${sheet.name}!${ref}: formula has no cached value`);
      }
      if (cell.type === "date" && !/^-?\d+(\.\d+)?$/.test(cell.value)) {
        issues.push(`${sheet.name}!${ref}: date cell is not stored as numeric serial`);
      }
    }
  }
  return issues.join("\n") || "No issues found.";
}

function queryExcelNodes(document: OfficekitDocument, selector: string) {
  const normalized = selector.trim();
  if (normalized.startsWith("/")) {
    return [materializePath(document, normalized)];
  }
  if (normalized === "sheet" || normalized === "sheets") {
    return document.excel!.sheets.map((sheet) => materializePath(document, `/${sheet.name}`));
  }
  if (normalized === "namedrange" || normalized === "namedranges") {
    return (document.excel?.namedRanges ?? []).map((_, index) => materializePath(document, `/namedrange[${index + 1}]`));
  }
  if (normalized === "row") {
    return document.excel!.sheets.flatMap((sheet) => [...groupCellsByRow(sheet).keys()].map((row) => materializePath(document, `/${sheet.name}/row[${row}]`)));
  }
  if (normalized === "column" || normalized === "col") {
    return document.excel!.sheets.flatMap((sheet) =>
      [...new Set(Object.keys(sheet.cells).map((ref) => /^([A-Z]+)/.exec(ref)?.[1] ?? "A"))]
        .sort((left, right) => columnNameToIndex(left) - columnNameToIndex(right))
        .map((column) => materializePath(document, `/${sheet.name}/col[${column}]`)),
    );
  }
  if (normalized === "cell" || normalized === "cells") {
    return document.excel!.sheets.flatMap((sheet) =>
      Object.keys(sheet.cells)
        .sort(compareCellRefs)
        .map((ref) => materializePath(document, `/${sheet.name}/${ref}`)),
    );
  }
  if (normalized === "formula" || normalized === "cell[formula]") {
    return document.excel!.sheets.flatMap((sheet) =>
      Object.entries(sheet.cells)
        .filter(([, cell]) => Boolean(cell.formula))
        .sort(([left], [right]) => compareCellRefs(left, right))
        .map(([ref]) => materializePath(document, `/${sheet.name}/${ref}`)),
    );
  }
  return [];
}

function renderExcelRaw(zip: Map<string, Buffer>, options: RawDocumentOptions, document: OfficekitDocument) {
  const partPath = options.partPath ?? "/";
  if (partPath === "/" || partPath === "/workbook") {
    return requireEntry(zip, "xl/workbook.xml");
  }
  if (partPath === "/styles") {
    const styles = zip.get("xl/styles.xml");
    return styles ? styles.toString("utf8") : "(no styles)";
  }
  if (partPath === "/sharedstrings") {
    const sharedStrings = zip.get("xl/sharedStrings.xml");
    return sharedStrings ? sharedStrings.toString("utf8") : "(no shared strings)";
  }
  const drawingMatch = /^\/([^/]+)\/drawing$/i.exec(partPath);
  if (drawingMatch) {
    const sheetName = drawingMatch[1];
    const workbookRels = parseRelationships(requireEntry(zip, "xl/_rels/workbook.xml.rels"));
    const workbookXml = requireEntry(zip, "xl/workbook.xml");
    const relationshipId = [...workbookXml.matchAll(/<(?:\w+:)?sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)]
      .find((match) => decodeXml(match[1]).toLowerCase() === sheetName.toLowerCase())?.[2];
    if (!relationshipId) throw new OfficekitError(`Sheet '${sheetName}' not found.`, "not_found");
    const worksheetPath = normalizeZipPath("xl", workbookRels.get(relationshipId) ?? "");
    const worksheetRelsEntry = getRelationshipsEntryName(worksheetPath);
    const worksheetRels = zip.get(worksheetRelsEntry);
    if (!worksheetRels) throw new OfficekitError(`Sheet '${sheetName}' has no drawings.`, "not_found");
    const drawingTarget = parseRelationshipEntries(worksheetRels.toString("utf8")).find((entry) => entry.type?.endsWith("/drawing"))?.target;
    if (!drawingTarget) throw new OfficekitError(`Sheet '${sheetName}' has no drawings.`, "not_found");
    return requireEntry(zip, normalizeZipPath(path.posix.dirname(worksheetPath), drawingTarget));
  }
  const chartMatch = /^\/([^/]+)\/chart\[(\d+)\]$/i.exec(partPath);
  if (chartMatch) {
    return resolveChartXml(zip, chartMatch[1], Number(chartMatch[2]));
  }
  const globalChartMatch = /^\/chart\[(\d+)\]$/i.exec(partPath);
  if (globalChartMatch) {
    return resolveGlobalChartXml(zip, Number(globalChartMatch[1]));
  }
  const sheetMatch = /^\/([^/]+)$/i.exec(partPath);
  if (sheetMatch) {
    const sheet = ensureSheet(document, sheetMatch[1]);
    if (options.startRow !== undefined || options.endRow !== undefined || options.cols?.length) {
      return renderFilteredSheetXml(sheet, options);
    }
    const workbookXml = requireEntry(zip, "xl/workbook.xml");
    const workbookRels = parseRelationships(requireEntry(zip, "xl/_rels/workbook.xml.rels"));
    const relationshipId = [...workbookXml.matchAll(/<(?:\w+:)?sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)]
      .find((match) => decodeXml(match[1]).toLowerCase() === sheet.name.toLowerCase())?.[2];
    if (!relationshipId) throw new OfficekitError(`Sheet '${sheet.name}' not found.`, "not_found");
    return requireEntry(zip, normalizeZipPath("xl", workbookRels.get(relationshipId) ?? ""));
  }
  throw new UsageError(`Unsupported Excel raw part '${partPath}'.`, "Use /workbook, /styles, /sharedstrings, /Sheet1, /Sheet1/drawing, /Sheet1/chart[1], or /chart[1].");
}

function renderFilteredSheetXml(sheet: ExcelSheet, options: RawDocumentOptions) {
  const clone: ExcelSheet = {
    ...sheet,
    cells: Object.fromEntries(
      Object.entries(sheet.cells).filter(([ref]) => {
        const { column, row } = parseCellAddress(ref);
        if (options.startRow !== undefined && row < options.startRow) return false;
        if (options.endRow !== undefined && row > options.endRow) return false;
        if (options.cols?.length && !options.cols.includes(column)) return false;
        return true;
      }),
    ),
  };
  return renderSheetXml(clone);
}

function resolveChartXml(zip: Map<string, Buffer>, sheetName: string, index: number) {
  const workbookXml = requireEntry(zip, "xl/workbook.xml");
  const workbookRels = parseRelationships(requireEntry(zip, "xl/_rels/workbook.xml.rels"));
  const relationshipId = [...workbookXml.matchAll(/<(?:\w+:)?sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)]
    .find((match) => decodeXml(match[1]).toLowerCase() === sheetName.toLowerCase())?.[2];
  if (!relationshipId) throw new OfficekitError(`Sheet '${sheetName}' not found.`, "not_found");
  const worksheetPath = normalizeZipPath("xl", workbookRels.get(relationshipId) ?? "");
  const worksheetRels = parseRelationshipEntries(requireEntry(zip, getRelationshipsEntryName(worksheetPath)));
  const drawingTarget = worksheetRels.find((entry) => entry.type?.endsWith("/drawing"))?.target;
  if (!drawingTarget) throw new OfficekitError(`Sheet '${sheetName}' has no charts.`, "not_found");
  const drawingPath = normalizeZipPath(path.posix.dirname(worksheetPath), drawingTarget);
  const drawingXml = requireEntry(zip, drawingPath);
  const drawingRels = parseRelationshipEntries(requireEntry(zip, getRelationshipsEntryName(drawingPath)));
  const chartIds = [...drawingXml.matchAll(/<c:chart\b[^>]*r:id="([^"]+)"/g)].map((match) => match[1]);
  const chartId = chartIds[index - 1];
  if (!chartId) throw new OfficekitError(`Chart ${index} does not exist in sheet '${sheetName}'.`, "not_found");
  const chartTarget = drawingRels.find((entry) => entry.id === chartId)?.target;
  if (!chartTarget) throw new OfficekitError(`Chart ${index} relationship is missing.`, "invalid_ooxml");
  return requireEntry(zip, normalizeZipPath(path.posix.dirname(drawingPath), chartTarget));
}

function resolveGlobalChartXml(zip: Map<string, Buffer>, index: number) {
  const workbookXml = requireEntry(zip, "xl/workbook.xml");
  const sheetNames = [...workbookXml.matchAll(/<(?:\w+:)?sheet\b[^>]*name="([^"]+)"/g)].map((match) => decodeXml(match[1]));
  let seen = 0;
  for (const sheetName of sheetNames) {
    try {
      let sheetIndex = 1;
      while (true) {
        const chartXml = resolveChartXml(zip, sheetName, sheetIndex);
        seen += 1;
        if (seen === index) return chartXml;
        sheetIndex += 1;
      }
    } catch {
      continue;
    }
  }
  throw new OfficekitError(`Chart ${index} does not exist.`, "not_found");
}

function groupCellsByRow(sheet: ExcelSheet) {
  const rows = new Map<number, Array<[string, ExcelCell]>>();
  for (const [ref, cell] of Object.entries(sheet.cells).sort(([left], [right]) => compareCellRefs(left, right))) {
    const rowNumber = parseCellAddress(ref).row;
    const row = rows.get(rowNumber) ?? [];
    row.push([ref, cell]);
    rows.set(rowNumber, row);
  }
  return rows;
}

export function parseProps(argv: string[]) {
  const props: Record<string, string> = {};
  let type: string | undefined;
  let json = false;
  const rest: string[] = [];

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token === "--type") {
      type = argv[index + 1];
      index += 1;
      continue;
    }
    if (token === "--json") {
      json = true;
      continue;
    }
    if (token === "--prop") {
      const pair = argv[index + 1] ?? "";
      const [key, ...valueParts] = pair.split("=");
      props[key] = valueParts.join("=");
      index += 1;
      continue;
    }
    rest.push(token);
  }

  return { type, props, json, rest };
}

function createBlankDocument(format: SupportedFormat): OfficekitDocument {
  const base = {
    product: "officekit" as const,
    lineage: LINEAGE,
    format,
    version: 1 as const,
    updatedAt: new Date().toISOString(),
  };
  if (format === "word") return { ...base, word: { body: [] } };
  if (format === "excel") return { ...base, excel: { sheets: [{ name: "Sheet1", cells: {} as Record<string, ExcelCell> }] } };
  return { ...base, powerpoint: { slides: [] as PptSlide[] } };
}

function stampDocument(document: OfficekitDocument) {
  document.updatedAt = new Date().toISOString();
}

async function persistDocument(filePath: string, document: OfficekitDocument) {
  await mkdir(path.dirname(filePath), { recursive: true });
  const entries = buildDocumentEntries(document);
  await writeFile(filePath, createStoredZip(entries));
}

async function loadDocument(filePath: string): Promise<OfficekitDocument> {
  const zip = readStoredZip(await readFile(filePath));
  const metadata = zip.get(METADATA_PATH);
  if (!metadata) {
    return parseExternalDocument(zip, filePath);
  }
  return normalizeDocument(JSON.parse(metadata.toString("utf8")) as OfficekitDocument);
}

function buildDocumentEntries(document: OfficekitDocument) {
  const entries = [
    { name: METADATA_PATH, data: Buffer.from(JSON.stringify(document, null, 2), "utf8") },
  ];

  if (document.format === "word") {
    return [
      ...entries,
      { name: "[Content_Types].xml", data: Buffer.from(renderWordContentTypes(), "utf8") },
      { name: "_rels/.rels", data: Buffer.from(renderWordRels(), "utf8") },
      { name: "word/document.xml", data: Buffer.from(renderWordDocumentXml(document), "utf8") },
    ];
  }

  if (document.format === "excel") {
    return [
      ...entries,
      { name: "[Content_Types].xml", data: Buffer.from(renderExcelContentTypes(document), "utf8") },
      { name: "_rels/.rels", data: Buffer.from(renderExcelRels(), "utf8") },
      { name: "xl/workbook.xml", data: Buffer.from(renderWorkbookXml(document), "utf8") },
      { name: "xl/_rels/workbook.xml.rels", data: Buffer.from(renderWorkbookRels(document), "utf8") },
      ...(document.excel?.styleSheetXml
        ? [{ name: "xl/styles.xml", data: Buffer.from(document.excel.styleSheetXml, "utf8") }]
        : []),
      ...document.excel!.sheets.map((sheet, index) => ({ name: `xl/worksheets/sheet${index + 1}.xml`, data: Buffer.from(renderSheetXml(sheet), "utf8") })),
    ];
  }

  return [
    ...entries,
    { name: "[Content_Types].xml", data: Buffer.from(renderPptContentTypes(document), "utf8") },
    { name: "_rels/.rels", data: Buffer.from(renderPptRels(), "utf8") },
    { name: "ppt/presentation.xml", data: Buffer.from(renderPresentationXml(document), "utf8") },
    { name: "ppt/_rels/presentation.xml.rels", data: Buffer.from(renderPresentationRels(document), "utf8") },
    ...document.powerpoint!.slides.map((slide, index) => ({ name: `ppt/slides/slide${index + 1}.xml`, data: Buffer.from(renderSlideXml(slide), "utf8") })),
  ];
}

function materializePath(document: OfficekitDocument, targetPath: string) {
  if (targetPath === "/" || targetPath === "") {
    return document;
  }

  if (document.format === "word") {
    if (targetPath === "/body") {
      return {
        body: document.word!.body,
        paragraphs: getWordParagraphs(document),
        tables: getWordTables(document),
      };
    }
    const match = /^\/body\/p\[(\d+)\]$/.exec(targetPath);
    const tableMatch = /^\/body\/table\[(\d+)\]$/.exec(targetPath);
    const tableCellMatch = /^\/body\/table\[(\d+)\]\/cell\[(\d+),(\d+)\]$/.exec(targetPath);
    if (match) {
      const paragraph = resolveWordParagraph(document, Number(match[1]));
      if (!paragraph) throw new OfficekitError(`Paragraph ${match[1]} does not exist.`, "not_found");
      return paragraph;
    }
    if (tableMatch) {
      const table = resolveWordTable(document, Number(tableMatch[1]));
      if (!table) throw new OfficekitError(`Table ${tableMatch[1]} does not exist.`, "not_found");
      return table;
    }
    if (tableCellMatch) {
      const table = resolveWordTable(document, Number(tableCellMatch[1]));
      const row = table?.rows[Number(tableCellMatch[2]) - 1];
      const cell = row?.cells[Number(tableCellMatch[3]) - 1];
      if (!cell) {
        throw new OfficekitError(
          `Table cell ${tableCellMatch[2]},${tableCellMatch[3]} does not exist in table ${tableCellMatch[1]}.`,
          "not_found",
        );
      }
      return cell;
    }
  }

  if (document.format === "excel") {
    if (targetPath === "/" || targetPath === "/workbook") return document.excel;
    if (/^\/namedrange\[(.+)\]$/i.test(targetPath)) {
      return resolveNamedRange(document, targetPath);
    }
    const rowMatch = /^\/([^/]+)\/row\[(\d+)\]$/i.exec(targetPath);
    if (rowMatch) {
      const sheet = ensureSheet(document, rowMatch[1]);
      const rowNumber = Number(rowMatch[2]);
      return {
        path: targetPath,
        row: rowNumber,
        cells: getRowCells(sheet, rowNumber),
      };
    }
    const colMatch = /^\/([^/]+)\/col\[([A-Z]+)\]$/i.exec(targetPath);
    if (colMatch) {
      const sheet = ensureSheet(document, colMatch[1]);
      const column = colMatch[2].toUpperCase();
      return {
        path: targetPath,
        column,
        cells: getColumnCells(sheet, column),
        ...(sheet.tabColor ? { tabColor: sheet.tabColor } : {}),
      };
    }
    const rangeMatch = /^\/([^/]+)\/([A-Z]+\d+):([A-Z]+\d+)$/i.exec(targetPath);
    if (rangeMatch) {
      const sheet = ensureSheet(document, rangeMatch[1]);
      return {
        path: targetPath,
        sheet: sheet.name,
        cells: getRangeCells(sheet, rangeMatch[2].toUpperCase(), rangeMatch[3].toUpperCase()),
      };
    }
    const autoFilterMatch = /^\/([^/]+)\/autofilter$/i.exec(targetPath);
    if (autoFilterMatch) {
      const sheet = ensureSheet(document, autoFilterMatch[1]);
      return {
        path: targetPath,
        ref: sheet.autoFilter ?? null,
      };
    }
    const { sheet, cellRef } = resolveExcelPath(document, targetPath);
    if (!cellRef) return sheet;
    const cell = sheet.cells[cellRef];
    return cell ? { ref: cellRef, ...cell } : { ref: cellRef, value: null };
  }

  if (document.format === "powerpoint") {
    const slideMatch = /^\/slide\[(\d+)\]$/.exec(targetPath);
    if (slideMatch) {
      const slide = document.powerpoint!.slides[Number(slideMatch[1]) - 1];
      if (!slide) throw new OfficekitError(`Slide ${slideMatch[1]} does not exist.`, "not_found");
      return slide;
    }
    const shapeMatch = /^\/slide\[(\d+)\]\/shape\[(\d+)\]$/.exec(targetPath);
    if (shapeMatch) {
      const slide = document.powerpoint!.slides[Number(shapeMatch[1]) - 1];
      const shape = slide?.shapes[Number(shapeMatch[2]) - 1];
      if (!shape) throw new OfficekitError(`Shape ${shapeMatch[2]} does not exist.`, "not_found");
      return shape;
    }
  }

  throw new OfficekitError(`Unsupported path '${targetPath}' for ${document.format}.`, "unsupported_path");
}

function ensureSheet(document: OfficekitDocument, name: string) {
  const existing = document.excel!.sheets.find((sheet) => sheet.name === name);
  if (existing) return existing;
  const sheet: ExcelSheet = { name, cells: {} };
  document.excel!.sheets.push(sheet);
  return sheet;
}

function resolveNamedRange(document: OfficekitDocument, targetPath: string) {
  const ranges = document.excel?.namedRanges ?? [];
  const index = resolveNamedRangeIndex(ranges, targetPath);
  const range = ranges[index];
  if (!range) {
    throw new OfficekitError(`Named range '${targetPath}' does not exist.`, "not_found");
  }
  return range;
}

function resolveNamedRangeIndex(ranges: ExcelNamedRange[], targetPath: string) {
  const selector = /^\/namedrange\[(.+)\]$/i.exec(targetPath)?.[1] ?? "";
  if (/^\d+$/.test(selector)) {
    const index = Number(selector) - 1;
    if (index < 0 || index >= ranges.length) {
      throw new OfficekitError(`Named range index ${selector} is out of range.`, "not_found");
    }
    return index;
  }
  const index = ranges.findIndex((range) => range.name.toLowerCase() === selector.toLowerCase());
  if (index === -1) {
    throw new OfficekitError(`Named range '${selector}' not found.`, "not_found");
  }
  return index;
}

function nextAvailableRowIndex(sheet: ExcelSheet) {
  const refs = Object.keys(sheet.cells);
  if (refs.length === 0) return 1;
  return (
    Math.max(
      ...refs.map((ref) => Number(/\d+/.exec(ref)?.[0] ?? "0")),
    ) + 1
  );
}

function resolveExcelPath(document: OfficekitDocument, targetPath: string) {
  const cellMatch = /^\/([^/]+)\/([A-Z]+\d+)$/i.exec(targetPath);
  if (cellMatch) {
    return {
      sheet: ensureSheet(document, cellMatch[1]),
      cellRef: cellMatch[2].toUpperCase(),
    };
  }
  const sheetName = targetPath.replace(/^\//, "") || "Sheet1";
  return { sheet: ensureSheet(document, sheetName), cellRef: "" };
}

function getRowCells(sheet: ExcelSheet, rowNumber: number) {
  return Object.entries(sheet.cells)
    .filter(([ref]) => Number(/\d+/.exec(ref)?.[0] ?? "0") === rowNumber)
    .sort(([left], [right]) => compareCellRefs(left, right))
    .map(([ref, cell]) => ({ ref, ...cell }));
}

function getColumnCells(sheet: ExcelSheet, column: string) {
  return Object.entries(sheet.cells)
    .filter(([ref]) => /^([A-Z]+)/.exec(ref)?.[1] === column)
    .sort(([left], [right]) => compareCellRefs(left, right))
    .map(([ref, cell]) => ({ ref, ...cell }));
}

function getRangeCells(sheet: ExcelSheet, startRef: string, endRef: string) {
  const start = parseCellAddress(startRef);
  const end = parseCellAddress(endRef);
  const startColumn = columnNameToIndex(start.column);
  const endColumn = columnNameToIndex(end.column);
  const cells: Array<{ ref: string } & ExcelCell> = [];
  for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row += 1) {
    for (let column = Math.min(startColumn, endColumn); column <= Math.max(startColumn, endColumn); column += 1) {
      const ref = `${indexToColumnName(column)}${row}`;
      const cell = sheet.cells[ref] ?? { value: "" };
      cells.push({ ref, ...cell });
    }
  }
  return cells;
}

function enumerateRangeRefs(startRef: string, endRef: string) {
  return getRangeCells({ name: "", cells: {} }, startRef, endRef).map((cell) => cell.ref);
}

function compareCellRefs(left: string, right: string) {
  const leftCell = parseCellAddress(left);
  const rightCell = parseCellAddress(right);
  const leftColumn = columnNameToIndex(leftCell.column);
  const rightColumn = columnNameToIndex(rightCell.column);
  if (leftCell.row !== rightCell.row) {
    return leftCell.row - rightCell.row;
  }
  return leftColumn - rightColumn;
}

function applySheetProperties(document: OfficekitDocument, sheet: ExcelSheet, props: Record<string, string>) {
  const currentName = sheet.name;
  for (const [key, value] of Object.entries(props)) {
    const normalized = key.toLowerCase();
    switch (normalized) {
      case "name":
        if (document.excel!.sheets.some((candidate) => candidate !== sheet && candidate.name.toLowerCase() === value.toLowerCase())) {
          throw new OfficekitError(`Sheet '${value}' already exists.`, "duplicate_sheet");
        }
        sheet.name = value;
        break;
      case "freeze":
        sheet.freezeTopLeftCell = value;
        break;
      case "zoom":
        sheet.zoom = Number(value);
        break;
      case "gridlines":
        sheet.showGridLines = isTruthy(value);
        break;
      case "headings":
        sheet.showHeadings = isTruthy(value);
        break;
      case "tabcolor":
        sheet.tabColor = normalizeColorValue(value);
        break;
      case "header":
        sheet.header = value;
        break;
      case "footer":
        sheet.footer = value;
        break;
      case "orientation":
        sheet.orientation = value.toLowerCase();
        break;
      case "papersize":
        sheet.paperSize = Number(value);
        break;
      case "fittopage":
        sheet.fitToPage = value;
        break;
      case "protect":
        sheet.protection = isTruthy(value);
        break;
      case "autofilter":
        sheet.autoFilter = value;
        break;
      case "rowbreaks":
        sheet.rowBreaks = parseBreakList(value);
        break;
      case "colbreaks":
        sheet.colBreaks = parseBreakList(value);
        break;
      default:
        throw new UsageError(
          `Unsupported Excel sheet property '${key}'.`,
          "Supported: name, freeze, zoom, gridlines, headings, tabColor, header, footer, orientation, paperSize, fitToPage, protect, autoFilter, rowBreaks, colBreaks.",
        );
    }
  }
  if (sheet.name !== currentName) {
    for (const range of document.excel?.namedRanges ?? []) {
      if (range.scope?.toLowerCase() === currentName.toLowerCase()) {
        range.scope = sheet.name;
      }
      if (range.ref.startsWith(`${currentName}!`)) {
        range.ref = `${sheet.name}!${range.ref.slice(currentName.length + 1)}`;
      }
    }
  }
}

function parseBreakList(value: string) {
  return value
    .split(",")
    .map((item) => Number(item.trim()))
    .filter((item) => Number.isFinite(item) && item > 0);
}

function resolveSlide(document: OfficekitDocument, targetPath: string) {
  const slideMatch = /^\/slide\[(\d+)\]$/.exec(targetPath);
  if (!slideMatch) {
    throw new UsageError("PowerPoint paths currently support / and /slide[n].", "Use / for slide creation or /slide[1] for shape insertion.");
  }
  const slide = document.powerpoint!.slides[Number(slideMatch[1]) - 1];
  if (!slide) throw new OfficekitError(`Slide ${slideMatch[1]} does not exist.`, "not_found");
  return slide;
}

function createWordParagraph(text: string): WordParagraphNode {
  return {
    type: "paragraph",
    text,
  };
}

function createWordTable(rows: number, cols: number): WordTableNode {
  return {
    type: "table",
    rows: Array.from({ length: rows }, () => ({
      cells: Array.from({ length: cols }, () => ({ text: "" })),
    })),
  };
}

function normalizeWordState(word: NonNullable<OfficekitDocument["word"]>) {
  if (word.body?.length) {
    return {
      body: word.body.map((node) => normalizeWordBodyNode(node)),
    };
  }

  return {
    body: [
      ...(word.paragraphs ?? []).map((paragraph) => createWordParagraph(paragraph.text ?? "")),
      ...(word.tables ?? []).map((table) => normalizeWordTableNode(table)),
    ],
  };
}

function normalizeWordBodyNode(node: WordBodyNode | WordParagraph | WordTable) {
  if ("type" in node && node.type === "table") {
    return normalizeWordTableNode(node);
  }
  if ("type" in node && node.type === "paragraph") {
    return createWordParagraph(node.text ?? "");
  }
  if ("rows" in node) {
    return normalizeWordTableNode(node);
  }
  return createWordParagraph(node.text ?? "");
}

function normalizeWordTableNode(table: WordTable): WordTableNode {
  return {
    type: "table",
    rows: (table.rows ?? []).map((row) => ({
      cells: (row.cells ?? []).map((cell) => ({ text: cell.text ?? "" })),
    })),
  };
}

function getWordParagraphs(document: OfficekitDocument): WordParagraphNode[] {
  return document.word!.body.filter((node): node is WordParagraphNode => node.type === "paragraph");
}

function getWordTables(document: OfficekitDocument): WordTableNode[] {
  return document.word!.body.filter((node): node is WordTableNode => node.type === "table");
}

function resolveWordParagraph(document: OfficekitDocument, index: number) {
  return getWordParagraphs(document)[index - 1];
}

function resolveWordTable(document: OfficekitDocument, index: number) {
  return getWordTables(document)[index - 1];
}

function removeWordBodyNode(document: OfficekitDocument, type: WordBodyNode["type"], index: number) {
  let seen = 0;
  const bodyIndex = document.word!.body.findIndex((node) => {
    if (node.type !== type) return false;
    seen += 1;
    return seen === index;
  });
  if (bodyIndex === -1) {
    const label = type === "paragraph" ? "Paragraph" : "Table";
    throw new OfficekitError(`${label} ${index} does not exist.`, "not_found");
  }
  document.word!.body.splice(bodyIndex, 1);
}

function renderWordContentTypes() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
}

function renderWordRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

function renderWordDocumentXml(document: OfficekitDocument) {
  const body = document.word!.body
    .map((node) => (
      node.type === "paragraph"
        ? `<w:p><w:r><w:t xml:space="preserve">${escapeXml(node.text)}</w:t></w:r></w:p>`
        : renderWordTableXml(node)
    ))
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${body}
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function renderExcelContentTypes(document: OfficekitDocument) {
  const sheetOverrides = document.excel!.sheets
    .map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
    .join("\n  ");
  const stylesOverride = document.excel?.styleSheetXml
    ? '\n  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetOverrides}
  ${stylesOverride}
</Types>`;
}

function renderExcelRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

function renderWorkbookXml(document: OfficekitDocument) {
  const workbookPr = renderWorkbookProperties(document.excel?.settings);
  const calcPr = renderCalculationProperties(document.excel?.settings);
  const workbookProtection = renderWorkbookProtection(document.excel?.settings);
  const definedNames = renderDefinedNames(document.excel?.namedRanges, document.excel?.sheets ?? []);
  const sheets = document.excel!.sheets
    .map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${workbookPr}
  ${workbookProtection}
  <sheets>${sheets}</sheets>
  ${definedNames}
  ${calcPr}
</workbook>`;
}

function renderWorkbookRels(document: OfficekitDocument) {
  const rels = [
    ...document.excel!.sheets.map(
      (_, index) =>
        `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
    ),
    ...(document.excel?.styleSheetXml
      ? [`<Relationship Id="rId${document.excel!.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`]
      : []),
  ].join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

function renderWorkbookProperties(settings?: ExcelWorkbookSettings) {
  if (!settings) {
    return "";
  }
  const attrs = [
    settings.date1904 !== undefined ? `date1904="${settings.date1904 ? 1 : 0}"` : "",
    settings.codeName ? `codeName="${escapeXml(settings.codeName)}"` : "",
    settings.filterPrivacy !== undefined ? `filterPrivacy="${settings.filterPrivacy ? 1 : 0}"` : "",
    settings.showObjects ? `showObjects="${escapeXml(settings.showObjects)}"` : "",
    settings.backupFile !== undefined ? `backupFile="${settings.backupFile ? 1 : 0}"` : "",
    settings.dateCompatibility !== undefined ? `dateCompatibility="${settings.dateCompatibility ? 1 : 0}"` : "",
  ].filter(Boolean);

  return attrs.length > 0 ? `<workbookPr ${attrs.join(" ")}/>` : "";
}

function renderCalculationProperties(settings?: ExcelWorkbookSettings) {
  if (!settings) {
    return "";
  }
  const attrs = [
    settings.calcMode ? `calcMode="${escapeXml(settings.calcMode)}"` : "",
    settings.iterate !== undefined ? `iterate="${settings.iterate ? 1 : 0}"` : "",
    settings.iterateCount !== undefined ? `iterateCount="${settings.iterateCount}"` : "",
    settings.iterateDelta !== undefined ? `iterateDelta="${settings.iterateDelta}"` : "",
    settings.fullPrecision !== undefined ? `fullPrecision="${settings.fullPrecision ? 1 : 0}"` : "",
    settings.fullCalcOnLoad !== undefined ? `fullCalcOnLoad="${settings.fullCalcOnLoad ? 1 : 0}"` : "",
    settings.refMode ? `refMode="${escapeXml(settings.refMode)}"` : "",
  ].filter(Boolean);

  return attrs.length > 0 ? `<calcPr ${attrs.join(" ")}/>` : "";
}

function renderWorkbookProtection(settings?: ExcelWorkbookSettings) {
  if (!settings) {
    return "";
  }
  const attrs = [
    settings.lockStructure !== undefined ? `lockStructure="${settings.lockStructure ? 1 : 0}"` : "",
    settings.lockWindows !== undefined ? `lockWindows="${settings.lockWindows ? 1 : 0}"` : "",
  ].filter(Boolean);

  return attrs.length > 0 ? `<workbookProtection ${attrs.join(" ")}/>` : "";
}

function renderDefinedNames(namedRanges: ExcelNamedRange[] | undefined, sheets: ExcelSheet[]) {
  if (!namedRanges || namedRanges.length === 0) {
    return "";
  }
  const items = namedRanges
    .map((range) => {
      const scopeIndex = range.scope
        ? sheets.findIndex((sheet) => sheet.name.toLowerCase() === range.scope!.toLowerCase())
        : -1;
      const attrs = [
        `name="${escapeXml(range.name)}"`,
        ...(scopeIndex >= 0 ? [`localSheetId="${scopeIndex}"`] : []),
        ...(range.comment ? [`comment="${escapeXml(range.comment)}"`] : []),
      ];
      return `<definedName ${attrs.join(" ")}>${escapeXml(range.ref)}</definedName>`;
    })
    .join("");
  return `<definedNames>${items}</definedNames>`;
}

function renderSheetXml(sheet: ExcelSheet) {
  const entries = Object.entries(sheet.cells).sort(([a], [b]) => a.localeCompare(b));
  const rows = new Map<number, string[]>();
  for (const [ref, cell] of entries) {
    const row = Number(/\d+/.exec(ref)?.[0] ?? "1");
    const cells = rows.get(row) ?? [];
    cells.push(renderExcelCellXml(ref, cell));
    rows.set(row, cells);
  }
  const xmlRows = [...rows.entries()].sort(([a], [b]) => a - b).map(([rowIndex, cells]) => `<row r="${rowIndex}">${cells.join("")}</row>`).join("");
  const sheetViewAttrs = [
    sheet.zoom !== undefined ? ` zoomScale="${sheet.zoom}"` : "",
    sheet.showGridLines === false ? ` showGridLines="0"` : "",
    sheet.showHeadings === false ? ` showRowColHeaders="0"` : "",
  ].join("");
  const pane = sheet.freezeTopLeftCell
    ? `<pane ySplit="1" topLeftCell="${escapeXml(sheet.freezeTopLeftCell)}" state="frozen" activePane="bottomLeft"/>`
    : "";
  const sheetViews = pane || sheetViewAttrs
    ? `<sheetViews><sheetView workbookViewId="0"${sheetViewAttrs}>${pane}</sheetView></sheetViews>`
    : "";
  const sheetPr = sheet.tabColor
    ? `<sheetPr><tabColor rgb="${escapeXml(normalizeColorValue(sheet.tabColor))}"/></sheetPr>`
    : "";
  const autoFilter = sheet.autoFilter ? `<autoFilter ref="${escapeXml(sheet.autoFilter)}"/>` : "";
  const pageSetupAttrs = [
    sheet.orientation ? ` orientation="${escapeXml(sheet.orientation)}"` : "",
    sheet.paperSize !== undefined ? ` paperSize="${sheet.paperSize}"` : "",
    ...(sheet.fitToPage
      ? (() => {
          const [width, height] = sheet.fitToPage!.split("x");
          return [
            ` fitToWidth="${Number(width ?? "1")}"`,
            ` fitToHeight="${Number(height ?? "1")}"`,
          ];
        })()
      : []),
  ].join("");
  const pageSetup = pageSetupAttrs ? `<pageSetup${pageSetupAttrs}/>` : "";
  const headerFooter =
    sheet.header || sheet.footer
      ? `<headerFooter>${sheet.header ? `<oddHeader>${escapeXml(sheet.header)}</oddHeader>` : ""}${sheet.footer ? `<oddFooter>${escapeXml(sheet.footer)}</oddFooter>` : ""}</headerFooter>`
      : "";
  const protection = sheet.protection ? `<sheetProtection sheet="1"/>` : "";
  const rowBreaks = sheet.rowBreaks?.length
    ? `<rowBreaks count="${sheet.rowBreaks.length}" manualBreakCount="${sheet.rowBreaks.length}">${sheet.rowBreaks.map((row) => `<brk id="${row}" man="1"/>`).join("")}</rowBreaks>`
    : "";
  const colBreaks = sheet.colBreaks?.length
    ? `<colBreaks count="${sheet.colBreaks.length}" manualBreakCount="${sheet.colBreaks.length}">${sheet.colBreaks.map((column) => `<brk id="${column}" man="1"/>`).join("")}</colBreaks>`
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheetPr}
  ${sheetViews}
  <sheetData>${xmlRows}</sheetData>
  ${autoFilter}
  ${protection}
  ${pageSetup}
  ${headerFooter}
  ${rowBreaks}
  ${colBreaks}
</worksheet>`;
}

function renderExcelCellXml(ref: string, cell: ExcelCell) {
  const styleAttr = cell.styleId ? ` s="${escapeXml(cell.styleId)}"` : "";
  if (cell.formula) {
    const valueXml = cell.value !== "" ? `<v>${escapeXml(cell.value)}</v>` : "";
    return `<c r="${ref}"${styleAttr}><f>${escapeXml(normalizeFormula(cell.formula))}</f>${valueXml}</c>`;
  }
  if (cell.type === "boolean") {
    return `<c r="${ref}"${styleAttr} t="b"><v>${escapeXml(cell.value)}</v></c>`;
  }
  if (cell.type === "number" || cell.type === "date") {
    return `<c r="${ref}"${styleAttr}><v>${escapeXml(cell.value)}</v></c>`;
  }
  return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t>${escapeXml(cell.value)}</t></is></c>`;
}

function renderPptContentTypes(document: OfficekitDocument) {
  const slides = document.powerpoint!.slides
    .map((_, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`)
    .join("\n  ");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${slides}
</Types>`;
}

function renderPptRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
}

function renderPresentationXml(document: OfficekitDocument) {
  const slideIds = document.powerpoint!.slides
    .map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst>${slideIds}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
}

function renderPresentationRels(document: OfficekitDocument) {
  const rels = document.powerpoint!.slides
    .map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

function renderSlideXml(slide: PptSlide) {
  const titleShape = renderShapeXml(2, slide.title, 685800, 457200, 10972800, 914400);
  const contentShapes = slide.shapes.map((shape, index) => renderShapeXml(3 + index, shape.text, 914400, 1600200 + index * 914400, 10058400, 685800)).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${titleShape}
      ${contentShapes}
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

function renderShapeXml(id: number, text: string, x: number, y: number, cx: number, cy: number) {
  return `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr>
    <p:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p>
    </p:txBody>
  </p:sp>`;
}

function escapeHtml(value: string) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function escapeXml(value: string) {
  return escapeHtml(value).replaceAll('"', '&quot;').replaceAll("'", "&apos;");
}

function renderWordTableXml(table: WordTable) {
  const rows = table.rows
    .map(
      (row) => `<w:tr>${row.cells
        .map((cell) => `<w:tc><w:p><w:r><w:t xml:space="preserve">${escapeXml(cell.text)}</w:t></w:r></w:p></w:tc>`)
        .join("")}</w:tr>`,
    )
    .join("");
  return `<w:tbl>${rows}</w:tbl>`;
}

function renderWordTableHtml(table: WordTable) {
  const rows = table.rows
    .map((row) => `<tr>${row.cells.map((cell) => `<td>${escapeHtml(cell.text)}</td>`).join("")}</tr>`)
    .join("");
  return `<table>${rows}</table>`;
}

function parseExternalDocument(zip: Map<string, Buffer>, filePath: string): OfficekitDocument {
  const format = assertFormat(filePath);
  if (format === "word") {
    return parseWordDocument(zip);
  }
  if (format === "excel") {
    return parseExcelDocument(zip);
  }
  return parsePowerPointDocument(zip);
}

function normalizeDocument(document: OfficekitDocument): OfficekitDocument {
  if (document.word) {
    document.word = normalizeWordState(document.word);
  }
  if (document.excel) {
    document.excel = {
      sheets: (document.excel.sheets ?? []).map((sheet) => ({
        ...sheet,
        cells: Object.fromEntries(
          Object.entries(sheet.cells ?? {}).map(([ref, cell]) => [ref, normalizeExcelCell(cell)]),
        ),
        ...(sheet.autoFilter ? { autoFilter: sheet.autoFilter } : {}),
        ...(sheet.freezeTopLeftCell ? { freezeTopLeftCell: sheet.freezeTopLeftCell } : {}),
        ...(sheet.zoom !== undefined ? { zoom: sheet.zoom } : {}),
        ...(sheet.showGridLines !== undefined ? { showGridLines: sheet.showGridLines } : {}),
        ...(sheet.showHeadings !== undefined ? { showHeadings: sheet.showHeadings } : {}),
        ...(sheet.tabColor ? { tabColor: sheet.tabColor } : {}),
        ...(sheet.header ? { header: sheet.header } : {}),
        ...(sheet.footer ? { footer: sheet.footer } : {}),
        ...(sheet.orientation ? { orientation: sheet.orientation } : {}),
        ...(sheet.paperSize !== undefined ? { paperSize: sheet.paperSize } : {}),
        ...(sheet.fitToPage ? { fitToPage: sheet.fitToPage } : {}),
        ...(sheet.protection !== undefined ? { protection: sheet.protection } : {}),
        ...(sheet.rowBreaks?.length ? { rowBreaks: [...sheet.rowBreaks] } : {}),
        ...(sheet.colBreaks?.length ? { colBreaks: [...sheet.colBreaks] } : {}),
      })),
      ...(document.excel.settings ? { settings: document.excel.settings } : {}),
      ...(document.excel.styleSheetXml ? { styleSheetXml: document.excel.styleSheetXml } : {}),
      ...(document.excel.namedRanges ? { namedRanges: document.excel.namedRanges } : {}),
    };
  }
  return document;
}

function parseWordDocument(zip: Map<string, Buffer>): OfficekitDocument {
  const xml = requireEntry(zip, "word/document.xml");
  const body = /<w:body\b[^>]*>([\s\S]*?)<w:sectPr\b[^>]*\/?>/.exec(xml)?.[1] ?? "";
  const bodyNodes: WordBodyNode[] = [];
  for (const match of body.matchAll(/<w:(p|tbl)\b[\s\S]*?<\/w:\1>/g)) {
    if (match[1] === "p") {
      const text = extractTextRuns(match[0]);
      bodyNodes.push(createWordParagraph(text));
    } else {
      bodyNodes.push(parseWordTable(match[0]));
    }
  }
  return {
    product: "officekit",
    lineage: LINEAGE,
    format: "word",
    version: 1,
    updatedAt: new Date().toISOString(),
    word: {
      body: bodyNodes,
    },
  };
}

function parseWordTable(xml: string): WordTableNode {
  const rows = [...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((rowMatch) => ({
    cells: [...rowMatch[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((cellMatch) => ({
      text: extractTextRuns(cellMatch[0]),
    })),
  }));
  return { type: "table", rows };
}

function parseExcelDocument(zip: Map<string, Buffer>): OfficekitDocument {
  const workbookXml = requireEntry(zip, "xl/workbook.xml");
  const workbookRelsXml = requireEntry(zip, "xl/_rels/workbook.xml.rels");
  const relationshipMap = parseRelationships(workbookRelsXml);
  const workbookSettings = parseWorkbookSettings(workbookXml);
  const styleSheetXml = zip.get("xl/styles.xml")?.toString("utf8");
  const sheets = [...workbookXml.matchAll(/<(?:\w+:)?sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g)].map((match) => {
    const name = decodeXml(match[1]);
    const target = relationshipMap.get(match[2]);
    if (!target) {
      throw new OfficekitError(`Workbook relationship '${match[2]}' is missing.`, "invalid_ooxml");
    }
    const entryName = normalizeZipPath("xl", target);
    const sheetXml = requireEntry(zip, entryName);
    return {
      name,
      cells: parseSheetCells(sheetXml, zip),
      ...parseSheetFeatures(sheetXml),
    };
  });

  return {
    product: "officekit",
    lineage: LINEAGE,
    format: "excel",
    version: 1,
    updatedAt: new Date().toISOString(),
    excel: {
      sheets,
      ...(Object.keys(workbookSettings).length > 0 ? { settings: workbookSettings } : {}),
      ...(styleSheetXml ? { styleSheetXml } : {}),
      ...(parseDefinedNames(workbookXml, sheets).length > 0 ? { namedRanges: parseDefinedNames(workbookXml, sheets) } : {}),
    },
  };
}

function parsePowerPointDocument(zip: Map<string, Buffer>): OfficekitDocument {
  const presentationXml = requireEntry(zip, "ppt/presentation.xml");
  const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
  const relationshipMap = parseRelationships(relsXml);
  const slides = [...presentationXml.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"[^>]*\/?>/g)].map((match) => {
    const target = relationshipMap.get(match[1]);
    if (!target) {
      throw new OfficekitError(`Presentation relationship '${match[1]}' is missing.`, "invalid_ooxml");
    }
    const slideEntryName = normalizeZipPath("ppt", target);
    const slideXml = requireEntry(zip, slideEntryName);
    const { title, shapes } = parsePowerPointSlide(slideXml);
    const { layoutName, layoutType, themeName } = parseSlideContext(zip, slideEntryName);
    return {
      title,
      layoutName,
      layoutType,
      themeName,
      shapes,
    };
  });

  return {
    product: "officekit",
    lineage: LINEAGE,
    format: "powerpoint",
    version: 1,
    updatedAt: new Date().toISOString(),
    powerpoint: { slides },
  };
}

function parseRelationships(xml: string) {
  const relationships = new Map<string, string>();
  for (const relationship of parseRelationshipEntries(xml)) {
    relationships.set(relationship.id, relationship.target);
  }
  return relationships;
}

function parseRelationshipEntries(xml: string) {
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

function parsePowerPointSlide(xml: string) {
  const shapes = [...xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)]
    .map((match) => parsePowerPointShape(match[0]))
    .filter((shape): shape is PptShape => shape !== null);
  const titleIndex =
    shapes.findIndex((shape) => shape.kind === "title" || shape.kind === "ctrTitle") ??
    -1;
  const fallbackTitleIndex = titleIndex >= 0 ? titleIndex : 0;
  const title = shapes[fallbackTitleIndex]?.text ?? "Untitled slide";
  return {
    title,
    shapes: shapes.filter((_, index) => index !== fallbackTitleIndex),
  };
}

function parsePowerPointShape(xml: string): PptShape | null {
  const text = extractTextRuns(xml).trim();
  if (!text) {
    return null;
  }
  const name = /<p:cNvPr\b[^>]*name="([^"]*)"/.exec(xml)?.[1];
  const kind = /<p:ph\b[^>]*type="([^"]+)"/.exec(xml)?.[1];
  return {
    text,
    kind,
    name: name ? decodeXml(name) : undefined,
  };
}

function parseSlideContext(zip: Map<string, Buffer>, slideEntryName: string) {
  const slideRels = readRelationships(zip, getRelationshipsEntryName(slideEntryName));
  const layoutTarget = slideRels.find((relationship) => relationship.type?.endsWith("/slideLayout"))?.target;
  if (!layoutTarget) {
    return {};
  }

  const layoutEntryName = normalizeZipPath(path.posix.dirname(slideEntryName), layoutTarget);
  const layoutXml = requireEntry(zip, layoutEntryName);
  const layoutName = decodeXml(/<p:cSld\b[^>]*name="([^"]*)"/.exec(layoutXml)?.[1] ?? "");
  const layoutType = /<p:sldLayout\b[^>]*type="([^"]+)"/.exec(layoutXml)?.[1];
  const layoutRels = readRelationships(zip, getRelationshipsEntryName(layoutEntryName));
  const masterTarget = layoutRels.find((relationship) => relationship.type?.endsWith("/slideMaster"))?.target;
  const themeName = masterTarget ? parseThemeName(zip, layoutEntryName, masterTarget) : undefined;

  return {
    layoutName: layoutName || undefined,
    layoutType,
    themeName,
  };
}

function parseThemeName(zip: Map<string, Buffer>, layoutEntryName: string, masterTarget: string) {
  const masterEntryName = normalizeZipPath(path.posix.dirname(layoutEntryName), masterTarget);
  const masterRels = readRelationships(zip, getRelationshipsEntryName(masterEntryName));
  const themeTarget = masterRels.find((relationship) => relationship.type?.endsWith("/theme"))?.target;
  if (!themeTarget) {
    return undefined;
  }
  const themeXml = requireEntry(zip, normalizeZipPath(path.posix.dirname(masterEntryName), themeTarget));
  return decodeXml(/<a:theme\b[^>]*name="([^"]*)"/.exec(themeXml)?.[1] ?? "") || undefined;
}

function readRelationships(zip: Map<string, Buffer>, entryName: string) {
  const rels = zip.get(entryName);
  if (!rels) {
    return [];
  }
  return parseRelationshipEntries(rels.toString("utf8"));
}

function parseSheetCells(xml: string, zip: Map<string, Buffer>) {
  const sharedStrings = parseSharedStrings(zip);
  const cells: Record<string, ExcelCell> = {};
  for (const match of xml.matchAll(/<(?:\w+:)?c\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>/g)) {
    const attributes = match[1];
    const body = match[2];
    const refMatch = /r="([^"]+)"/.exec(attributes);
    if (!refMatch) continue;
    const ref = refMatch[1].toUpperCase();
    const styleId = /s="([^"]+)"/.exec(attributes)?.[1];
    const typeMatch = /t="([^"]+)"/.exec(attributes);
    const type = typeMatch?.[1] ?? "";
    const formula = (/<(?:\w+:)?f\b[^>]*>([\s\S]*?)<\/(?:\w+:)?f>/.exec(body)?.[1] ?? "").trim();
    let value = "";
    if (type === "inlineStr") {
      value = extractTexts(body).join("");
    } else if (type === "s") {
      const index = Number((/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/.exec(body)?.[1] ?? "0").trim());
      value = sharedStrings[index] ?? "";
    } else {
      value = decodeXml((/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/.exec(body)?.[1] ?? "").trim());
    }
    cells[ref] = {
      value,
      ...(styleId ? { styleId } : {}),
      ...(type === "b"
        ? { type: "boolean" as const }
        : type === "inlineStr" || type === "s" || type === "str"
          ? { type: "string" as const }
          : formula
            ? {}
            : { type: "number" as const }),
      ...(formula ? { formula: decodeXml(formula) } : {}),
    };
  }
  return cells;
}

function parseSheetFeatures(xml: string) {
  const autoFilter = /<(?:\w+:)?autoFilter\b[^>]*ref="([^"]+)"/.exec(xml)?.[1];
  const freezeTopLeftCell = /<(?:\w+:)?pane\b[^>]*topLeftCell="([^"]+)"/.exec(xml)?.[1];
  const zoom = /<(?:\w+:)?sheetView\b[^>]*zoomScale="([^"]+)"/.exec(xml)?.[1];
  const showGridLines = /<(?:\w+:)?sheetView\b[^>]*showGridLines="([^"]+)"/.exec(xml)?.[1];
  const showHeadings = /<(?:\w+:)?sheetView\b[^>]*showRowColHeaders="([^"]+)"/.exec(xml)?.[1];
  const tabColor = /<(?:\w+:)?tabColor\b[^>]*rgb="([^"]+)"/.exec(xml)?.[1];
  const header = /<(?:\w+:)?oddHeader>([\s\S]*?)<\/(?:\w+:)?oddHeader>/.exec(xml)?.[1];
  const footer = /<(?:\w+:)?oddFooter>([\s\S]*?)<\/(?:\w+:)?oddFooter>/.exec(xml)?.[1];
  const orientation = /<(?:\w+:)?pageSetup\b[^>]*orientation="([^"]+)"/.exec(xml)?.[1];
  const paperSize = /<(?:\w+:)?pageSetup\b[^>]*paperSize="([^"]+)"/.exec(xml)?.[1];
  const fitToWidth = /<(?:\w+:)?pageSetup\b[^>]*fitToWidth="([^"]+)"/.exec(xml)?.[1];
  const fitToHeight = /<(?:\w+:)?pageSetup\b[^>]*fitToHeight="([^"]+)"/.exec(xml)?.[1];
  const protection = /<(?:\w+:)?sheetProtection\b[^>]*sheet="([^"]+)"/.exec(xml)?.[1];
  const rowBreaks = [...xml.matchAll(/<(?:\w+:)?rowBreaks\b[\s\S]*?<brk\b[^>]*id="([^"]+)"/g)].map((match) => Number(match[1]));
  const colBreaks = [...xml.matchAll(/<(?:\w+:)?colBreaks\b[\s\S]*?<brk\b[^>]*id="([^"]+)"/g)].map((match) => Number(match[1]));
  return {
    ...(autoFilter ? { autoFilter } : {}),
    ...(freezeTopLeftCell ? { freezeTopLeftCell } : {}),
    ...(zoom ? { zoom: Number(zoom) } : {}),
    ...(showGridLines !== undefined ? { showGridLines: isTruthy(showGridLines) } : {}),
    ...(showHeadings !== undefined ? { showHeadings: isTruthy(showHeadings) } : {}),
    ...(tabColor ? { tabColor: decodeXml(tabColor) } : {}),
    ...(header ? { header: decodeXml(header) } : {}),
    ...(footer ? { footer: decodeXml(footer) } : {}),
    ...(orientation ? { orientation: decodeXml(orientation) } : {}),
    ...(paperSize ? { paperSize: Number(paperSize) } : {}),
    ...(fitToWidth || fitToHeight ? { fitToPage: `${fitToWidth ?? "1"}x${fitToHeight ?? "1"}` } : {}),
    ...(protection !== undefined ? { protection: isTruthy(protection) } : {}),
    ...(rowBreaks.length > 0 ? { rowBreaks } : {}),
    ...(colBreaks.length > 0 ? { colBreaks } : {}),
  };
}

function parseWorkbookSettings(xml: string): ExcelWorkbookSettings {
  const attrs = /<(?:\w+:)?workbookPr\b([^>]*)\/?>/.exec(xml)?.[1];
  const calcAttrs = /<(?:\w+:)?calcPr\b([^>]*)\/?>/.exec(xml)?.[1];
  const protectionAttrs = /<(?:\w+:)?workbookProtection\b([^>]*)\/?>/.exec(xml)?.[1];
  return {
    ...parseWorkbookPropertyAttributes(attrs),
    ...parseCalculationPropertyAttributes(calcAttrs),
    ...parseWorkbookProtectionAttributes(protectionAttrs),
  };
}

function parseDefinedNames(workbookXml: string, sheets: ExcelSheet[]) {
  return [...workbookXml.matchAll(/<(?:\w+:)?definedName\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?definedName>/g)].map((match) => {
    const attrs = match[1];
    const ref = decodeXml(match[2]);
    const name = decodeXml(/name="([^"]+)"/.exec(attrs)?.[1] ?? "");
    const localSheetId = /localSheetId="([^"]+)"/.exec(attrs)?.[1];
    const comment = /comment="([^"]+)"/.exec(attrs)?.[1];
    return {
      name,
      ref,
      ...(localSheetId !== undefined && sheets[Number(localSheetId)] ? { scope: sheets[Number(localSheetId)].name } : {}),
      ...(comment ? { comment: decodeXml(comment) } : {}),
    };
  });
}

function parseSharedStrings(zip: Map<string, Buffer>) {
  const shared = zip.get("xl/sharedStrings.xml");
  if (!shared) return [];
  return [...shared.toString("utf8").matchAll(/<(?:\w+:)?si\b[\s\S]*?<\/(?:\w+:)?si>/g)].map((match) => extractTexts(match[0]).join(""));
}

function extractTextRuns(xml: string) {
  return extractTexts(xml).join("");
}

function extractTexts(xml: string) {
  return [...xml.matchAll(/<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)].map((match) => decodeXml(match[1]));
}

function normalizeZipPath(baseDir: string, target: string) {
  const normalized = target.replace(/\\/g, "/");
  if (normalized.startsWith("/")) {
    return path.posix.normalize(normalized.slice(1));
  }
  return path.posix.normalize(path.posix.join(baseDir, normalized));
}

function getRelationshipsEntryName(entryName: string) {
  const directory = path.posix.dirname(entryName);
  const basename = path.posix.basename(entryName);
  return path.posix.join(directory, "_rels", `${basename}.rels`);
}

function requireEntry(zip: Map<string, Buffer>, entryName: string) {
  const buffer = zip.get(entryName);
  if (!buffer) {
    throw new OfficekitError(`OOXML entry '${entryName}' is missing.`, "invalid_ooxml");
  }
  return buffer.toString("utf8");
}

function decodeXml(value: string) {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function normalizeExcelCell(cell: string | ExcelCell | undefined): ExcelCell {
  if (typeof cell === "string") {
    return { value: cell, type: "string" };
  }
  return {
    value: cell?.value ?? "",
    ...(cell?.styleId ? { styleId: cell.styleId } : {}),
    ...(cell?.type ? { type: cell.type } : {}),
    ...(cell?.formula ? { formula: normalizeFormula(cell.formula) } : {}),
  };
}

function mergeExcelCell(existing: string | ExcelCell | undefined, props: Record<string, string>): ExcelCell {
  const base = normalizeExcelCell(existing);
  const formula = props.formula === undefined ? base.formula : normalizeFormula(props.formula);
  const styleId = props.styleId ?? props.style ?? base.styleId;
  const explicitType = props.type?.toLowerCase();
  const type =
    explicitType === "number" || explicitType === "boolean" || explicitType === "date" || explicitType === "string"
      ? (explicitType as ExcelCell["type"])
      : base.type;
  return {
    value: props.value ?? props.text ?? base.value,
    ...(styleId ? { styleId } : {}),
    ...(type ? { type } : {}),
    ...(formula ? { formula } : {}),
  };
}

function normalizeFormula(formula: string) {
  return formula.replace(/^=/, "");
}

function mergeWorkbookSettings(
  existing: ExcelWorkbookSettings | undefined,
  props: Record<string, string>,
): ExcelWorkbookSettings {
  const next: ExcelWorkbookSettings = { ...(existing ?? {}) };

  if (props.date1904 !== undefined) {
    next.date1904 = isTruthy(props.date1904);
  }
  if (props.codeName !== undefined || props.codename !== undefined) {
    next.codeName = props.codeName ?? props.codename;
  }
  if (props.filterPrivacy !== undefined || props.filterprivacy !== undefined) {
    next.filterPrivacy = isTruthy(props.filterPrivacy ?? props.filterprivacy ?? "false");
  }
  if (props.showObjects !== undefined || props.showobjects !== undefined) {
    next.showObjects = (props.showObjects ?? props.showobjects)?.toLowerCase();
  }
  if (props.backupFile !== undefined || props.backupfile !== undefined) {
    next.backupFile = isTruthy(props.backupFile ?? props.backupfile ?? "false");
  }
  if (props.dateCompatibility !== undefined || props.datecompatibility !== undefined) {
    next.dateCompatibility = isTruthy(props.dateCompatibility ?? props.datecompatibility ?? "false");
  }
  if (props["calc.mode"] !== undefined || props.calcmode !== undefined) {
    next.calcMode = normalizeCalcMode(props["calc.mode"] ?? props.calcmode ?? "");
  }
  if (props["calc.iterate"] !== undefined || props.iterate !== undefined) {
    next.iterate = isTruthy(props["calc.iterate"] ?? props.iterate ?? "false");
  }
  if (props["calc.iterateCount"] !== undefined || props.iteratecount !== undefined) {
    next.iterateCount = Number(props["calc.iterateCount"] ?? props.iteratecount);
  }
  if (props["calc.iterateDelta"] !== undefined || props.iteratedelta !== undefined) {
    next.iterateDelta = Number(props["calc.iterateDelta"] ?? props.iteratedelta);
  }
  if (props["calc.fullPrecision"] !== undefined || props.fullprecision !== undefined) {
    next.fullPrecision = isTruthy(props["calc.fullPrecision"] ?? props.fullprecision ?? "false");
  }
  if (props["calc.fullCalcOnLoad"] !== undefined || props.fullcalconload !== undefined) {
    next.fullCalcOnLoad = isTruthy(props["calc.fullCalcOnLoad"] ?? props.fullcalconload ?? "false");
  }
  if (props["calc.refMode"] !== undefined || props.refmode !== undefined) {
    next.refMode = normalizeRefMode(props["calc.refMode"] ?? props.refmode ?? "");
  }
  if (props["workbook.lockStructure"] !== undefined || props.lockstructure !== undefined) {
    next.lockStructure = isTruthy(props["workbook.lockStructure"] ?? props.lockstructure ?? "false");
  }
  if (props["workbook.lockWindows"] !== undefined || props.lockwindows !== undefined) {
    next.lockWindows = isTruthy(props["workbook.lockWindows"] ?? props.lockwindows ?? "false");
  }

  return next;
}

function isTruthy(value: string) {
  return /^(1|true|yes|on)$/i.test(value.trim());
}

function normalizeColorValue(value: string) {
  const cleaned = value.trim().replace(/^#/, "").toUpperCase();
  if (cleaned.length === 6) {
    return `FF${cleaned}`;
  }
  return cleaned;
}

function parseWorkbookPropertyAttributes(attrs?: string): ExcelWorkbookSettings {
  if (!attrs) return {};
  const date1904 = /date1904="([^"]+)"/.exec(attrs)?.[1];
  const codeName = /codeName="([^"]+)"/.exec(attrs)?.[1];
  const filterPrivacy = /filterPrivacy="([^"]+)"/.exec(attrs)?.[1];
  const showObjects = /showObjects="([^"]+)"/.exec(attrs)?.[1];
  const backupFile = /backupFile="([^"]+)"/.exec(attrs)?.[1];
  const dateCompatibility = /dateCompatibility="([^"]+)"/.exec(attrs)?.[1];
  return {
    ...(date1904 !== undefined ? { date1904: isTruthy(date1904) } : {}),
    ...(codeName ? { codeName: decodeXml(codeName) } : {}),
    ...(filterPrivacy !== undefined ? { filterPrivacy: isTruthy(filterPrivacy) } : {}),
    ...(showObjects ? { showObjects: decodeXml(showObjects) } : {}),
    ...(backupFile !== undefined ? { backupFile: isTruthy(backupFile) } : {}),
    ...(dateCompatibility !== undefined ? { dateCompatibility: isTruthy(dateCompatibility) } : {}),
  };
}

function parseCalculationPropertyAttributes(attrs?: string): ExcelWorkbookSettings {
  if (!attrs) return {};
  const calcMode = /calcMode="([^"]+)"/.exec(attrs)?.[1];
  const iterate = /iterate="([^"]+)"/.exec(attrs)?.[1];
  const iterateCount = /iterateCount="([^"]+)"/.exec(attrs)?.[1];
  const iterateDelta = /iterateDelta="([^"]+)"/.exec(attrs)?.[1];
  const fullPrecision = /fullPrecision="([^"]+)"/.exec(attrs)?.[1];
  const fullCalcOnLoad = /fullCalcOnLoad="([^"]+)"/.exec(attrs)?.[1];
  const refMode = /refMode="([^"]+)"/.exec(attrs)?.[1];
  return {
    ...(calcMode ? { calcMode: decodeXml(calcMode) } : {}),
    ...(iterate !== undefined ? { iterate: isTruthy(iterate) } : {}),
    ...(iterateCount !== undefined ? { iterateCount: Number(iterateCount) } : {}),
    ...(iterateDelta !== undefined ? { iterateDelta: Number(iterateDelta) } : {}),
    ...(fullPrecision !== undefined ? { fullPrecision: isTruthy(fullPrecision) } : {}),
    ...(fullCalcOnLoad !== undefined ? { fullCalcOnLoad: isTruthy(fullCalcOnLoad) } : {}),
    ...(refMode ? { refMode: decodeXml(refMode) } : {}),
  };
}

function parseWorkbookProtectionAttributes(attrs?: string): ExcelWorkbookSettings {
  if (!attrs) return {};
  const lockStructure = /lockStructure="([^"]+)"/.exec(attrs)?.[1];
  const lockWindows = /lockWindows="([^"]+)"/.exec(attrs)?.[1];
  return {
    ...(lockStructure !== undefined ? { lockStructure: isTruthy(lockStructure) } : {}),
    ...(lockWindows !== undefined ? { lockWindows: isTruthy(lockWindows) } : {}),
  };
}

function normalizeCalcMode(value: string) {
  const normalized = value.trim().toLowerCase();
  if (normalized === "automatic") return "auto";
  if (normalized === "autoexcepttables" || normalized === "autonoexcepttables" || normalized === "autonotable") {
    return "autoNoTable";
  }
  return normalized;
}

function normalizeRefMode(value: string) {
  const normalized = value.trim().toUpperCase();
  return normalized === "R1C1" ? "R1C1" : "A1";
}

function parseDelimitedRows(content: string, delimiter: string) {
  const rows: string[][] = [];
  if (!content) return rows;
  if (content.charCodeAt(0) === 0xfeff) {
    content = content.slice(1);
  }

  const currentRow: string[] = [];
  let field = "";
  let inQuotes = false;

  for (let index = 0; index < content.length; index += 1) {
    const char = content[index];
    if (inQuotes) {
      if (char === '"') {
        if (content[index + 1] === '"') {
          field += '"';
          index += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += char;
      }
      continue;
    }

    if (char === '"') {
      inQuotes = true;
      continue;
    }
    if (char === delimiter) {
      currentRow.push(field);
      field = "";
      continue;
    }
    if (char === "\n" || char === "\r") {
      if (char === "\r" && content[index + 1] === "\n") {
        index += 1;
      }
      currentRow.push(field);
      field = "";
      if (!(currentRow.length === 1 && currentRow[0] === "")) {
        rows.push([...currentRow]);
      }
      currentRow.length = 0;
      continue;
    }
    field += char;
  }

  if (field.length > 0 || currentRow.length > 0) {
    currentRow.push(field);
    if (!(currentRow.length === 1 && currentRow[0] === "")) {
      rows.push([...currentRow]);
    }
  }

  return rows;
}

function parseCellAddress(value: string) {
  const match = /^([A-Z]+)(\d+)$/.exec(value);
  if (!match) {
    throw new UsageError(`Invalid cell address '${value}'.`, "Use an address like A1.");
  }
  return { column: match[1], row: Number(match[2]) };
}

function columnNameToIndex(column: string) {
  let result = 0;
  for (const char of column) {
    result = result * 26 + (char.charCodeAt(0) - 64);
  }
  return result;
}

function indexToColumnName(index: number) {
  let value = index;
  let column = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    value = Math.floor((value - 1) / 26);
  }
  return column;
}

function inferImportedCell(rawValue: string): ExcelCell {
  if (rawValue === "") return { value: "" };
  if (rawValue.startsWith("=")) {
    return { value: "", formula: normalizeFormula(rawValue) };
  }
  if (/^(true|false)$/i.test(rawValue)) {
    return { value: rawValue.toUpperCase() === "TRUE" ? "1" : "0", type: "boolean" };
  }
  const isoDate = tryParseIsoDate(rawValue);
  if (isoDate) {
    return { value: isoDate, type: "date" };
  }
  if (!Number.isNaN(Number(rawValue))) {
    return { value: rawValue, type: "number" };
  }
  return { value: rawValue, type: "string" };
}

function tryParseIsoDate(value: string) {
  const formats = [
    /^(\d{4})-(\d{2})-(\d{2})$/,
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/,
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})Z$/,
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})\.(\d{3})$/,
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})\.(\d{3})Z$/,
    /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/,
  ];

  for (const format of formats) {
    const match = format.exec(value);
    if (!match) continue;
    const [, year, month, day, hour = "0", minute = "0", second = "0", millis = "0"] = match;
    const date = new Date(Date.UTC(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second),
      Number(millis),
    ));
    if (Number.isNaN(date.getTime())) continue;
    return toOADate(date).toString();
  }

  return null;
}

function toOADate(date: Date) {
  const oaEpoch = Date.UTC(1899, 11, 30, 0, 0, 0, 0);
  return (date.getTime() - oaEpoch) / 86400000;
}
