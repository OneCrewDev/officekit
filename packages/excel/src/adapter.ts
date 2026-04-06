import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { OfficekitError, UsageError } from "../../core/src/errors.js";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";

export interface ExcelCommandOptions {
  type?: string;
  props: Record<string, string>;
  json?: boolean;
}

export interface ExcelImportOptions {
  delimiter: string;
  hasHeader: boolean;
  startCell: string;
}

export interface ExcelCellModel {
  value: string;
  formula?: string;
  styleId?: string;
  type?: "string" | "number" | "boolean" | "date";
}

export interface ExcelWorkbookSettings {
  date1904?: boolean;
  codeName?: string;
  filterPrivacy?: boolean;
  showObjects?: string;
  backupFile?: boolean;
  dateCompatibility?: boolean;
  calcMode?: string;
  iterate?: boolean;
  iterateCount?: number;
  iterateDelta?: number;
  fullPrecision?: boolean;
  fullCalcOnLoad?: boolean;
  refMode?: string;
  lockStructure?: boolean;
  lockWindows?: boolean;
}

export interface ExcelNamedRangeModel {
  name: string;
  ref: string;
  scope?: string;
  comment?: string;
}

export interface ExcelValidationModel {
  type?: string;
  sqref?: string;
  formula1?: string;
  formula2?: string;
  operator?: string;
  allowBlank?: boolean;
  showError?: boolean;
  errorTitle?: string;
  error?: string;
  showInput?: boolean;
  promptTitle?: string;
  prompt?: string;
}

export interface ExcelCommentModel {
  ref: string;
  author?: string;
  text: string;
}

export interface ExcelTableModel {
  name?: string;
  displayName?: string;
  ref?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
  styleName?: string;
}

export interface ExcelChartModel {
  title?: string;
  path: string;
  sheet?: string;
  chartType?: string;
  legend?: string | boolean;
  dataLabels?: string;
  categoryAxisTitle?: string;
  valueAxisTitle?: string;
  seriesNames?: string[];
}

export interface ExcelPivotTableModel {
  name?: string;
  path: string;
}

export interface ExcelSparklineModel {
  location?: string;
  sourceRange?: string;
  type?: string;
}

export interface ExcelShapeModel {
  name?: string;
  text?: string;
  kind: "shape" | "picture";
}

export interface ExcelSheetModel {
  name: string;
  xml: string;
  entryName: string;
  relId: string;
  relationshipTarget: string;
  cells: Record<string, ExcelCellModel>;
  autoFilter?: string;
  freezeTopLeftCell?: string;
  zoom?: number;
  showGridLines?: boolean;
  showHeadings?: boolean;
  tabColor?: string;
  header?: string;
  footer?: string;
  orientation?: string;
  paperSize?: number;
  fitToPage?: string;
  protection?: boolean;
  rowBreaks?: number[];
  colBreaks?: number[];
}

interface RelationshipEntry {
  id: string;
  target: string;
  type?: string;
}

interface ExcelWorkbookState {
  zip: Map<string, Buffer>;
  workbookXml: string;
  workbookEntryName: string;
  workbookRelsXml: string;
  workbookRelsEntryName: string;
  sheets: ExcelSheetModel[];
  settings: ExcelWorkbookSettings;
  namedRanges: ExcelNamedRangeModel[];
  styleSheetXml?: string;
  metadata: Record<string, string>;
  officekitMetadata?: {
    excel?: {
      sheets?: Array<{
        name: string;
        cells?: Record<string, ExcelCellModel>;
        autoFilter?: string;
        freezeTopLeftCell?: string;
        zoom?: number;
        showGridLines?: boolean;
        showHeadings?: boolean;
        tabColor?: string;
        header?: string;
        footer?: string;
        orientation?: string;
        paperSize?: number;
        fitToPage?: string;
        protection?: boolean;
        rowBreaks?: number[];
        colBreaks?: number[];
      }>;
      namedRanges?: ExcelNamedRangeModel[];
      settings?: ExcelWorkbookSettings;
    };
  };
}

const METADATA_PATH = "officekit/document.json";
const LINEAGE = "officekit is migrated from OfficeCLI and currently persists metadata-backed OOXML vertical slices.";

export async function createExcelDocument(filePath: string) {
  const state = createBlankWorkbookState();
  await writeWorkbookState(filePath, state);
  return {
    format: "excel" as const,
    filePath,
    document: materializeWorkbookRoot(state),
  };
}

export async function addExcelNode(filePath: string, targetPath: string, options: ExcelCommandOptions) {
  const state = await loadWorkbookState(filePath);
  if (options.type === "namedrange" || options.type === "definedname") {
    if (targetPath !== "/" && targetPath !== "/workbook") {
      throw new UsageError(
        "Excel add namedrange currently supports only the workbook root.",
        "Use: officekit add book.xlsx / --type namedrange --prop name=Range1 --prop ref=Sheet1!A1",
      );
    }
    const name = options.props.name ?? "";
    const ref = options.props.ref ?? "";
    if (!name || !ref) {
      throw new UsageError("Excel namedrange requires --prop name and --prop ref.");
    }
    if (state.namedRanges.some((range) => range.name.toLowerCase() === name.toLowerCase())) {
      throw new OfficekitError(`Named range '${name}' already exists.`, "duplicate_named_range");
    }
    state.namedRanges.push({
      name,
      ref,
      ...(options.props.scope ? { scope: options.props.scope } : {}),
      ...(options.props.comment ? { comment: options.props.comment } : {}),
    });
    updateWorkbookXml(state);
    await writeWorkbookState(filePath, state);
    return resolveNamedRangeNode(state, `/namedrange[${name}]`);
  }

  if (options.type === "sheet") {
    if (targetPath !== "/" && targetPath !== "/workbook") {
      throw new UsageError(
        "Excel add sheet currently supports only the workbook root.",
        "Use: officekit add book.xlsx / --type sheet --prop name=Sheet2",
      );
    }
    const requestedName = options.props.name ?? `Sheet${state.sheets.length + 1}`;
    if (state.sheets.some((sheet) => sheet.name.toLowerCase() === requestedName.toLowerCase())) {
      throw new OfficekitError(`Sheet '${requestedName}' already exists.`, "duplicate_sheet");
    }
    addSheet(state, requestedName);
    await writeWorkbookState(filePath, state);
    return materializeSheetNode(state, `/${requestedName}`);
  }

  if (options.type === "row") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const rowIndex = Math.max(1, Number(options.props.index ?? nextAvailableRowIndex(sheet)));
    const cols = Math.max(1, Number(options.props.cols ?? "1"));
    for (let columnOffset = 0; columnOffset < cols; columnOffset += 1) {
      const ref = `${indexToColumnName(columnOffset + 1)}${rowIndex}`;
      if (!sheet.cells[ref]) {
        sheet.cells[ref] = { value: "" };
      }
    }
    updateSheetXml(sheet);
    await writeWorkbookState(filePath, state);
    return { path: `/${sheet.name}/row[${rowIndex}]`, cells: collectRowCells(sheet, rowIndex) };
  }

  if (options.type === "validation" || options.type === "datavalidation") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const validation = addValidation(sheet, options.props);
    await writeWorkbookState(filePath, state);
    const validations = parseValidations(sheet.xml);
    return { ...validation, ...(validation.type ? { validationType: validation.type } : {}), path: `/${sheet.name}/validation[${validations.length}]`, type: "validation" };
  }

  if (options.type === "comment" || options.type === "note") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const ref = (options.props.ref ?? options.props.cell ?? "").toUpperCase();
    if (!ref) {
      throw new UsageError("Excel comment requires --prop ref=A1 or --prop cell=A1.");
    }
    const comment = addComment(state, sheet, {
      ref,
      text: options.props.text ?? options.props.value ?? "",
      author: options.props.author ?? "officekit",
    });
    await writeWorkbookState(filePath, state);
    return { ...comment, path: `/${sheet.name}/comment[${getSheetComments(state, sheet).length}]`, type: "comment" };
  }

  if (options.type === "autofilter") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const range = (options.props.range ?? options.props.ref ?? "").toUpperCase();
    if (!range) {
      throw new UsageError("Excel add autofilter requires --prop range=SheetRange.", "Use: officekit add book.xlsx /Sheet1 --type autofilter --prop range=A1:B10");
    }
    sheet.autoFilter = range;
    updateSheetXml(sheet);
    await writeWorkbookState(filePath, state);
    return { path: `/${sheet.name}/autofilter`, type: "autofilter", range };
  }

  if (options.type === "rowbreak" || options.type === "colbreak" || options.type === "pagebreak") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const requestedKind = options.type === "pagebreak"
      ? ((options.props.col ?? options.props.column) !== undefined ? "colbreak" : "rowbreak")
      : options.type;
    const numericValue = requestedKind === "colbreak"
      ? Number(options.props.col ?? options.props.column ?? options.props.id ?? "")
      : Number(options.props.row ?? options.props.id ?? "");
    if (!Number.isFinite(numericValue) || numericValue < 1) {
      throw new UsageError(
        requestedKind === "colbreak" ? "Excel add colbreak requires --prop col=<number>." : "Excel add rowbreak requires --prop row=<number>.",
      );
    }
    if (requestedKind === "colbreak") {
      sheet.colBreaks = [...new Set([...(sheet.colBreaks ?? []), numericValue])].sort((a, b) => a - b);
      updateSheetXml(sheet);
      await writeWorkbookState(filePath, state);
      return { path: `/${sheet.name}/colbreak[${sheet.colBreaks.indexOf(numericValue) + 1}]`, type: "colbreak", id: numericValue, manual: true };
    }
    sheet.rowBreaks = [...new Set([...(sheet.rowBreaks ?? []), numericValue])].sort((a, b) => a - b);
    updateSheetXml(sheet);
    await writeWorkbookState(filePath, state);
    return { path: `/${sheet.name}/rowbreak[${sheet.rowBreaks.indexOf(numericValue) + 1}]`, type: "rowbreak", id: numericValue, manual: true };
  }

  if (options.type === "table" || options.type === "listobject") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const table = addTable(state, sheet, options.props);
    await writeWorkbookState(filePath, state);
    return { ...table, path: `/${sheet.name}/table[${getSheetTables(state, sheet).length}]`, type: "table" };
  }

  if (options.type === "sparkline") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const location = (options.props.location ?? options.props.cell ?? options.props.ref ?? "").toUpperCase();
    if (!location) {
      throw new UsageError("Excel sparkline requires --prop cell=C2 or --prop location=C2.");
    }
    const sourceRange = options.props.sourceRange ?? options.props.sourcerange ?? options.props.range ?? options.props.data;
    if (!sourceRange) {
      throw new UsageError("Excel sparkline requires --prop range=A1:B1 or --prop sourceRange=A1:B1.");
    }
    const sparkline = addSparkline(sheet, {
      location,
      sourceRange: sourceRange.includes("!") ? sourceRange : `${sheet.name}!${sourceRange}`,
      type: options.props.type?.toLowerCase(),
    });
    await writeWorkbookState(filePath, state);
    return { ...sparkline, ...(sparkline.type ? { sparklineType: sparkline.type } : {}), path: `/${sheet.name}/sparkline[${parseSparklines(sheet.xml).length}]`, type: "sparkline" };
  }

  if (options.type !== "cell") {
    throw new UsageError(
      "Excel add currently supports: sheet, row, cell, namedrange, validation, comment, autofilter, rowbreak, colbreak, table, or sparkline.",
      "Use / with --type sheet|namedrange, or /Sheet1 with --type row|cell|validation|comment|autofilter|rowbreak|colbreak|table|sparkline.",
    );
  }

  const sheet = ensureSheetState(state, normalizeSheetPath(targetPath) || options.props.sheet || "Sheet1");
  const ref = (options.props.ref ?? options.props.cell ?? "A1").toUpperCase();
  sheet.cells[ref] = mergeExcelCell(sheet.cells[ref], options.props);
  applyCellStyleProps(state, sheet.cells[ref], options.props);
  updateSheetXml(sheet);
  await writeWorkbookState(filePath, state);
  return materializeCellNode(sheet, ref);
}

export async function importExcelDelimitedData(
  filePath: string,
  parentPath: string,
  content: string,
  options: ExcelImportOptions,
) {
  const state = await loadWorkbookState(filePath);
  const sheet = ensureSheetState(state, normalizeSheetPath(parentPath));
  const rows = parseDelimitedRows(content, options.delimiter);
  if (rows.length === 0) {
    return { importedRows: 0, importedCols: 0, sheet: sheet.name, startCell: options.startCell.toUpperCase() };
  }

  const { column, row } = parseCellAddress(options.startCell.toUpperCase());
  const startColumnIndex = columnNameToIndex(column);
  let maxColumns = 0;

  for (const [rowOffset, values] of rows.entries()) {
    maxColumns = Math.max(maxColumns, values.length);
    for (const [columnOffset, rawValue] of values.entries()) {
      const ref = `${indexToColumnName(startColumnIndex + columnOffset)}${row + rowOffset}`;
      sheet.cells[ref] = inferImportedCell(rawValue);
    }
  }

  if (options.hasHeader && rows.length > 0) {
    const endColumn = indexToColumnName(startColumnIndex + Math.max(maxColumns - 1, 0));
    const endRow = row + rows.length - 1;
    sheet.autoFilter = `${column}${row}:${endColumn}${endRow}`;
    sheet.freezeTopLeftCell = `${column}${row + 1}`;
  }
  updateSheetXml(sheet);
  await writeWorkbookState(filePath, state);
  return {
    importedRows: rows.length,
    importedCols: maxColumns,
    sheet: sheet.name,
    startCell: options.startCell.toUpperCase(),
    ...(sheet.autoFilter ? { autoFilter: sheet.autoFilter } : {}),
    ...(sheet.freezeTopLeftCell ? { freezeTopLeftCell: sheet.freezeTopLeftCell } : {}),
  };
}

export async function setExcelNode(filePath: string, targetPath: string, options: ExcelCommandOptions) {
  const state = await loadWorkbookState(filePath);
  if (/^\/namedrange\[(.+)\]$/i.test(targetPath)) {
    const range = resolveNamedRangeNode(state, targetPath) as ExcelNamedRangeModel;
    for (const [key, value] of Object.entries(options.props)) {
      switch (key.toLowerCase()) {
        case "name":
          range.name = value;
          break;
        case "ref":
          range.ref = value;
          break;
        case "comment":
          range.comment = value;
          break;
        case "scope":
          range.scope = value.toLowerCase() === "workbook" ? undefined : value;
          break;
        default:
          throw new UsageError(`Unsupported namedrange property '${key}'.`, "Supported: name, ref, comment, scope.");
      }
    }
    updateWorkbookXml(state);
    await writeWorkbookState(filePath, state);
    return range;
  }

  if (targetPath === "/" || targetPath === "/workbook") {
    state.settings = mergeWorkbookSettings(state.settings, options.props);
    updateWorkbookXml(state);
    await writeWorkbookState(filePath, state);
    return materializeWorkbookRoot(state);
  }

  const sheetOnlyMatch = /^\/([^/]+)$/.exec(targetPath);
  if (sheetOnlyMatch) {
    const sheet = ensureSheetState(state, sheetOnlyMatch[1]);
    applySheetProperties(sheet, options.props);
    await writeWorkbookState(filePath, state);
    return materializeSheetNode(state, targetPath);
  }

  const validationMatch = /^\/([^/]+)\/validation\[(\d+)\]$/i.exec(targetPath);
  if (validationMatch) {
    const sheet = ensureSheetState(state, validationMatch[1]);
    const index = Number(validationMatch[2]);
    const validations = parseValidations(sheet.xml);
    const validation = validations[index - 1];
    if (!validation) {
      throw new OfficekitError(`Validation ${index} does not exist.`, "not_found");
    }
    const next = mergeValidation(validation, options.props);
    sheet.xml = replaceSheetValidations(sheet.xml, validations.map((item, itemIndex) => itemIndex === index - 1 ? next : item));
    await writeWorkbookState(filePath, state);
    return next;
  }

  const commentMatch = /^\/([^/]+)\/comment\[(\d+)\]$/i.exec(targetPath);
  if (commentMatch) {
    const sheet = ensureSheetState(state, commentMatch[1]);
    const next = setComment(state, sheet, Number(commentMatch[2]), options.props);
    await writeWorkbookState(filePath, state);
    return next;
  }

  const tableMatch = /^\/([^/]+)\/table\[(\d+)\]$/i.exec(targetPath);
  if (tableMatch) {
    const sheet = ensureSheetState(state, tableMatch[1]);
    const next = setTable(state, sheet, Number(tableMatch[2]), options.props);
    await writeWorkbookState(filePath, state);
    return next;
  }

  const drawingObjectMatch = /^\/([^/]+)\/(shape|picture)\[(\d+)\]$/i.exec(targetPath);
  if (drawingObjectMatch) {
    const sheet = ensureSheetState(state, drawingObjectMatch[1]);
    const next = setDrawingObject(state, sheet, drawingObjectMatch[2].toLowerCase() as "shape" | "picture", Number(drawingObjectMatch[3]), options.props);
    await writeWorkbookState(filePath, state);
    return next;
  }

  const sparklineMatch = /^\/([^/]+)\/sparkline\[(\d+)\]$/i.exec(targetPath);
  if (sparklineMatch) {
    const sheet = ensureSheetState(state, sparklineMatch[1]);
    const next = setSparkline(sheet, Number(sparklineMatch[2]), options.props);
    await writeWorkbookState(filePath, state);
    return next;
  }

  const chartMatch = /^\/([^/]+)\/chart\[(\d+)\](?:\/series\[(\d+)\])?$/i.exec(targetPath);
  if (chartMatch) {
    const sheet = ensureSheetState(state, chartMatch[1]);
    const next = setChart(state, sheet, Number(chartMatch[2]), chartMatch[3] ? Number(chartMatch[3]) : undefined, options.props);
    await writeWorkbookState(filePath, state);
    return next;
  }

  const pivotMatch = /^\/([^/]+)\/pivottable\[(\d+)\]$/i.exec(targetPath);
  if (pivotMatch) {
    const sheet = ensureSheetState(state, pivotMatch[1]);
    const next = setPivotTable(state, sheet, Number(pivotMatch[2]), options.props);
    await writeWorkbookState(filePath, state);
    return next;
  }

  const { sheet, range } = resolveExcelTarget(state, targetPath);
  if (!range) {
    throw new UsageError("Excel set requires a cell, range, or supported object path.");
  }
  if (range.includes(":")) {
    for (const ref of expandRange(range)) {
      sheet.cells[ref] = mergeExcelCell(sheet.cells[ref], options.props);
      applyCellStyleProps(state, sheet.cells[ref], options.props);
    }
  } else {
    sheet.cells[range] = mergeExcelCell(sheet.cells[range], options.props);
    applyCellStyleProps(state, sheet.cells[range], options.props);
  }
  updateSheetXml(sheet);
  await writeWorkbookState(filePath, state);
  return range.includes(":")
    ? materializeRangeNode(sheet, range)
    : materializeCellNode(sheet, range);
}

export async function removeExcelNode(filePath: string, targetPath: string) {
  const state = await loadWorkbookState(filePath);
  if (/^\/namedrange\[(.+)\]$/i.test(targetPath)) {
    const selector = /^\/namedrange\[(.+)\]$/i.exec(targetPath)?.[1] ?? "";
    const nextRanges = state.namedRanges.filter((range, index) => {
      if (/^\d+$/.test(selector)) {
        return index !== Number(selector) - 1;
      }
      return range.name.toLowerCase() !== selector.toLowerCase();
    });
    if (nextRanges.length === state.namedRanges.length) {
      throw new OfficekitError(`Named range '${selector}' not found.`, "not_found");
    }
    state.namedRanges = nextRanges;
    updateWorkbookXml(state);
    await writeWorkbookState(filePath, state);
    return { ok: true, targetPath };
  }

  const validationMatch = /^\/([^/]+)\/validation\[(\d+)\]$/i.exec(targetPath);
  if (validationMatch) {
    const sheet = ensureSheetState(state, validationMatch[1]);
    const validations = parseValidations(sheet.xml);
    const index = Number(validationMatch[2]) - 1;
    if (!validations[index]) throw new OfficekitError(`Validation ${validationMatch[2]} does not exist.`, "not_found");
    sheet.xml = replaceSheetValidations(sheet.xml, validations.filter((_, itemIndex) => itemIndex !== index));
    await writeWorkbookState(filePath, state);
    return { ok: true, targetPath };
  }

  const commentMatch = /^\/([^/]+)\/comment\[(\d+)\]$/i.exec(targetPath);
  if (commentMatch) {
    const sheet = ensureSheetState(state, commentMatch[1]);
    removeComment(state, sheet, Number(commentMatch[2]));
    await writeWorkbookState(filePath, state);
    return { ok: true, targetPath };
  }

  const tableMatch = /^\/([^/]+)\/table\[(\d+)\]$/i.exec(targetPath);
  if (tableMatch) {
    const sheet = ensureSheetState(state, tableMatch[1]);
    removeTable(state, sheet, Number(tableMatch[2]));
    await writeWorkbookState(filePath, state);
    return { ok: true, targetPath };
  }

  const sparklineMatch = /^\/([^/]+)\/sparkline\[(\d+)\]$/i.exec(targetPath);
  if (sparklineMatch) {
    const sheet = ensureSheetState(state, sparklineMatch[1]);
    removeSparkline(sheet, Number(sparklineMatch[2]));
    await writeWorkbookState(filePath, state);
    return { ok: true, targetPath };
  }

  const drawingObjectMatch = /^\/([^/]+)\/(shape|picture)\[(\d+)\]$/i.exec(targetPath);
  if (drawingObjectMatch) {
    const sheet = ensureSheetState(state, drawingObjectMatch[1]);
    removeDrawingObject(state, sheet, drawingObjectMatch[2].toLowerCase() as "shape" | "picture", Number(drawingObjectMatch[3]));
    await writeWorkbookState(filePath, state);
    return { ok: true, targetPath };
  }

  const { sheet, range } = resolveExcelTarget(state, targetPath);
  if (!range) {
    throw new UsageError("Excel remove currently supports cell, range, or namedrange paths.");
  }
  for (const ref of expandRange(range)) {
    delete sheet.cells[ref];
  }
  updateSheetXml(sheet);
  await writeWorkbookState(filePath, state);
  return { ok: true, targetPath };
}

export async function getExcelNode(filePath: string, targetPath: string) {
  const state = await loadWorkbookState(filePath);
  return materializeExcelPath(state, targetPath);
}

export async function queryExcelNodes(filePath: string, selector: string) {
  const state = await loadWorkbookState(filePath);
  if (selector.startsWith("/")) {
    return [materializeExcelPath(state, selector)];
  }

  const normalized = selector.trim().toLowerCase();
  const nodes: unknown[] = [];
  if (normalized === "sheet" || normalized === "sheets") {
    return state.sheets.map((sheet) => materializeSheetNode(state, `/${sheet.name}`));
  }
  if (normalized === "namedrange" || normalized === "namedranges") {
    return state.namedRanges.map((range) => ({
      ...range,
      path: `/namedrange[${range.name}]`,
      type: "namedrange",
    }));
  }
  if (normalized === "cell" || normalized === "cells") {
    for (const sheet of state.sheets) {
      for (const ref of Object.keys(sheet.cells).sort()) {
        nodes.push(materializeCellNode(sheet, ref));
      }
    }
    return nodes;
  }
  if (normalized === "formula" || normalized === "formulas") {
    for (const sheet of state.sheets) {
      for (const ref of Object.keys(sheet.cells).sort()) {
      if (sheet.cells[ref]?.formula) {
          nodes.push(materializeCellNode(sheet, ref));
        }
      }
    }
    return nodes;
  }
  if (normalized === "validation" || normalized === "validations") {
    for (const sheet of state.sheets) {
      parseValidations(sheet.xml).forEach((validation, index) => {
        nodes.push({ ...validation, ...(validation.type ? { validationType: validation.type } : {}), path: `/${sheet.name}/validation[${index + 1}]`, type: "validation" });
      });
    }
    return nodes;
  }
  if (normalized === "comment" || normalized === "comments") {
    for (const sheet of state.sheets) {
      getSheetComments(state, sheet).forEach((comment, index) => {
        nodes.push({ ...comment, path: `/${sheet.name}/comment[${index + 1}]`, type: "comment" });
      });
    }
    return nodes;
  }
  if (normalized === "shape" || normalized === "shapes") {
    for (const sheet of state.sheets) {
      getDrawingShapes(state, sheet)
        .filter((item) => item.kind === "shape")
        .forEach((shape, index) => {
          nodes.push({ ...shape, path: `/${sheet.name}/shape[${index + 1}]`, type: "shape" });
        });
    }
    return nodes;
  }
  if (normalized === "picture" || normalized === "pictures") {
    for (const sheet of state.sheets) {
      getDrawingShapes(state, sheet)
        .filter((item) => item.kind === "picture")
        .forEach((picture, index) => {
          nodes.push({ ...picture, path: `/${sheet.name}/picture[${index + 1}]`, type: "picture" });
        });
    }
    return nodes;
  }
  if (normalized === "table" || normalized === "tables") {
    for (const sheet of state.sheets) {
      getSheetTables(state, sheet).forEach((table, index) => {
        nodes.push({ ...table, path: `/${sheet.name}/table[${index + 1}]`, type: "table" });
      });
    }
    return nodes;
  }
  if (normalized === "chart" || normalized === "charts") {
    for (const sheet of state.sheets) {
      getSheetCharts(state, sheet).forEach((chart, index) => {
        nodes.push({ ...chart, path: `/${sheet.name}/chart[${index + 1}]`, type: "chart" });
      });
    }
    return nodes;
  }
  if (normalized === "pivottable" || normalized === "pivot" || normalized === "pivots") {
    for (const sheet of state.sheets) {
      getSheetPivots(state, sheet).forEach((pivot, index) => {
        nodes.push({ ...pivot, path: `/${sheet.name}/pivottable[${index + 1}]`, type: "pivottable" });
      });
    }
    return nodes;
  }
  if (normalized === "sparkline" || normalized === "sparklines") {
    for (const sheet of state.sheets) {
      parseSparklines(sheet.xml).forEach((sparkline, index) => {
        nodes.push({ ...sparkline, ...(sparkline.type ? { sparklineType: sparkline.type } : {}), path: `/${sheet.name}/sparkline[${index + 1}]`, type: "sparkline" });
      });
    }
    return nodes;
  }
  throw new UsageError(`Unsupported Excel query selector '${selector}'.`, "Supported selectors: sheet, namedrange, cell, formula, validation, comment, table, chart, pivottable, sparkline, shape, picture.");
}

export async function viewExcelDocument(filePath: string, mode: string) {
  const state = await loadWorkbookState(filePath);
  const normalizedMode = mode.toLowerCase();
  if (normalizedMode === "json") {
    return { mode, output: JSON.stringify(materializeWorkbookRoot(state), null, 2) };
  }
  if (normalizedMode === "outline") {
    return { mode, output: renderOutline(state) };
  }
  if (normalizedMode === "text") {
    return { mode, output: renderTextView(state) };
  }
  if (normalizedMode === "annotated") {
    return { mode, output: renderAnnotatedView(state) };
  }
  if (normalizedMode === "stats") {
    return { mode, output: renderStatsView(state) };
  }
  if (normalizedMode === "issues") {
    return { mode, output: renderIssuesView(state) };
  }
  if (normalizedMode === "html") {
    return { mode, output: renderHtmlView(state) };
  }
  throw new UsageError(`Unsupported Excel view mode '${mode}'.`, "Use outline, text, annotated, stats, issues, html, or json.");
}

export async function rawExcelDocument(
  filePath: string,
  partPath = "/",
  options?: { startRow?: number; endRow?: number; cols?: string[] },
) {
  const state = await loadWorkbookState(filePath);
  if (partPath === "/" || partPath === "/workbook") {
    return state.workbookXml;
  }
  if (partPath === "/styles") {
    return state.styleSheetXml ?? "(no styles)";
  }
  if (partPath === "/sharedstrings") {
    const shared = state.zip.get("xl/sharedStrings.xml");
    return shared ? shared.toString("utf8") : "(no shared strings)";
  }
  const drawingMatch = /^\/([^/]+)\/drawing$/i.exec(partPath);
  if (drawingMatch) {
    const sheet = ensureSheetState(state, drawingMatch[1]);
    const drawingPath = resolveDrawingPath(state, sheet);
    if (!drawingPath) {
      throw new UsageError(`Sheet '${sheet.name}' has no drawing part.`);
    }
    return requireEntry(state.zip, drawingPath);
  }
  const chartMatch = /^\/([^/]+)\/chart\[(\d+)\]$/i.exec(partPath);
  if (chartMatch) {
    const sheet = ensureSheetState(state, chartMatch[1]);
    const chart = getSheetCharts(state, sheet)[Number(chartMatch[2]) - 1];
    const drawingPath = resolveDrawingPath(state, sheet);
    if (!chart?.path || !drawingPath) {
      throw new OfficekitError(`Chart ${chartMatch[2]} not found.`, "not_found");
    }
    return requireEntry(state.zip, normalizeZipPath(path.posix.dirname(drawingPath), chart.path));
  }
  const globalChartMatch = /^\/chart\[(\d+)\]$/i.exec(partPath);
  if (globalChartMatch) {
    const charts = state.sheets.flatMap((sheet) => getSheetCharts(state, sheet));
    const chart = charts[Number(globalChartMatch[1]) - 1];
    if (!chart?.path) {
      throw new OfficekitError(`Chart ${globalChartMatch[1]} not found.`, "not_found");
    }
    const ownerSheet = state.sheets.find((sheet) => getSheetCharts(state, sheet).some((candidate) => candidate.path === chart.path));
    const drawingPath = ownerSheet ? resolveDrawingPath(state, ownerSheet) : undefined;
    if (!drawingPath) {
      throw new OfficekitError(`Chart ${globalChartMatch[1]} drawing is missing.`, "invalid_ooxml");
    }
    return requireEntry(state.zip, normalizeZipPath(path.posix.dirname(drawingPath), chart.path));
  }
  const sheetMatch = /^\/([^/]+)$/i.exec(partPath);
  if (sheetMatch) {
    const sheet = ensureSheetState(state, sheetMatch[1]);
    if (options?.startRow !== undefined || options?.endRow !== undefined || options?.cols?.length) {
      return filterSheetRaw(sheet.xml, options);
    }
    return sheet.xml;
  }
  throw new UsageError(`Unsupported Excel raw part '${partPath}'.`, "Use /workbook, /styles, /sharedstrings, /Sheet1, /Sheet1/drawing, /Sheet1/chart[1], or /chart[1].");
}

export function renderExcelHtmlFromRoot(root: unknown) {
  if (!root || typeof root !== "object" || !("sheets" in root)) {
    return `<section data-format="excel"><table><tbody><tr><td colspan="2"><em>Empty workbook</em></td></tr></tbody></table></section>`;
  }
  const workbook = root as { sheets?: Array<{ name: string; cells?: Record<string, ExcelCellModel> }> };
  const rows = (workbook.sheets ?? []).flatMap((sheet) =>
    Object.entries(sheet.cells ?? {}).map(([ref, cell]) => {
      const detail = [cell.value, cell.formula ? `formula=${cell.formula}` : ""].filter(Boolean).join(" · ");
      return `<tr><th>${escapeHtml(sheet.name)}!${escapeHtml(ref)}</th><td>${escapeHtml(detail)}</td></tr>`;
    }),
  );
  return `<section data-format="excel"><table><tbody>${rows.join("") || '<tr><td colspan="2"><em>Empty workbook</em></td></tr>'}</tbody></table></section>`;
}

export function summarizeExcelCheck(filePath: string) {
  return loadWorkbookState(filePath).then((state) => ({
    ok: true,
    format: "excel",
    summary: renderOutline(state),
  }));
}

function createBlankWorkbookState(): ExcelWorkbookState {
  const zip = new Map<string, Buffer>();
  const state: ExcelWorkbookState = {
    zip,
    workbookXml: "",
    workbookEntryName: "xl/workbook.xml",
    workbookRelsXml: "",
    workbookRelsEntryName: "xl/_rels/workbook.xml.rels",
    sheets: [],
    settings: {},
    namedRanges: [],
    metadata: {},
  };
  addSheet(state, "Sheet1");
  updateWorkbookXml(state);
  return state;
}

async function loadWorkbookState(filePath: string): Promise<ExcelWorkbookState> {
  const zip = readStoredZip(await readFile(filePath));
  const workbookEntryName = "xl/workbook.xml";
  const workbookRelsEntryName = "xl/_rels/workbook.xml.rels";
  const workbookXml = requireEntry(zip, workbookEntryName);
  const workbookRelsXml = requireEntry(zip, workbookRelsEntryName);
  const relationships = parseRelationshipEntries(workbookRelsXml);
  const settings = parseWorkbookSettings(workbookXml);
  const officekitMetadata = parseOfficekitMetadata(zip);
  const namedRanges = parseDefinedNames(workbookXml, []);
  const sheets = [...workbookXml.matchAll(/<(?:\w+:)?sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g)].map((match) => {
    const name = decodeXml(match[1]);
    const relId = match[2];
    const target = relationships.find((relationship) => relationship.id === relId)?.target;
    if (!target) {
      throw new OfficekitError(`Workbook relationship '${relId}' is missing.`, "invalid_ooxml");
    }
    const entryName = normalizeZipPath("xl", target);
    const xml = requireEntry(zip, entryName);
    const features = parseSheetFeatures(xml);
    const metadataSheet = officekitMetadata?.excel?.sheets?.find((item) => item.name.toLowerCase() === name.toLowerCase());
    const parsedCells = parseSheetCells(xml, zip);
    return {
      name,
      relId,
      relationshipTarget: target,
      entryName,
      xml,
      cells: overlayMetadataCells(parsedCells, metadataSheet?.cells),
      ...features,
      ...(metadataSheet?.autoFilter ? { autoFilter: metadataSheet.autoFilter } : {}),
      ...(metadataSheet?.freezeTopLeftCell ? { freezeTopLeftCell: metadataSheet.freezeTopLeftCell } : {}),
      ...(metadataSheet?.zoom !== undefined ? { zoom: metadataSheet.zoom } : {}),
      ...(metadataSheet?.showGridLines !== undefined ? { showGridLines: metadataSheet.showGridLines } : {}),
      ...(metadataSheet?.showHeadings !== undefined ? { showHeadings: metadataSheet.showHeadings } : {}),
      ...(metadataSheet?.tabColor ? { tabColor: metadataSheet.tabColor } : {}),
      ...(metadataSheet?.header ? { header: metadataSheet.header } : {}),
      ...(metadataSheet?.footer ? { footer: metadataSheet.footer } : {}),
      ...(metadataSheet?.orientation ? { orientation: metadataSheet.orientation } : {}),
      ...(metadataSheet?.paperSize !== undefined ? { paperSize: metadataSheet.paperSize } : {}),
      ...(metadataSheet?.fitToPage ? { fitToPage: metadataSheet.fitToPage } : {}),
      ...(metadataSheet?.protection !== undefined ? { protection: metadataSheet.protection } : {}),
      ...(metadataSheet?.rowBreaks?.length ? { rowBreaks: [...metadataSheet.rowBreaks] } : {}),
      ...(metadataSheet?.colBreaks?.length ? { colBreaks: [...metadataSheet.colBreaks] } : {}),
    } satisfies ExcelSheetModel;
  });
  const scopedNamedRanges = officekitMetadata?.excel?.namedRanges ?? parseDefinedNames(workbookXml, sheets);
  return {
    zip,
    workbookXml,
    workbookEntryName,
    workbookRelsXml,
    workbookRelsEntryName,
    sheets,
    settings: officekitMetadata?.excel?.settings ?? settings,
    namedRanges: scopedNamedRanges,
    styleSheetXml: zip.get("xl/styles.xml")?.toString("utf8"),
    metadata: parsePackageProperties(zip),
    officekitMetadata,
  };
}

async function writeWorkbookState(filePath: string, state: ExcelWorkbookState) {
  await mkdir(path.dirname(filePath), { recursive: true });
  state.zip.set(state.workbookEntryName, Buffer.from(state.workbookXml, "utf8"));
  state.zip.set(state.workbookRelsEntryName, Buffer.from(state.workbookRelsXml, "utf8"));
  state.zip.set("[Content_Types].xml", Buffer.from(buildContentTypesXml(state), "utf8"));
  state.zip.set("_rels/.rels", Buffer.from(buildRootRelsXml(state), "utf8"));
  state.zip.set(METADATA_PATH, Buffer.from(JSON.stringify({
    product: "officekit",
    lineage: LINEAGE,
    format: "excel",
    version: 1,
    updatedAt: new Date().toISOString(),
    excel: {
      sheets: state.sheets.map((sheet) => ({
        name: sheet.name,
        cells: sheet.cells,
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
        ...(sheet.rowBreaks?.length ? { rowBreaks: sheet.rowBreaks } : {}),
        ...(sheet.colBreaks?.length ? { colBreaks: sheet.colBreaks } : {}),
      })),
      ...(Object.keys(state.settings).length > 0 ? { settings: state.settings } : {}),
      ...(state.styleSheetXml ? { styleSheetXml: state.styleSheetXml } : {}),
      ...(state.namedRanges.length > 0 ? { namedRanges: state.namedRanges } : {}),
    },
  }, null, 2), "utf8"));
  if (!state.zip.has("docProps/core.xml")) {
    state.zip.set("docProps/core.xml", Buffer.from(buildCorePropertiesXml(state.metadata), "utf8"));
  }
  if (state.styleSheetXml) {
    state.zip.set("xl/styles.xml", Buffer.from(state.styleSheetXml, "utf8"));
  }
  for (const sheet of state.sheets) {
    state.zip.set(sheet.entryName, Buffer.from(sheet.xml, "utf8"));
  }
  const entries = [...state.zip.entries()].map(([name, data]) => ({ name, data }));
  await writeFile(filePath, createStoredZip(entries));
}

function materializeExcelPath(state: ExcelWorkbookState, targetPath: string): unknown {
  if (targetPath === "/" || targetPath === "/workbook") {
    return materializeWorkbookRoot(state);
  }
  if (targetPath === "/styles") {
    return { path: targetPath, type: "styles", xml: state.styleSheetXml ?? null };
  }
  if (targetPath === "/sharedstrings") {
    return { path: targetPath, type: "sharedstrings", count: parseSharedStrings(state.zip).length };
  }
  if (/^\/namedrange\[(.+)\]$/i.test(targetPath)) {
    return resolveNamedRangeNode(state, targetPath);
  }
  const validationMatch = /^\/([^/]+)\/validation\[(\d+)\]$/i.exec(targetPath);
  if (validationMatch) {
    const sheet = ensureSheetState(state, validationMatch[1]);
    const validation = parseValidations(sheet.xml)[Number(validationMatch[2]) - 1];
    if (!validation) throw new OfficekitError(`Validation ${validationMatch[2]} does not exist.`, "not_found");
    return { ...validation, ...(validation.type ? { validationType: validation.type } : {}), path: targetPath, type: "validation" };
  }
  const commentMatch = /^\/([^/]+)\/comment\[(\d+)\]$/i.exec(targetPath);
  if (commentMatch) {
    const sheet = ensureSheetState(state, commentMatch[1]);
    const comment = getSheetComments(state, sheet)[Number(commentMatch[2]) - 1];
    if (!comment) throw new OfficekitError(`Comment ${commentMatch[2]} does not exist.`, "not_found");
    return { ...comment, path: targetPath, type: "comment" };
  }
  const tableMatch = /^\/([^/]+)\/table\[(\d+)\]$/i.exec(targetPath);
  if (tableMatch) {
    const sheet = ensureSheetState(state, tableMatch[1]);
    const table = getSheetTables(state, sheet)[Number(tableMatch[2]) - 1];
    if (!table) throw new OfficekitError(`Table ${tableMatch[2]} does not exist.`, "not_found");
    return { ...table, path: targetPath, type: "table" };
  }
  const chartMatch = /^\/([^/]+)\/chart\[(\d+)\](?:\/series\[(\d+)\])?$/i.exec(targetPath);
  if (chartMatch) {
    const sheet = ensureSheetState(state, chartMatch[1]);
    const chart = getSheetCharts(state, sheet)[Number(chartMatch[2]) - 1];
    if (!chart) throw new OfficekitError(`Chart ${chartMatch[2]} does not exist.`, "not_found");
    return { ...chart, path: targetPath, type: chartMatch[3] ? "chart-series" : "chart", series: chartMatch[3] ? { index: Number(chartMatch[3]) } : undefined };
  }
  const pivotMatch = /^\/([^/]+)\/pivottable\[(\d+)\]$/i.exec(targetPath);
  if (pivotMatch) {
    const sheet = ensureSheetState(state, pivotMatch[1]);
    const pivot = getSheetPivots(state, sheet)[Number(pivotMatch[2]) - 1];
    if (!pivot) throw new OfficekitError(`Pivot table ${pivotMatch[2]} does not exist.`, "not_found");
    return { ...pivot, path: targetPath, type: "pivottable" };
  }
  const sparklineMatch = /^\/([^/]+)\/sparkline\[(\d+)\]$/i.exec(targetPath);
  if (sparklineMatch) {
    const sheet = ensureSheetState(state, sparklineMatch[1]);
    const sparkline = parseSparklines(sheet.xml)[Number(sparklineMatch[2]) - 1];
    if (!sparkline) throw new OfficekitError(`Sparkline ${sparklineMatch[2]} does not exist.`, "not_found");
    return { ...sparkline, ...(sparkline.type ? { sparklineType: sparkline.type } : {}), path: targetPath, type: "sparkline" };
  }
  const shapeMatch = /^\/([^/]+)\/(shape|picture)\[(\d+)\]$/i.exec(targetPath);
  if (shapeMatch) {
    const sheet = ensureSheetState(state, shapeMatch[1]);
    const shapes = shapeMatch[2].toLowerCase() === "picture" ? getDrawingShapes(state, sheet).filter((item) => item.kind === "picture") : getDrawingShapes(state, sheet).filter((item) => item.kind === "shape");
    const shape = shapes[Number(shapeMatch[3]) - 1];
    if (!shape) throw new OfficekitError(`${shapeMatch[2]} ${shapeMatch[3]} does not exist.`, "not_found");
    return { ...shape, path: targetPath, type: shapeMatch[2].toLowerCase() };
  }
  const rowBreakMatch = /^\/([^/]+)\/rowbreak\[(\d+)\]$/i.exec(targetPath);
  if (rowBreakMatch) {
    const sheet = ensureSheetState(state, rowBreakMatch[1]);
    const rowBreak = parseBreaks(sheet.xml, "row")[Number(rowBreakMatch[2]) - 1];
    if (!rowBreak) throw new OfficekitError(`Row break ${rowBreakMatch[2]} does not exist.`, "not_found");
    return { ...rowBreak, path: targetPath, type: "rowbreak" };
  }
  const colBreakMatch = /^\/([^/]+)\/colbreak\[(\d+)\]$/i.exec(targetPath);
  if (colBreakMatch) {
    const sheet = ensureSheetState(state, colBreakMatch[1]);
    const colBreak = parseBreaks(sheet.xml, "col")[Number(colBreakMatch[2]) - 1];
    if (!colBreak) throw new OfficekitError(`Column break ${colBreakMatch[2]} does not exist.`, "not_found");
    return { ...colBreak, path: targetPath, type: "colbreak" };
  }
  const rowMatch = /^\/([^/]+)\/row\[(\d+)\]$/i.exec(targetPath);
  if (rowMatch) {
    const sheet = ensureSheetState(state, rowMatch[1]);
    return {
      path: targetPath,
      type: "row",
      index: Number(rowMatch[2]),
      cells: collectRowCells(sheet, Number(rowMatch[2])),
    };
  }
  const colMatch = /^\/([^/]+)\/col\[([A-Z0-9]+)\]$/i.exec(targetPath);
  if (colMatch) {
    const sheet = ensureSheetState(state, colMatch[1]);
    return {
      path: targetPath,
      type: "column",
      index: colMatch[2],
      cells: collectColumnCells(sheet, colMatch[2]),
    };
  }
  const { sheet, range } = resolveExcelTarget(state, targetPath);
  if (!range) {
    return materializeSheetNode(state, targetPath);
  }
  return range.includes(":") ? materializeRangeNode(sheet, range) : materializeCellNode(sheet, range);
}

function materializeWorkbookRoot(state: ExcelWorkbookState) {
  return {
    path: "/workbook",
    type: "workbook",
    sheets: state.sheets.map((sheet) => ({
      name: sheet.name,
      cells: sheet.cells,
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
      ...(sheet.rowBreaks?.length ? { rowBreaks: sheet.rowBreaks } : {}),
      ...(sheet.colBreaks?.length ? { colBreaks: sheet.colBreaks } : {}),
    })),
    ...(Object.keys(state.settings).length > 0 ? { settings: state.settings } : {}),
    ...(state.styleSheetXml ? { styleSheetXml: state.styleSheetXml } : {}),
    ...(state.namedRanges.length > 0 ? { namedRanges: state.namedRanges } : {}),
    metadata: state.metadata,
  };
}

function materializeSheetNode(state: ExcelWorkbookState, targetPath: string) {
  const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
  return {
    path: `/${sheet.name}`,
    type: "sheet",
    name: sheet.name,
    childCount: Object.keys(sheet.cells).length,
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
    ...(sheet.rowBreaks?.length ? { rowBreaks: sheet.rowBreaks } : {}),
    ...(sheet.colBreaks?.length ? { colBreaks: sheet.colBreaks } : {}),
  };
}

function materializeRangeNode(sheet: ExcelSheetModel, range: string) {
  return {
    path: `/${sheet.name}/${range}`,
    type: "range",
    cells: expandRange(range).map((ref) => materializeCellNode(sheet, ref)),
  };
}

function materializeCellNode(sheet: ExcelSheetModel, ref: string) {
  const cell = sheet.cells[ref];
  if (!cell) {
    return { path: `/${sheet.name}/${ref}`, ref, type: "cell", value: null };
  }
  const evaluatedValue = cell.formula && cell.value === "" ? evaluateFormulaForDisplay(sheet, ref) : undefined;
  return {
    path: `/${sheet.name}/${ref}`,
    ref,
    type: "cell",
    ...cell,
    ...(evaluatedValue !== undefined ? { evaluatedValue } : {}),
  };
}

function resolveNamedRangeNode(state: ExcelWorkbookState, targetPath: string) {
  const selector = /^\/namedrange\[(.+)\]$/i.exec(targetPath)?.[1] ?? "";
  const index = /^\d+$/.test(selector)
    ? Number(selector) - 1
    : state.namedRanges.findIndex((range) => range.name.toLowerCase() === selector.toLowerCase());
  const range = state.namedRanges[index];
  if (!range) {
    throw new OfficekitError(`Named range '${selector}' not found.`, "not_found");
  }
  return range;
}

function ensureSheetState(state: ExcelWorkbookState, name: string) {
  const existing = state.sheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase());
  if (existing) return existing;
  addSheet(state, name);
  return state.sheets[state.sheets.length - 1];
}

function addSheet(state: ExcelWorkbookState, name: string) {
  const nextIndex = state.sheets.length + 1;
  const relId = `rId${nextIndex}`;
  const entryName = `xl/worksheets/sheet${nextIndex}.xml`;
  const relationshipTarget = `worksheets/sheet${nextIndex}.xml`;
  const sheet: ExcelSheetModel = {
    name,
    relId,
    entryName,
    relationshipTarget,
    xml: buildSheetXml({ cells: {} }),
    cells: {},
  };
  state.sheets.push(sheet);
  updateWorkbookXml(state);
}

function updateWorkbookXml(state: ExcelWorkbookState) {
  const workbookPr = renderWorkbookProperties(state.settings);
  const workbookProtection = renderWorkbookProtection(state.settings);
  const calcPr = renderCalculationProperties(state.settings);
  const definedNames = renderDefinedNames(state.namedRanges, state.sheets);
  const sheetsXml = state.sheets
    .map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="${sheet.relId}"/>`)
    .join("");
  state.workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${workbookPr}
  ${workbookProtection}
  <sheets>${sheetsXml}</sheets>
  ${definedNames}
  ${calcPr}
</workbook>`;
  const sheetRels = state.sheets
    .map(
      (sheet) => `<Relationship Id="${sheet.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${sheet.relationshipTarget}"/>`,
    )
    .join("");
  const styleRel = state.styleSheetXml
    ? '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    : "";
  state.workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheetRels}${styleRel}</Relationships>`;
}

function updateSheetXml(sheet: ExcelSheetModel) {
  const nextXml = buildSheetXml({
    cells: sheet.cells,
    autoFilter: sheet.autoFilter,
    freezeTopLeftCell: sheet.freezeTopLeftCell,
    zoom: sheet.zoom,
    showGridLines: sheet.showGridLines,
    showHeadings: sheet.showHeadings,
    tabColor: sheet.tabColor,
    header: sheet.header,
    footer: sheet.footer,
    orientation: sheet.orientation,
    paperSize: sheet.paperSize,
    fitToPage: sheet.fitToPage,
    protection: sheet.protection,
    rowBreaks: sheet.rowBreaks,
    colBreaks: sheet.colBreaks,
  });
  sheet.xml = mergeSheetXmlPreservingExtras(sheet.xml, nextXml);
}

function buildSheetXml(sheet: {
  cells: Record<string, ExcelCellModel>;
  autoFilter?: string;
  freezeTopLeftCell?: string;
  zoom?: number;
  showGridLines?: boolean;
  showHeadings?: boolean;
  tabColor?: string;
  header?: string;
  footer?: string;
  orientation?: string;
  paperSize?: number;
  fitToPage?: string;
  protection?: boolean;
  rowBreaks?: number[];
  colBreaks?: number[];
}) {
  const entries = Object.entries(sheet.cells).sort(([a], [b]) => compareCellRefs(a, b));
  const rows = new Map<number, string[]>();
  for (const [ref, cell] of entries) {
    const row = Number(/\d+/.exec(ref)?.[0] ?? "1");
    const cells = rows.get(row) ?? [];
    cells.push(renderExcelCellXml(ref, cell));
    rows.set(row, cells);
  }
  const xmlRows = [...rows.entries()].sort(([a], [b]) => a - b).map(([rowIndex, cells]) => `<row r="${rowIndex}">${cells.join("")}</row>`).join("");
  const sheetViewAttrs = [
    sheet.zoom !== undefined ? `zoomScale="${sheet.zoom}"` : "",
    sheet.showGridLines === false ? 'showGridLines="0"' : "",
    sheet.showHeadings === false ? 'showRowColHeaders="0"' : "",
  ].filter(Boolean).join(" ");
  const pane = sheet.freezeTopLeftCell
    ? `<pane ySplit="1" topLeftCell="${escapeXml(sheet.freezeTopLeftCell)}" state="frozen" activePane="bottomLeft"/>`
    : "";
  const sheetViews = pane || sheetViewAttrs
    ? `<sheetViews><sheetView workbookViewId="0"${sheetViewAttrs ? ` ${sheetViewAttrs}` : ""}>${pane}</sheetView></sheetViews>`
    : "";
  const sheetPr = sheet.tabColor ? `<sheetPr><tabColor rgb="${escapeXml(normalizeArgbColor(sheet.tabColor))}"/></sheetPr>` : "";
  const autoFilter = sheet.autoFilter ? `<autoFilter ref="${escapeXml(sheet.autoFilter)}"/>` : "";
  const pageSetupAttrs = [
    sheet.orientation ? `orientation="${escapeXml(sheet.orientation)}"` : "",
    sheet.paperSize !== undefined ? `paperSize="${sheet.paperSize}"` : "",
    ...(sheet.fitToPage
      ? (() => {
          const [width, height] = sheet.fitToPage.split("x");
          return [`fitToWidth="${width}"`, `fitToHeight="${height ?? "1"}"`];
        })()
      : []),
  ].filter(Boolean).join(" ");
  const pageSetup = pageSetupAttrs ? `<pageSetup ${pageSetupAttrs}/>` : "";
  const headerFooter = sheet.header || sheet.footer
    ? `<headerFooter>${sheet.header ? `<oddHeader>${escapeXml(sheet.header)}</oddHeader>` : ""}${sheet.footer ? `<oddFooter>${escapeXml(sheet.footer)}</oddFooter>` : ""}</headerFooter>`
    : "";
  const sheetProtection = sheet.protection ? `<sheetProtection sheet="1"/>` : "";
  const rowBreaks = sheet.rowBreaks?.length
    ? `<rowBreaks count="${sheet.rowBreaks.length}" manualBreakCount="${sheet.rowBreaks.length}">${sheet.rowBreaks.map((row) => `<brk id="${row}" man="1"/>`).join("")}</rowBreaks>`
    : "";
  const colBreaks = sheet.colBreaks?.length
    ? `<colBreaks count="${sheet.colBreaks.length}" manualBreakCount="${sheet.colBreaks.length}">${sheet.colBreaks.map((col) => `<brk id="${col}" man="1"/>`).join("")}</colBreaks>`
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheetPr}
  ${sheetViews}
  <sheetData>${xmlRows}</sheetData>
  ${autoFilter}
  ${sheetProtection}
  ${pageSetup}
  ${headerFooter}
  ${rowBreaks}
  ${colBreaks}
</worksheet>`;
}

function mergeSheetXmlPreservingExtras(previousXml: string, nextXml: string) {
  let xml = previousXml;
  const nextSheetPr = /<(?:\w+:)?sheetPr\b[\s\S]*?<\/(?:\w+:)?sheetPr>/.exec(nextXml)?.[0] ?? "";
  const nextSheetViews = /<(?:\w+:)?sheetViews\b[\s\S]*?<\/(?:\w+:)?sheetViews>/.exec(nextXml)?.[0] ?? "";
  const nextSheetData = /<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/.exec(nextXml)?.[0] ?? "<sheetData/>";
  const nextAutoFilter = /<(?:\w+:)?autoFilter\b[^>]*\/?>/.exec(nextXml)?.[0] ?? "";
  const nextSheetProtection = /<(?:\w+:)?sheetProtection\b[^>]*\/?>/.exec(nextXml)?.[0] ?? "";
  const nextPageSetup = /<(?:\w+:)?pageSetup\b[^>]*\/?>/.exec(nextXml)?.[0] ?? "";
  const nextHeaderFooter = /<(?:\w+:)?headerFooter\b[\s\S]*?<\/(?:\w+:)?headerFooter>/.exec(nextXml)?.[0] ?? "";
  const nextRowBreaks = /<(?:\w+:)?rowBreaks\b[\s\S]*?<\/(?:\w+:)?rowBreaks>/.exec(nextXml)?.[0] ?? "";
  const nextColBreaks = /<(?:\w+:)?colBreaks\b[\s\S]*?<\/(?:\w+:)?colBreaks>/.exec(nextXml)?.[0] ?? "";

  xml = replaceOrInsert(xml, /<(?:\w+:)?sheetPr\b[\s\S]*?<\/(?:\w+:)?sheetPr>/, nextSheetPr, /<(?:\w+:)?worksheet\b[^>]*>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?sheetViews\b[\s\S]*?<\/(?:\w+:)?sheetViews>/, nextSheetViews, /<(?:\w+:)?sheetPr\b[\s\S]*?<\/(?:\w+:)?sheetPr>|<(?:\w+:)?worksheet\b[^>]*>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/, nextSheetData, /<(?:\w+:)?sheetViews\b[\s\S]*?<\/(?:\w+:)?sheetViews>|<(?:\w+:)?sheetPr\b[\s\S]*?<\/(?:\w+:)?sheetPr>|<(?:\w+:)?worksheet\b[^>]*>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?autoFilter\b[^>]*\/?>/, nextAutoFilter, /<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?sheetProtection\b[^>]*\/?>/, nextSheetProtection, /<(?:\w+:)?autoFilter\b[^>]*\/?>|<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?pageSetup\b[^>]*\/?>/, nextPageSetup, /<(?:\w+:)?sheetProtection\b[^>]*\/?>|<(?:\w+:)?autoFilter\b[^>]*\/?>|<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?headerFooter\b[\s\S]*?<\/(?:\w+:)?headerFooter>/, nextHeaderFooter, /<(?:\w+:)?pageSetup\b[^>]*\/?>|<(?:\w+:)?sheetProtection\b[^>]*\/?>|<(?:\w+:)?autoFilter\b[^>]*\/?>|<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?rowBreaks\b[\s\S]*?<\/(?:\w+:)?rowBreaks>/, nextRowBreaks, /<(?:\w+:)?headerFooter\b[\s\S]*?<\/(?:\w+:)?headerFooter>|<(?:\w+:)?pageSetup\b[^>]*\/?>|<(?:\w+:)?sheetProtection\b[^>]*\/?>|<(?:\w+:)?autoFilter\b[^>]*\/?>|<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?colBreaks\b[\s\S]*?<\/(?:\w+:)?colBreaks>/, nextColBreaks, /<(?:\w+:)?rowBreaks\b[\s\S]*?<\/(?:\w+:)?rowBreaks>|<(?:\w+:)?headerFooter\b[\s\S]*?<\/(?:\w+:)?headerFooter>|<(?:\w+:)?pageSetup\b[^>]*\/?>|<(?:\w+:)?sheetProtection\b[^>]*\/?>|<(?:\w+:)?autoFilter\b[^>]*\/?>|<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/);
  return xml;
}

function replaceOrInsert(xml: string, pattern: RegExp, replacement: string, anchorPattern: RegExp) {
  if (pattern.test(xml)) {
    return replacement ? xml.replace(pattern, replacement) : xml.replace(pattern, "");
  }
  if (!replacement) {
    return xml;
  }
  const anchor = anchorPattern.exec(xml);
  if (!anchor || anchor.index === undefined) {
    return xml;
  }
  const insertAt = anchor.index + anchor[0].length;
  return `${xml.slice(0, insertAt)}${replacement}${xml.slice(insertAt)}`;
}

function resolveExcelTarget(state: ExcelWorkbookState, targetPath: string) {
  const rangeMatch = /^\/([^/]+)\/([A-Z]+\d+(?::[A-Z]+\d+)?)$/i.exec(targetPath);
  if (rangeMatch) {
    return {
      sheet: ensureSheetState(state, rangeMatch[1]),
      range: rangeMatch[2].toUpperCase(),
    };
  }
  return {
    sheet: ensureSheetState(state, normalizeSheetPath(targetPath)),
    range: "",
  };
}

function renderOutline(state: ExcelWorkbookState) {
  const lines: string[] = [];
  for (const sheet of state.sheets) {
    lines.push(`Sheet ${sheet.name}`);
    for (const ref of Object.keys(sheet.cells).sort(compareCellRefs)) {
      const cell = sheet.cells[ref];
      lines.push(`  ${ref}: ${cell.value}${cell.formula ? ` (formula=${cell.formula})` : ""}`);
    }
    parseValidations(sheet.xml).forEach((validation, index) => {
      lines.push(`  Validation ${index + 1}: ${validation.sqref ?? ""}${validation.type ? ` [${validation.type}]` : ""}`);
    });
    getSheetComments(state, sheet).forEach((comment, index) => {
      lines.push(`  Comment ${index + 1}: ${comment.ref} = ${comment.text}`);
    });
    getSheetTables(state, sheet).forEach((table, index) => {
      lines.push(`  Table ${index + 1}: ${table.name ?? table.displayName ?? table.ref ?? "Unnamed table"}`);
    });
    getSheetCharts(state, sheet).forEach((chart, index) => {
      lines.push(`  Chart ${index + 1}: ${chart.title ?? "Untitled chart"}`);
    });
    getSheetPivots(state, sheet).forEach((pivot, index) => {
      lines.push(`  Pivot ${index + 1}: ${pivot.name ?? `Pivot ${index + 1}`}`);
    });
    parseSparklines(sheet.xml).forEach((sparkline, index) => {
      lines.push(`  Sparkline ${index + 1}: ${sparkline.location ?? "Unknown"}`);
    });
  }
  if (state.namedRanges.length > 0) {
    lines.push("Named ranges");
    state.namedRanges.forEach((range) => lines.push(`  ${range.name}: ${range.ref}`));
  }
  return lines.join("\n") || "Workbook is empty.";
}

function renderTextView(state: ExcelWorkbookState) {
  const lines: string[] = [];
  for (const sheet of state.sheets) {
    lines.push(`=== Sheet: ${sheet.name} ===`);
    const rowMap = new Map<number, Map<string, ExcelCellModel>>();
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      const row = Number(/\d+/.exec(ref)?.[0] ?? "1");
      const col = /^[A-Z]+/.exec(ref)?.[0] ?? "A";
      const byRow = rowMap.get(row) ?? new Map<string, ExcelCellModel>();
      byRow.set(col, cell);
      rowMap.set(row, byRow);
    }
    for (const [rowIndex, cells] of [...rowMap.entries()].sort(([a], [b]) => a - b)) {
      const ordered = [...cells.entries()]
        .sort(([a], [b]) => columnNameToIndex(a) - columnNameToIndex(b))
        .map(([column]) => {
          const materialized = materializeCellNode(sheet, `${column}${rowIndex}`) as { value?: string | null; evaluatedValue?: string };
          return materialized.evaluatedValue ?? materialized.value ?? "";
        });
      lines.push(`[/${sheet.name}/row[${rowIndex}]] ${ordered.join("\t")}`);
    }
  }
  return lines.join("\n").trimEnd();
}

function renderAnnotatedView(state: ExcelWorkbookState) {
  const lines: string[] = [];
  for (const sheet of state.sheets) {
    lines.push(`=== Sheet: ${sheet.name} ===`);
    for (const ref of Object.keys(sheet.cells).sort(compareCellRefs)) {
      const cell = sheet.cells[ref];
      const materialized = materializeCellNode(sheet, ref) as { value?: string | null; evaluatedValue?: string };
      const value = materialized.evaluatedValue ?? materialized.value ?? "";
      const annotation = cell.formula ? `=${cell.formula}` : cell.type ?? "number";
      const warn = !cell.value && !cell.formula ? " empty" : cell.formula && value === "" ? " unevaluated-formula" : "";
      lines.push(`  ${ref}: [${value}] <- ${annotation}${warn ? ` !${warn}` : ""}`);
    }
  }
  return lines.join("\n").trimEnd();
}

function renderStatsView(state: ExcelWorkbookState) {
  let totalCells = 0;
  let emptyCells = 0;
  let formulaCells = 0;
  let errorCells = 0;
  const typeCounts = new Map<string, number>();
  for (const sheet of state.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      totalCells += 1;
      if (!cell.value) emptyCells += 1;
      if (cell.formula) formulaCells += 1;
      if (/^#(REF|VALUE|NAME\?|DIV\/0!)/.test(cell.value)) errorCells += 1;
      const type = cell.type ?? (cell.formula ? "formula" : "number");
      typeCounts.set(type, (typeCounts.get(type) ?? 0) + 1);
    }
  }
  const lines = [
    `Sheets: ${state.sheets.length}`,
    `Total Cells: ${totalCells}`,
    `Empty Cells: ${emptyCells}`,
    `Formula Cells: ${formulaCells}`,
    `Error Cells: ${errorCells}`,
    "",
    "Data Type Distribution:",
  ];
  for (const [type, count] of [...typeCounts.entries()].sort((a, b) => b[1] - a[1])) {
    lines.push(`  ${type}: ${count}`);
  }
  return lines.join("\n").trimEnd();
}

function renderIssuesView(state: ExcelWorkbookState) {
  const issues: string[] = [];
  for (const sheet of state.sheets) {
    const refs = Object.keys(sheet.cells);
    const duplicateRefs = refs.filter((ref, index) => refs.indexOf(ref) !== index);
    duplicateRefs.forEach((ref) => issues.push(`/${sheet.name}/${ref}: duplicate cell reference`));
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      if (cell.formula && !/^[A-Z0-9_+\-*/(),:\s.]+$/i.test(cell.formula)) {
        issues.push(`/${sheet.name}/${ref}: formula contains unsupported tokens`);
      }
      if (/^#(REF|VALUE|NAME\?|DIV\/0!)/.test(cell.value)) {
        issues.push(`/${sheet.name}/${ref}: formula display error '${cell.value}'`);
      }
    }
  }
  return issues.length > 0 ? issues.join("\n") : "No issues found.";
}

function renderHtmlView(state: ExcelWorkbookState) {
  return renderExcelHtmlFromRoot(materializeWorkbookRoot(state));
}

function applySheetProperties(sheet: ExcelSheetModel, props: Record<string, string>) {
  if (props.freeze !== undefined) {
    sheet.freezeTopLeftCell = props.freeze;
  }
  if (props.zoom !== undefined) {
    sheet.zoom = Number(props.zoom);
  }
  if (props.gridlines !== undefined) {
    sheet.showGridLines = isTruthy(props.gridlines);
  }
  if (props.headings !== undefined) {
    sheet.showHeadings = isTruthy(props.headings);
  }
  if (props.tabColor !== undefined || props.tabcolor !== undefined) {
    sheet.tabColor = normalizeArgbColor(props.tabColor ?? props.tabcolor ?? "");
  }
  if (props.autoFilter !== undefined || props.autofilter !== undefined) {
    sheet.autoFilter = props.autoFilter ?? props.autofilter;
  }
  if (props.orientation !== undefined) {
    sheet.orientation = props.orientation.toLowerCase();
  }
  if (props.paperSize !== undefined || props.papersize !== undefined) {
    sheet.paperSize = Number(props.paperSize ?? props.papersize);
  }
  if (props.fitToPage !== undefined || props.fittopage !== undefined) {
    sheet.fitToPage = props.fitToPage ?? props.fittopage;
  }
  if (props.header !== undefined) {
    sheet.header = props.header;
  }
  if (props.footer !== undefined) {
    sheet.footer = props.footer;
  }
  if (props.protect !== undefined || props.protection !== undefined) {
    sheet.protection = isTruthy(props.protect ?? props.protection ?? "false");
  }
  if (props.rowBreaks !== undefined || props.rowbreaks !== undefined) {
    sheet.rowBreaks = parseBreakList(props.rowBreaks ?? props.rowbreaks ?? "");
  }
  if (props.colBreaks !== undefined || props.colbreaks !== undefined) {
    sheet.colBreaks = parseBreakList(props.colBreaks ?? props.colbreaks ?? "");
  }
  updateSheetXml(sheet);
}

function mergeValidation(existing: ExcelValidationModel, props: Record<string, string>): ExcelValidationModel {
  const next = { ...existing };
  for (const [key, value] of Object.entries(props)) {
    switch (key.toLowerCase()) {
      case "sqref":
        next.sqref = value;
        break;
      case "type":
        next.type = value;
        break;
      case "formula1":
        next.formula1 = value;
        break;
      case "formula2":
        next.formula2 = value;
        break;
      case "operator":
        next.operator = value;
        break;
      case "allowblank":
        next.allowBlank = isTruthy(value);
        break;
      case "showerror":
        next.showError = isTruthy(value);
        break;
      case "errortitle":
        next.errorTitle = value;
        break;
      case "error":
        next.error = value;
        break;
      case "showinput":
        next.showInput = isTruthy(value);
        break;
      case "prompttitle":
        next.promptTitle = value;
        break;
      case "prompt":
        next.prompt = value;
        break;
      default:
        throw new UsageError(`Unsupported validation property '${key}'.`);
    }
  }
  return next;
}

function addValidation(sheet: ExcelSheetModel, props: Record<string, string>) {
  const sqref = props.sqref ?? props.ref;
  if (!sqref) {
    throw new UsageError("Excel validation requires --prop sqref or --prop ref.");
  }
  const base: ExcelValidationModel = {
    sqref,
    allowBlank: props.allowBlank === undefined && props.allowblank === undefined ? true : undefined,
    showError: props.showError === undefined && props.showerror === undefined ? true : undefined,
    showInput: props.showInput === undefined && props.showinput === undefined ? true : undefined,
  };
  const validation = mergeValidation(
    base,
    Object.fromEntries(Object.entries(props).filter(([key]) => !["sqref", "ref"].includes(key.toLowerCase()))),
  );
  const validations = parseValidations(sheet.xml);
  validations.push(validation);
  sheet.xml = replaceSheetValidations(sheet.xml, validations);
  return validation;
}

function addComment(state: ExcelWorkbookState, sheet: ExcelSheetModel, props: Record<string, string>) {
  const ref = (props.ref ?? props.cell ?? "").toUpperCase();
  if (!ref) {
    throw new UsageError("Excel comment requires --prop ref=A1 or --prop cell=A1.");
  }
  const nextComment: ExcelCommentModel = {
    ref,
    text: props.text ?? props.value ?? "",
    author: props.author ?? "officekit",
  };
  let commentsPath = resolveCommentsPath(state, sheet);
  if (!commentsPath) {
    commentsPath = nextIndexedPartPath(state.zip, "xl/comments", ".xml");
    state.zip.set(commentsPath, Buffer.from(renderCommentsXml([]), "utf8"));
    appendRelationship(
      state.zip,
      getRelationshipsEntryName(sheet.entryName),
      sheet.entryName,
      commentsPath,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    );
  }
  const comments = getSheetComments(state, sheet);
  comments.push(nextComment);
  state.zip.set(commentsPath, Buffer.from(renderCommentsXml(comments), "utf8"));
  return nextComment;
}

function addTable(
  state: ExcelWorkbookState,
  sheet: ExcelSheetModel,
  props: Record<string, string>,
) {
  const ref = (props.ref ?? props.range ?? "").toUpperCase();
  if (!ref) {
    throw new UsageError("Excel table requires --prop ref=A1:B5 or --prop range=A1:B5.");
  }
  const [startRef, endRef = startRef] = ref.split(":");
  const startAddress = parseCellAddress(startRef);
  const endAddress = parseCellAddress(endRef);
  const columnCount = Math.max(1, columnNameToIndex(endAddress.column) - columnNameToIndex(startAddress.column) + 1);
  const existingIds = [...state.zip.keys()]
    .filter((name) => /^xl\/tables\/table\d+\.xml$/i.test(name))
    .map((name) => Number(/table(\d+)\.xml$/i.exec(name)?.[1] ?? "0"));
  const tableId = (existingIds.length > 0 ? Math.max(...existingIds) : 0) + 1;
  const tableEntry = nextIndexedPartPath(state.zip, "xl/tables/table", ".xml");
  const name = props.name ?? `Table${tableId}`;
  const displayName = props.displayName ?? props.displayname ?? name;
  const styleName = props.style ?? props.stylename ?? "TableStyleMedium2";
  const headerRow = props.headerRow === undefined && props.headerrow === undefined ? true : isTruthy(props.headerRow ?? props.headerrow ?? "false");
  const totalsRow = isTruthy(props.totalRow ?? props.totalrow ?? props.totalsrow ?? "false");
  const columnNames = resolveTableColumnNames(sheet, props, startAddress, columnCount, headerRow);
  state.zip.set(
    tableEntry,
    Buffer.from(renderTableXml({ id: tableId, name, displayName, ref, styleName, headerRow, totalsRow, columnNames }), "utf8"),
  );
  sheet.xml = ensureWorksheetNamespaces(sheet.xml, {
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  });
  const relId = appendRelationship(
    state.zip,
    getRelationshipsEntryName(sheet.entryName),
    sheet.entryName,
    tableEntry,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
  );
  sheet.xml = upsertTablePartReference(sheet.xml, relId);
  return {
    name,
    displayName,
    ref,
    headerRow,
    totalsRow,
    styleName,
    path: path.posix.relative(path.posix.dirname(sheet.entryName), tableEntry),
  };
}

function addSparkline(sheet: ExcelSheetModel, props: Record<string, string>) {
  const location = (props.location ?? props.cell ?? props.ref ?? "").toUpperCase();
  if (!location) {
    throw new UsageError("Excel sparkline requires --prop cell=C2 or --prop location=C2.");
  }
  const rawSourceRange = props.sourceRange ?? props.sourcerange ?? props.range ?? props.data;
  if (!rawSourceRange) {
    throw new UsageError("Excel sparkline requires --prop range=A1:B1 or --prop sourceRange=A1:B1.");
  }
  const sourceRange = rawSourceRange.includes("!") ? rawSourceRange : `${sheet.name}!${rawSourceRange}`;
  const type = (props.type ?? "line").toLowerCase();
  const sparklineXml = `<x14:sparklineGroup${type !== "line" ? ` type="${escapeXml(type)}"` : ""}><x14:sparklines><x14:sparkline><xm:f>${escapeXml(sourceRange)}</xm:f><xm:sqref>${escapeXml(location)}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>`;
  sheet.xml = ensureWorksheetNamespaces(sheet.xml, {
    x14: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    xm: "http://schemas.microsoft.com/office/excel/2006/main",
  });
  if (/<x14:sparklineGroups>/.test(sheet.xml)) {
    sheet.xml = sheet.xml.replace(/<\/x14:sparklineGroups>/, `${sparklineXml}</x14:sparklineGroups>`);
  } else if (/<(?:\w+:)?extLst>/.test(sheet.xml)) {
    sheet.xml = sheet.xml.replace(
      /<\/(?:\w+:)?extLst>/,
      `<ext uri="{05C60535-1F16-4fd2-B633-E4A46CF9E463}"><x14:sparklineGroups>${sparklineXml}</x14:sparklineGroups></ext></extLst>`,
    );
  } else {
    sheet.xml = sheet.xml.replace(
      /<\/(?:\w+:)?worksheet>/,
      `<extLst><ext uri="{05C60535-1F16-4fd2-B633-E4A46CF9E463}"><x14:sparklineGroups>${sparklineXml}</x14:sparklineGroups></ext></extLst></worksheet>`,
    );
  }
  return parseSparklines(sheet.xml).at(-1)!;
}

function setComment(state: ExcelWorkbookState, sheet: ExcelSheetModel, index: number, props: Record<string, string>) {
  const commentsPath = resolveCommentsPath(state, sheet);
  if (!commentsPath) {
    throw new OfficekitError(`Comment ${index} does not exist.`, "not_found");
  }
  const commentsXml = requireEntry(state.zip, commentsPath);
  const comments = getSheetComments(state, sheet);
  const comment = comments[index - 1];
  if (!comment) {
    throw new OfficekitError(`Comment ${index} does not exist.`, "not_found");
  }
  const next = {
    ...comment,
    ...(props.ref ? { ref: props.ref } : {}),
    ...(props.author ? { author: props.author } : {}),
    ...(props.text || props.value ? { text: props.text ?? props.value ?? comment.text } : {}),
  };
  const rebuilt = renderCommentsXml(comments.map((item, itemIndex) => itemIndex === index - 1 ? next : item));
  state.zip.set(commentsPath, Buffer.from(rebuilt, "utf8"));
  return next;
}

function setTable(state: ExcelWorkbookState, sheet: ExcelSheetModel, index: number, props: Record<string, string>) {
  const tables = getSheetTables(state, sheet);
  const table = tables[index - 1];
  if (!table) {
    throw new OfficekitError(`Table ${index} does not exist.`, "not_found");
  }
  const next = {
    ...table,
    ...(props.name ? { name: props.name } : {}),
    ...(props.displayname ? { displayName: props.displayname } : {}),
    ...(props.ref ? { ref: props.ref.toUpperCase() } : {}),
    ...(props.headerrow !== undefined ? { headerRow: isTruthy(props.headerrow) } : {}),
    ...(props.totalsrow !== undefined ? { totalsRow: isTruthy(props.totalsrow) } : {}),
    ...(props.style || props.stylename ? { styleName: props.style ?? props.stylename } : {}),
  };
  const xmlPath = normalizeZipPath(path.posix.dirname(sheet.entryName), table.path);
  const original = requireEntry(state.zip, xmlPath);
  const updated = original
    .replace(/\bname="[^"]*"/, next.name ? `name="${escapeXml(next.name)}"` : 'name="Table1"')
    .replace(/\bdisplayName="[^"]*"/, next.displayName ? `displayName="${escapeXml(next.displayName)}"` : 'displayName="Table1"')
    .replace(/\bref="[^"]*"/, next.ref ? `ref="${escapeXml(next.ref)}"` : 'ref="A1:A1"');
  const withCounts = updated
    .replace(/\bheaderRowCount="[^"]*"/, `headerRowCount="${next.headerRow === false ? 0 : 1}"`)
    .replace(/\btotalsRowShown="[^"]*"/, `totalsRowShown="${next.totalsRow ? 1 : 0}"`);
  state.zip.set(xmlPath, Buffer.from(withCounts, "utf8"));
  return next;
}

function setSparkline(sheet: ExcelSheetModel, index: number, props: Record<string, string>) {
  const sparklines = [...sheet.xml.matchAll(/<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g)];
  let sparklineIndex = 0;
  let updatedSparkline: ExcelSparklineModel | undefined;
  for (const groupMatch of sparklines) {
    const fullGroup = groupMatch[0];
    const groupAttrs = groupMatch[1];
    const sparklineBlocks = [...fullGroup.matchAll(/<x14:sparkline\b[\s\S]*?<\/x14:sparkline>/g)];
    for (const sparklineBlock of sparklineBlocks) {
      sparklineIndex += 1;
      if (sparklineIndex !== index) {
        continue;
      }
      let nextGroup = fullGroup;
      let nextSparkline = sparklineBlock[0];
      if (props.location !== undefined) {
        nextSparkline = nextSparkline.replace(/<xm:sqref>[\s\S]*?<\/xm:sqref>/, `<xm:sqref>${escapeXml(props.location)}</xm:sqref>`);
      }
      if (props.sourceRange !== undefined || props.range !== undefined) {
        nextSparkline = nextSparkline.replace(/<xm:f>[\s\S]*?<\/xm:f>/, `<xm:f>${escapeXml(props.sourceRange ?? props.range ?? "")}</xm:f>`);
      }
      if (props.type !== undefined) {
        const nextType = props.type.toLowerCase();
        if (/type="[^"]+"/.test(nextGroup)) {
          nextGroup = nextGroup.replace(/type="[^"]+"/, `type="${escapeXml(nextType)}"`);
        } else {
          nextGroup = nextGroup.replace(/<x14:sparklineGroup\b/, `<x14:sparklineGroup type="${escapeXml(nextType)}"`);
        }
      }
      nextGroup = nextGroup.replace(sparklineBlock[0], nextSparkline);
      sheet.xml = sheet.xml.replace(fullGroup, nextGroup);
      updatedSparkline = {
        ...parseSparklines(sheet.xml)[index - 1],
      };
      return updatedSparkline;
    }
  }
  throw new OfficekitError(`Sparkline ${index} does not exist.`, "not_found");
}

function setChart(
  state: ExcelWorkbookState,
  sheet: ExcelSheetModel,
  index: number,
  seriesIndex: number | undefined,
  props: Record<string, string>,
) {
  const chart = getSheetCharts(state, sheet)[index - 1];
  if (!chart) {
    throw new OfficekitError(`Chart ${index} does not exist.`, "not_found");
  }
  const xmlPath = normalizeZipPath(path.posix.dirname(resolveDrawingPath(state, sheet) ?? sheet.entryName), chart.path);
  let xml = requireEntry(state.zip, xmlPath);
  if (seriesIndex !== undefined) {
    const seriesMatches = [...xml.matchAll(/<c:ser\b[\s\S]*?<\/c:ser>/g)];
    const series = seriesMatches[seriesIndex - 1];
    if (!series) {
      throw new OfficekitError(`Chart series ${seriesIndex} does not exist.`, "not_found");
    }
    let nextSeries = series[0];
    if (props.name !== undefined || props.title !== undefined || props.text !== undefined) {
      const value = props.name ?? props.title ?? props.text ?? "";
      if (/<c:tx>[\s\S]*?<\/c:tx>/.test(nextSeries)) {
        nextSeries = nextSeries.replace(/<c:tx>[\s\S]*?<\/c:tx>/, `<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>${escapeXml(value)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`);
      } else {
        nextSeries = nextSeries.replace(/<c:idx\b[^>]*\/>/, `$&<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>${escapeXml(value)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`);
      }
    }
    xml = xml.replace(series[0], nextSeries);
  } else {
    if (props.title !== undefined || props.name !== undefined || props.text !== undefined) {
      const value = props.title ?? props.name ?? props.text ?? "";
      if (/<c:title>[\s\S]*?<\/c:title>/.test(xml)) {
        xml = xml.replace(/<c:title>[\s\S]*?<\/c:title>/, `<c:title><c:tx><c:rich><a:p><a:r><a:t>${escapeXml(value)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`);
      } else {
        xml = xml.replace(/<c:chart\b[^>]*>/, `$&<c:title><c:tx><c:rich><a:p><a:r><a:t>${escapeXml(value)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`);
      }
    }
    if (props.legend !== undefined) {
      const normalized = props.legend.toLowerCase();
      if (normalized === "false" || normalized === "none") {
        xml = xml.replace(/<c:legend\b[\s\S]*?<\/c:legend>/, "");
      } else {
        const legendPositionMap: Record<string, string> = { top: "t", left: "l", right: "r", bottom: "b" };
        const legendXml = `<c:legend><c:legendPos val="${legendPositionMap[normalized] ?? "b"}"/><c:overlay val="0"/></c:legend>`;
        if (/<c:legend\b[\s\S]*?<\/c:legend>/.test(xml)) {
          xml = xml.replace(/<c:legend\b[\s\S]*?<\/c:legend>/, legendXml);
        } else {
          xml = xml.replace(/<\/c:plotArea>/, `</c:plotArea>${legendXml}`);
        }
      }
    }
    if (props.datalabels !== undefined || props.labels !== undefined) {
      const value = (props.datalabels ?? props.labels ?? "").toLowerCase();
      const dataLabelsXml = value === "none"
        ? ""
        : `<c:dLbls><c:showLegendKey val="0"/><c:showValue val="${value.includes("value") || value === "true" ? 1 : 0}"/><c:showCategoryName val="${value.includes("category") ? 1 : 0}"/><c:showSeriesName val="${value.includes("series") ? 1 : 0}"/><c:showPercent val="${value.includes("percent") ? 1 : 0}"/></c:dLbls>`;
      if (/<c:dLbls\b[\s\S]*?<\/c:dLbls>/.test(xml)) {
        xml = xml.replace(/<c:dLbls\b[\s\S]*?<\/c:dLbls>/, dataLabelsXml);
      } else if (dataLabelsXml) {
        xml = xml.replace(/(<c:(?:barChart|lineChart|pieChart|areaChart)\b[\s\S]*?<c:ser\b[\s\S]*?<\/c:ser>)/, `$1${dataLabelsXml}`);
      }
    }
    if (props.categoryAxisTitle !== undefined || props.cataxistitle !== undefined) {
      xml = setAxisTitle(xml, "catAx", props.categoryAxisTitle ?? props.cataxistitle ?? "");
    }
    if (props.valueAxisTitle !== undefined || props.valaxistitle !== undefined) {
      xml = setAxisTitle(xml, "valAx", props.valueAxisTitle ?? props.valaxistitle ?? "");
    }
  }
  state.zip.set(xmlPath, Buffer.from(xml, "utf8"));
  return seriesIndex !== undefined
    ? { ...getSheetCharts(state, sheet)[index - 1], series: { index: seriesIndex } }
    : getSheetCharts(state, sheet)[index - 1];
}

function setPivotTable(state: ExcelWorkbookState, sheet: ExcelSheetModel, index: number, props: Record<string, string>) {
  const pivot = getSheetPivots(state, sheet)[index - 1];
  if (!pivot) {
    throw new OfficekitError(`Pivot table ${index} does not exist.`, "not_found");
  }
  const xmlPath = normalizeZipPath(path.posix.dirname(sheet.entryName), pivot.path);
  let xml = requireEntry(state.zip, xmlPath);
  if (props.name !== undefined) {
    if (/\bname="[^"]+"/.test(xml)) {
      xml = xml.replace(/\bname="[^"]+"/, `name="${escapeXml(props.name)}"`);
    } else {
      xml = xml.replace(/<pivotTableDefinition\b/, `<pivotTableDefinition name="${escapeXml(props.name)}"`);
    }
  }
  const booleanPivotProps: Array<[string, string]> = [
    ["rowGrandTotals", "rowGrandTotals"],
    ["colGrandTotals", "colGrandTotals"],
    ["compact", "compact"],
    ["compactData", "compactData"],
    ["outline", "outline"],
  ];
  for (const [propKey, attrName] of booleanPivotProps) {
    const rawValue = props[propKey] ?? props[propKey.toLowerCase()];
    if (rawValue === undefined) continue;
    const attrValue = isTruthy(rawValue) ? "1" : "0";
    if (new RegExp(`\\b${attrName}="[^"]+"`).test(xml)) {
      xml = xml.replace(new RegExp(`\\b${attrName}="[^"]+"`), `${attrName}="${attrValue}"`);
    } else {
      xml = xml.replace(/<pivotTableDefinition\b/, `<pivotTableDefinition ${attrName}="${attrValue}"`);
    }
  }
  state.zip.set(xmlPath, Buffer.from(xml, "utf8"));
  return {
    ...pivot,
    ...(props.name !== undefined ? { name: props.name } : {}),
    ...Object.fromEntries(
      booleanPivotProps
        .filter(([propKey]) => props[propKey] !== undefined || props[propKey.toLowerCase()] !== undefined)
        .map(([propKey]) => [propKey, isTruthy(props[propKey] ?? props[propKey.toLowerCase()] ?? "false")]),
    ),
  };
}

function setAxisTitle(xml: string, axisTag: "catAx" | "valAx", title: string) {
  const axisPattern = new RegExp(`<c:${axisTag}\\b([\\s\\S]*?)<\\/c:${axisTag}>`);
  const axisMatch = axisPattern.exec(xml);
  if (!axisMatch) {
    return xml;
  }
  const axisXml = axisMatch[0];
  const titleXml = title
    ? `<c:title><c:tx><c:rich><a:p><a:r><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`
    : "";
  const nextAxisXml = /<c:title>[\s\S]*?<\/c:title>/.test(axisXml)
    ? axisXml.replace(/<c:title>[\s\S]*?<\/c:title>/, titleXml)
    : axisXml.replace(/>/, `>${titleXml}`);
  return xml.replace(axisXml, nextAxisXml);
}

function setDrawingObject(
  state: ExcelWorkbookState,
  sheet: ExcelSheetModel,
  kind: "shape" | "picture",
  index: number,
  props: Record<string, string>,
) {
  const drawingPath = resolveDrawingPath(state, sheet);
  if (!drawingPath) {
    throw new OfficekitError(`Sheet '${sheet.name}' has no drawing part.`, "not_found");
  }
  const xml = requireEntry(state.zip, drawingPath);
  const block = getDrawingAnchorBlocks(xml, kind)[index - 1];
  if (!block) {
    throw new OfficekitError(`${kind} ${index} does not exist.`, "not_found");
  }
  let nextBlock = block;
  if (props.name !== undefined) {
    nextBlock = nextBlock.replace(/name="[^"]*"/, `name="${escapeXml(props.name)}"`);
  }
  if (props.alt !== undefined || props.description !== undefined) {
    const description = props.alt ?? props.description ?? "";
    if (/descr="[^"]*"/.test(nextBlock)) {
      nextBlock = nextBlock.replace(/descr="[^"]*"/, `descr="${escapeXml(description)}"`);
    } else {
      nextBlock = nextBlock.replace(/<xdr:cNvPr\b/, `<xdr:cNvPr descr="${escapeXml(description)}"`);
    }
  }
  if (kind === "shape" && (props.text !== undefined || props.value !== undefined)) {
    const text = props.text ?? props.value ?? "";
    if (/<a:t>[\s\S]*?<\/a:t>/.test(nextBlock)) {
      nextBlock = nextBlock.replace(/<a:t>[\s\S]*?<\/a:t>/, `<a:t>${escapeXml(text)}</a:t>`);
    }
  }
  const anchorReplacements: Array<[string, string, string]> = [
    ["x", "<xdr:col>", "</xdr:col>"],
    ["y", "<xdr:row>", "</xdr:row>"],
    ["width", "<xdr:col>", "</xdr:col>"],
    ["height", "<xdr:row>", "</xdr:row>"],
  ];
  for (const [propKey, openTag, closeTag] of anchorReplacements) {
    if (props[propKey] === undefined) continue;
    const numeric = Number(props[propKey]);
    if (Number.isNaN(numeric)) continue;
    if (propKey === "x" || propKey === "y") {
      const tag = propKey === "x" ? "col" : "row";
      nextBlock = nextBlock.replace(new RegExp(`(<xdr:from>[\\s\\S]*?<xdr:${tag}>)([\\s\\S]*?)(<\\/xdr:${tag}>)`), `$1${numeric}$3`);
    } else {
      const tag = propKey === "width" ? "col" : "row";
      nextBlock = nextBlock.replace(new RegExp(`(<xdr:to>[\\s\\S]*?<xdr:${tag}>)([\\s\\S]*?)(<\\/xdr:${tag}>)`), `$1${numeric}$3`);
    }
  }
  state.zip.set(drawingPath, Buffer.from(xml.replace(block, nextBlock), "utf8"));
  const updated = getDrawingShapes(state, sheet).filter((item) => item.kind === kind)[index - 1];
  return updated;
}

function removeComment(state: ExcelWorkbookState, sheet: ExcelSheetModel, index: number) {
  const commentsPath = resolveCommentsPath(state, sheet);
  if (!commentsPath) {
    throw new OfficekitError(`Comment ${index} does not exist.`, "not_found");
  }
  const comments = getSheetComments(state, sheet);
  if (!comments[index - 1]) {
    throw new OfficekitError(`Comment ${index} does not exist.`, "not_found");
  }
  state.zip.set(commentsPath, Buffer.from(renderCommentsXml(comments.filter((_, itemIndex) => itemIndex !== index - 1)), "utf8"));
}

function removeTable(state: ExcelWorkbookState, sheet: ExcelSheetModel, index: number) {
  const tables = getSheetTables(state, sheet);
  const table = tables[index - 1];
  if (!table) {
    throw new OfficekitError(`Table ${index} does not exist.`, "not_found");
  }
  const xmlPath = normalizeZipPath(path.posix.dirname(sheet.entryName), table.path);
  state.zip.delete(xmlPath);
  const relsPath = getRelationshipsEntryName(sheet.entryName);
  const relsXml = requireEntry(state.zip, relsPath);
  state.zip.set(relsPath, Buffer.from(relsXml.replace(new RegExp(`<Relationship\\b[^>]*Target="${escapeXmlForRegex(table.path)}"[^>]*/>`, "g"), ""), "utf8"));
}

function removeSparkline(sheet: ExcelSheetModel, index: number) {
  const groups = [...sheet.xml.matchAll(/<x14:sparklineGroup\b[\s\S]*?<\/x14:sparklineGroup>/g)];
  let current = 0;
  for (const group of groups) {
    const sparklineBlocks = [...group[0].matchAll(/<x14:sparkline\b[\s\S]*?<\/x14:sparkline>/g)];
    for (const sparklineBlock of sparklineBlocks) {
      current += 1;
      if (current !== index) continue;
      const nextGroup = group[0].replace(sparklineBlock[0], "");
      sheet.xml = sheet.xml.replace(group[0], /<x14:sparkline\b/.test(nextGroup) ? nextGroup : "");
      return;
    }
  }
  throw new OfficekitError(`Sparkline ${index} does not exist.`, "not_found");
}

function removeDrawingObject(state: ExcelWorkbookState, sheet: ExcelSheetModel, kind: "shape" | "picture", index: number) {
  const drawingPath = resolveDrawingPath(state, sheet);
  if (!drawingPath) {
    throw new OfficekitError(`Sheet '${sheet.name}' has no drawing part.`, "not_found");
  }
  const xml = requireEntry(state.zip, drawingPath);
  const block = getDrawingAnchorBlocks(xml, kind)[index - 1];
  if (!block) {
    throw new OfficekitError(`${kind} ${index} does not exist.`, "not_found");
  }
  state.zip.set(drawingPath, Buffer.from(xml.replace(block, ""), "utf8"));
}

function getDrawingAnchorBlocks(xml: string, kind: "shape" | "picture") {
  const anchors = [...xml.matchAll(/<xdr:twoCellAnchor\b[\s\S]*?<\/xdr:twoCellAnchor>/g)].map((match) => match[0]);
  return anchors.filter((anchor) => kind === "picture" ? /<xdr:pic\b/.test(anchor) : /<xdr:sp\b/.test(anchor));
}

function normalizeSheetPath(targetPath: string) {
  return targetPath.replace(/^\//, "") || "Sheet1";
}

function collectRowCells(sheet: ExcelSheetModel, rowIndex: number) {
  return Object.keys(sheet.cells)
    .filter((ref) => Number(/\d+/.exec(ref)?.[0] ?? "0") === rowIndex)
    .sort(compareCellRefs)
    .map((ref) => materializeCellNode(sheet, ref));
}

function collectColumnCells(sheet: ExcelSheetModel, column: string) {
  const normalized = /^\d+$/.test(column) ? indexToColumnName(Number(column)) : column.toUpperCase();
  return Object.keys(sheet.cells)
    .filter((ref) => /^[A-Z]+/.exec(ref)?.[0] === normalized)
    .sort(compareCellRefs)
    .map((ref) => materializeCellNode(sheet, ref));
}

function nextAvailableRowIndex(sheet: ExcelSheetModel) {
  const refs = Object.keys(sheet.cells);
  if (refs.length === 0) return 1;
  return Math.max(...refs.map((ref) => Number(/\d+/.exec(ref)?.[0] ?? "0"))) + 1;
}

function parseSheetCells(xml: string, zip: Map<string, Buffer>) {
  const sharedStrings = parseSharedStrings(zip);
  const cells: Record<string, ExcelCellModel> = {};
  for (const match of xml.matchAll(/<(?:\w+:)?c\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>/g)) {
    const attributes = match[1];
    const body = match[2];
    const ref = /r="([^"]+)"/.exec(attributes)?.[1]?.toUpperCase();
    if (!ref) continue;
    const styleId = /s="([^"]+)"/.exec(attributes)?.[1];
    const type = /t="([^"]+)"/.exec(attributes)?.[1] ?? "";
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
  const pane = /<(?:\w+:)?pane\b([^>]*)\/?>/.exec(xml)?.[1];
  const sheetView = /<(?:\w+:)?sheetView\b([^>]*)>/.exec(xml)?.[1];
  const topLeftCell = /topLeftCell="([^"]+)"/.exec(pane ?? "")?.[1];
  const zoom = /zoomScale="([^"]+)"/.exec(sheetView ?? "")?.[1];
  const showGridLines = /showGridLines="([^"]+)"/.exec(sheetView ?? "")?.[1];
  const showHeadings = /showRowColHeaders="([^"]+)"/.exec(sheetView ?? "")?.[1];
  const tabColor = /<(?:\w+:)?tabColor\b[^>]*rgb="([^"]+)"/.exec(xml)?.[1];
  const orientation = /<(?:\w+:)?pageSetup\b[^>]*orientation="([^"]+)"/.exec(xml)?.[1];
  const paperSize = /<(?:\w+:)?pageSetup\b[^>]*paperSize="([^"]+)"/.exec(xml)?.[1];
  const fitToWidth = /<(?:\w+:)?pageSetup\b[^>]*fitToWidth="([^"]+)"/.exec(xml)?.[1];
  const fitToHeight = /<(?:\w+:)?pageSetup\b[^>]*fitToHeight="([^"]+)"/.exec(xml)?.[1];
  const header = /<(?:\w+:)?oddHeader>([\s\S]*?)<\/(?:\w+:)?oddHeader>/.exec(xml)?.[1];
  const footer = /<(?:\w+:)?oddFooter>([\s\S]*?)<\/(?:\w+:)?oddFooter>/.exec(xml)?.[1];
  const protection = /<(?:\w+:)?sheetProtection\b[^>]*sheet="([^"]+)"/.exec(xml)?.[1];
  const rowBreaks = [...xml.matchAll(/<(?:\w+:)?rowBreaks\b[\s\S]*?<brk\b[^>]*id="([^"]+)"/g)].map((match) => Number(match[1]));
  const colBreaks = [...xml.matchAll(/<(?:\w+:)?colBreaks\b[\s\S]*?<brk\b[^>]*id="([^"]+)"/g)].map((match) => Number(match[1]));
  return {
    ...(autoFilter ? { autoFilter } : {}),
    ...(topLeftCell ? { freezeTopLeftCell: topLeftCell } : {}),
    ...(zoom ? { zoom: Number(zoom) } : {}),
    ...(showGridLines !== undefined ? { showGridLines: isTruthy(showGridLines) } : {}),
    ...(showHeadings !== undefined ? { showHeadings: isTruthy(showHeadings) } : {}),
    ...(tabColor ? { tabColor } : {}),
    ...(orientation ? { orientation } : {}),
    ...(paperSize ? { paperSize: Number(paperSize) } : {}),
    ...(fitToWidth || fitToHeight ? { fitToPage: `${fitToWidth ?? "1"}x${fitToHeight ?? "1"}` } : {}),
    ...(header ? { header: decodeXml(header) } : {}),
    ...(footer ? { footer: decodeXml(footer) } : {}),
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

function parseDefinedNames(workbookXml: string, sheets: Array<{ name: string }>) {
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

function parseValidations(sheetXml: string): ExcelValidationModel[] {
  return [...sheetXml.matchAll(/<(?:\w+:)?dataValidation\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?dataValidation>/g)].map((match) => {
    const attrs = match[1];
    const body = match[2];
    return {
      ...(parseAttr(attrs, "sqref") ? { sqref: parseAttr(attrs, "sqref") } : {}),
      ...(parseAttr(attrs, "type") ? { type: parseAttr(attrs, "type") } : {}),
      ...(parseAttr(attrs, "operator") ? { operator: parseAttr(attrs, "operator") } : {}),
      ...(parseAttr(attrs, "allowBlank") !== undefined ? { allowBlank: isTruthy(parseAttr(attrs, "allowBlank") ?? "false") } : {}),
      ...(parseAttr(attrs, "showErrorMessage") !== undefined ? { showError: isTruthy(parseAttr(attrs, "showErrorMessage") ?? "false") } : {}),
      ...(parseAttr(attrs, "errorTitle") ? { errorTitle: parseAttr(attrs, "errorTitle") } : {}),
      ...(parseAttr(attrs, "error") ? { error: parseAttr(attrs, "error") } : {}),
      ...(parseAttr(attrs, "showInputMessage") !== undefined ? { showInput: isTruthy(parseAttr(attrs, "showInputMessage") ?? "false") } : {}),
      ...(parseAttr(attrs, "promptTitle") ? { promptTitle: parseAttr(attrs, "promptTitle") } : {}),
      ...(parseAttr(attrs, "prompt") ? { prompt: parseAttr(attrs, "prompt") } : {}),
      ...(extractTagText(body, "formula1") ? { formula1: extractTagText(body, "formula1") } : {}),
      ...(extractTagText(body, "formula2") ? { formula2: extractTagText(body, "formula2") } : {}),
    };
  });
}

function replaceSheetValidations(sheetXml: string, validations: ExcelValidationModel[]) {
  const rendered = validations.length > 0
    ? `<dataValidations count="${validations.length}">${validations.map(renderValidationXml).join("")}</dataValidations>`
    : "";
  return replaceOrInsert(sheetXml, /<(?:\w+:)?dataValidations\b[\s\S]*?<\/(?:\w+:)?dataValidations>/, rendered, /<(?:\w+:)?autoFilter\b[^>]*\/?>|<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/);
}

function renderValidationXml(validation: ExcelValidationModel) {
  const attrs = [
    validation.sqref ? `sqref="${escapeXml(validation.sqref)}"` : "",
    validation.type ? `type="${escapeXml(validation.type)}"` : "",
    validation.operator ? `operator="${escapeXml(validation.operator)}"` : "",
    validation.allowBlank !== undefined ? `allowBlank="${validation.allowBlank ? 1 : 0}"` : "",
    validation.showError !== undefined ? `showErrorMessage="${validation.showError ? 1 : 0}"` : "",
    validation.errorTitle ? `errorTitle="${escapeXml(validation.errorTitle)}"` : "",
    validation.error ? `error="${escapeXml(validation.error)}"` : "",
    validation.showInput !== undefined ? `showInputMessage="${validation.showInput ? 1 : 0}"` : "",
    validation.promptTitle ? `promptTitle="${escapeXml(validation.promptTitle)}"` : "",
    validation.prompt ? `prompt="${escapeXml(validation.prompt)}"` : "",
  ].filter(Boolean).join(" ");
  const body = [
    validation.formula1 ? `<formula1>${escapeXml(validation.formula1)}</formula1>` : "",
    validation.formula2 ? `<formula2>${escapeXml(validation.formula2)}</formula2>` : "",
  ].join("");
  return `<dataValidation ${attrs}>${body}</dataValidation>`;
}

function getSheetComments(state: ExcelWorkbookState, sheet: ExcelSheetModel): ExcelCommentModel[] {
  const commentsPath = resolveCommentsPath(state, sheet);
  if (!commentsPath) {
    return [];
  }
  const xml = requireEntry(state.zip, commentsPath);
  const authors = [...xml.matchAll(/<(?:\w+:)?author>([\s\S]*?)<\/(?:\w+:)?author>/g)].map((match) => decodeXml(match[1]));
  return [...xml.matchAll(/<(?:\w+:)?comment\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?comment>/g)].map((match) => {
    const attrs = match[1];
    const body = match[2];
    const authorId = Number(/authorId="([^"]+)"/.exec(attrs)?.[1] ?? "0");
    return {
      ref: /ref="([^"]+)"/.exec(attrs)?.[1] ?? "A1",
      author: authors[authorId],
      text: extractTexts(body).join(""),
    };
  });
}

function renderCommentsXml(comments: ExcelCommentModel[]) {
  const authors = [...new Set(comments.map((comment) => comment.author ?? "officekit"))];
  const items = comments.map((comment) => {
    const authorId = authors.indexOf(comment.author ?? "officekit");
    return `<comment ref="${escapeXml(comment.ref)}" authorId="${authorId}"><text><r><t>${escapeXml(comment.text)}</t></r></text></comment>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>${authors.map((author) => `<author>${escapeXml(author)}</author>`).join("")}</authors>
  <commentList>${items}</commentList>
</comments>`;
}

function resolveCommentsPath(state: ExcelWorkbookState, sheet: ExcelSheetModel) {
  const worksheetRelsPath = getRelationshipsEntryName(sheet.entryName);
  const rels = readRelationships(state.zip, worksheetRelsPath);
  const commentRel = rels.find((relationship) => relationship.type?.endsWith("/comments"));
  return commentRel ? normalizeZipPath(path.posix.dirname(sheet.entryName), commentRel.target) : undefined;
}

function getSheetTables(state: ExcelWorkbookState, sheet: ExcelSheetModel): Array<ExcelTableModel & { path: string }> {
  const worksheetRelsPath = getRelationshipsEntryName(sheet.entryName);
  const rels = readRelationships(state.zip, worksheetRelsPath).filter((relationship) => relationship.type?.endsWith("/table"));
  return rels.map((relationship) => {
    const entryName = normalizeZipPath(path.posix.dirname(sheet.entryName), relationship.target);
    const xml = requireEntry(state.zip, entryName);
    return {
      path: relationship.target,
      ...(parseAttr(xml, "name") ? { name: parseAttr(xml, "name") } : {}),
      ...(parseAttr(xml, "displayName") ? { displayName: parseAttr(xml, "displayName") } : {}),
      ...(parseAttr(xml, "ref") ? { ref: parseAttr(xml, "ref") } : {}),
      ...(parseAttr(xml, "headerRowCount") ? { headerRow: parseAttr(xml, "headerRowCount") !== "0" } : {}),
      ...(parseAttr(xml, "totalsRowShown") ? { totalsRow: parseAttr(xml, "totalsRowShown") === "1" } : {}),
      ...(parseAttr(/<(?:\w+:)?tableStyleInfo\b([^>]*)/.exec(xml)?.[1] ?? "", "name") ? { styleName: parseAttr(/<(?:\w+:)?tableStyleInfo\b([^>]*)/.exec(xml)?.[1] ?? "", "name") } : {}),
    };
  });
}

function getSheetCharts(state: ExcelWorkbookState, sheet: ExcelSheetModel) {
  const drawingPath = resolveDrawingPath(state, sheet);
  if (!drawingPath) {
    return [] as Array<ExcelChartModel & { path: string }>;
  }
  const drawingXml = requireEntry(state.zip, drawingPath);
  const drawingRels = readRelationships(state.zip, getRelationshipsEntryName(drawingPath));
  return [...drawingXml.matchAll(/<c:chart\b[^>]*r:id="([^"]+)"/g)].map((match) => {
    const rel = drawingRels.find((relationship) => relationship.id === match[1]);
    const chartPath = rel ? normalizeZipPath(path.posix.dirname(drawingPath), rel.target) : undefined;
    const chartXml = chartPath ? requireEntry(state.zip, chartPath) : "";
    const seriesNames = [...chartXml.matchAll(/<c:ser\b[\s\S]*?<c:tx>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>[\s\S]*?<\/c:tx>[\s\S]*?<\/c:ser>/g)].map((seriesMatch) => decodeXml(seriesMatch[1]).trim());
    const legendPos = /<c:legendPos\b[^>]*val="([^"]+)"/.exec(chartXml)?.[1];
    const dataLabels = /<c:dLbls\b[\s\S]*?<(?:c:showValue)\b[^>]*val="([^"]+)"/.exec(chartXml)?.[1];
    return {
      path: rel?.target ?? "",
      sheet: sheet.name,
      title: decodeXml(/<c:title>[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml)?.[1] ?? "").trim() || undefined,
      ...(chartXml.includes("<c:barChart") ? { chartType: "bar" } : chartXml.includes("<c:lineChart") ? { chartType: "line" } : chartXml.includes("<c:pieChart") ? { chartType: "pie" } : chartXml.includes("<c:areaChart") ? { chartType: "area" } : {}),
      ...(legendPos ? { legend: legendPos } : chartXml.includes("<c:legend") ? { legend: true } : {}),
      ...(dataLabels ? { dataLabels: isTruthy(dataLabels) ? "value" : "none" } : {}),
      ...(decodeXml(/<c:catAx\b[\s\S]*?<c:title>[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml)?.[1] ?? "").trim() ? { categoryAxisTitle: decodeXml(/<c:catAx\b[\s\S]*?<c:title>[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml)?.[1] ?? "").trim() } : {}),
      ...(decodeXml(/<c:valAx\b[\s\S]*?<c:title>[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml)?.[1] ?? "").trim() ? { valueAxisTitle: decodeXml(/<c:valAx\b[\s\S]*?<c:title>[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml)?.[1] ?? "").trim() } : {}),
      ...(seriesNames.length > 0 ? { seriesNames } : {}),
    };
  });
}

function getSheetPivots(state: ExcelWorkbookState, sheet: ExcelSheetModel) {
  const worksheetRelsPath = getRelationshipsEntryName(sheet.entryName);
  const rels = readRelationships(state.zip, worksheetRelsPath).filter((relationship) => relationship.type?.endsWith("/pivotTable"));
  return rels.map((relationship) => {
    const entryName = normalizeZipPath(path.posix.dirname(sheet.entryName), relationship.target);
    const xml = requireEntry(state.zip, entryName);
    return {
      path: relationship.target,
      name: parseAttr(xml, "name") ?? undefined,
      ...(parseAttr(xml, "rowGrandTotals") !== undefined ? { rowGrandTotals: isTruthy(parseAttr(xml, "rowGrandTotals") ?? "false") } : {}),
      ...(parseAttr(xml, "colGrandTotals") !== undefined ? { colGrandTotals: isTruthy(parseAttr(xml, "colGrandTotals") ?? "false") } : {}),
      ...(parseAttr(xml, "compact") !== undefined ? { compact: isTruthy(parseAttr(xml, "compact") ?? "false") } : {}),
      ...(parseAttr(xml, "compactData") !== undefined ? { compactData: isTruthy(parseAttr(xml, "compactData") ?? "false") } : {}),
      ...(parseAttr(xml, "outline") !== undefined ? { outline: isTruthy(parseAttr(xml, "outline") ?? "false") } : {}),
    };
  });
}

function parseSparklines(sheetXml: string) {
  return [...sheetXml.matchAll(/<x14:sparklineGroup\b([^>]*)>[\s\S]*?<x14:sparkline\b[\s\S]*?<xm:f>([\s\S]*?)<\/xm:f>[\s\S]*?<xm:sqref>([\s\S]*?)<\/xm:sqref>[\s\S]*?<\/x14:sparkline>[\s\S]*?<\/x14:sparklineGroup>/g)].map((match) => ({
    ...(parseAttr(match[1], "type") ? { type: parseAttr(match[1], "type") } : {}),
    sourceRange: decodeXml(match[2]).trim(),
    location: decodeXml(match[3]).trim(),
  }));
}

function getDrawingShapes(state: ExcelWorkbookState, sheet: ExcelSheetModel): ExcelShapeModel[] {
  const drawingPath = resolveDrawingPath(state, sheet);
  if (!drawingPath) {
    return [];
  }
  const xml = requireEntry(state.zip, drawingPath);
  const shapes = [...xml.matchAll(/<xdr:(sp|pic)\b[\s\S]*?<\/xdr:\1>/g)].map((match) => {
    const block = match[0];
    return {
      kind: match[1] === "pic" ? "picture" : "shape",
      name: /name="([^"]+)"/.exec(block)?.[1],
      text: extractTexts(block).join("").trim() || undefined,
    } satisfies ExcelShapeModel;
  });
  return shapes;
}

function parseBreaks(sheetXml: string, type: "row" | "col") {
  const tag = type === "row" ? "rowBreaks" : "colBreaks";
  return [...sheetXml.matchAll(new RegExp(`<(?:\\w+:)?${tag}\\b[\\s\\S]*?<(?:\\w+:)?brk\\b([^>]*)\\/?>`, "g"))].map((match) => ({
    id: Number(parseAttr(match[1], "id") ?? "0"),
    manual: isTruthy(parseAttr(match[1], "man") ?? "false"),
  }));
}

function resolveDrawingPath(state: ExcelWorkbookState, sheet: ExcelSheetModel) {
  const drawingRelId = /<(?:\w+:)?drawing\b[^>]*r:id="([^"]+)"/.exec(sheet.xml)?.[1];
  if (!drawingRelId) return undefined;
  const worksheetRelsPath = getRelationshipsEntryName(sheet.entryName);
  const rel = readRelationships(state.zip, worksheetRelsPath).find((relationship) => relationship.id === drawingRelId);
  return rel ? normalizeZipPath(path.posix.dirname(sheet.entryName), rel.target) : undefined;
}

function parseSharedStrings(zip: Map<string, Buffer>) {
  const shared = zip.get("xl/sharedStrings.xml");
  if (!shared) return [];
  return [...shared.toString("utf8").matchAll(/<(?:\w+:)?si\b[\s\S]*?<\/(?:\w+:)?si>/g)].map((match) => extractTexts(match[0]).join(""));
}

function parsePackageProperties(zip: Map<string, Buffer>) {
  const core = zip.get("docProps/core.xml")?.toString("utf8");
  if (!core) return {};
  return {
    ...(extractTagText(core, "dc:title") ? { title: extractTagText(core, "dc:title")! } : {}),
    ...(extractTagText(core, "dc:creator") ? { author: extractTagText(core, "dc:creator")! } : {}),
    ...(extractTagText(core, "dc:subject") ? { subject: extractTagText(core, "dc:subject")! } : {}),
    ...(extractTagText(core, "dc:description") ? { description: extractTagText(core, "dc:description")! } : {}),
  };
}

function parseOfficekitMetadata(zip: Map<string, Buffer>) {
  const metadata = zip.get(METADATA_PATH)?.toString("utf8");
  if (!metadata) return undefined;
  try {
    return JSON.parse(metadata) as ExcelWorkbookState["officekitMetadata"];
  } catch {
    return undefined;
  }
}

function overlayMetadataCells(
  parsedCells: Record<string, ExcelCellModel>,
  metadataCells?: Record<string, ExcelCellModel>,
) {
  if (!metadataCells) return parsedCells;
  const next: Record<string, ExcelCellModel> = {};
  for (const [ref, cell] of Object.entries(parsedCells)) {
    next[ref] = {
      ...cell,
      ...(metadataCells[ref]?.type ? { type: metadataCells[ref].type } : {}),
    };
  }
  for (const [ref, cell] of Object.entries(metadataCells)) {
    if (!next[ref]) {
      next[ref] = cell;
    }
  }
  return next;
}

function renderExcelCellXml(ref: string, cell: ExcelCellModel) {
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

function mergeExcelCell(existing: ExcelCellModel | undefined, props: Record<string, string>): ExcelCellModel {
  const base = normalizeExcelCell(existing);
  const formula = props.formula === undefined ? base.formula : normalizeFormula(props.formula);
  const styleId = props.styleId ?? props.style ?? base.styleId;
  const explicitType = props.type?.toLowerCase();
  const type =
    explicitType === "number" || explicitType === "boolean" || explicitType === "date" || explicitType === "string"
      ? (explicitType as ExcelCellModel["type"])
      : base.type;
  return {
    value: props.value ?? props.text ?? base.value,
    ...(styleId ? { styleId } : {}),
    ...(type ? { type } : {}),
    ...(formula ? { formula } : {}),
  };
}

function applyCellStyleProps(state: ExcelWorkbookState, cell: ExcelCellModel, props: Record<string, string>) {
  if (!hasStyleProps(props)) {
    return;
  }
  const styleId = registerStyle(state, props);
  cell.styleId = String(styleId);
}

function hasStyleProps(props: Record<string, string>) {
  return Object.keys(props).some((key) => {
    const lower = key.toLowerCase();
    return lower === "fill"
      || lower === "numfmt"
      || lower === "format"
      || lower === "numberformat"
      || lower.startsWith("font.")
      || lower.startsWith("alignment.");
  });
}

function registerStyle(state: ExcelWorkbookState, props: Record<string, string>) {
  const stylesheet = parseStylesheet(state.styleSheetXml ?? buildDefaultStylesheetXml());
  const fontXml = buildFontXml(props);
  const fillXml = buildFillXml(props);
  const borderXml = `<border><left/><right/><top/><bottom/><diagonal/></border>`;
  const numFmtCode = props.numFmt ?? props.numfmt ?? props.format ?? props.numberformat;
  const numFmtId = numFmtCode ? ensureNumFmt(stylesheet, numFmtCode) : 0;
  const fontId = ensureFragment(stylesheet.fonts, fontXml, `<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`);
  const fillId = ensureFragment(stylesheet.fills, fillXml, `<fill><patternFill patternType="none"/></fill>`);
  const borderId = ensureFragment(stylesheet.borders, borderXml, `<border><left/><right/><top/><bottom/><diagonal/></border>`);
  const alignmentXml = buildAlignmentXml(props);
  const xfXml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''}${fontXml !== DEFAULT_FONT_XML ? ' applyFont="1"' : ''}${fillXml !== DEFAULT_FILL_XML ? ' applyFill="1"' : ''}${alignmentXml ? ' applyAlignment="1"' : ''}>${alignmentXml}</xf>`;
  const xfId = ensureFragment(stylesheet.cellXfs, xfXml, `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`);
  state.styleSheetXml = serializeStylesheet(stylesheet);
  return xfId;
}

function normalizeExcelCell(cell: ExcelCellModel | undefined): ExcelCellModel {
  return {
    value: cell?.value ?? "",
    ...(cell?.styleId ? { styleId: cell.styleId } : {}),
    ...(cell?.type ? { type: cell.type } : {}),
    ...(cell?.formula ? { formula: normalizeFormula(cell.formula) } : {}),
  };
}

function formatCellDisplayValue(cell: ExcelCellModel, evaluate?: () => string | undefined) {
  if (cell.type === "boolean") {
    return cell.value === "1" ? "TRUE" : "FALSE";
  }
  if (cell.formula && cell.value === "" && evaluate) {
    return evaluate() ?? "";
  }
  return cell.value;
}

function normalizeFormula(formula: string) {
  return formula.replace(/^=/, "");
}

function evaluateFormulaForDisplay(sheet: ExcelSheetModel, ref: string) {
  const visited = new Set<string>();
  const numeric = evaluateFormulaExpression(sheet, ref, visited);
  if (numeric === undefined || Number.isNaN(numeric)) {
    return undefined;
  }
  return Number.isInteger(numeric) ? String(numeric) : String(Number(numeric.toFixed(10)));
}

function evaluateFormulaExpression(sheet: ExcelSheetModel, ref: string, visited: Set<string>): number | undefined {
  const key = `${sheet.name}!${ref}`;
  if (visited.has(key)) {
    return undefined;
  }
  visited.add(key);
  const cell = sheet.cells[ref];
  if (!cell) {
    visited.delete(key);
    return 0;
  }
  if (!cell.formula) {
    const value = coerceCellToNumber(cell);
    visited.delete(key);
    return value;
  }

  let expression = normalizeFormula(cell.formula);
  const aggregateMatch = /^(SUM|AVERAGE|MIN|MAX)\(([^()]+)\)$/i.exec(expression.trim());
  if (aggregateMatch) {
    const result = foldFormulaArgs(aggregateMatch[2], sheet, visited, aggregateMatch[1].toUpperCase() === "SUM"
      ? (values) => values.reduce((sum, value) => sum + value, 0)
      : aggregateMatch[1].toUpperCase() === "AVERAGE"
        ? (values) => values.length > 0 ? values.reduce((sum, value) => sum + value, 0) / values.length : 0
        : aggregateMatch[1].toUpperCase() === "MIN"
          ? (values) => values.length > 0 ? Math.min(...values) : 0
          : (values) => values.length > 0 ? Math.max(...values) : 0);
    visited.delete(key);
    return result;
  }
  const functionEvaluators: Record<string, (args: string) => number | undefined> = {
    SUM: (args) => foldFormulaArgs(args, sheet, visited, (values) => values.reduce((sum, value) => sum + value, 0)),
    AVERAGE: (args) => foldFormulaArgs(args, sheet, visited, (values) => values.length > 0 ? values.reduce((sum, value) => sum + value, 0) / values.length : 0),
    MIN: (args) => foldFormulaArgs(args, sheet, visited, (values) => values.length > 0 ? Math.min(...values) : 0),
    MAX: (args) => foldFormulaArgs(args, sheet, visited, (values) => values.length > 0 ? Math.max(...values) : 0),
  };

  let replaced = true;
  while (replaced) {
    replaced = false;
    expression = expression.replace(/\b(SUM|AVERAGE|MIN|MAX)\(([^()]*)\)/gi, (match, fn, args) => {
      const result = functionEvaluators[fn.toUpperCase()]?.(args);
      if (result === undefined) {
        return match;
      }
      replaced = true;
      return String(result);
    });
  }

  expression = expression.replace(/\b([A-Z]+[0-9]+)\b/g, (match, refValue) => {
    const value = evaluateFormulaExpression(sheet, refValue.toUpperCase(), visited);
    return value === undefined ? match : String(value);
  });

  visited.delete(key);
  if (!/^[0-9+\-*/().,\s]+$/.test(expression)) {
    return undefined;
  }
  try {
    // eslint-disable-next-line no-new-func
    const result = Function(`return (${expression});`)();
    return typeof result === "number" && Number.isFinite(result) ? result : undefined;
  } catch {
    return undefined;
  }
}

function foldFormulaArgs(
  args: string,
  sheet: ExcelSheetModel,
  visited: Set<string>,
  reducer: (values: number[]) => number,
) {
  const values = args
    .split(",")
    .flatMap((part) => {
      const value = part.trim();
      if (/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/i.test(value)) {
        return expandRange(value.toUpperCase()).map((ref) => evaluateFormulaExpression(sheet, ref, visited)).filter((item): item is number => item !== undefined);
      }
      if (/^[A-Z]+[0-9]+$/i.test(value)) {
        const evaluated = evaluateFormulaExpression(sheet, value.toUpperCase(), visited);
        return evaluated !== undefined ? [evaluated] : [];
      }
      const numeric = Number(value);
      return Number.isFinite(numeric) ? [numeric] : [];
    });
  return reducer(values);
}

function coerceCellToNumber(cell: ExcelCellModel) {
  if (cell.type === "boolean") {
    return cell.value === "1" ? 1 : 0;
  }
  const numeric = Number(cell.value);
  return Number.isFinite(numeric) ? numeric : undefined;
}

function mergeWorkbookSettings(existing: ExcelWorkbookSettings | undefined, props: Record<string, string>): ExcelWorkbookSettings {
  const next: ExcelWorkbookSettings = { ...(existing ?? {}) };
  if (props.date1904 !== undefined) next.date1904 = isTruthy(props.date1904);
  if (props.codeName !== undefined || props.codename !== undefined) next.codeName = props.codeName ?? props.codename;
  if (props.filterPrivacy !== undefined || props.filterprivacy !== undefined) next.filterPrivacy = isTruthy(props.filterPrivacy ?? props.filterprivacy ?? "false");
  if (props.showObjects !== undefined || props.showobjects !== undefined) next.showObjects = (props.showObjects ?? props.showobjects)?.toLowerCase();
  if (props.backupFile !== undefined || props.backupfile !== undefined) next.backupFile = isTruthy(props.backupFile ?? props.backupfile ?? "false");
  if (props.dateCompatibility !== undefined || props.datecompatibility !== undefined) next.dateCompatibility = isTruthy(props.dateCompatibility ?? props.datecompatibility ?? "false");
  if (props["calc.mode"] !== undefined || props.calcmode !== undefined) next.calcMode = normalizeCalcMode(props["calc.mode"] ?? props.calcmode ?? "");
  if (props["calc.iterate"] !== undefined || props.iterate !== undefined) next.iterate = isTruthy(props["calc.iterate"] ?? props.iterate ?? "false");
  if (props["calc.iterateCount"] !== undefined || props.iteratecount !== undefined) next.iterateCount = Number(props["calc.iterateCount"] ?? props.iteratecount);
  if (props["calc.iterateDelta"] !== undefined || props.iteratedelta !== undefined) next.iterateDelta = Number(props["calc.iterateDelta"] ?? props.iteratedelta);
  if (props["calc.fullPrecision"] !== undefined || props.fullprecision !== undefined) next.fullPrecision = isTruthy(props["calc.fullPrecision"] ?? props.fullprecision ?? "false");
  if (props["calc.fullCalcOnLoad"] !== undefined || props.fullcalconload !== undefined) next.fullCalcOnLoad = isTruthy(props["calc.fullCalcOnLoad"] ?? props.fullcalconload ?? "false");
  if (props["calc.refMode"] !== undefined || props.refmode !== undefined) next.refMode = normalizeRefMode(props["calc.refMode"] ?? props.refmode ?? "");
  if (props["workbook.lockStructure"] !== undefined || props.lockstructure !== undefined) next.lockStructure = isTruthy(props["workbook.lockStructure"] ?? props.lockstructure ?? "false");
  if (props["workbook.lockWindows"] !== undefined || props.lockwindows !== undefined) next.lockWindows = isTruthy(props["workbook.lockWindows"] ?? props.lockwindows ?? "false");
  return next;
}

function renderWorkbookProperties(settings?: ExcelWorkbookSettings) {
  if (!settings) return "";
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

const DEFAULT_FONT_XML = `<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`;
const DEFAULT_FILL_XML = `<fill><patternFill patternType="none"/></fill>`;

interface ParsedStylesheet {
  numFmts: string[];
  fonts: string[];
  fills: string[];
  borders: string[];
  cellXfs: string[];
}

function buildDefaultStylesheetXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">${DEFAULT_FONT_XML}</fonts>
  <fills count="2">${DEFAULT_FILL_XML}<fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`;
}

function parseStylesheet(xml: string): ParsedStylesheet {
  return {
    numFmts: extractStyleSection(xml, "numFmts", "numFmt"),
    fonts: extractStyleSection(xml, "fonts", "font"),
    fills: extractStyleSection(xml, "fills", "fill"),
    borders: extractStyleSection(xml, "borders", "border"),
    cellXfs: extractStyleSection(xml, "cellXfs", "xf"),
  };
}

function extractStyleSection(xml: string, containerTag: string, itemTag: string) {
  const section = new RegExp(`<${containerTag}\\b[^>]*>([\\s\\S]*?)<\\/${containerTag}>`).exec(xml)?.[1] ?? "";
  return [...section.matchAll(new RegExp(`<${itemTag}\\b[\\s\\S]*?<\\/${itemTag}>|<${itemTag}\\b[^>]*/>`, "g"))].map((match) => match[0]);
}

function ensureFragment(collection: string[], fragment: string, fallback: string) {
  const normalized = fragment || fallback;
  const existing = collection.findIndex((item) => item === normalized);
  if (existing >= 0) {
    return existing;
  }
  collection.push(normalized);
  return collection.length - 1;
}

function ensureNumFmt(stylesheet: ParsedStylesheet, formatCode: string) {
  const existing = stylesheet.numFmts.findIndex((item) => new RegExp(`formatCode="${escapeXmlForRegex(escapeXml(formatCode))}"`).test(item));
  if (existing >= 0) {
    const id = /numFmtId="([^"]+)"/.exec(stylesheet.numFmts[existing])?.[1];
    return Number(id ?? "164");
  }
  const nextId = Math.max(163, ...stylesheet.numFmts.map((item) => Number(/numFmtId="([^"]+)"/.exec(item)?.[1] ?? "163"))) + 1;
  stylesheet.numFmts.push(`<numFmt numFmtId="${nextId}" formatCode="${escapeXml(formatCode)}"/>`);
  return nextId;
}

function buildFontXml(props: Record<string, string>) {
  const fontName = props["font.name"] ?? props.font ?? "Calibri";
  const fontSize = props["font.size"] ?? "11";
  const color = props["font.color"];
  const bold = isTruthy(props["font.bold"] ?? "false");
  const italic = isTruthy(props["font.italic"] ?? "false");
  const underline = props["font.underline"];
  const strike = isTruthy(props["font.strike"] ?? props["font.strikethrough"] ?? "false");
  return `<font>${bold ? "<b/>" : ""}${italic ? "<i/>" : ""}${strike ? "<strike/>" : ""}${underline ? `<u${underline !== "true" ? ` val="${escapeXml(underline)}"` : ""}/>` : ""}<sz val="${escapeXml(fontSize)}"/>${color ? `<color rgb="${escapeXml(normalizeArgbColor(color))}"/>` : '<color theme="1"/>'}<name val="${escapeXml(fontName)}"/><family val="2"/></font>`;
}

function buildFillXml(props: Record<string, string>) {
  const fill = props.fill ?? props.bgColor ?? props.bgcolor;
  if (!fill) {
    return DEFAULT_FILL_XML;
  }
  return `<fill><patternFill patternType="solid"><fgColor rgb="${escapeXml(normalizeArgbColor(fill))}"/><bgColor indexed="64"/></patternFill></fill>`;
}

function buildAlignmentXml(props: Record<string, string>) {
  const horizontal = props["alignment.horizontal"] ?? props.halign;
  const vertical = props["alignment.vertical"] ?? props.valign;
  const wrapText = props["alignment.wrapText"] ?? props["alignment.wraptext"] ?? props.wrapText ?? props.wraptext;
  const attrs = [
    horizontal ? `horizontal="${escapeXml(horizontal)}"` : "",
    vertical ? `vertical="${escapeXml(vertical)}"` : "",
    wrapText !== undefined ? `wrapText="${isTruthy(wrapText) ? 1 : 0}"` : "",
  ].filter(Boolean).join(" ");
  return attrs ? `<alignment ${attrs}/>` : "";
}

function serializeStylesheet(stylesheet: ParsedStylesheet) {
  const numFmts = stylesheet.numFmts.length > 0 ? `<numFmts count="${stylesheet.numFmts.length}">${stylesheet.numFmts.join("")}</numFmts>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${numFmts}
  <fonts count="${stylesheet.fonts.length}">${stylesheet.fonts.join("")}</fonts>
  <fills count="${stylesheet.fills.length}">${stylesheet.fills.join("")}</fills>
  <borders count="${stylesheet.borders.length}">${stylesheet.borders.join("")}</borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${stylesheet.cellXfs.length}">${stylesheet.cellXfs.join("")}</cellXfs>
</styleSheet>`;
}

function renderCalculationProperties(settings?: ExcelWorkbookSettings) {
  if (!settings) return "";
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
  if (!settings) return "";
  const attrs = [
    settings.lockStructure !== undefined ? `lockStructure="${settings.lockStructure ? 1 : 0}"` : "",
    settings.lockWindows !== undefined ? `lockWindows="${settings.lockWindows ? 1 : 0}"` : "",
  ].filter(Boolean);
  return attrs.length > 0 ? `<workbookProtection ${attrs.join(" ")}/>` : "";
}

function renderDefinedNames(namedRanges: ExcelNamedRangeModel[], sheets: ExcelSheetModel[]) {
  if (namedRanges.length === 0) return "";
  const items = namedRanges.map((range) => {
    const scopeIndex = range.scope ? sheets.findIndex((sheet) => sheet.name.toLowerCase() === range.scope!.toLowerCase()) : -1;
    const attrs = [
      `name="${escapeXml(range.name)}"`,
      ...(scopeIndex >= 0 ? [`localSheetId="${scopeIndex}"`] : []),
      ...(range.comment ? [`comment="${escapeXml(range.comment)}"`] : []),
    ];
    return `<definedName ${attrs.join(" ")}>${escapeXml(range.ref)}</definedName>`;
  }).join("");
  return `<definedNames>${items}</definedNames>`;
}

function parseWorkbookPropertyAttributes(attrs?: string): ExcelWorkbookSettings {
  if (!attrs) return {};
  return {
    ...(parseAttr(attrs, "date1904") !== undefined ? { date1904: isTruthy(parseAttr(attrs, "date1904") ?? "false") } : {}),
    ...(parseAttr(attrs, "codeName") ? { codeName: decodeXml(parseAttr(attrs, "codeName") ?? "") } : {}),
    ...(parseAttr(attrs, "filterPrivacy") !== undefined ? { filterPrivacy: isTruthy(parseAttr(attrs, "filterPrivacy") ?? "false") } : {}),
    ...(parseAttr(attrs, "showObjects") ? { showObjects: decodeXml(parseAttr(attrs, "showObjects") ?? "") } : {}),
    ...(parseAttr(attrs, "backupFile") !== undefined ? { backupFile: isTruthy(parseAttr(attrs, "backupFile") ?? "false") } : {}),
    ...(parseAttr(attrs, "dateCompatibility") !== undefined ? { dateCompatibility: isTruthy(parseAttr(attrs, "dateCompatibility") ?? "false") } : {}),
  };
}

function parseCalculationPropertyAttributes(attrs?: string): ExcelWorkbookSettings {
  if (!attrs) return {};
  return {
    ...(parseAttr(attrs, "calcMode") ? { calcMode: decodeXml(parseAttr(attrs, "calcMode") ?? "") } : {}),
    ...(parseAttr(attrs, "iterate") !== undefined ? { iterate: isTruthy(parseAttr(attrs, "iterate") ?? "false") } : {}),
    ...(parseAttr(attrs, "iterateCount") !== undefined ? { iterateCount: Number(parseAttr(attrs, "iterateCount")) } : {}),
    ...(parseAttr(attrs, "iterateDelta") !== undefined ? { iterateDelta: Number(parseAttr(attrs, "iterateDelta")) } : {}),
    ...(parseAttr(attrs, "fullPrecision") !== undefined ? { fullPrecision: isTruthy(parseAttr(attrs, "fullPrecision") ?? "false") } : {}),
    ...(parseAttr(attrs, "fullCalcOnLoad") !== undefined ? { fullCalcOnLoad: isTruthy(parseAttr(attrs, "fullCalcOnLoad") ?? "false") } : {}),
    ...(parseAttr(attrs, "refMode") ? { refMode: decodeXml(parseAttr(attrs, "refMode") ?? "") } : {}),
  };
}

function parseWorkbookProtectionAttributes(attrs?: string): ExcelWorkbookSettings {
  if (!attrs) return {};
  return {
    ...(parseAttr(attrs, "lockStructure") !== undefined ? { lockStructure: isTruthy(parseAttr(attrs, "lockStructure") ?? "false") } : {}),
    ...(parseAttr(attrs, "lockWindows") !== undefined ? { lockWindows: isTruthy(parseAttr(attrs, "lockWindows") ?? "false") } : {}),
  };
}

function filterSheetRaw(sheetXml: string, options?: { startRow?: number; endRow?: number; cols?: string[] }) {
  if (!options) return sheetXml;
  const cells = parseSheetCells(sheetXml, new Map());
  const rows = new Map<number, string[]>();
  for (const [ref, cell] of Object.entries(cells)) {
    const column = /^[A-Z]+/.exec(ref)?.[0] ?? "A";
    const row = Number(/\d+/.exec(ref)?.[0] ?? "1");
    if (options.startRow !== undefined && row < options.startRow) continue;
    if (options.endRow !== undefined && row > options.endRow) continue;
    if (options.cols?.length && !options.cols.map((item) => item.toUpperCase()).includes(column)) continue;
    const list = rows.get(row) ?? [];
    list.push(renderExcelCellXml(ref, cell));
    rows.set(row, list);
  }
  const xmlRows = [...rows.entries()].sort(([a], [b]) => a - b).map(([rowIndex, values]) => `<row r="${rowIndex}">${values.join("")}</row>`).join("");
  return sheetXml.replace(/<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/, `<sheetData>${xmlRows}</sheetData>`);
}

function parseRelationshipEntries(xml: string): RelationshipEntry[] {
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/g)].map((match) => ({
    id: parseAttr(match[1], "Id") ?? "",
    target: parseAttr(match[1], "Target") ?? "",
    type: parseAttr(match[1], "Type") ?? undefined,
  })).filter((relationship) => relationship.id && relationship.target);
}

function readRelationships(zip: Map<string, Buffer>, entryName: string) {
  const rels = zip.get(entryName);
  if (!rels) return [];
  return parseRelationshipEntries(rels.toString("utf8"));
}

function requireEntry(zip: Map<string, Buffer>, entryName: string) {
  const buffer = zip.get(entryName);
  if (!buffer) throw new OfficekitError(`OOXML entry '${entryName}' is missing.`, "invalid_ooxml");
  return buffer.toString("utf8");
}

function buildContentTypesXml(state: ExcelWorkbookState) {
  const overrides: Array<[string, string]> = [
    ...state.sheets.map((sheet) => [sheet.entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"] as [string, string]),
  ];
  for (const entryName of state.zip.keys()) {
    const override = resolveContentTypeOverride(entryName);
    if (override) overrides.push(override);
  }
  const seen = new Set<string>();
  const overrideXml = overrides
    .filter(([entryName]) => {
      if (seen.has(entryName)) return false;
      seen.add(entryName);
      return true;
    })
    .map(([entryName, contentType]) => `<Override PartName="/${entryName}" ContentType="${contentType}"/>`)
    .join("\n  ");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="json" ContentType="application/json"/>
  ${overrideXml}
</Types>`;
}

function buildRootRelsXml(state: ExcelWorkbookState) {
  const extraRelationships = state.zip.has("docProps/core.xml")
    ? '\n  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  ${extraRelationships}
</Relationships>`;
}

function buildCorePropertiesXml(metadata: Record<string, string>) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  ${metadata.title ? `<dc:title>${escapeXml(metadata.title)}</dc:title>` : ""}
  ${metadata.author ? `<dc:creator>${escapeXml(metadata.author)}</dc:creator>` : ""}
  ${metadata.subject ? `<dc:subject>${escapeXml(metadata.subject)}</dc:subject>` : ""}
  ${metadata.description ? `<dc:description>${escapeXml(metadata.description)}</dc:description>` : ""}
</cp:coreProperties>`;
}

function parseDelimitedRows(content: string, delimiter: string) {
  const rows: string[][] = [];
  if (!content) return rows;
  if (content.charCodeAt(0) === 0xfeff) content = content.slice(1);
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
      if (char === "\r" && content[index + 1] === "\n") index += 1;
      currentRow.push(field);
      field = "";
      if (!(currentRow.length === 1 && currentRow[0] === "")) rows.push([...currentRow]);
      currentRow.length = 0;
      continue;
    }
    field += char;
  }
  if (field.length > 0 || currentRow.length > 0) {
    currentRow.push(field);
    if (!(currentRow.length === 1 && currentRow[0] === "")) rows.push([...currentRow]);
  }
  return rows;
}

function inferImportedCell(rawValue: string): ExcelCellModel {
  if (rawValue === "") return { value: "" };
  if (rawValue.startsWith("=")) return { value: "", formula: normalizeFormula(rawValue) };
  if (/^(true|false)$/i.test(rawValue)) return { value: rawValue.toUpperCase() === "TRUE" ? "1" : "0", type: "boolean" };
  const isoDate = tryParseIsoDate(rawValue);
  if (isoDate) return { value: isoDate, type: "date" };
  if (!Number.isNaN(Number(rawValue))) return { value: rawValue, type: "number" };
  return { value: rawValue, type: "string" };
}

function tryParseIsoDate(rawValue: string) {
  const date = new Date(rawValue);
  if (Number.isNaN(date.getTime()) || !/^\d{4}-\d{2}-\d{2}/.test(rawValue)) {
    return null;
  }
  return String((date.getTime() - Date.UTC(1899, 11, 30)) / (24 * 60 * 60 * 1000));
}

function parseCellAddress(value: string) {
  const match = /^([A-Z]+)(\d+)$/.exec(value);
  if (!match) throw new UsageError(`Invalid cell address '${value}'.`, "Use an address like A1.");
  return { column: match[1], row: Number(match[2]) };
}

function expandRange(range: string) {
  const [start, end = start] = range.split(":");
  const startAddress = parseCellAddress(start);
  const endAddress = parseCellAddress(end);
  const refs: string[] = [];
  for (let row = startAddress.row; row <= endAddress.row; row += 1) {
    for (let col = columnNameToIndex(startAddress.column); col <= columnNameToIndex(endAddress.column); col += 1) {
      refs.push(`${indexToColumnName(col)}${row}`);
    }
  }
  return refs;
}

function compareCellRefs(a: string, b: string) {
  const aColumn = /^[A-Z]+/.exec(a)?.[0] ?? "A";
  const bColumn = /^[A-Z]+/.exec(b)?.[0] ?? "A";
  const aRow = Number(/\d+/.exec(a)?.[0] ?? "0");
  const bRow = Number(/\d+/.exec(b)?.[0] ?? "0");
  return aRow === bRow ? columnNameToIndex(aColumn) - columnNameToIndex(bColumn) : aRow - bRow;
}

function columnNameToIndex(column: string) {
  let result = 0;
  for (const char of column.toUpperCase()) {
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

function parseAttr(source: string, name: string) {
  return new RegExp(`${name}="([^"]+)"`).exec(source)?.[1];
}

function extractTagText(xml: string, tag: string) {
  return new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`).exec(xml)?.[1]
    ?? new RegExp(`<(?:\\w+:)?${tag}[^>]*>([\\s\\S]*?)<\\/(?:\\w+:)?${tag}>`).exec(xml)?.[1];
}

function extractTexts(xml: string) {
  return [...xml.matchAll(/<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)].map((match) => decodeXml(match[1]));
}

function decodeXml(value: string) {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function normalizeZipPath(baseDir: string, target: string) {
  const normalized = target.replace(/\\/g, "/");
  if (normalized.startsWith("/")) return path.posix.normalize(normalized.slice(1));
  return path.posix.normalize(path.posix.join(baseDir, normalized));
}

function getRelationshipsEntryName(entryName: string) {
  const directory = path.posix.dirname(entryName);
  const basename = path.posix.basename(entryName);
  return path.posix.join(directory, "_rels", `${basename}.rels`);
}

function ensureWorksheetNamespaces(xml: string, namespaces: Record<string, string>) {
  return xml.replace(/<(?:\w+:)?worksheet\b([^>]*)>/, (match, attrs) => {
    let next = match;
    for (const [prefix, uri] of Object.entries(namespaces)) {
      if (new RegExp(`xmlns:${prefix}=`).test(attrs)) continue;
      next = next.replace(/>$/, ` xmlns:${prefix}="${uri}">`);
    }
    return next;
  });
}

function nextIndexedPartPath(zip: Map<string, Buffer>, prefix: string, suffix: string) {
  const pattern = new RegExp(`^${escapeXmlForRegex(prefix)}(\\d+)${escapeXmlForRegex(suffix)}$`, "i");
  const nextIndex = [...zip.keys()]
    .map((name) => Number(pattern.exec(name)?.[1] ?? "0"))
    .reduce((max, value) => Math.max(max, value), 0) + 1;
  return `${prefix}${nextIndex}${suffix}`;
}

function appendRelationship(
  zip: Map<string, Buffer>,
  relsPath: string,
  ownerEntryName: string,
  targetEntryName: string,
  type: string,
) {
  const relsXml = zip.get(relsPath)?.toString("utf8") ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const relationships = parseRelationshipEntries(relsXml);
  const target = path.posix.relative(path.posix.dirname(ownerEntryName), targetEntryName).replace(/\\/g, "/");
  const existing = relationships.find((relationship) => relationship.target === target && relationship.type === type);
  if (existing) {
    return existing.id;
  }
  const nextId = `rId${relationships.reduce((max, relationship) => Math.max(max, Number(/^rId(\d+)$/.exec(relationship.id)?.[1] ?? "0")), 0) + 1}`;
  const nextXml = relsXml.replace(
    /<\/Relationships>/,
    `  <Relationship Id="${nextId}" Type="${type}" Target="${escapeXml(target)}"/>\n</Relationships>`,
  );
  zip.set(relsPath, Buffer.from(nextXml, "utf8"));
  return nextId;
}

function resolveTableColumnNames(
  sheet: ExcelSheetModel,
  props: Record<string, string>,
  startAddress: { column: string; row: number },
  columnCount: number,
  headerRow: boolean,
) {
  if (props.columns) {
    const provided = props.columns.split(",").map((item) => item.trim()).filter(Boolean);
    return Array.from({ length: columnCount }, (_, index) => provided[index] || `Column${index + 1}`);
  }
  return Array.from({ length: columnCount }, (_, index) => {
    if (!headerRow) return `Column${index + 1}`;
    const ref = `${indexToColumnName(columnNameToIndex(startAddress.column) + index)}${startAddress.row}`;
    const text = sheet.cells[ref]?.value?.trim();
    return text || `Column${index + 1}`;
  });
}

function renderTableXml(input: {
  id: number;
  name: string;
  displayName: string;
  ref: string;
  styleName: string;
  headerRow: boolean;
  totalsRow: boolean;
  columnNames?: string[];
  columns?: string[];
}) {
  const columnNames = input.columnNames ?? input.columns ?? ["Column1"];
  const totalAttrs = input.totalsRow ? ' totalsRowShown="1" totalsRowCount="1"' : ' totalsRowShown="0"';
  const headerAttr = ` headerRowCount="${input.headerRow ? 1 : 0}"`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${input.id}" name="${escapeXml(input.name)}" displayName="${escapeXml(input.displayName)}" ref="${escapeXml(input.ref)}"${totalAttrs}${headerAttr}>
  <autoFilter ref="${escapeXml(input.ref)}"/>
  <tableColumns count="${columnNames.length}">
    ${columnNames.map((columnName, index) => `<tableColumn id="${index + 1}" name="${escapeXml(columnName)}"/>`).join("")}
  </tableColumns>
  <tableStyleInfo name="${escapeXml(input.styleName)}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`;
}

function upsertTablePartReference(sheetXml: string, relId: string) {
  const existing = [...sheetXml.matchAll(/<(?:\w+:)?tablePart\b[^>]*r:id="([^"]+)"[^>]*\/?>/g)].map((match) => match[1]);
  const nextIds = existing.includes(relId) ? existing : [...existing, relId];
  const tablePartsXml = nextIds.length > 0
    ? `<tableParts count="${nextIds.length}">${nextIds.map((id) => `<tablePart r:id="${escapeXml(id)}"/>`).join("")}</tableParts>`
    : "";
  return replaceOrInsert(
    sheetXml,
    /<(?:\w+:)?tableParts\b[\s\S]*?<\/(?:\w+:)?tableParts>/,
    tablePartsXml,
    /<(?:\w+:)?drawing\b[^>]*\/?>|<(?:\w+:)?extLst\b[\s\S]*?<\/(?:\w+:)?extLst>|<(?:\w+:)?colBreaks\b[\s\S]*?<\/(?:\w+:)?colBreaks>|<(?:\w+:)?rowBreaks\b[\s\S]*?<\/(?:\w+:)?rowBreaks>|<(?:\w+:)?headerFooter\b[\s\S]*?<\/(?:\w+:)?headerFooter>|<(?:\w+:)?pageSetup\b[^>]*\/?>|<(?:\w+:)?sheetProtection\b[^>]*\/?>|<(?:\w+:)?autoFilter\b[^>]*\/?>|<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/,
  );
}

function resolveContentTypeOverride(entryName: string) {
  if (entryName === "xl/workbook.xml") return [entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"] as [string, string];
  if (entryName === "xl/styles.xml") return [entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"] as [string, string];
  if (entryName === "xl/sharedStrings.xml") return [entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"] as [string, string];
  if (entryName === "docProps/core.xml") return [entryName, "application/vnd.openxmlformats-package.core-properties+xml"] as [string, string];
  if (/^xl\/comments\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"] as [string, string];
  if (/^xl\/tables\/table\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"] as [string, string];
  if (/^xl\/drawings\/drawing\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.drawing+xml"] as [string, string];
  if (/^xl\/charts\/chart\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"] as [string, string];
  if (/^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"] as [string, string];
  return undefined;
}

function normalizeCalcMode(value: string) {
  const normalized = value.trim().toLowerCase();
  if (normalized === "automatic") return "auto";
  if (normalized === "autoexcepttables" || normalized === "autonoexcepttables" || normalized === "autonotable") return "autoNoTable";
  return normalized;
}

function normalizeRefMode(value: string) {
  const normalized = value.trim().toUpperCase();
  return normalized === "R1C1" ? "R1C1" : "A1";
}

function normalizeArgbColor(value: string) {
  const normalized = value.replace(/^#/, "").toUpperCase();
  return normalized.length === 6 ? `FF${normalized}` : normalized;
}

function parseBreakList(value: string) {
  return value
    .split(",")
    .map((item) => Number(item.trim()))
    .filter((item) => !Number.isNaN(item));
}

function isTruthy(value: string) {
  return /^(1|true|yes|on)$/i.test(value.trim());
}

function escapeHtml(value: string) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function escapeXml(value: string) {
  return escapeHtml(value).replaceAll('"', "&quot;").replaceAll("'", "&apos;");
}

function escapeXmlForRegex(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
