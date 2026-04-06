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

interface ExcelConditionalFormattingModel {
  sqref: string;
  cfType: "databar" | "colorscale" | "iconset" | "formula" | "topn" | "aboveaverage" | "uniquevalues" | "duplicatevalues" | "containstext" | "dateoccurring";
  priority?: number;
  color?: string;
  min?: string;
  max?: string;
  minColor?: string;
  midColor?: string;
  maxColor?: string;
  iconset?: string;
  reverse?: boolean;
  showvalue?: boolean;
  formula?: string;
  dxfId?: number;
  fontColor?: string;
  fontBold?: boolean;
  fill?: string;
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  above?: boolean;
  text?: string;
  period?: string;
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
  axisMin?: number;
  axisMax?: number;
  majorUnit?: number;
  minorUnit?: number;
  axisNumberFormat?: string;
  styleId?: number;
  plotAreaFill?: string;
  chartAreaFill?: string;
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

  if (options.type === "cf" || options.type === "conditionalformatting" || options.type === "databar" || options.type === "colorscale" || options.type === "iconset" || options.type === "formulacf" || options.type === "topn" || options.type === "aboveaverage" || options.type === "uniquevalues" || options.type === "duplicatevalues" || options.type === "containstext" || options.type === "dateoccurring") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const cf = addConditionalFormatting(state, sheet, options.type, options.props);
    await writeWorkbookState(filePath, state);
    return { ...cf, path: `/${sheet.name}/cf[${parseConditionalFormatting(sheet.xml).length}]`, type: "conditionalformatting" };
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

  if (options.type === "chart") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const chart = addChart(state, sheet, options.props);
    await writeWorkbookState(filePath, state);
    return { ...chart, path: `/${sheet.name}/chart[${getSheetCharts(state, sheet).length}]`, type: "chart" };
  }

  if (options.type === "pivottable" || options.type === "pivot") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const pivot = addPivotTable(state, sheet, options.props);
    await writeWorkbookState(filePath, state);
    return { ...pivot, path: `/${sheet.name}/pivottable[${getSheetPivots(state, sheet).length}]`, type: "pivottable" };
  }

  if (options.type === "picture" || options.type === "image") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const picture = await addPicture(state, sheet, options.props);
    await writeWorkbookState(filePath, state);
    return { ...picture, path: `/${sheet.name}/picture[${getDrawingShapes(state, sheet).filter((item) => item.kind === "picture").length}]`, type: "picture" };
  }

  if (options.type === "shape" || options.type === "textbox") {
    const sheet = ensureSheetState(state, normalizeSheetPath(targetPath));
    const shape = addShape(state, sheet, options.props);
    await writeWorkbookState(filePath, state);
    return { ...shape, path: `/${sheet.name}/shape[${getDrawingShapes(state, sheet).filter((item) => item.kind === "shape").length}]`, type: "shape" };
  }

  if (options.type !== "cell") {
    throw new UsageError(
      "Excel add currently supports: sheet, row, cell, namedrange, validation, comment, autofilter, rowbreak, colbreak, table, sparkline, chart, pivottable, picture, or shape.",
      "Use / with --type sheet|namedrange, or /Sheet1 with --type row|cell|validation|comment|autofilter|rowbreak|colbreak|table|sparkline|chart|pivottable|picture|shape.",
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

  const cfMatch = /^\/([^/]+)\/cf\[(\d+)\]$/i.exec(targetPath);
  if (cfMatch) {
    const sheet = ensureSheetState(state, cfMatch[1]);
    const next = setConditionalFormatting(state, sheet, Number(cfMatch[2]), options.props);
    await writeWorkbookState(filePath, state);
    return { ...next, path: targetPath, type: "conditionalformatting" };
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

  const cfMatch = /^\/([^/]+)\/cf\[(\d+)\]$/i.exec(targetPath);
  if (cfMatch) {
    const sheet = ensureSheetState(state, cfMatch[1]);
    removeConditionalFormatting(sheet, Number(cfMatch[2]));
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
  if (normalized === "cf" || normalized === "conditionalformatting" || normalized === "conditionalformattings" || normalized === "cfs") {
    for (const sheet of state.sheets) {
      parseConditionalFormatting(sheet.xml).forEach((cf, index) => {
        nodes.push({ ...cf, path: `/${sheet.name}/cf[${index + 1}]`, type: "conditionalformatting" });
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
  throw new UsageError(`Unsupported Excel query selector '${selector}'.`, "Supported selectors: sheet, namedrange, cell, formula, validation, cf, comment, table, chart, pivottable, sparkline, shape, picture.");
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
  const cfMatch = /^\/([^/]+)\/cf\[(\d+)\]$/i.exec(targetPath);
  if (cfMatch) {
    const sheet = ensureSheetState(state, cfMatch[1]);
    const cf = parseConditionalFormatting(sheet.xml)[Number(cfMatch[2]) - 1];
    if (!cf) throw new OfficekitError(`Conditional formatting ${cfMatch[2]} does not exist.`, "not_found");
    return { ...cf, path: targetPath, type: "conditionalformatting" };
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
  return range.includes(":") ? materializeRangeNode(sheet, range, state) : materializeCellNode(sheet, range, state);
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

function materializeRangeNode(sheet: ExcelSheetModel, range: string, state?: ExcelWorkbookState) {
  return {
    path: `/${sheet.name}/${range}`,
    type: "range",
    cells: expandRange(range).map((ref) => materializeCellNode(sheet, ref, state)),
  };
}

function materializeCellNode(sheet: ExcelSheetModel, ref: string, state?: ExcelWorkbookState) {
  const cell = sheet.cells[ref];
  if (!cell) {
    return { path: `/${sheet.name}/${ref}`, ref, type: "cell", value: null };
  }
  const evaluatedValue = cell.formula && cell.value === "" ? evaluateFormulaForDisplay(state, sheet, ref) : undefined;
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
    parseConditionalFormatting(sheet.xml).forEach((cf, index) => {
      lines.push(`  CF ${index + 1}: ${cf.sqref} [${cf.cfType}]`);
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
          const materialized = materializeCellNode(sheet, `${column}${rowIndex}`, state) as { value?: string | null; evaluatedValue?: string };
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
      const materialized = materializeCellNode(sheet, ref, state) as { value?: string | null; evaluatedValue?: string };
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

function addConditionalFormatting(state: ExcelWorkbookState, sheet: ExcelSheetModel, requestedType: string, props: Record<string, string>) {
  const normalizedType = requestedType.toLowerCase() === "cf" || requestedType.toLowerCase() === "conditionalformatting"
    ? (props.type ?? "databar").toLowerCase()
    : requestedType.toLowerCase();
  const sqref = props.sqref ?? props.range ?? props.ref ?? "A1:A10";
  const rules = parseConditionalFormatting(sheet.xml);
  const priority = rules.length + 1;
  const dxfId = hasConditionalFormattingStyleProps(props) ? registerDifferentialFormat(state, props) : undefined;
  let next: ExcelConditionalFormattingModel;
  if (normalizedType === "databar") {
    next = {
      sqref,
      cfType: "databar",
      priority,
      ...(props.color ? { color: normalizeArgbColor(props.color) } : { color: "FF638EC6" }),
      ...(props.min !== undefined ? { min: props.min } : {}),
      ...(props.max !== undefined ? { max: props.max } : {}),
      ...(dxfId !== undefined ? { dxfId } : {}),
    };
  } else if (normalizedType === "colorscale") {
    next = {
      sqref,
      cfType: "colorscale",
      priority,
      minColor: normalizeArgbColor(props.mincolor ?? "FFF8696B"),
      ...(props.midcolor ? { midColor: normalizeArgbColor(props.midcolor) } : {}),
      maxColor: normalizeArgbColor(props.maxcolor ?? "FF63BE7B"),
      ...(dxfId !== undefined ? { dxfId } : {}),
    };
  } else if (normalizedType === "iconset") {
    next = {
      sqref,
      cfType: "iconset",
      priority,
      iconset: props.iconset ?? props.icons ?? "3TrafficLights1",
      ...(props.reverse !== undefined ? { reverse: isTruthy(props.reverse) } : {}),
      ...(props.showvalue !== undefined ? { showvalue: isTruthy(props.showvalue) } : {}),
      ...(dxfId !== undefined ? { dxfId } : {}),
    };
  } else if (normalizedType === "formulacf" || normalizedType === "formula") {
    if (!props.formula) {
      throw new UsageError("Formula-based conditional formatting requires --prop formula=...");
    }
    next = {
      sqref,
      cfType: "formula",
      priority,
      formula: props.formula,
      ...(dxfId !== undefined ? { dxfId } : {}),
      ...extractConditionalFormattingStyleProps(props),
    };
  } else if (normalizedType === "topn" || normalizedType === "top10") {
    next = {
      sqref,
      cfType: "topn",
      priority,
      rank: Number(props.rank ?? "10"),
      ...(props.percent !== undefined ? { percent: isTruthy(props.percent) } : {}),
      ...(props.bottom !== undefined ? { bottom: isTruthy(props.bottom) } : {}),
      ...(dxfId !== undefined ? { dxfId } : {}),
      ...extractConditionalFormattingStyleProps(props),
    };
  } else if (normalizedType === "aboveaverage") {
    next = {
      sqref,
      cfType: "aboveaverage",
      priority,
      ...(props.above !== undefined ? { above: isTruthy(props.above) } : { above: true }),
      ...(dxfId !== undefined ? { dxfId } : {}),
      ...extractConditionalFormattingStyleProps(props),
    };
  } else if (normalizedType === "uniquevalues") {
    next = {
      sqref,
      cfType: "uniquevalues",
      priority,
      ...(dxfId !== undefined ? { dxfId } : {}),
      ...extractConditionalFormattingStyleProps(props),
    };
  } else if (normalizedType === "duplicatevalues") {
    next = {
      sqref,
      cfType: "duplicatevalues",
      priority,
      ...(dxfId !== undefined ? { dxfId } : {}),
      ...extractConditionalFormattingStyleProps(props),
    };
  } else if (normalizedType === "containstext") {
    if (!props.text) {
      throw new UsageError("ContainsText conditional formatting requires --prop text=...");
    }
    next = {
      sqref,
      cfType: "containstext",
      priority,
      text: props.text,
      ...(dxfId !== undefined ? { dxfId } : {}),
      ...extractConditionalFormattingStyleProps(props),
    };
  } else if (normalizedType === "dateoccurring" || normalizedType === "timeperiod") {
    next = {
      sqref,
      cfType: "dateoccurring",
      priority,
      period: normalizeTimePeriod(props.period ?? "today"),
      ...(dxfId !== undefined ? { dxfId } : {}),
      ...extractConditionalFormattingStyleProps(props),
    };
  } else {
    throw new UsageError(`Unsupported conditional formatting type '${requestedType}'.`, "Use databar, colorscale, iconset, formulacf, topn, aboveaverage, uniquevalues, duplicatevalues, containstext, dateoccurring, or cf --prop type=...");
  }
  rules.push(next);
  sheet.xml = replaceSheetConditionalFormatting(sheet.xml, rules);
  return next;
}

function setConditionalFormatting(state: ExcelWorkbookState, sheet: ExcelSheetModel, index: number, props: Record<string, string>) {
  const rules = parseConditionalFormatting(sheet.xml);
  const current = rules[index - 1];
  if (!current) {
    throw new OfficekitError(`Conditional formatting ${index} does not exist.`, "not_found");
  }
  const next: ExcelConditionalFormattingModel = { ...current };
  const dxfRelevantProps = Object.fromEntries(Object.entries(props).filter(([key]) => isConditionalFormattingStyleKey(key)));
  if (Object.keys(dxfRelevantProps).length > 0) {
    next.dxfId = registerDifferentialFormat(state, { ...serializeConditionalFormattingStyleProps(current), ...dxfRelevantProps });
    Object.assign(next, extractConditionalFormattingStyleProps({ ...serializeConditionalFormattingStyleProps(current), ...dxfRelevantProps }));
  }
  for (const [key, value] of Object.entries(props)) {
    switch (key.toLowerCase()) {
      case "sqref":
      case "range":
      case "ref":
        next.sqref = value;
        break;
      case "color":
        next.color = normalizeArgbColor(value);
        break;
      case "mincolor":
        next.minColor = normalizeArgbColor(value);
        break;
      case "midcolor":
        next.midColor = normalizeArgbColor(value);
        break;
      case "maxcolor":
        next.maxColor = normalizeArgbColor(value);
        break;
      case "iconset":
      case "icons":
        next.iconset = value;
        break;
      case "reverse":
        next.reverse = isTruthy(value);
        break;
      case "showvalue":
        next.showvalue = isTruthy(value);
        break;
      case "formula":
        next.formula = value;
        break;
      case "rank":
        next.rank = Number(value);
        break;
      case "percent":
        next.percent = isTruthy(value);
        break;
      case "bottom":
        next.bottom = isTruthy(value);
        break;
      case "above":
        next.above = isTruthy(value);
        break;
      case "text":
        next.text = value;
        break;
      case "period":
        next.period = normalizeTimePeriod(value);
        break;
      case "font.color":
      case "font.bold":
      case "fill":
        break;
      default:
        throw new UsageError(`Unsupported conditional formatting property '${key}'.`, "Supported: sqref/range/ref, color, mincolor, midcolor, maxcolor, iconset/icons, reverse, showvalue, formula, rank, percent, bottom, above, text, period, font.color, font.bold, fill.");
    }
  }
  rules[index - 1] = next;
  sheet.xml = replaceSheetConditionalFormatting(sheet.xml, rules);
  return next;
}

function removeConditionalFormatting(sheet: ExcelSheetModel, index: number) {
  const rules = parseConditionalFormatting(sheet.xml);
  if (!rules[index - 1]) {
    throw new OfficekitError(`Conditional formatting ${index} does not exist.`, "not_found");
  }
  sheet.xml = replaceSheetConditionalFormatting(sheet.xml, rules.filter((_, itemIndex) => itemIndex !== index - 1));
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

function addChart(state: ExcelWorkbookState, sheet: ExcelSheetModel, props: Record<string, string>) {
  const drawingPath = ensureDrawingPart(state, sheet);
  const chartPath = nextIndexedPartPath(state.zip, "xl/charts/chart", ".xml");
  const drawingRelId = appendRelationship(
    state.zip,
    getRelationshipsEntryName(drawingPath),
    drawingPath,
    chartPath,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  );
  const chartData = resolveChartAddData(state, sheet, props);
  state.zip.set(chartPath, Buffer.from(buildChartXml(chartData), "utf8"));

  const drawingXml = requireEntry(state.zip, drawingPath);
  const nextAnchorId = Math.max(
    1,
    ...[...drawingXml.matchAll(/<xdr:cNvPr\b[^>]*id="(\d+)"/g)].map((match) => Number(match[1])),
  ) + 1;
  const fromCol = Number(props.x ?? "0");
  const fromRow = Number(props.y ?? "0");
  const toCol = fromCol + Number(props.width ?? "8");
  const toRow = fromRow + Number(props.height ?? "15");
  const chartName = props.title ?? props.name ?? `Chart ${nextAnchorId}`;
  const anchorXml = `  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>${fromCol}</xdr:col><xdr:row>${fromRow}</xdr:row></xdr:from>
    <xdr:to><xdr:col>${toCol}</xdr:col><xdr:row>${toRow}</xdr:row></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="${nextAnchorId}" name="${escapeXml(chartName)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="${escapeXml(drawingRelId)}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;
  state.zip.set(drawingPath, Buffer.from(drawingXml.replace(/<\/xdr:wsDr>/, `${anchorXml}\n</xdr:wsDr>`), "utf8"));
  return getSheetCharts(state, sheet).at(-1) ?? { title: chartData.title, path: path.posix.relative(path.posix.dirname(drawingPath), chartPath), sheet: sheet.name };
}

function addPivotTable(state: ExcelWorkbookState, sheet: ExcelSheetModel, props: Record<string, string>) {
  const source = props.source ?? props.src;
  if (!source) {
    throw new UsageError("Excel pivottable requires --prop source=Sheet1!A1:D10.");
  }
  const pivotIndex = Math.max(
    0,
    ...[...state.zip.keys()].map((entryName) => Number(/^xl\/pivotTables\/pivotTable(\d+)\.xml$/i.exec(entryName)?.[1] ?? "0")),
  ) + 1;
  const pivotPath = `xl/pivotTables/pivotTable${pivotIndex}.xml`;
  const name = props.name ?? `PivotTable${pivotIndex}`;
  const rowGrandTotals = props.rowGrandTotals === undefined && props.rowgrandtotals === undefined ? true : isTruthy(props.rowGrandTotals ?? props.rowgrandtotals ?? "false");
  const colGrandTotals = props.colGrandTotals === undefined && props.colgrandtotals === undefined ? true : isTruthy(props.colGrandTotals ?? props.colgrandtotals ?? "false");
  const compact = props.compact === undefined ? true : isTruthy(props.compact);
  const compactData = props.compactData === undefined && props.compactdata === undefined ? true : isTruthy(props.compactData ?? props.compactdata ?? "false");
  const outline = props.outline === undefined ? false : isTruthy(props.outline);
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${escapeXml(name)}" dataCaption="Values" rowGrandTotals="${rowGrandTotals ? 1 : 0}" colGrandTotals="${colGrandTotals ? 1 : 0}" compact="${compact ? 1 : 0}" compactData="${compactData ? 1 : 0}" outline="${outline ? 1 : 0}" location="${escapeXml(props.position ?? props.pos ?? "H1")}" source="${escapeXml(source)}"/>`;
  state.zip.set(pivotPath, Buffer.from(xml, "utf8"));
  appendRelationship(
    state.zip,
    getRelationshipsEntryName(sheet.entryName),
    sheet.entryName,
    pivotPath,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
  );
  return getSheetPivots(state, sheet).at(-1) ?? { name, path: path.posix.relative(path.posix.dirname(sheet.entryName), pivotPath) };
}

async function addPicture(state: ExcelWorkbookState, sheet: ExcelSheetModel, props: Record<string, string>) {
  const sourcePath = props.path ?? props.src;
  if (!sourcePath) {
    throw new UsageError("Excel picture requires --prop path=<image> or --prop src=<image>.");
  }
  const imageBytes = await readFile(sourcePath);
  const extension = normalizeImageExtension(path.extname(sourcePath));
  const mediaPath = nextIndexedPartPath(state.zip, "xl/media/image", `.${extension}`);
  state.zip.set(mediaPath, imageBytes);

  const drawingPath = ensureDrawingPart(state, sheet);
  const drawingRelId = appendRelationship(
    state.zip,
    getRelationshipsEntryName(drawingPath),
    drawingPath,
    mediaPath,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  );
  const drawingXml = requireEntry(state.zip, drawingPath);
  const nextAnchorId = nextDrawingObjectId(drawingXml);
  const { fromCol, fromRow, toCol, toRow } = resolveAnchorBounds(props, { x: 0, y: 0, width: 5, height: 5 });
  const name = props.name ?? `Picture ${nextAnchorId}`;
  const description = props.alt ?? props.description ?? "";
  const anchorXml = `  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>${fromCol}</xdr:col><xdr:row>${fromRow}</xdr:row></xdr:from>
    <xdr:to><xdr:col>${toCol}</xdr:col><xdr:row>${toRow}</xdr:row></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="${nextAnchorId}" name="${escapeXml(name)}"${description ? ` descr="${escapeXml(description)}"` : ""}/><xdr:cNvPicPr/><xdr:nvPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="${escapeXml(drawingRelId)}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;
  state.zip.set(drawingPath, Buffer.from(drawingXml.replace(/<\/xdr:wsDr>/, `${anchorXml}\n</xdr:wsDr>`), "utf8"));
  return getDrawingShapes(state, sheet).filter((item) => item.kind === "picture").at(-1) ?? { kind: "picture" as const, name };
}

function addShape(state: ExcelWorkbookState, sheet: ExcelSheetModel, props: Record<string, string>) {
  const drawingPath = ensureDrawingPart(state, sheet);
  const drawingXml = requireEntry(state.zip, drawingPath);
  const nextAnchorId = nextDrawingObjectId(drawingXml);
  const { fromCol, fromRow, toCol, toRow } = resolveAnchorBounds(props, { x: 1, y: 1, width: 5, height: 3 });
  const name = props.name ?? `Shape ${nextAnchorId}`;
  const text = props.text ?? props.value ?? "";
  const fillXml = props.fill && props.fill.toLowerCase() !== "none"
    ? `<a:solidFill><a:srgbClr val="${escapeXml(stripArgb(normalizeArgbColor(props.fill)))}"/></a:solidFill>`
    : props.fill?.toLowerCase() === "none"
      ? "<a:noFill/>"
      : "";
  const lineXml = props.line
    ? props.line.toLowerCase() === "none"
      ? "<a:ln><a:noFill/></a:ln>"
      : `<a:ln><a:solidFill><a:srgbClr val="${escapeXml(stripArgb(normalizeArgbColor(props.line)))}"/></a:solidFill></a:ln>`
    : "";
  const bodyPrAttrs = props.margin ? ` lIns="${Math.round(Number(props.margin) * 12700)}" rIns="${Math.round(Number(props.margin) * 12700)}" tIns="${Math.round(Number(props.margin) * 12700)}" bIns="${Math.round(Number(props.margin) * 12700)}"` : "";
  const paragraphProps = props.align
    ? `<a:pPr algn="${normalizeShapeAlign(props.align)}"/>`
    : "";
  const runProps = [
    props.size ? ` sz="${Math.round(Number(props.size) * 100)}"` : "",
    props.bold && isTruthy(props.bold) ? ` b="1"` : "",
    props.italic && isTruthy(props.italic) ? ` i="1"` : "",
  ].join("");
  const textFillXml = props.color ? `<a:solidFill><a:srgbClr val="${escapeXml(stripArgb(normalizeArgbColor(props.color)))}"/></a:solidFill>` : "";
  const fontXml = props.font ? `<a:latin typeface="${escapeXml(props.font)}"/><a:ea typeface="${escapeXml(props.font)}"/>` : "";
  const anchorXml = `  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>${fromCol}</xdr:col><xdr:row>${fromRow}</xdr:row></xdr:from>
    <xdr:to><xdr:col>${toCol}</xdr:col><xdr:row>${toRow}</xdr:row></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="${nextAnchorId}" name="${escapeXml(name)}"/><xdr:cNvSpPr/><xdr:nvPr/></xdr:nvSpPr>
      <xdr:spPr>${fillXml}${lineXml}</xdr:spPr>
      <xdr:txBody><a:bodyPr${bodyPrAttrs}/><a:lstStyle/><a:p>${paragraphProps}<a:r><a:rPr lang="en-US"${runProps}>${textFillXml}${fontXml}</a:rPr><a:t>${escapeXml(text)}</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;
  state.zip.set(drawingPath, Buffer.from(drawingXml.replace(/<\/xdr:wsDr>/, `${anchorXml}\n</xdr:wsDr>`), "utf8"));
  return getDrawingShapes(state, sheet).filter((item) => item.kind === "shape").at(-1) ?? { kind: "shape" as const, name, text };
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
    if (props.axismin !== undefined || props.min !== undefined) {
      xml = setAxisValue(xml, "minVal", props.axismin ?? props.min ?? "0");
    }
    if (props.axismax !== undefined || props.max !== undefined) {
      xml = setAxisValue(xml, "maxVal", props.axismax ?? props.max ?? "0");
    }
    if (props.majorunit !== undefined) {
      xml = setAxisValue(xml, "majorUnit", props.majorunit);
    }
    if (props.minorunit !== undefined) {
      xml = setAxisValue(xml, "minorUnit", props.minorunit);
    }
    if (props.axisnumfmt !== undefined || props.axisnumberformat !== undefined) {
      const formatCode = props.axisnumfmt ?? props.axisnumberformat ?? "General";
      if (/<c:numFmt\b[^>]*formatCode="[^"]+"/.test(xml)) {
        xml = xml.replace(/<c:numFmt\b[^>]*formatCode="[^"]+"[^>]*sourceLinked="[^"]+"\/>/, `<c:numFmt formatCode="${escapeXml(formatCode)}" sourceLinked="0"/>`);
      } else {
        xml = xml.replace(/(<c:valAx\b[\s\S]*?<c:axId\b[^>]*\/>)/, `$1<c:numFmt formatCode="${escapeXml(formatCode)}" sourceLinked="0"/>`);
      }
    }
    if (props.style !== undefined || props.styleid !== undefined) {
      const styleValue = Number(props.style ?? props.styleid);
      if (Number.isFinite(styleValue)) {
        if (/<c:style\b[^>]*val="[^"]+"/.test(xml)) {
          xml = xml.replace(/<c:style\b[^>]*val="[^"]+"\/>/, `<c:style val="${styleValue}"/>`);
        } else {
          xml = xml.replace(/<c:chartSpace\b[^>]*>/, `$&<c:style val="${styleValue}"/>`);
        }
      }
    }
    if (props.plotfill !== undefined || props.plotareafill !== undefined) {
      xml = setChartFill(xml, "plotArea", props.plotfill ?? props.plotareafill ?? "");
    }
    if (props.chartfill !== undefined || props.chartareafill !== undefined) {
      xml = setChartAreaFill(xml, props.chartfill ?? props.chartareafill ?? "");
    }
    if (props.colors !== undefined) {
      const colors = props.colors.split(",").map((value) => value.trim()).filter(Boolean);
      xml = setSeriesColors(xml, colors);
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

function setAxisValue(xml: string, tag: "minVal" | "maxVal" | "majorUnit" | "minorUnit", value: string) {
  const targetTag = `c:${tag}`;
  const axisPattern = /<c:valAx\b[\s\S]*?<\/c:valAx>/;
  const axisXml = axisPattern.exec(xml)?.[0];
  if (!axisXml) {
    return xml;
  }
  const nextAxisXml = new RegExp(`<${targetTag}\\b[^>]*val="[^"]+"\\s*\\/?>`).test(axisXml)
    ? axisXml.replace(new RegExp(`<${targetTag}\\b[^>]*val="[^"]+"\\s*\\/?>`), `<${targetTag} val="${escapeXml(value)}"/>`)
    : axisXml.replace(/<\/c:valAx>/, `<${targetTag} val="${escapeXml(value)}"/></c:valAx>`);
  return xml.replace(axisXml, nextAxisXml);
}

function setChartFill(xml: string, scope: "plotArea", color: string) {
  const fillXml = `<c:spPr><a:solidFill><a:srgbClr val="${escapeXml(color.replace(/^#/, "").toUpperCase())}"/></a:solidFill></c:spPr>`;
  const pattern = /<c:plotArea\b[\s\S]*?<\/c:plotArea>/;
  const plotArea = pattern.exec(xml)?.[0];
  if (!plotArea) return xml;
  const nextPlotArea = /<c:spPr\b[\s\S]*?<\/c:spPr>/.test(plotArea)
    ? plotArea.replace(/<c:spPr\b[\s\S]*?<\/c:spPr>/, fillXml)
    : plotArea.replace(/<c:plotArea\b[^>]*>/, `$&${fillXml}`);
  return xml.replace(plotArea, nextPlotArea);
}

function setChartAreaFill(xml: string, color: string) {
  const fillXml = `<c:spPr><a:solidFill><a:srgbClr val="${escapeXml(color.replace(/^#/, "").toUpperCase())}"/></a:solidFill></c:spPr>`;
  const chartSpacePrefix = /<c:chartSpace\b[\s\S]*?<c:chart\b/.exec(xml)?.[0] ?? "";
  if (/<c:spPr\b[\s\S]*?<\/c:spPr>/.test(chartSpacePrefix)) {
    const nextPrefix = chartSpacePrefix.replace(/<c:spPr\b[\s\S]*?<\/c:spPr>/, fillXml);
    return xml.replace(chartSpacePrefix, nextPrefix);
  }
  return xml.replace(/<c:chartSpace\b[^>]*>/, `$&${fillXml}`);
}

function setSeriesColors(xml: string, colors: string[]) {
  const seriesMatches = [...xml.matchAll(/<c:ser\b[\s\S]*?<\/c:ser>/g)];
  let nextXml = xml;
  for (const [index, series] of seriesMatches.entries()) {
    const color = colors[index];
    if (!color) continue;
    const normalized = color.replace(/^#/, "").toUpperCase();
    const shapeProps = `<c:spPr><a:solidFill><a:srgbClr val="${escapeXml(normalized)}"/></a:solidFill></c:spPr>`;
    const nextSeries = /<c:spPr\b[\s\S]*?<\/c:spPr>/.test(series[0])
      ? series[0].replace(/<c:spPr\b[\s\S]*?<\/c:spPr>/, shapeProps)
      : series[0].replace(/<c:tx>[\s\S]*?<\/c:tx>/, `$&${shapeProps}`);
    nextXml = nextXml.replace(series[0], nextSeries);
  }
  return nextXml;
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

function parseConditionalFormatting(sheetXml: string): ExcelConditionalFormattingModel[] {
  return [...sheetXml.matchAll(/<(?:\w+:)?conditionalFormatting\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?conditionalFormatting>/g)].flatMap((match) => {
    const sqref = parseAttr(match[1], "sqref") ?? "";
    const body = match[2];
    return [...body.matchAll(/<(?:\w+:)?cfRule\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?cfRule>/g)].map((ruleMatch) => {
      const attrs = ruleMatch[1];
      const ruleBody = ruleMatch[2];
      const type = (parseAttr(attrs, "type") ?? "").toLowerCase();
      const common = {
        sqref,
        ...(parseAttr(attrs, "priority") ? { priority: Number(parseAttr(attrs, "priority")) } : {}),
        ...(parseAttr(attrs, "dxfId") ? { dxfId: Number(parseAttr(attrs, "dxfId")) } : {}),
      };
      if (type === "databar") {
        const cfvo = [...ruleBody.matchAll(/<(?:\w+:)?cfvo\b([^>]*)\/?>/g)].map((item) => ({ type: parseAttr(item[1], "type"), val: parseAttr(item[1], "val") }));
        return {
          ...common,
          cfType: "databar" as const,
          ...(cfvo[0]?.type === "num" && cfvo[0]?.val ? { min: cfvo[0].val } : {}),
          ...(cfvo[1]?.type === "num" && cfvo[1]?.val ? { max: cfvo[1].val } : {}),
          ...(parseAttr(/<(?:\w+:)?color\b([^>]*)\/?>/.exec(ruleBody)?.[1] ?? "", "rgb") ? { color: parseAttr(/<(?:\w+:)?color\b([^>]*)\/?>/.exec(ruleBody)?.[1] ?? "", "rgb") } : {}),
        };
      }
      if (type === "colorscale") {
        const colors = [...ruleBody.matchAll(/<(?:\w+:)?color\b([^>]*)\/?>/g)].map((item) => parseAttr(item[1], "rgb")).filter(Boolean) as string[];
        return {
          ...common,
          cfType: "colorscale" as const,
          ...(colors[0] ? { minColor: colors[0] } : {}),
          ...(colors[1] && colors.length === 3 ? { midColor: colors[1] } : {}),
          ...(colors.at(-1) ? { maxColor: colors.at(-1)! } : {}),
        };
      }
      if (type === "iconset") {
        const iconAttrs = /<(?:\w+:)?iconSet\b([^>]*)>/.exec(ruleBody)?.[1] ?? "";
        return {
          ...common,
          cfType: "iconset" as const,
          ...(parseAttr(iconAttrs, "iconSet") ? { iconset: parseAttr(iconAttrs, "iconSet") } : {}),
          ...(parseAttr(iconAttrs, "reverse") !== undefined ? { reverse: isTruthy(parseAttr(iconAttrs, "reverse") ?? "false") } : {}),
          ...(parseAttr(iconAttrs, "showValue") !== undefined ? { showvalue: isTruthy(parseAttr(iconAttrs, "showValue") ?? "false") } : {}),
        };
      }
      if (type === "expression") {
        return {
          ...common,
          cfType: "formula" as const,
          ...(extractTagText(ruleBody, "formula") ? { formula: decodeXmlRecursive(extractTagText(ruleBody, "formula")!) } : {}),
        };
      }
      if (type === "top10") {
        return {
          ...common,
          cfType: "topn" as const,
          ...(parseAttr(attrs, "rank") ? { rank: Number(parseAttr(attrs, "rank")) } : {}),
          ...(parseAttr(attrs, "percent") !== undefined ? { percent: isTruthy(parseAttr(attrs, "percent") ?? "false") } : {}),
          ...(parseAttr(attrs, "bottom") !== undefined ? { bottom: isTruthy(parseAttr(attrs, "bottom") ?? "false") } : {}),
        };
      }
      if (type === "aboveaverage") {
        return {
          ...common,
          cfType: "aboveaverage" as const,
          ...(parseAttr(attrs, "aboveAverage") !== undefined ? { above: isTruthy(parseAttr(attrs, "aboveAverage") ?? "false") } : { above: true }),
        };
      }
      if (type === "uniquevalues") {
        return { ...common, cfType: "uniquevalues" as const };
      }
      if (type === "duplicatevalues") {
        return { ...common, cfType: "duplicatevalues" as const };
      }
      if (type === "containstext") {
        return {
          ...common,
          cfType: "containstext" as const,
          ...(parseAttr(attrs, "text") ? { text: decodeXml(parseAttr(attrs, "text") ?? "") } : {}),
          ...(extractTagText(ruleBody, "formula") ? { formula: decodeXmlRecursive(extractTagText(ruleBody, "formula")!) } : {}),
        };
      }
      if (type === "timeperiod") {
        return {
          ...common,
          cfType: "dateoccurring" as const,
          ...(parseAttr(attrs, "timePeriod") ? { period: decodeXml(parseAttr(attrs, "timePeriod") ?? "") } : {}),
        };
      }
      return {
        ...common,
        cfType: "databar" as const,
      };
    }).filter((item) => item.sqref);
  });
}

function replaceSheetConditionalFormatting(sheetXml: string, rules: ExcelConditionalFormattingModel[]) {
  const rendered = rules.length > 0
    ? rules.map(renderConditionalFormattingXml).join("")
    : "";
  const withoutExisting = sheetXml.replace(/<(?:\w+:)?conditionalFormatting\b[\s\S]*?<\/(?:\w+:)?conditionalFormatting>/g, "");
  return replaceOrInsert(withoutExisting, /<(?:\w+:)?conditionalFormatting\b[\s\S]*?<\/(?:\w+:)?conditionalFormatting>/, rendered, /<(?:\w+:)?dataValidations\b[\s\S]*?<\/(?:\w+:)?dataValidations>|<(?:\w+:)?autoFilter\b[^>]*\/?>|<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/);
}

function renderConditionalFormattingXml(rule: ExcelConditionalFormattingModel, index: number) {
  const priority = rule.priority ?? index + 1;
  if (rule.cfType === "databar") {
    return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="dataBar" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}><dataBar><cfvo type="${rule.min !== undefined ? "num" : "min"}"${rule.min !== undefined ? ` val="${escapeXml(rule.min)}"` : ""}/><cfvo type="${rule.max !== undefined ? "num" : "max"}"${rule.max !== undefined ? ` val="${escapeXml(rule.max)}"` : ""}/><color rgb="${escapeXml(rule.color ?? "FF638EC6")}"/></dataBar></cfRule></conditionalFormatting>`;
  }
  if (rule.cfType === "colorscale") {
    const mid = rule.midColor ? `<cfvo type="percentile" val="50"/><color rgb="${escapeXml(rule.midColor)}"/>` : "";
    return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="colorScale" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}><colorScale><cfvo type="min"/><cfvo type="max"/>${mid ? "" : ""}<color rgb="${escapeXml(rule.minColor ?? "FFF8696B")}"/>${mid}<color rgb="${escapeXml(rule.maxColor ?? "FF63BE7B")}"/></colorScale></cfRule></conditionalFormatting>`
      .replace("<cfvo type=\"min\"/><cfvo type=\"max\"/>", rule.midColor ? `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>` : `<cfvo type="min"/><cfvo type="max"/>`);
  }
  if (rule.cfType === "formula") {
    return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="expression" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}><formula>${escapeXml(rule.formula ?? "")}</formula></cfRule></conditionalFormatting>`;
  }
  if (rule.cfType === "topn") {
    return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="top10" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}${rule.rank !== undefined ? ` rank="${rule.rank}"` : ""}${rule.percent ? ' percent="1"' : ""}${rule.bottom ? ' bottom="1"' : ""}></cfRule></conditionalFormatting>`;
  }
  if (rule.cfType === "aboveaverage") {
    return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="aboveAverage" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}${rule.above === false ? ' aboveAverage="0"' : ""}></cfRule></conditionalFormatting>`;
  }
  if (rule.cfType === "uniquevalues") {
    return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="uniqueValues" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}></cfRule></conditionalFormatting>`;
  }
  if (rule.cfType === "duplicatevalues") {
    return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="duplicateValues" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}></cfRule></conditionalFormatting>`;
  }
  if (rule.cfType === "containstext") {
    const firstCell = rule.sqref.split(":")[0].replace(/\$/g, "");
    const formula = rule.formula ?? `NOT(ISERROR(SEARCH("${rule.text ?? ""}",${firstCell})))`;
    return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="containsText" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}${rule.text ? ` text="${escapeXml(rule.text)}"` : ""} operator="containsText"><formula>${escapeXml(formula)}</formula></cfRule></conditionalFormatting>`;
  }
  if (rule.cfType === "dateoccurring") {
    return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="timePeriod" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}${rule.period ? ` timePeriod="${escapeXml(normalizeTimePeriod(rule.period))}"` : ""}></cfRule></conditionalFormatting>`;
  }
  const iconSetName = rule.iconset ?? "3TrafficLights1";
  const iconCount = getIconSetThresholdCount(iconSetName);
  const thresholdXml = Array.from({ length: iconCount }, (_, itemIndex) => {
    const value = itemIndex === 0 ? "0" : String(Math.floor(itemIndex * 100 / iconCount));
    return `<cfvo type="percent" val="${value}"/>`;
  }).join("");
  return `<conditionalFormatting sqref="${escapeXml(rule.sqref)}"><cfRule type="iconSet" priority="${priority}"${rule.dxfId !== undefined ? ` dxfId="${rule.dxfId}"` : ""}><iconSet iconSet="${escapeXml(iconSetName)}"${rule.reverse ? ' reverse="1"' : ""}${rule.showvalue !== undefined ? ` showValue="${rule.showvalue ? 1 : 0}"` : ""}>${thresholdXml}</iconSet></cfRule></conditionalFormatting>`;
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
    const axisMin = /<c:minVal\b[^>]*val="([^"]+)"/.exec(chartXml)?.[1];
    const axisMax = /<c:maxVal\b[^>]*val="([^"]+)"/.exec(chartXml)?.[1];
    const majorUnit = /<c:majorUnit\b[^>]*val="([^"]+)"/.exec(chartXml)?.[1];
    const minorUnit = /<c:minorUnit\b[^>]*val="([^"]+)"/.exec(chartXml)?.[1];
    const axisNumberFormat = /<c:numFmt\b[^>]*formatCode="([^"]+)"/.exec(chartXml)?.[1];
    const styleId = /<c:style\b[^>]*val="([^"]+)"/.exec(chartXml)?.[1];
    const plotAreaXml = /<c:plotArea\b[\s\S]*?<\/c:plotArea>/.exec(chartXml)?.[0] ?? "";
    const chartSpacePrefix = /<c:chartSpace\b[\s\S]*?<c:chart\b/.exec(chartXml)?.[0] ?? "";
    const plotAreaFill = /<c:spPr\b[\s\S]*?<a:solidFill>\s*<a:srgbClr val="([^"]+)"/.exec(plotAreaXml)?.[1];
    const chartAreaFill = /<c:spPr\b[\s\S]*?<a:solidFill>\s*<a:srgbClr val="([^"]+)"/.exec(chartSpacePrefix)?.[1];
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
      ...(axisMin !== undefined ? { axisMin: Number(axisMin) } : {}),
      ...(axisMax !== undefined ? { axisMax: Number(axisMax) } : {}),
      ...(majorUnit !== undefined ? { majorUnit: Number(majorUnit) } : {}),
      ...(minorUnit !== undefined ? { minorUnit: Number(minorUnit) } : {}),
      ...(axisNumberFormat ? { axisNumberFormat } : {}),
      ...(styleId !== undefined ? { styleId: Number(styleId) } : {}),
      ...(plotAreaFill ? { plotAreaFill } : {}),
      ...(chartAreaFill ? { chartAreaFill } : {}),
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

function ensureDrawingPart(state: ExcelWorkbookState, sheet: ExcelSheetModel) {
  const existing = resolveDrawingPath(state, sheet);
  if (existing) return existing;
  const drawingPath = nextIndexedPartPath(state.zip, "xl/drawings/drawing", ".xml");
  state.zip.set(drawingPath, Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
</xdr:wsDr>`, "utf8"));
  sheet.xml = ensureWorksheetNamespaces(sheet.xml, {
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  });
  const relId = appendRelationship(
    state.zip,
    getRelationshipsEntryName(sheet.entryName),
    sheet.entryName,
    drawingPath,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
  );
  sheet.xml = replaceOrInsert(
    sheet.xml,
    /<(?:\w+:)?drawing\b[^>]*\/?>/,
    `<drawing r:id="${escapeXml(relId)}"/>`,
    /<(?:\w+:)?tableParts\b[\s\S]*?<\/(?:\w+:)?tableParts>|<(?:\w+:)?extLst\b[\s\S]*?<\/(?:\w+:)?extLst>|<(?:\w+:)?colBreaks\b[\s\S]*?<\/(?:\w+:)?colBreaks>|<(?:\w+:)?rowBreaks\b[\s\S]*?<\/(?:\w+:)?rowBreaks>|<(?:\w+:)?headerFooter\b[\s\S]*?<\/(?:\w+:)?headerFooter>|<(?:\w+:)?pageSetup\b[^>]*\/?>|<(?:\w+:)?sheetProtection\b[^>]*\/?>|<(?:\w+:)?autoFilter\b[^>]*\/?>|<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/,
  );
  return drawingPath;
}

function nextDrawingObjectId(drawingXml: string) {
  return Math.max(0, ...[...drawingXml.matchAll(/<xdr:cNvPr\b[^>]*id="(\d+)"/g)].map((match) => Number(match[1]))) + 1;
}

function resolveAnchorBounds(
  props: Record<string, string>,
  defaults: { x: number; y: number; width: number; height: number },
) {
  const fromCol = Number(props.x ?? defaults.x);
  const fromRow = Number(props.y ?? defaults.y);
  const width = Number(props.width ?? defaults.width);
  const height = Number(props.height ?? defaults.height);
  return {
    fromCol,
    fromRow,
    toCol: fromCol + width,
    toRow: fromRow + height,
  };
}

function normalizeImageExtension(extension: string) {
  const normalized = extension.replace(/^\./, "").toLowerCase();
  if (normalized === "jpg") return "jpeg";
  if (["png", "jpeg", "gif"].includes(normalized)) return normalized;
  throw new UsageError(`Unsupported Excel picture extension '.${normalized || "unknown"}'.`, "Use png, jpg/jpeg, or gif.");
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
      || lower === "bgcolor"
      || lower === "numfmt"
      || lower === "format"
      || lower === "numberformat"
      || lower === "border"
      || lower === "wrap"
      || lower === "wraptext"
      || lower === "halign"
      || lower === "valign"
      || lower === "rotation"
      || lower === "indent"
      || lower === "shrinktofit"
      || lower === "locked"
      || lower === "formulahidden"
      || lower.startsWith("font.")
      || lower.startsWith("alignment.")
      || lower.startsWith("border.");
  });
}

function registerStyle(state: ExcelWorkbookState, props: Record<string, string>) {
  const stylesheet = parseStylesheet(state.styleSheetXml ?? buildDefaultStylesheetXml());
  const fontXml = buildFontXml(props);
  const fillXml = buildFillXml(props);
  const borderXml = buildBorderXml(props);
  const numFmtCode = props.numFmt ?? props.numfmt ?? props.format ?? props.numberformat;
  const numFmtId = numFmtCode ? ensureNumFmt(stylesheet, numFmtCode) : 0;
  const fontId = ensureFragment(stylesheet.fonts, fontXml, `<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`);
  const fillId = ensureFragment(stylesheet.fills, fillXml, `<fill><patternFill patternType="none"/></fill>`);
  const borderId = ensureFragment(stylesheet.borders, borderXml, `<border><left/><right/><top/><bottom/><diagonal/></border>`);
  const alignmentXml = buildAlignmentXml(props);
  const protectionXml = buildProtectionXml(props);
  const xfXml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''}${fontXml !== DEFAULT_FONT_XML ? ' applyFont="1"' : ''}${fillXml !== DEFAULT_FILL_XML ? ' applyFill="1"' : ''}${borderXml !== `<border><left/><right/><top/><bottom/><diagonal/></border>` ? ' applyBorder="1"' : ''}${alignmentXml ? ' applyAlignment="1"' : ''}${protectionXml ? ' applyProtection="1"' : ''}>${alignmentXml}${protectionXml}</xf>`;
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

function evaluateFormulaForDisplay(state: ExcelWorkbookState | undefined, sheet: ExcelSheetModel, ref: string) {
  const formula = sheet.cells[ref]?.formula;
  if (formula) {
    const normalized = normalizeFormula(formula);
    const ifMatch = /^IF\((.*)\)$/i.exec(normalized.trim());
    if (ifMatch) {
      const result = evaluateIfFast(state, ifMatch[1], sheet, new Set());
      if (result !== undefined) return String(result);
    }
    const textFormula = evaluateTextFormulaForDisplay(state, normalized.trim(), sheet);
    if (textFormula !== undefined) {
      return textFormula;
    }
    const lookupFormula = evaluateLookupFormulaForDisplay(state, normalized.trim(), sheet);
    if (lookupFormula !== undefined) {
      return lookupFormula;
    }
    const countaMatch = /^COUNTA\((.*)\)$/i.exec(normalized.trim());
    if (countaMatch) {
      return String(countFormulaArgs(state, countaMatch[1], sheet, new Set()));
    }
    const sumProductMatch = /^SUMPRODUCT\((.*)\)$/i.exec(normalized.trim());
    if (sumProductMatch) {
      return String(sumProductFormulaArgs(state, sumProductMatch[1], sheet, new Set()));
    }
    const conditionalAggregation = evaluateConditionalAggregationFormula(state, normalized.trim(), sheet);
    if (conditionalAggregation !== undefined) {
      return conditionalAggregation;
    }
  }
  const visited = new Set<string>();
  const numeric = evaluateFormulaExpression(state, sheet, ref, visited);
  if (numeric === undefined || Number.isNaN(numeric)) {
    return undefined;
  }
  return Number.isInteger(numeric) ? String(numeric) : String(Number(numeric.toFixed(10)));
}

function evaluateIfFast(state: ExcelWorkbookState | undefined, args: string, sheet: ExcelSheetModel, visited: Set<string>) {
  const [condition, whenTrue = "0", whenFalse = "0"] = splitFormulaArgs(args);
  if (!condition) return undefined;
  const truthy = evaluateCondition(state, condition.trim(), sheet, visited);
  if (truthy === undefined) {
    return undefined;
  }
  return truthy
    ? evaluateInlineFormulaArg(state, whenTrue.trim(), sheet, visited)
    : evaluateInlineFormulaArg(state, whenFalse.trim(), sheet, visited);
}

function evaluateFormulaExpression(
  state: ExcelWorkbookState | undefined,
  sheet: ExcelSheetModel,
  ref: string,
  visited: Set<string>,
): number | undefined {
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
  const ifMatch = /^IF\((.*)\)$/i.exec(expression.trim());
  if (ifMatch) {
    const result = evaluateIfFormula(state, ifMatch[1], sheet, visited);
    visited.delete(key);
    return result;
  }
  const aggregateMatch = /^(SUM|AVERAGE|MIN|MAX)\(([^()]+)\)$/i.exec(expression.trim());
  if (aggregateMatch) {
    const result = foldFormulaArgs(state, aggregateMatch[2], sheet, visited, aggregateMatch[1].toUpperCase() === "SUM"
      ? (values: number[]) => values.reduce((sum: number, value: number) => sum + value, 0)
      : aggregateMatch[1].toUpperCase() === "AVERAGE"
        ? (values: number[]) => values.length > 0 ? values.reduce((sum: number, value: number) => sum + value, 0) / values.length : 0
        : aggregateMatch[1].toUpperCase() === "MIN"
          ? (values: number[]) => values.length > 0 ? Math.min(...values) : 0
          : (values: number[]) => values.length > 0 ? Math.max(...values) : 0);
    visited.delete(key);
    return result;
  }
  const functionEvaluators: Record<string, (args: string) => number | undefined> = {
    SUM: (args) => foldFormulaArgs(state, args, sheet, visited, (values) => values.reduce((sum, value) => sum + value, 0)),
    AVERAGE: (args) => foldFormulaArgs(state, args, sheet, visited, (values) => values.length > 0 ? values.reduce((sum, value) => sum + value, 0) / values.length : 0),
    MIN: (args) => foldFormulaArgs(state, args, sheet, visited, (values) => values.length > 0 ? Math.min(...values) : 0),
    MAX: (args) => foldFormulaArgs(state, args, sheet, visited, (values) => values.length > 0 ? Math.max(...values) : 0),
    COUNT: (args) => countNumericFormulaArgs(state, args, sheet, visited),
    COUNTA: (args) => countFormulaArgs(state, args, sheet, visited),
    SUMPRODUCT: (args) => sumProductFormulaArgs(state, args, sheet, visited),
    IF: (args) => evaluateIfFormula(state, args, sheet, visited),
    ABS: (args) => {
      const value = firstNumericFormulaArg(state, args, sheet, visited);
      return value === undefined ? undefined : Math.abs(value);
    },
    ROUND: (args) => evaluateRoundFormula(state, args, sheet, visited, "round"),
    ROUNDUP: (args) => evaluateRoundFormula(state, args, sheet, visited, "up"),
    ROUNDDOWN: (args) => evaluateRoundFormula(state, args, sheet, visited, "down"),
    MOD: (args) => evaluateBinaryNumericFormula(state, args, sheet, visited, (left, right) => right === 0 ? undefined : left - right * Math.floor(left / right)),
    POWER: (args) => evaluateBinaryNumericFormula(state, args, sheet, visited, (left, right) => Math.pow(left, right)),
    SQRT: (args) => {
      const value = firstNumericFormulaArg(state, args, sheet, visited);
      return value === undefined || value < 0 ? undefined : Math.sqrt(value);
    },
  };

  let replaced = true;
  while (replaced) {
    replaced = false;
    expression = expression.replace(/\b(SUM|AVERAGE|MIN|MAX|COUNT|COUNTA|SUMPRODUCT|IF|ABS|ROUND|ROUNDUP|ROUNDDOWN|MOD|POWER|SQRT)\(([^()]*)\)/gi, (match, fn, args) => {
      const result = functionEvaluators[fn.toUpperCase()]?.(args);
      if (result === undefined) {
        return match;
      }
      replaced = true;
      return String(result);
    });
  }

  expression = expression.replace(/\b([A-Z]+[0-9]+)\b/g, (match, refValue) => {
    const value = evaluateFormulaExpression(state, sheet, refValue.toUpperCase(), visited);
    return value === undefined ? match : String(value);
  });

  expression = expression.replace(/(?:'([^']+)'|([A-Za-z0-9_ ]+))!([A-Z]+\d+)/g, (match, quotedSheet, plainSheet, refValue) => {
    const targetSheetName = (quotedSheet ?? plainSheet ?? "").trim();
    const targetSheet = state?.sheets.find((item) => item.name.toLowerCase() === targetSheetName.toLowerCase());
    if (!targetSheet) return match;
    const value = evaluateFormulaExpression(state, targetSheet, refValue.toUpperCase(), visited);
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
  state: ExcelWorkbookState | undefined,
  args: string,
  sheet: ExcelSheetModel,
  visited: Set<string>,
  reducer: (values: number[]) => number,
) {
  const values = extractFormulaArgValues(state, args, sheet, visited);
  return reducer(values);
}

function splitFormulaArgs(args: string) {
  const values: string[] = [];
  let current = "";
  let depth = 0;
  let inString = false;
  for (let index = 0; index < args.length; index += 1) {
    const char = args[index];
    if (char === '"') {
      inString = !inString;
      current += char;
      continue;
    }
    if (!inString && char === "(") depth += 1;
    if (!inString && char === ")") depth -= 1;
    if (!inString && depth === 0 && char === ",") {
      values.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  if (current !== "") values.push(current);
  return values;
}

function countFormulaArgs(state: ExcelWorkbookState | undefined, args: string, sheet: ExcelSheetModel, visited: Set<string>) {
  return splitFormulaArgs(args)
    .flatMap((part): Array<string | ExcelCellModel> => {
      const value = part.trim();
      if (/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/i.test(value)) {
        return expandRange(value.toUpperCase()).map((ref) => sheet.cells[ref]).filter(Boolean);
      }
      if (/^[A-Z]+[0-9]+$/i.test(value)) {
        return [sheet.cells[value.toUpperCase()]].filter(Boolean);
      }
      const crossSheetRange = /^(?:'([^']+)'|([A-Za-z0-9_ ]+))!([A-Z]+\d+):([A-Z]+\d+)$/i.exec(value);
      if (crossSheetRange) {
        const targetSheetName = (crossSheetRange[1] ?? crossSheetRange[2] ?? "").trim();
        const targetSheet = state?.sheets.find((item) => item.name.toLowerCase() === targetSheetName.toLowerCase());
        return targetSheet ? expandRange(`${crossSheetRange[3].toUpperCase()}:${crossSheetRange[4].toUpperCase()}`).map((ref) => targetSheet.cells[ref]).filter(Boolean) : [];
      }
      return value !== "" ? [value] : [];
    })
    .filter((value: string | ExcelCellModel) => typeof value === "string" ? value !== "" : (value.value ?? "") !== "")
    .length;
}

function sumProductFormulaArgs(state: ExcelWorkbookState | undefined, args: string, sheet: ExcelSheetModel, visited: Set<string>) {
  const arrays = splitFormulaArgs(args).map((part) => extractFormulaArgValues(state, part, sheet, visited));
  if (arrays.length === 0) return 0;
  const length = Math.min(...arrays.map((group) => group.length));
  let total = 0;
  for (let index = 0; index < length; index += 1) {
    total += arrays.reduce((product, group) => product * (group[index] ?? 0), 1);
  }
  return total;
}

function countNumericFormulaArgs(state: ExcelWorkbookState | undefined, args: string, sheet: ExcelSheetModel, visited: Set<string>) {
  return extractFormulaArgValues(state, args, sheet, visited).length;
}

function firstNumericFormulaArg(state: ExcelWorkbookState | undefined, args: string, sheet: ExcelSheetModel, visited: Set<string>) {
  return extractFormulaArgValues(state, args, sheet, visited)[0];
}

function evaluateBinaryNumericFormula(
  state: ExcelWorkbookState | undefined,
  args: string,
  sheet: ExcelSheetModel,
  visited: Set<string>,
  evaluator: (left: number, right: number) => number | undefined,
) {
  const values = splitFormulaArgs(args).map((part) => firstNumericFormulaArg(state, part, sheet, visited));
  if (values[0] === undefined || values[1] === undefined) {
    return undefined;
  }
  return evaluator(values[0], values[1]);
}

function evaluateRoundFormula(
  state: ExcelWorkbookState | undefined,
  args: string,
  sheet: ExcelSheetModel,
  visited: Set<string>,
  mode: "round" | "up" | "down",
) {
  const parts = splitFormulaArgs(args);
  const value = firstNumericFormulaArg(state, parts[0] ?? "", sheet, visited);
  const digits = firstNumericFormulaArg(state, parts[1] ?? "0", sheet, visited) ?? 0;
  if (value === undefined) {
    return undefined;
  }
  const factor = Math.pow(10, digits);
  if (mode === "up") {
    return Math.ceil(Math.abs(value) * factor) / factor * Math.sign(value);
  }
  if (mode === "down") {
    return Math.floor(Math.abs(value) * factor) / factor * Math.sign(value);
  }
  return Math.round(value * factor) / factor;
}

function evaluateTextFormulaForDisplay(state: ExcelWorkbookState | undefined, expression: string, sheet: ExcelSheetModel) {
  const direct = /^(LEN|LEFT|RIGHT|MID|LOWER|UPPER|TRIM|CONCAT|CONCATENATE)\((.*)\)$/i.exec(expression);
  if (!direct) return undefined;
  const fn = direct[1].toUpperCase();
  const args = splitFormulaArgs(direct[2]);
  const resolveText = (part: string) => evaluateTextFormulaArg(state, part.trim(), sheet);
  if (fn === "LEN") {
    return String(resolveText(args[0] ?? "").length);
  }
  if (fn === "LEFT") {
    const text = resolveText(args[0] ?? "");
    const count = Number(firstNumericFormulaArg(state, args[1] ?? "1", sheet, new Set()) ?? 1);
    return text.slice(0, count);
  }
  if (fn === "RIGHT") {
    const text = resolveText(args[0] ?? "");
    const count = Number(firstNumericFormulaArg(state, args[1] ?? "1", sheet, new Set()) ?? 1);
    return count >= text.length ? text : text.slice(-count);
  }
  if (fn === "MID") {
    const text = resolveText(args[0] ?? "");
    const start = Math.max(1, Number(firstNumericFormulaArg(state, args[1] ?? "1", sheet, new Set()) ?? 1)) - 1;
    const count = Math.max(0, Number(firstNumericFormulaArg(state, args[2] ?? "0", sheet, new Set()) ?? 0));
    return text.substring(start, start + count);
  }
  if (fn === "LOWER") return resolveText(args[0] ?? "").toLowerCase();
  if (fn === "UPPER") return resolveText(args[0] ?? "").toUpperCase();
  if (fn === "TRIM") return resolveText(args[0] ?? "").trim().replace(/\s+/g, " ");
  if (fn === "CONCAT" || fn === "CONCATENATE") {
    return args.map((part) => resolveText(part)).join("");
  }
  return undefined;
}

function evaluateTextFormulaArg(state: ExcelWorkbookState | undefined, arg: string, sheet: ExcelSheetModel) {
  if (/^".*"$/.test(arg)) {
    return arg.slice(1, -1);
  }
  const cellMatch = /^(?:'([^']+)'|([A-Za-z0-9_ ]+))!([A-Z]+\d+)$/i.exec(arg);
  if (cellMatch) {
    const targetSheetName = (cellMatch[1] ?? cellMatch[2] ?? "").trim();
    const targetSheet = state?.sheets.find((item) => item.name.toLowerCase() === targetSheetName.toLowerCase());
    return targetSheet?.cells[cellMatch[3].toUpperCase()]?.value ?? "";
  }
  if (/^[A-Z]+\d+$/i.test(arg)) {
    return sheet.cells[arg.toUpperCase()]?.value ?? "";
  }
  return arg;
}

function evaluateIfFormula(state: ExcelWorkbookState | undefined, args: string, sheet: ExcelSheetModel, visited: Set<string>) {
  const [condition, whenTrue = "0", whenFalse = "0"] = splitFormulaArgs(args);
  if (!condition) return undefined;
  const conditionExpression = replaceFormulaRefsWithValues(state, condition, sheet, visited);
  try {
    // eslint-disable-next-line no-new-func
    const truthy = Boolean(Function(`return (${conditionExpression});`)());
    return truthy
      ? evaluateInlineFormulaArg(state, whenTrue.trim(), sheet, visited)
      : evaluateInlineFormulaArg(state, whenFalse.trim(), sheet, visited);
  } catch {
    return undefined;
  }
}

function replaceFormulaRefsWithValues(
  state: ExcelWorkbookState | undefined,
  expression: string,
  sheet: ExcelSheetModel,
  visited: Set<string>,
) {
  return expression
    .replace(/(?:'([^']+)'|([A-Za-z0-9_ ]+))!([A-Z]+\d+)/g, (match, quotedSheet, plainSheet, refValue) => {
      const targetSheetName = (quotedSheet ?? plainSheet ?? "").trim();
      const targetSheet = state?.sheets.find((item) => item.name.toLowerCase() === targetSheetName.toLowerCase());
      return String(targetSheet ? (evaluateFormulaExpression(state, targetSheet, refValue.toUpperCase(), visited) ?? match) : match);
    })
    .replace(/\b([A-Z]+[0-9]+)\b/g, (match, refValue) => String(evaluateFormulaExpression(state, sheet, refValue.toUpperCase(), visited) ?? match));
}

function evaluateInlineFormulaArg(state: ExcelWorkbookState | undefined, arg: string, sheet: ExcelSheetModel, visited: Set<string>) {
  if (/^[A-Z]+[0-9]+$/i.test(arg)) {
    return evaluateFormulaExpression(state, sheet, arg.toUpperCase(), visited);
  }
  const crossSheetCell = /^(?:'([^']+)'|([A-Za-z0-9_ ]+))!([A-Z]+\d+)$/i.exec(arg);
  if (crossSheetCell) {
    const targetSheetName = (crossSheetCell[1] ?? crossSheetCell[2] ?? "").trim();
    const targetSheet = state?.sheets.find((item) => item.name.toLowerCase() === targetSheetName.toLowerCase());
    return targetSheet ? evaluateFormulaExpression(state, targetSheet, crossSheetCell[3].toUpperCase(), visited) : undefined;
  }
  const numeric = Number(arg.replace(/^"|"$/g, ""));
  return Number.isFinite(numeric) ? numeric : undefined;
}

function resolveScalarValue(state: ExcelWorkbookState | undefined, arg: string, sheet: ExcelSheetModel) {
  const trimmed = arg.trim();
  const crossSheetCell = /^(?:'([^']+)'|([A-Za-z0-9_ ]+))!([A-Z]+\d+)$/i.exec(trimmed);
  if (crossSheetCell) {
    const targetSheetName = (crossSheetCell[1] ?? crossSheetCell[2] ?? "").trim();
    const targetSheet = state?.sheets.find((item) => item.name.toLowerCase() === targetSheetName.toLowerCase());
    if (!targetSheet) return undefined;
    return formatResolvedCellValue(targetSheet.cells[crossSheetCell[3].toUpperCase()]);
  }
  if (/^[A-Z]+\d+$/i.test(trimmed)) {
    return formatResolvedCellValue(sheet.cells[trimmed.toUpperCase()]);
  }
  if (/^".*"$/.test(trimmed)) {
    return trimmed.slice(1, -1);
  }
  return trimmed;
}

function resolveRangeReference(state: ExcelWorkbookState | undefined, arg: string, sheet: ExcelSheetModel) {
  const trimmed = arg.trim();
  const crossSheetRange = /^(?:'([^']+)'|([A-Za-z0-9_ ]+))!([A-Z]+\d+):([A-Z]+\d+)$/i.exec(trimmed);
  if (crossSheetRange) {
    const targetSheetName = (crossSheetRange[1] ?? crossSheetRange[2] ?? "").trim();
    const targetSheet = state?.sheets.find((item) => item.name.toLowerCase() === targetSheetName.toLowerCase());
    if (!targetSheet) return undefined;
    return buildRangeMatrix(targetSheet, `${crossSheetRange[3].toUpperCase()}:${crossSheetRange[4].toUpperCase()}`);
  }
  if (/^[A-Z]+\d+:[A-Z]+\d+$/i.test(trimmed)) {
    return buildRangeMatrix(sheet, trimmed.toUpperCase());
  }
  return undefined;
}

function buildRangeMatrix(sheet: ExcelSheetModel, range: string) {
  const refs = expandRange(range);
  const [startRef, endRef = startRef] = range.split(":");
  const start = parseCellAddress(startRef);
  const end = parseCellAddress(endRef);
  const rows = end.row - start.row + 1;
  const cols = columnNameToIndex(end.column) - columnNameToIndex(start.column) + 1;
  const cells: ExcelCellModel[][] = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const row: ExcelCellModel[] = [];
    for (let colIndex = 0; colIndex < cols; colIndex += 1) {
      row.push(sheet.cells[refs[rowIndex * cols + colIndex]] ?? { value: "" });
    }
    cells.push(row);
  }
  return { rows, cols, cells };
}

function flattenRange(range: { cells: ExcelCellModel[][] }) {
  return range.cells.flat();
}

function formatResolvedCellValue(cell?: ExcelCellModel) {
  if (!cell) return "";
  if (cell.type === "boolean") {
    return cell.value === "1" ? "TRUE" : "FALSE";
  }
  return cell.value;
}

function compareFormulaValues(left: string, right: string) {
  const leftNumber = Number(left);
  const rightNumber = Number(right);
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) {
    return leftNumber === rightNumber ? 0 : leftNumber < rightNumber ? -1 : 1;
  }
  return left === right ? 0 : left.localeCompare(right);
}

function matchesFormulaCriteria(cell: ExcelCellModel, criteria: string) {
  const value = formatResolvedCellValue(cell);
  const directNumber = Number(criteria);
  if (Number.isFinite(directNumber)) {
    return Number(value) === directNumber;
  }
  const operatorMatch = /^(<=|>=|<>|=|<|>)(.*)$/.exec(criteria);
  if (!operatorMatch) {
    return value === criteria;
  }
  const operand = operatorMatch[2].trim().replace(/^"|"$/g, "");
  const leftNumber = Number(value);
  const rightNumber = Number(operand);
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) {
    switch (operatorMatch[1]) {
      case "<": return leftNumber < rightNumber;
      case "<=": return leftNumber <= rightNumber;
      case ">": return leftNumber > rightNumber;
      case ">=": return leftNumber >= rightNumber;
      case "<>": return leftNumber !== rightNumber;
      default: return leftNumber === rightNumber;
    }
  }
  switch (operatorMatch[1]) {
    case "<>": return value !== operand;
    case "=": return value === operand;
    default: return false;
  }
}

function evaluateLookupFormulaForDisplay(state: ExcelWorkbookState | undefined, formula: string, sheet: ExcelSheetModel) {
  const indexMatch = /^INDEX\((.+)\)$/i.exec(formula);
  if (indexMatch) {
    const args = splitFormulaArgs(indexMatch[1]);
    if (args.length < 2) return undefined;
    const range = resolveRangeReference(state, args[0].trim(), sheet);
    const rowIndex = Number(resolveScalarValue(state, args[1].trim(), sheet) ?? NaN);
    const colIndex = args[2] ? Number(resolveScalarValue(state, args[2].trim(), sheet) ?? NaN) : 1;
    if (!range || !Number.isFinite(rowIndex) || !Number.isFinite(colIndex)) return undefined;
    const cell = range.cells[rowIndex - 1]?.[colIndex - 1];
    return cell ? formatResolvedCellValue(cell) : undefined;
  }

  const matchMatch = /^MATCH\((.+)\)$/i.exec(formula);
  if (matchMatch) {
    const args = splitFormulaArgs(matchMatch[1]);
    if (args.length < 2) return undefined;
    const lookupValue = resolveScalarValue(state, args[0].trim(), sheet);
    const range = resolveRangeReference(state, args[1].trim(), sheet);
    if (lookupValue === undefined || !range) return undefined;
    const flat = range.rows === 1
      ? range.cells[0]
      : range.cols === 1
        ? range.cells.map((row) => row[0])
        : [];
    const index = flat.findIndex((cell) => compareFormulaValues(formatResolvedCellValue(cell), lookupValue) === 0);
    return index >= 0 ? String(index + 1) : undefined;
  }

  const vlookupMatch = /^VLOOKUP\((.+)\)$/i.exec(formula);
  if (vlookupMatch) {
    const args = splitFormulaArgs(vlookupMatch[1]);
    if (args.length < 3) return undefined;
    const lookupValue = resolveScalarValue(state, args[0].trim(), sheet);
    const range = resolveRangeReference(state, args[1].trim(), sheet);
    const columnIndex = Number(resolveScalarValue(state, args[2].trim(), sheet) ?? NaN);
    const exact = args[3] ? !isTruthy(String(resolveScalarValue(state, args[3].trim(), sheet) ?? "true")) : false;
    if (lookupValue === undefined || !range || !Number.isFinite(columnIndex)) return undefined;
    let foundRow = -1;
    for (let rowIndex = 0; rowIndex < range.rows; rowIndex += 1) {
      const firstCell = range.cells[rowIndex][0];
      const compare = compareFormulaValues(formatResolvedCellValue(firstCell), lookupValue);
      if (exact) {
        if (compare === 0) {
          foundRow = rowIndex;
          break;
        }
      } else if (compare <= 0) {
        foundRow = rowIndex;
      } else {
        break;
      }
    }
    const cell = foundRow >= 0 ? range.cells[foundRow]?.[columnIndex - 1] : undefined;
    return cell ? formatResolvedCellValue(cell) : undefined;
  }

  const hlookupMatch = /^HLOOKUP\((.+)\)$/i.exec(formula);
  if (hlookupMatch) {
    const args = splitFormulaArgs(hlookupMatch[1]);
    if (args.length < 3) return undefined;
    const lookupValue = resolveScalarValue(state, args[0].trim(), sheet);
    const range = resolveRangeReference(state, args[1].trim(), sheet);
    const rowIndex = Number(resolveScalarValue(state, args[2].trim(), sheet) ?? NaN);
    const exact = args[3] ? !isTruthy(String(resolveScalarValue(state, args[3].trim(), sheet) ?? "true")) : false;
    if (lookupValue === undefined || !range || !Number.isFinite(rowIndex)) return undefined;
    let foundCol = -1;
    for (let colIndex = 0; colIndex < range.cols; colIndex += 1) {
      const firstCell = range.cells[0][colIndex];
      const compare = compareFormulaValues(formatResolvedCellValue(firstCell), lookupValue);
      if (exact) {
        if (compare === 0) {
          foundCol = colIndex;
          break;
        }
      } else if (compare <= 0) {
        foundCol = colIndex;
      } else {
        break;
      }
    }
    const cell = foundCol >= 0 ? range.cells[rowIndex - 1]?.[foundCol] : undefined;
    return cell ? formatResolvedCellValue(cell) : undefined;
  }

  return undefined;
}

function evaluateConditionalAggregationFormula(state: ExcelWorkbookState | undefined, formula: string, sheet: ExcelSheetModel) {
  const sumIfMatch = /^SUMIF\((.+)\)$/i.exec(formula);
  if (sumIfMatch) {
    const args = splitFormulaArgs(sumIfMatch[1]);
    if (args.length < 2) return undefined;
    const criteriaRange = resolveRangeReference(state, args[0].trim(), sheet);
    const criteria = resolveScalarValue(state, args[1].trim(), sheet);
    const sumRange = args[2] ? resolveRangeReference(state, args[2].trim(), sheet) ?? criteriaRange : criteriaRange;
    if (!criteriaRange || !sumRange || criteria === undefined) return undefined;
    let total = 0;
    for (let index = 0; index < Math.min(flattenRange(criteriaRange).length, flattenRange(sumRange).length); index += 1) {
      if (matchesFormulaCriteria(flattenRange(criteriaRange)[index], criteria)) {
        total += Number(formatResolvedCellValue(flattenRange(sumRange)[index]) || 0);
      }
    }
    return String(total);
  }

  const countIfMatch = /^COUNTIF\((.+)\)$/i.exec(formula);
  if (countIfMatch) {
    const args = splitFormulaArgs(countIfMatch[1]);
    if (args.length < 2) return undefined;
    const criteriaRange = resolveRangeReference(state, args[0].trim(), sheet);
    const criteria = resolveScalarValue(state, args[1].trim(), sheet);
    if (!criteriaRange || criteria === undefined) return undefined;
    const count = flattenRange(criteriaRange).filter((cell) => matchesFormulaCriteria(cell, criteria)).length;
    return String(count);
  }

  const averageIfMatch = /^AVERAGEIF\((.+)\)$/i.exec(formula);
  if (averageIfMatch) {
    const args = splitFormulaArgs(averageIfMatch[1]);
    if (args.length < 2) return undefined;
    const criteriaRange = resolveRangeReference(state, args[0].trim(), sheet);
    const criteria = resolveScalarValue(state, args[1].trim(), sheet);
    const averageRange = args[2] ? resolveRangeReference(state, args[2].trim(), sheet) ?? criteriaRange : criteriaRange;
    if (!criteriaRange || !averageRange || criteria === undefined) return undefined;
    const values: number[] = [];
    const criteriaCells = flattenRange(criteriaRange);
    const averageCells = flattenRange(averageRange);
    for (let index = 0; index < Math.min(criteriaCells.length, averageCells.length); index += 1) {
      if (matchesFormulaCriteria(criteriaCells[index], criteria)) {
        const numeric = Number(formatResolvedCellValue(averageCells[index]));
        if (Number.isFinite(numeric)) values.push(numeric);
      }
    }
    if (values.length === 0) return undefined;
    return String(values.reduce((sum, value) => sum + value, 0) / values.length);
  }

  return undefined;
}

function evaluateCondition(state: ExcelWorkbookState | undefined, condition: string, sheet: ExcelSheetModel, visited: Set<string>) {
  const match = /^(.*?)(>=|<=|<>|=|>|<)(.*)$/.exec(condition);
  if (!match) {
    const direct = evaluateInlineFormulaArg(state, condition, sheet, visited);
    return direct !== undefined ? direct !== 0 : undefined;
  }
  const left = evaluateConditionOperand(state, match[1].trim(), sheet, visited);
  const right = evaluateConditionOperand(state, match[3].trim(), sheet, visited);
  if (left === undefined || right === undefined) return undefined;
  switch (match[2]) {
    case ">":
      return left > right;
    case "<":
      return left < right;
    case ">=":
      return left >= right;
    case "<=":
      return left <= right;
    case "=":
      return left === right;
    case "<>":
      return left !== right;
    default:
      return undefined;
  }
}

function evaluateConditionOperand(state: ExcelWorkbookState | undefined, operand: string, sheet: ExcelSheetModel, visited: Set<string>) {
  const direct = evaluateInlineFormulaArg(state, operand, sheet, visited);
  if (direct !== undefined) return direct;
  const stripped = operand.replace(/^"|"$/g, "");
  const numeric = Number(stripped);
  return Number.isFinite(numeric) ? numeric : undefined;
}

function extractFormulaArgValues(state: ExcelWorkbookState | undefined, args: string, sheet: ExcelSheetModel, visited: Set<string>) {
  return splitFormulaArgs(args)
    .flatMap((part) => {
      const value = part.trim();
      const crossSheetRange = /^(?:'([^']+)'|([A-Za-z0-9_ ]+))!([A-Z]+\d+):([A-Z]+\d+)$/i.exec(value);
      if (crossSheetRange) {
        const targetSheetName = (crossSheetRange[1] ?? crossSheetRange[2] ?? "").trim();
        const targetSheet = state?.sheets.find((item) => item.name.toLowerCase() === targetSheetName.toLowerCase());
        if (!targetSheet) return [];
        return expandRange(`${crossSheetRange[3].toUpperCase()}:${crossSheetRange[4].toUpperCase()}`)
          .map((ref) => evaluateFormulaExpression(state, targetSheet, ref, visited))
          .filter((item): item is number => item !== undefined);
      }
      const crossSheetCell = /^(?:'([^']+)'|([A-Za-z0-9_ ]+))!([A-Z]+\d+)$/i.exec(value);
      if (crossSheetCell) {
        const targetSheetName = (crossSheetCell[1] ?? crossSheetCell[2] ?? "").trim();
        const targetSheet = state?.sheets.find((item) => item.name.toLowerCase() === targetSheetName.toLowerCase());
        if (!targetSheet) return [];
        const evaluated = evaluateFormulaExpression(state, targetSheet, crossSheetCell[3].toUpperCase(), visited);
        return evaluated !== undefined ? [evaluated] : [];
      }
      if (/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/i.test(value)) {
        return expandRange(value.toUpperCase()).map((ref) => evaluateFormulaExpression(state, sheet, ref, visited)).filter((item): item is number => item !== undefined);
      }
      if (/^[A-Z]+[0-9]+$/i.test(value)) {
        const evaluated = evaluateFormulaExpression(state, sheet, value.toUpperCase(), visited);
        return evaluated !== undefined ? [evaluated] : [];
      }
      const numeric = Number(value);
      return Number.isFinite(numeric) ? [numeric] : [];
    });
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
  dxfs: string[];
}

function buildDefaultStylesheetXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">${DEFAULT_FONT_XML}</fonts>
  <fills count="2">${DEFAULT_FILL_XML}<fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <dxfs count="0"></dxfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`;
}

function parseStylesheet(xml: string): ParsedStylesheet {
  return {
    numFmts: extractStyleSection(xml, "numFmts", "numFmt"),
    fonts: extractStyleSection(xml, "fonts", "font"),
    fills: extractStyleSection(xml, "fills", "fill"),
    borders: extractStyleSection(xml, "borders", "border"),
    dxfs: extractStyleSection(xml, "dxfs", "dxf"),
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
  const textRotation = props["alignment.rotation"] ?? props.rotation;
  const indent = props["alignment.indent"] ?? props.indent;
  const shrinkToFit = props["alignment.shrinkToFit"] ?? props["alignment.shrinktofit"] ?? props.shrinktofit;
  const attrs = [
    horizontal ? `horizontal="${escapeXml(horizontal)}"` : "",
    vertical ? `vertical="${escapeXml(vertical)}"` : "",
    wrapText !== undefined ? `wrapText="${isTruthy(wrapText) ? 1 : 0}"` : "",
    textRotation !== undefined ? `textRotation="${Math.max(0, Math.min(180, Number(textRotation)))}"` : "",
    indent !== undefined ? `indent="${Math.max(0, Number(indent))}"` : "",
    shrinkToFit !== undefined ? `shrinkToFit="${isTruthy(shrinkToFit) ? 1 : 0}"` : "",
  ].filter(Boolean).join(" ");
  return attrs ? `<alignment ${attrs}/>` : "";
}

function buildBorderXml(props: Record<string, string>) {
  const defaultXml = `<border><left/><right/><top/><bottom/><diagonal/></border>`;
  const allStyle = props.border;
  const allColor = props["border.color"];
  const sideStyle = (side: "left" | "right" | "top" | "bottom" | "diagonal") => props[`border.${side}`] ?? allStyle;
  const sideColor = (side: "left" | "right" | "top" | "bottom" | "diagonal") => props[`border.${side}.color`] ?? allColor;
  const diagonalUp = props["border.diagonalUp"];
  const diagonalDown = props["border.diagonalDown"];
  const createSide = (tag: string, style?: string, color?: string) => {
    if (!style || style.toLowerCase() === "none") return `<${tag}/>`;
    return `<${tag} style="${escapeXml(style.toLowerCase())}">${color ? `<color rgb="${escapeXml(normalizeArgbColor(color))}"/>` : ""}</${tag}>`;
  };
  const xml = `<border${diagonalUp !== undefined ? ` diagonalUp="${isTruthy(diagonalUp) ? 1 : 0}"` : ""}${diagonalDown !== undefined ? ` diagonalDown="${isTruthy(diagonalDown) ? 1 : 0}"` : ""}>${createSide("left", sideStyle("left"), sideColor("left"))}${createSide("right", sideStyle("right"), sideColor("right"))}${createSide("top", sideStyle("top"), sideColor("top"))}${createSide("bottom", sideStyle("bottom"), sideColor("bottom"))}${createSide("diagonal", sideStyle("diagonal"), sideColor("diagonal"))}</border>`;
  return xml === `<border><left/><right/><top/><bottom/><diagonal/></border>` ? defaultXml : xml;
}

function buildProtectionXml(props: Record<string, string>) {
  const locked = props.locked;
  const hidden = props.formulahidden;
  if (locked === undefined && hidden === undefined) {
    return "";
  }
  const attrs = [
    locked !== undefined ? `locked="${isTruthy(locked) ? 1 : 0}"` : "",
    hidden !== undefined ? `hidden="${isTruthy(hidden) ? 1 : 0}"` : "",
  ].filter(Boolean).join(" ");
  return `<protection ${attrs}/>`;
}

function isConditionalFormattingStyleKey(key: string) {
  const lower = key.toLowerCase();
  return lower === "fill" || lower === "font.color" || lower === "font.bold";
}

function hasConditionalFormattingStyleProps(props: Record<string, string>) {
  return Object.keys(props).some((key) => isConditionalFormattingStyleKey(key));
}

function extractConditionalFormattingStyleProps(props: Record<string, string>) {
  return {
    ...(props["font.color"] ? { fontColor: normalizeArgbColor(props["font.color"]) } : {}),
    ...(props["font.bold"] !== undefined ? { fontBold: isTruthy(props["font.bold"]) } : {}),
    ...(props.fill ? { fill: normalizeArgbColor(props.fill) } : {}),
  };
}

function serializeConditionalFormattingStyleProps(rule: ExcelConditionalFormattingModel): Record<string, string> {
  return {
    ...(rule.fontColor ? { "font.color": rule.fontColor } : {}),
    ...(rule.fontBold !== undefined ? { "font.bold": rule.fontBold ? "true" : "false" } : {}),
    ...(rule.fill ? { fill: rule.fill } : {}),
  };
}

function registerDifferentialFormat(state: ExcelWorkbookState, props: Record<string, string>) {
  const stylesheet = parseStylesheet(state.styleSheetXml ?? buildDefaultStylesheetXml());
  const dxfXml = buildDifferentialFormatXml(props);
  const dxfId = ensureFragment(stylesheet.dxfs, dxfXml, dxfXml);
  state.styleSheetXml = serializeStylesheet(stylesheet);
  return dxfId;
}

function buildDifferentialFormatXml(props: Record<string, string>) {
  const fontColor = props["font.color"];
  const fontBold = props["font.bold"] !== undefined && isTruthy(props["font.bold"]);
  const fill = props.fill;
  const fontXml = fontColor || fontBold
    ? `<font>${fontBold ? "<b/>" : ""}${fontColor ? `<color rgb="${escapeXml(normalizeArgbColor(fontColor))}"/>` : ""}</font>`
    : "";
  const fillXml = fill
    ? `<fill><patternFill patternType="solid"><bgColor rgb="${escapeXml(normalizeArgbColor(fill))}"/></patternFill></fill>`
    : "";
  return `<dxf>${fontXml}${fillXml}</dxf>`;
}

function normalizeTimePeriod(value: string) {
  const normalized = value.trim().toLowerCase();
  return normalized === "last7days" ? "last7Days"
    : normalized === "thisweek" ? "thisWeek"
    : normalized === "lastweek" ? "lastWeek"
    : normalized === "nextweek" ? "nextWeek"
    : normalized === "thismonth" ? "thisMonth"
    : normalized === "lastmonth" ? "lastMonth"
    : normalized === "nextmonth" ? "nextMonth"
    : normalized;
}

function serializeStylesheet(stylesheet: ParsedStylesheet) {
  const numFmts = stylesheet.numFmts.length > 0 ? `<numFmts count="${stylesheet.numFmts.length}">${stylesheet.numFmts.join("")}</numFmts>` : "";
  const dxfs = `<dxfs count="${stylesheet.dxfs.length}">${stylesheet.dxfs.join("")}</dxfs>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${numFmts}
  <fonts count="${stylesheet.fonts.length}">${stylesheet.fonts.join("")}</fonts>
  <fills count="${stylesheet.fills.length}">${stylesheet.fills.join("")}</fills>
  <borders count="${stylesheet.borders.length}">${stylesheet.borders.join("")}</borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  ${dxfs}
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

function decodeXmlRecursive(value: string) {
  let current = value;
  for (let index = 0; index < 5; index += 1) {
    const next = decodeXml(current);
    if (next === current) {
      return next;
    }
    current = next;
  }
  return current;
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

function resolveChartAddData(state: ExcelWorkbookState, sheet: ExcelSheetModel, props: Record<string, string>) {
  const chartType = (props.charttype ?? props.type ?? "column").toLowerCase();
  const title = props.title ?? props.name;
  const dataRange = props.datarange ?? props.dataRange ?? props.range;
  if (dataRange) {
    const { sheet: sourceSheet, range } = resolveChartRangeSource(state, sheet, dataRange);
    const { categories, series } = parseChartRangeData(sourceSheet, range);
    if (series.length === 0) {
      throw new UsageError("Excel chart dataRange must contain at least one series row.");
    }
    return { chartType, title, categories, series };
  }
  const categories = (props.categories ?? "").split(",").map((item) => item.trim()).filter(Boolean);
  const seriesFromData = parseChartDataProp(props.data);
  const seriesFromNamedProps = Object.entries(props)
    .filter(([key]) => /^series\d+$/i.test(key))
    .map(([, value], index) => parseChartSeriesString(value, `Series ${index + 1}`));
  const series = [...seriesFromData, ...seriesFromNamedProps];
  if (series.length === 0) {
    throw new UsageError("Excel chart requires --prop data=Series:1,2,3 or --prop dataRange=Sheet1!A1:D4.");
  }
  const normalizedCategories = categories.length > 0 ? categories : Array.from({ length: series[0].values.length }, (_, index) => `Category ${index + 1}`);
  return { chartType, title, categories: normalizedCategories, series };
}

function resolveChartRangeSource(state: ExcelWorkbookState, defaultSheet: ExcelSheetModel, source: string) {
  const [sheetPart, rangePart] = source.includes("!") ? source.split("!", 2) : [defaultSheet.name, source];
  return {
    sheet: ensureSheetState(state, sheetPart.replace(/^['"]|['"]$/g, "")),
    range: rangePart.toUpperCase(),
  };
}

function parseChartRangeData(sheet: ExcelSheetModel, range: string) {
  const refs = expandRange(range);
  const start = parseCellAddress(range.split(":")[0]);
  const end = parseCellAddress(range.split(":")[1] ?? range.split(":")[0]);
  const rows = new Map<number, string[]>();
  for (const ref of refs) {
    const address = parseCellAddress(ref);
    const rowValues = rows.get(address.row) ?? [];
    const value = sheet.cells[ref]?.value ?? "";
    rowValues.push(value);
    rows.set(address.row, rowValues);
  }
  const orderedRows = [...rows.entries()].sort(([a], [b]) => a - b).map(([, values]) => values);
  if (orderedRows.length < 2 || columnNameToIndex(end.column) - columnNameToIndex(start.column) < 1) {
    return { categories: [] as string[], series: [] as Array<{ name: string; values: number[] }> };
  }
  const categories = orderedRows[0].slice(1).map((value, index) => value || `Category ${index + 1}`);
  const series = orderedRows.slice(1).map((values, index) => ({
    name: values[0] || `Series ${index + 1}`,
    values: values.slice(1).map((value) => Number(value || "0")),
  }));
  return { categories, series };
}

function parseChartDataProp(raw: string | undefined) {
  if (!raw) return [] as Array<{ name: string; values: number[] }>;
  return raw
    .split(";")
    .map((chunk, index) => parseChartSeriesString(chunk, `Series ${index + 1}`))
    .filter((series) => series.values.length > 0);
}

function parseChartSeriesString(raw: string, fallbackName: string) {
  const [namePart, valuesPart = ""] = raw.split(":", 2);
  const values = valuesPart
    .split(",")
    .map((item) => Number(item.trim()))
    .filter((item) => !Number.isNaN(item));
  return {
    name: namePart.trim() || fallbackName,
    values,
  };
}

function buildChartXml(input: { chartType: string; title?: string; categories: string[]; series: Array<{ name: string; values: number[] }> }) {
  const chartTag = input.chartType === "line" ? "lineChart" : input.chartType === "pie" ? "pieChart" : input.chartType === "area" ? "areaChart" : "barChart";
  const seriesXml = input.series.map((series, index) => {
    const categoryPoints = input.categories.map((category, categoryIndex) => `<c:pt idx="${categoryIndex}"><c:v>${escapeXml(category)}</c:v></c:pt>`).join("");
    const valuePoints = series.values.map((value, valueIndex) => `<c:pt idx="${valueIndex}"><c:v>${value}</c:v></c:pt>`).join("");
    return `<c:ser>
          <c:idx val="${index}"/>
          <c:order val="${index}"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>${escapeXml(series.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache><c:ptCount val="${input.categories.length}"/>${categoryPoints}</c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${series.values.length}"/>${valuePoints}</c:numCache></c:numRef></c:val>
        </c:ser>`;
  }).join("");
  const titleXml = input.title ? `<c:title><c:tx><c:rich><a:p><a:r><a:t>${escapeXml(input.title)}</a:t></a:r></a:p></c:rich></c:tx></c:title>` : "";
  const chartBody = chartTag === "barChart"
    ? `<c:barDir val="col"/><c:grouping val="clustered"/>${seriesXml}<c:axId val="1"/><c:axId val="2"/>`
    : `${seriesXml}${chartTag === "pieChart" ? "" : '<c:axId val="1"/><c:axId val="2"/>'}`;
  const axesXml = chartTag === "pieChart"
    ? ""
    : `<c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    ${titleXml}
    <c:plotArea>
      <c:${chartTag}>
        ${chartBody}
      </c:${chartTag}>
      ${axesXml}
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
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
  if (/^xl\/media\/image\d+\.png$/i.test(entryName)) return [entryName, "image/png"] as [string, string];
  if (/^xl\/media\/image\d+\.jpe?g$/i.test(entryName)) return [entryName, "image/jpeg"] as [string, string];
  if (/^xl\/media\/image\d+\.gif$/i.test(entryName)) return [entryName, "image/gif"] as [string, string];
  if (/^xl\/comments\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"] as [string, string];
  if (/^xl\/tables\/table\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"] as [string, string];
  if (/^xl\/drawings\/drawing\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.drawing+xml"] as [string, string];
  if (/^xl\/charts\/chart\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"] as [string, string];
  if (/^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(entryName)) return [entryName, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"] as [string, string];
  return undefined;
}

function getIconSetThresholdCount(iconSetName: string) {
  const normalized = iconSetName.toLowerCase();
  if (normalized.startsWith("5")) return 5;
  if (normalized.startsWith("4")) return 4;
  return 3;
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

function stripArgb(value: string) {
  return value.length === 8 ? value.slice(2) : value;
}

function normalizeShapeAlign(value: string) {
  const normalized = value.trim().toLowerCase();
  if (normalized === "center" || normalized === "c" || normalized === "ctr") return "ctr";
  if (normalized === "right" || normalized === "r") return "r";
  if (normalized === "justify" || normalized === "j" || normalized === "justified") return "just";
  return "l";
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
