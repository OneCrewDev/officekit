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

  if (options.type !== "cell") {
    throw new UsageError(
      "Excel add currently supports: sheet, row, cell, or namedrange.",
      "Use / with --type sheet|namedrange, /Sheet1 with --type row, or /Sheet1 with --type cell.",
    );
  }

  const sheet = ensureSheetState(state, normalizeSheetPath(targetPath) || options.props.sheet || "Sheet1");
  const ref = (options.props.ref ?? options.props.cell ?? "A1").toUpperCase();
  sheet.cells[ref] = mergeExcelCell(sheet.cells[ref], options.props);
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

  const { sheet, range } = resolveExcelTarget(state, targetPath);
  if (!range) {
    throw new UsageError("Excel set requires a cell, range, or supported object path.");
  }
  if (range.includes(":")) {
    for (const ref of expandRange(range)) {
      sheet.cells[ref] = mergeExcelCell(sheet.cells[ref], options.props);
    }
  } else {
    sheet.cells[range] = mergeExcelCell(sheet.cells[range], options.props);
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
  if (normalized === "validation" || normalized === "validations") {
    for (const sheet of state.sheets) {
      parseValidations(sheet.xml).forEach((validation, index) => {
        nodes.push({ ...validation, path: `/${sheet.name}/validation[${index + 1}]`, type: "validation" });
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
  throw new UsageError(`Unsupported Excel query selector '${selector}'.`, "Supported selectors: sheet, namedrange, cell, validation, comment, table, chart, pivottable.");
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
    return requireEntry(state.zip, getSheetCharts(state, ensureSheetState(state, chartMatch[1]))[Number(chartMatch[2]) - 1]?.path
      ? normalizeZipPath("xl", getSheetCharts(state, ensureSheetState(state, chartMatch[1]))[Number(chartMatch[2]) - 1]!.path)
      : (() => { throw new OfficekitError(`Chart ${chartMatch[2]} not found.`, "not_found"); })());
  }
  const globalChartMatch = /^\/chart\[(\d+)\]$/i.exec(partPath);
  if (globalChartMatch) {
    const charts = state.sheets.flatMap((sheet) => getSheetCharts(state, sheet));
    const chart = charts[Number(globalChartMatch[1]) - 1];
    if (!chart?.path) {
      throw new OfficekitError(`Chart ${globalChartMatch[1]} not found.`, "not_found");
    }
    return requireEntry(state.zip, normalizeZipPath("xl", chart.path));
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
    return {
      name,
      relId,
      relationshipTarget: target,
      entryName,
      xml,
      cells: parseSheetCells(xml, zip),
      ...features,
    } satisfies ExcelSheetModel;
  });
  const scopedNamedRanges = parseDefinedNames(workbookXml, sheets);
  return {
    zip,
    workbookXml,
    workbookEntryName,
    workbookRelsXml,
    workbookRelsEntryName,
    sheets,
    settings,
    namedRanges: scopedNamedRanges,
    styleSheetXml: zip.get("xl/styles.xml")?.toString("utf8"),
    metadata: parsePackageProperties(zip),
  };
}

async function writeWorkbookState(filePath: string, state: ExcelWorkbookState) {
  await mkdir(path.dirname(filePath), { recursive: true });
  state.zip.set(state.workbookEntryName, Buffer.from(state.workbookXml, "utf8"));
  state.zip.set(state.workbookRelsEntryName, Buffer.from(state.workbookRelsXml, "utf8"));
  state.zip.set("[Content_Types].xml", Buffer.from(buildContentTypesXml(state), "utf8"));
  state.zip.set("_rels/.rels", Buffer.from(buildRootRelsXml(), "utf8"));
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
      })),
      ...(Object.keys(state.settings).length > 0 ? { settings: state.settings } : {}),
      ...(state.styleSheetXml ? { styleSheetXml: state.styleSheetXml } : {}),
      ...(state.namedRanges.length > 0 ? { namedRanges: state.namedRanges } : {}),
    },
  }, null, 2), "utf8"));
  if (!state.zip.has("docProps/core.xml")) {
    state.zip.set("docProps/core.xml", Buffer.from(buildCorePropertiesXml(state.metadata), "utf8"));
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
    return { ...validation, path: targetPath, type: "validation" };
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
    return { ...sparkline, path: targetPath, type: "sparkline" };
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
  return cell ? { path: `/${sheet.name}/${ref}`, ref, type: "cell", ...cell } : { path: `/${sheet.name}/${ref}`, ref, type: "cell", value: null };
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
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheetPr}
  ${sheetViews}
  <sheetData>${xmlRows}</sheetData>
  ${autoFilter}
</worksheet>`;
}

function mergeSheetXmlPreservingExtras(previousXml: string, nextXml: string) {
  let xml = previousXml;
  const nextSheetPr = /<(?:\w+:)?sheetPr\b[\s\S]*?<\/(?:\w+:)?sheetPr>/.exec(nextXml)?.[0] ?? "";
  const nextSheetViews = /<(?:\w+:)?sheetViews\b[\s\S]*?<\/(?:\w+:)?sheetViews>/.exec(nextXml)?.[0] ?? "";
  const nextSheetData = /<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/.exec(nextXml)?.[0] ?? "<sheetData/>";
  const nextAutoFilter = /<(?:\w+:)?autoFilter\b[^>]*\/?>/.exec(nextXml)?.[0] ?? "";

  xml = replaceOrInsert(xml, /<(?:\w+:)?sheetPr\b[\s\S]*?<\/(?:\w+:)?sheetPr>/, nextSheetPr, /<(?:\w+:)?worksheet\b[^>]*>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?sheetViews\b[\s\S]*?<\/(?:\w+:)?sheetViews>/, nextSheetViews, /<(?:\w+:)?sheetPr\b[\s\S]*?<\/(?:\w+:)?sheetPr>|<(?:\w+:)?worksheet\b[^>]*>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/, nextSheetData, /<(?:\w+:)?sheetViews\b[\s\S]*?<\/(?:\w+:)?sheetViews>|<(?:\w+:)?sheetPr\b[\s\S]*?<\/(?:\w+:)?sheetPr>|<(?:\w+:)?worksheet\b[^>]*>/);
  xml = replaceOrInsert(xml, /<(?:\w+:)?autoFilter\b[^>]*\/?>/, nextAutoFilter, /<(?:\w+:)?sheetData\b[\s\S]*?<\/(?:\w+:)?sheetData>/);
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
      const ordered = [...cells.entries()].sort(([a], [b]) => columnNameToIndex(a) - columnNameToIndex(b)).map(([, cell]) => formatCellDisplayValue(cell));
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
      const value = formatCellDisplayValue(cell);
      const annotation = cell.formula ? `=${cell.formula}` : cell.type ?? "number";
      const warn = !cell.value && !cell.formula ? " empty" : "";
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
  return issues.length > 0 ? issues.join("\n") : "No structural Excel issues detected.";
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
    sheet.tabColor = props.tabColor ?? props.tabcolor;
  }
  if (props.autofilter !== undefined) {
    sheet.autoFilter = props.autofilter;
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
  return {
    ...(autoFilter ? { autoFilter } : {}),
    ...(topLeftCell ? { freezeTopLeftCell: topLeftCell } : {}),
    ...(zoom ? { zoom: Number(zoom) } : {}),
    ...(showGridLines !== undefined ? { showGridLines: isTruthy(showGridLines) } : {}),
    ...(showHeadings !== undefined ? { showHeadings: isTruthy(showHeadings) } : {}),
    ...(tabColor ? { tabColor: stripAlpha(tabColor) } : {}),
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
    return {
      path: rel?.target ?? "",
      sheet: sheet.name,
      title: extractTexts(chartXml).join("").trim() || undefined,
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
    };
  });
}

function parseSparklines(sheetXml: string) {
  return [...sheetXml.matchAll(/<x14:sparkline\b[\s\S]*?<xm:f>([\s\S]*?)<\/xm:f>[\s\S]*?<xm:sqref>([\s\S]*?)<\/xm:sqref>[\s\S]*?<\/x14:sparkline>/g)].map((match) => ({
    sourceRange: decodeXml(match[1]).trim(),
    location: decodeXml(match[2]).trim(),
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

function normalizeExcelCell(cell: ExcelCellModel | undefined): ExcelCellModel {
  return {
    value: cell?.value ?? "",
    ...(cell?.styleId ? { styleId: cell.styleId } : {}),
    ...(cell?.type ? { type: cell.type } : {}),
    ...(cell?.formula ? { formula: normalizeFormula(cell.formula) } : {}),
  };
}

function formatCellDisplayValue(cell: ExcelCellModel) {
  if (cell.type === "boolean") {
    return cell.value === "1" ? "TRUE" : "FALSE";
  }
  return cell.value;
}

function normalizeFormula(formula: string) {
  return formula.replace(/^=/, "");
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
  const sheetOverrides = state.sheets
    .map((sheet) => `<Override PartName="/${sheet.entryName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
    .join("\n  ");
  const stylesOverride = state.styleSheetXml
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

function buildRootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
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

function stripAlpha(value: string) {
  return value.length === 8 ? value.slice(2) : value;
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
