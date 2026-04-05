import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { OfficekitError, UsageError } from "./errors.js";
import { assertFormat, type SupportedFormat } from "./formats.js";
import { createStoredZip, readStoredZip } from "./zip.js";

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

export interface ExcelCell {
  value: string;
  formula?: string;
  styleId?: string;
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

export interface ExcelSheet {
  name: string;
  cells: Record<string, ExcelCell>;
}

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
    sheets: ExcelSheet[];
    settings?: ExcelWorkbookSettings;
    styleSheetXml?: string;
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

export async function createDocument(filePath: string) {
  const format = assertFormat(filePath);
  const document = createBlankDocument(format);
  await persistDocument(filePath, document);
  return { format, filePath, document };
}

export async function addDocumentNode(filePath: string, targetPath: string, options: CommandOptions) {
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
    case "excel": {
      if (options.type !== "cell") {
        throw new UsageError("Excel add currently supports only: add <file.xlsx> /Sheet1 --type cell --prop ref=A1 --prop value=...", "Use --type cell with a sheet path.");
      }
      const sheetName = targetPath.replace(/^\//, "") || options.props.sheet || "Sheet1";
      const sheet = ensureSheet(document, sheetName);
      const ref = (options.props.ref ?? options.props.cell ?? "A1").toUpperCase();
      sheet.cells[ref] = mergeExcelCell(sheet.cells[ref], options.props);
      break;
    }
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

export async function setDocumentNode(filePath: string, targetPath: string, options: CommandOptions) {
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
    if (targetPath === "/" || targetPath === "/workbook") {
      document.excel ??= { sheets: [], settings: {} };
      document.excel.settings = mergeWorkbookSettings(document.excel.settings, options.props);
    } else {
    const { sheet, cellRef } = resolveExcelPath(document, targetPath);
      sheet.cells[cellRef] = mergeExcelCell(sheet.cells[cellRef], options.props);
    }
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
    const { sheet, cellRef } = resolveExcelPath(document, targetPath);
    delete sheet.cells[cellRef];
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
  const document = await loadDocument(filePath);
  return materializePath(document, targetPath);
}

export async function viewDocument(filePath: string, mode: string) {
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
  const document = await loadDocument(filePath);
  return {
    ok: true,
    format: document.format,
    summary: renderDocumentOutline(document),
  };
}

export async function rawDocument(filePath: string) {
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
    const rows = document.excel!.sheets.flatMap((sheet) =>
      Object.entries(sheet.cells).map(([ref, cell]) => {
        const detail = [cell.value, cell.formula ? `formula=${cell.formula}` : ""].filter(Boolean).join(" · ");
        return `<tr><th>${escapeHtml(sheet.name)}!${escapeHtml(ref)}</th><td>${escapeHtml(detail)}</td></tr>`;
      }),
    );
    return `<section data-format="excel"><table><tbody>${rows.join("") || '<tr><td colspan="2"><em>Empty workbook</em></td></tr>'}</tbody></table></section>`;
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
    if (targetPath === "/workbook") return document.excel;
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
  const sheets = document.excel!.sheets
    .map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${workbookPr}
  ${workbookProtection}
  <sheets>${sheets}</sheets>
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
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${xmlRows}</sheetData>
</worksheet>`;
}

function renderExcelCellXml(ref: string, cell: ExcelCell) {
  const styleAttr = cell.styleId ? ` s="${escapeXml(cell.styleId)}"` : "";
  if (cell.formula) {
    const valueXml = cell.value !== "" ? `<v>${escapeXml(cell.value)}</v>` : "";
    return `<c r="${ref}"${styleAttr}><f>${escapeXml(normalizeFormula(cell.formula))}</f>${valueXml}</c>`;
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
      })),
      ...(document.excel.settings ? { settings: document.excel.settings } : {}),
      ...(document.excel.styleSheetXml ? { styleSheetXml: document.excel.styleSheetXml } : {}),
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
  const sheets = [...workbookXml.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g)].map((match) => {
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
  for (const match of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const attributes = match[1];
    const body = match[2];
    const refMatch = /r="([^"]+)"/.exec(attributes);
    if (!refMatch) continue;
    const ref = refMatch[1].toUpperCase();
    const styleId = /s="([^"]+)"/.exec(attributes)?.[1];
    const typeMatch = /t="([^"]+)"/.exec(attributes);
    const type = typeMatch?.[1] ?? "";
    const formula = (/<f\b[^>]*>([\s\S]*?)<\/f>/.exec(body)?.[1] ?? "").trim();
    let value = "";
    if (type === "inlineStr") {
      value = extractTexts(body).join("");
    } else if (type === "s") {
      const index = Number((/<v>([\s\S]*?)<\/v>/.exec(body)?.[1] ?? "0").trim());
      value = sharedStrings[index] ?? "";
    } else {
      value = decodeXml((/<v>([\s\S]*?)<\/v>/.exec(body)?.[1] ?? "").trim());
    }
    cells[ref] = {
      value,
      ...(styleId ? { styleId } : {}),
      ...(formula ? { formula: decodeXml(formula) } : {}),
    };
  }
  return cells;
}

function parseWorkbookSettings(xml: string): ExcelWorkbookSettings {
  const attrs = /<workbookPr\b([^>]*)\/?>/.exec(xml)?.[1];
  const calcAttrs = /<calcPr\b([^>]*)\/?>/.exec(xml)?.[1];
  const protectionAttrs = /<workbookProtection\b([^>]*)\/?>/.exec(xml)?.[1];
  return {
    ...parseWorkbookPropertyAttributes(attrs),
    ...parseCalculationPropertyAttributes(calcAttrs),
    ...parseWorkbookProtectionAttributes(protectionAttrs),
  };
}

function parseSharedStrings(zip: Map<string, Buffer>) {
  const shared = zip.get("xl/sharedStrings.xml");
  if (!shared) return [];
  return [...shared.toString("utf8").matchAll(/<si\b[\s\S]*?<\/si>/g)].map((match) => extractTexts(match[0]).join(""));
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
    return { value: cell };
  }
  return {
    value: cell?.value ?? "",
    ...(cell?.styleId ? { styleId: cell.styleId } : {}),
    ...(cell?.formula ? { formula: normalizeFormula(cell.formula) } : {}),
  };
}

function mergeExcelCell(existing: string | ExcelCell | undefined, props: Record<string, string>): ExcelCell {
  const base = normalizeExcelCell(existing);
  const formula = props.formula === undefined ? base.formula : normalizeFormula(props.formula);
  const styleId = props.styleId ?? props.style ?? base.styleId;
  return {
    value: props.value ?? props.text ?? base.value,
    ...(styleId ? { styleId } : {}),
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
