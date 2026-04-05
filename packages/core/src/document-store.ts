import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { OfficekitError, UsageError } from "./errors.js";
import { assertFormat, type SupportedFormat } from "./formats.js";
import { createStoredZip, readStoredZip } from "./zip.js";

export interface WordParagraph {
  text: string;
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

export interface ExcelSheet {
  name: string;
  cells: Record<string, string>;
}

export interface PptShape {
  text: string;
}

export interface PptSlide {
  title: string;
  shapes: PptShape[];
}

export interface OfficekitDocument {
  product: "officekit";
  lineage: string;
  format: SupportedFormat;
  version: 1;
  updatedAt: string;
  word?: { paragraphs: WordParagraph[]; tables: WordTable[] };
  excel?: { sheets: ExcelSheet[] };
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
        document.word!.paragraphs.push({ text: options.props.text ?? "" });
        break;
      }
      if (options.type === "table") {
        const rows = Math.max(1, Number(options.props.rows ?? "2"));
        const cols = Math.max(1, Number(options.props.cols ?? "2"));
        document.word!.tables.push({
          rows: Array.from({ length: rows }, () => ({
            cells: Array.from({ length: cols }, () => ({ text: "" })),
          })),
        });
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
      sheet.cells[ref] = options.props.value ?? options.props.text ?? "";
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
      const paragraph = document.word!.paragraphs[Number(match[1]) - 1];
      if (!paragraph) throw new OfficekitError(`Paragraph ${match[1]} does not exist.`, "not_found");
      paragraph.text = options.props.text ?? paragraph.text;
    } else if (tableMatch) {
      const table = document.word!.tables[Number(tableMatch[1]) - 1];
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
    const { sheet, cellRef } = resolveExcelPath(document, targetPath);
    sheet.cells[cellRef] = options.props.value ?? options.props.text ?? sheet.cells[cellRef] ?? "";
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
      document.word!.paragraphs.splice(Number(match[1]) - 1, 1);
    } else if (tableMatch) {
      document.word!.tables.splice(Number(tableMatch[1]) - 1, 1);
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
    const paragraphs = document.word!.paragraphs.map((paragraph) => `<p>${escapeHtml(paragraph.text)}</p>`);
    const tables = document.word!.tables.map((table) => renderWordTableHtml(table));
    const body = [...paragraphs, ...tables].join("\n") || "<p><em>Empty document</em></p>";
    return `<article data-format="word">${body}</article>`;
  }

  if (document.format === "excel") {
    const rows = document.excel!.sheets.flatMap((sheet) => Object.entries(sheet.cells).map(([ref, value]) => `<tr><th>${escapeHtml(sheet.name)}!${escapeHtml(ref)}</th><td>${escapeHtml(value)}</td></tr>`));
    return `<section data-format="excel"><table><tbody>${rows.join("") || '<tr><td colspan="2"><em>Empty workbook</em></td></tr>'}</tbody></table></section>`;
  }

  const slides = document.powerpoint!.slides.map((slide, index) => `<section class="slide"><h2>Slide ${index + 1}: ${escapeHtml(slide.title)}</h2>${slide.shapes.map((shape) => `<p>${escapeHtml(shape.text)}</p>`).join("")}</section>`);
  return `<main data-format="powerpoint">${slides.join("") || '<section class="slide"><em>Empty deck</em></section>'}</main>`;
}

export function renderDocumentOutline(document: OfficekitDocument): string {
  if (document.format === "word") {
    const lines: string[] = [];
    for (const [index, paragraph] of document.word!.paragraphs.entries()) {
      lines.push(`Paragraph ${index + 1}: ${paragraph.text}`);
    }
    for (const [tableIndex, table] of document.word!.tables.entries()) {
      const rowCount = table.rows.length;
      const colCount = table.rows[0]?.cells.length ?? 0;
      lines.push(`Table ${tableIndex + 1}: ${rowCount}x${colCount}`);
      for (const [rowIndex, row] of table.rows.entries()) {
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
        lines.push(`  ${ref}: ${sheet.cells[ref]}`);
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
  if (format === "word") return { ...base, word: { paragraphs: [], tables: [] } };
  if (format === "excel") return { ...base, excel: { sheets: [{ name: "Sheet1", cells: {} as Record<string, string> }] } };
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
    if (targetPath === "/body") return document.word;
    const match = /^\/body\/p\[(\d+)\]$/.exec(targetPath);
    const tableMatch = /^\/body\/table\[(\d+)\]$/.exec(targetPath);
    const tableCellMatch = /^\/body\/table\[(\d+)\]\/cell\[(\d+),(\d+)\]$/.exec(targetPath);
    if (match) {
      const paragraph = document.word!.paragraphs[Number(match[1]) - 1];
      if (!paragraph) throw new OfficekitError(`Paragraph ${match[1]} does not exist.`, "not_found");
      return paragraph;
    }
    if (tableMatch) {
      const table = document.word!.tables[Number(tableMatch[1]) - 1];
      if (!table) throw new OfficekitError(`Table ${tableMatch[1]} does not exist.`, "not_found");
      return table;
    }
    if (tableCellMatch) {
      const table = document.word!.tables[Number(tableCellMatch[1]) - 1];
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
    return { ref: cellRef, value: sheet.cells[cellRef] ?? null };
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
  const paragraphs = document.word!.paragraphs
    .map((paragraph) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(paragraph.text)}</w:t></w:r></w:p>`)
    .join("");
  const tables = document.word!.tables.map((table) => renderWordTableXml(table)).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${paragraphs}
    ${tables}
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function renderExcelContentTypes(document: OfficekitDocument) {
  const sheetOverrides = document.excel!.sheets
    .map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
    .join("\n  ");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetOverrides}
</Types>`;
}

function renderExcelRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

function renderWorkbookXml(document: OfficekitDocument) {
  const sheets = document.excel!.sheets
    .map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets}</sheets>
</workbook>`;
}

function renderWorkbookRels(document: OfficekitDocument) {
  const rels = document.excel!.sheets
    .map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

function renderSheetXml(sheet: ExcelSheet) {
  const entries = Object.entries(sheet.cells).sort(([a], [b]) => a.localeCompare(b));
  const rows = new Map<number, string[]>();
  for (const [ref, value] of entries) {
    const row = Number(/\d+/.exec(ref)?.[0] ?? "1");
    const cells = rows.get(row) ?? [];
    cells.push(`<c r="${ref}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`);
    rows.set(row, cells);
  }
  const xmlRows = [...rows.entries()].sort(([a], [b]) => a - b).map(([rowIndex, cells]) => `<row r="${rowIndex}">${cells.join("")}</row>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${xmlRows}</sheetData>
</worksheet>`;
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
    document.word = {
      paragraphs: document.word.paragraphs ?? [],
      tables: document.word.tables ?? [],
    };
  }
  return document;
}

function parseWordDocument(zip: Map<string, Buffer>): OfficekitDocument {
  const xml = requireEntry(zip, "word/document.xml");
  const body = /<w:body\b[^>]*>([\s\S]*?)<w:sectPr\b[^>]*\/?>/.exec(xml)?.[1] ?? "";
  const paragraphs: WordParagraph[] = [];
  const tables: WordTable[] = [];
  for (const match of body.matchAll(/<w:(p|tbl)\b[\s\S]*?<\/w:\1>/g)) {
    if (match[1] === "p") {
      const text = extractTextRuns(match[0]);
      if (text.length > 0) {
        paragraphs.push({ text });
      }
    } else {
      tables.push(parseWordTable(match[0]));
    }
  }
  return {
    product: "officekit",
    lineage: LINEAGE,
    format: "word",
    version: 1,
    updatedAt: new Date().toISOString(),
    word: {
      paragraphs,
      tables,
    },
  };
}

function parseWordTable(xml: string): WordTable {
  const rows = [...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((rowMatch) => ({
    cells: [...rowMatch[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((cellMatch) => ({
      text: extractTextRuns(cellMatch[0]),
    })),
  }));
  return { rows };
}

function parseExcelDocument(zip: Map<string, Buffer>): OfficekitDocument {
  const workbookXml = requireEntry(zip, "xl/workbook.xml");
  const workbookRelsXml = requireEntry(zip, "xl/_rels/workbook.xml.rels");
  const relationshipMap = parseRelationships(workbookRelsXml);
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
    excel: { sheets },
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
    const slideXml = requireEntry(zip, normalizeZipPath("ppt", target));
    const texts = [...slideXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((textMatch) => decodeXml(textMatch[1]));
    const [title = "Untitled slide", ...shapeTexts] = texts;
    return {
      title,
      shapes: shapeTexts.filter(Boolean).map((text) => ({ text })),
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
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const attributes = match[1];
    const id = /Id="([^"]+)"/.exec(attributes)?.[1];
    const target = /Target="([^"]+)"/.exec(attributes)?.[1];
    if (id && target) {
      relationships.set(id, target);
    }
  }
  return relationships;
}

function parseSheetCells(xml: string, zip: Map<string, Buffer>) {
  const sharedStrings = parseSharedStrings(zip);
  const cells: Record<string, string> = {};
  for (const match of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const attributes = match[1];
    const body = match[2];
    const refMatch = /r="([^"]+)"/.exec(attributes);
    if (!refMatch) continue;
    const ref = refMatch[1].toUpperCase();
    const typeMatch = /t="([^"]+)"/.exec(attributes);
    const type = typeMatch?.[1] ?? "";
    let value = "";
    if (type === "inlineStr") {
      value = extractTexts(body).join("");
    } else if (type === "s") {
      const index = Number((/<v>([\s\S]*?)<\/v>/.exec(body)?.[1] ?? "0").trim());
      value = sharedStrings[index] ?? "";
    } else {
      value = decodeXml((/<v>([\s\S]*?)<\/v>/.exec(body)?.[1] ?? "").trim());
    }
    cells[ref] = value;
  }
  return cells;
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
