import { describe, expect, test } from "bun:test";
import { spawn } from "node:child_process";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { deflateRawSync } from "node:zlib";
import { createStoredZip, readStoredZip } from "@officekit/core";
import { runCli } from "./index.js";

describe("officekit CLI scaffold", () => {
  test("returns a JSON execution plan for Word create", async () => {
    const result = await runCli(["create", "demo.docx", "--plan", "--json"]);
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('"targetPackage": "packages/word"');
  });

  test("returns lineage summary for about", async () => {
    const result = await runCli(["about"]);
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain("migration of OfficeCLI");
  });

  test("keeps unsupported MCP explicit", async () => {
    const result = await runCli(["mcp", "--json"]);
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain("capability_excluded");
  });

  test("creates and mutates a Word document vertical slice", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-word-"));
    const filePath = path.join(dir, "demo.docx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/body", "--type", "paragraph", "--prop", "text=Hello vertical slice"]);
    const result = await runCli(["get", filePath, "/body/p[1]", "--json"]);
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain("Hello vertical slice");
    const rawBytes = await readFile(filePath);
    expect(rawBytes.length).toBeGreaterThan(50);
  });

  test("creates and mutates a Word table vertical slice", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-word-table-"));
    const filePath = path.join(dir, "table.docx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/body", "--type", "table", "--prop", "rows=2", "--prop", "cols=2"]);
    await runCli(["set", filePath, "/body/table[1]/cell[1,1]", "--prop", "text=Cell 11"]);
    await runCli(["set", filePath, "/body/table[1]/cell[2,2]", "--prop", "text=Cell 22"]);

    const table = await runCli(["get", filePath, "/body/table[1]", "--json"]);
    const outline = await runCli(["view", filePath, "outline"]);
    const html = await runCli(["view", filePath, "html"]);

    expect(table.stdout).toContain("Cell 11");
    expect(outline.stdout).toContain("Table 1: 2x2");
    expect(outline.stdout).toContain("R2C2: Cell 22");
    expect(html.stdout).toContain("<table>");
    expect(html.stdout).toContain("Cell 11");
  });

  test("preserves paragraph and table order in created Word documents", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-word-ordered-"));
    const filePath = path.join(dir, "ordered.docx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/body", "--type", "paragraph", "--prop", "text=Intro"]);
    await runCli(["add", filePath, "/body", "--type", "table", "--prop", "rows=1", "--prop", "cols=2"]);
    await runCli(["set", filePath, "/body/table[1]/cell[1,1]", "--prop", "text=Cell A"]);
    await runCli(["set", filePath, "/body/table[1]/cell[1,2]", "--prop", "text=Cell B"]);
    await runCli(["add", filePath, "/body", "--type", "paragraph", "--prop", "text=Outro"]);

    const outline = await runCli(["view", filePath, "outline"]);
    const html = await runCli(["view", filePath, "html"]);
    const xml = readStoredZip(await readFile(filePath)).get("word/document.xml")!.toString("utf8");
    const outlineText = outline.stdout ?? "";
    const htmlText = html.stdout ?? "";

    expect(outline.exitCode).toBe(0);
    expect(outlineText.indexOf("Paragraph 1: Intro")).toBeLessThan(outlineText.indexOf("Table 1: 1x2"));
    expect(outlineText.indexOf("Table 1: 1x2")).toBeLessThan(outlineText.indexOf("Paragraph 2: Outro"));
    expect(htmlText.indexOf("<p>Intro</p>")).toBeLessThan(htmlText.indexOf("<table>"));
    expect(htmlText.indexOf("<table>")).toBeLessThan(htmlText.indexOf("<p>Outro</p>"));
    expect(xml.indexOf("Intro")).toBeLessThan(xml.indexOf("<w:tbl>"));
    expect(xml.indexOf("<w:tbl>")).toBeLessThan(xml.indexOf("Outro"));
  });

  test("creates and mutates an Excel document vertical slice", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-"));
    const filePath = path.join(dir, "demo.xlsx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/Sheet1", "--type", "cell", "--prop", "ref=A1", "--prop", "value=42"]);
    const result = await runCli(["get", filePath, "/Sheet1/A1", "--json"]);
    expect(result.stdout).toContain("\"42\"");
  });

  test("adds Excel sheets and rows", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-sheet-row-"));
    const filePath = path.join(dir, "sheets.xlsx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/", "--type", "sheet", "--prop", "name=Analysis"]);
    await runCli(["add", filePath, "/Analysis", "--type", "row", "--prop", "index=3", "--prop", "cols=2"]);

    const workbook = await runCli(["get", filePath, "/workbook", "--json"]);
    const rowCell = await runCli(["get", filePath, "/Analysis/B3", "--json"]);
    const outline = await runCli(["view", filePath, "outline"]);

    expect(workbook.stdout).toContain('"name": "Analysis"');
    expect(rowCell.stdout).toContain('"ref": "B3"');
    expect(outline.stdout).toContain("Sheet Analysis");
  });

  test("preserves authored Excel formulas in created workbooks", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-formula-"));
    const filePath = path.join(dir, "formula.xlsx");
    await runCli(["create", filePath]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=21"]);
    await runCli(["set", filePath, "/Sheet1/B1", "--prop", "formula==SUM(A1:A1)", "--prop", "value=21"]);

    const result = await runCli(["get", filePath, "/Sheet1/B1", "--json"]);
    const outline = await runCli(["view", filePath, "outline"]);
    const xml = readStoredZip(await readFile(filePath)).get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(result.stdout).toContain('"formula": "SUM(A1:A1)"');
    expect(result.stdout).toContain('"value": "21"');
    expect(outline.stdout).toContain("B1: 21 (formula=SUM(A1:A1))");
    expect(xml).toContain("<f>SUM(A1:A1)</f>");
    expect(xml).toContain("<v>21</v>");
  });

  test("preserves authored Excel style ids on created workbooks", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-style-"));
    const filePath = path.join(dir, "style.xlsx");
    await runCli(["create", filePath]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=Styled", "--prop", "styleId=3"]);

    const result = await runCli(["get", filePath, "/Sheet1/A1", "--json"]);
    const xml = readStoredZip(await readFile(filePath)).get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(result.stdout).toContain('"styleId": "3"');
    expect(xml).toContain(' s="3"');
  });

  test("creates and mutates a PowerPoint document vertical slice", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-ppt-"));
    const filePath = path.join(dir, "demo.pptx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/", "--type", "slide", "--prop", "title=Roadmap"]);
    await runCli(["add", filePath, "/slide[1]", "--type", "shape", "--prop", "text=Q4 launch"]);
    const result = await runCli(["view", filePath, "outline"]);
    expect(result.stdout).toContain("Slide 1: Roadmap");
    expect(result.stdout).toContain("Shape 1: Q4 launch");
  });

  test("reads a metadata-free standard Word OOXML file", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-word-fallback-"));
    const filePath = path.join(dir, "fallback.docx");
    await writeFile(filePath, buildExternalWordZip("Imported paragraph"));

    const getResult = await runCli(["get", filePath, "/body/p[1]", "--json"]);
    const viewResult = await runCli(["view", filePath, "outline"]);
    const rawResult = await runCli(["raw", filePath]);

    expect(getResult.stdout).toContain("Imported paragraph");
    expect(viewResult.stdout).toContain("Paragraph 1: Imported paragraph");
    expect(rawResult.stdout).toContain('"format": "word"');
  });

  test("reads and mutates a metadata-free standard Word table OOXML file", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-word-table-fallback-"));
    const filePath = path.join(dir, "fallback-table.docx");
    await writeFile(filePath, buildExternalWordTableZip("A", "B"));

    const before = await runCli(["view", filePath, "outline"]);
    expect(before.stdout).toContain("Table 1: 1x2");
    expect(before.stdout).toContain("R1C1: A");

    await runCli(["set", filePath, "/body/table[1]/cell[1,2]", "--prop", "text=Updated"]);
    const after = await runCli(["get", filePath, "/body/table[1]/cell[1,2]", "--json"]);
    expect(after.stdout).toContain("Updated");
  });

  test("keeps mixed paragraph and table order for metadata-free Word OOXML files", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-word-mixed-fallback-"));
    const filePath = path.join(dir, "mixed.docx");
    await writeFile(filePath, buildExternalMixedWordZip("Imported intro", "Imported left", "Imported right", "Imported outro"));

    const before = await runCli(["view", filePath, "outline"]);
    const paragraph = await runCli(["get", filePath, "/body/p[2]", "--json"]);
    await runCli(["set", filePath, "/body/table[1]/cell[1,2]", "--prop", "text=Updated right"]);
    const after = await runCli(["view", filePath, "outline"]);
    const xml = readStoredZip(await readFile(filePath)).get("word/document.xml")!.toString("utf8");
    const beforeText = before.stdout ?? "";

    expect(before.exitCode).toBe(0);
    expect(beforeText.indexOf("Paragraph 1: Imported intro")).toBeLessThan(beforeText.indexOf("Table 1: 1x2"));
    expect(beforeText.indexOf("Table 1: 1x2")).toBeLessThan(beforeText.indexOf("Paragraph 2: Imported outro"));
    expect(paragraph.stdout).toContain("Imported outro");
    expect(after.stdout).toContain("R1C2: Updated right");
    expect(xml.indexOf("Imported intro")).toBeLessThan(xml.indexOf("<w:tbl>"));
    expect(xml.indexOf("<w:tbl>")).toBeLessThan(xml.indexOf("Imported outro"));
  });

  test("reads and mutates a metadata-free standard Excel OOXML file", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-fallback-"));
    const filePath = path.join(dir, "fallback.xlsx");
    await writeFile(filePath, buildExternalExcelZip("A1", "99"));

    const before = await runCli(["get", filePath, "/Sheet1/A1", "--json"]);
    expect(before.stdout).toContain('"99"');

    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=100"]);
    const after = await runCli(["get", filePath, "/Sheet1/A1", "--json"]);
    expect(after.stdout).toContain('"100"');

    const zipEntries = readStoredZip(await readFile(filePath));
    expect(zipEntries.has("officekit/document.json")).toBe(true);
  });

  test("reads metadata-free Excel formulas from OOXML workbooks", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-formula-fallback-"));
    const filePath = path.join(dir, "formula-fallback.xlsx");
    await writeFile(filePath, buildExternalExcelFormulaZip());

    const result = await runCli(["get", filePath, "/Sheet1/B1", "--json"]);
    const outline = await runCli(["view", filePath, "outline"]);

    expect(result.stdout).toContain('"formula": "SUM(A1:A1)"');
    expect(result.stdout).toContain('"value": "21"');
    expect(outline.stdout).toContain("B1: 21 (formula=SUM(A1:A1))");
  });

  test("preserves workbook settings and style ids from metadata-free Excel OOXML files", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-settings-fallback-"));
    const filePath = path.join(dir, "settings-fallback.xlsx");
    await writeFile(filePath, buildExternalExcelSettingsZip());

    const workbook = await runCli(["get", filePath, "/workbook", "--json"]);
    const cell = await runCli(["get", filePath, "/Sheet1/A1", "--json"]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=Updated"]);
    const zip = readStoredZip(await readFile(filePath));
    const workbookXml = zip.get("xl/workbook.xml")!.toString("utf8");
    const sheetXml = zip.get("xl/worksheets/sheet1.xml")!.toString("utf8");
    const stylesXml = zip.get("xl/styles.xml")!.toString("utf8");

    expect(workbook.stdout).toContain('"date1904": true');
    expect(workbook.stdout).toContain('"codeName": "WorkbookCode"');
    expect(cell.stdout).toContain('"styleId": "1"');
    expect(workbookXml).toContain('date1904="1"');
    expect(workbookXml).toContain('codeName="WorkbookCode"');
    expect(sheetXml).toContain(' s="1"');
    expect(stylesXml).toContain("cellXfs");
  });

  test("sets workbook settings on officekit-created workbooks", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-settings-create-"));
    const filePath = path.join(dir, "workbook-settings.xlsx");
    await runCli(["create", filePath]);
    await runCli([
      "set",
      filePath,
      "/workbook",
      "--prop",
      "date1904=true",
      "--prop",
      "codeName=OfficekitBook",
      "--prop",
      "filterPrivacy=true",
      "--prop",
      "showObjects=all",
      "--prop",
      "calc.mode=autoNoTable",
    ]);

    const workbook = await runCli(["get", filePath, "/workbook", "--json"]);
    const workbookXml = readStoredZip(await readFile(filePath)).get("xl/workbook.xml")!.toString("utf8");

    expect(workbook.stdout).toContain('"date1904": true');
    expect(workbook.stdout).toContain('"codeName": "OfficekitBook"');
    expect(workbook.stdout).toContain('"filterPrivacy": true');
    expect(workbook.stdout).toContain('"showObjects": "all"');
    expect(workbook.stdout).toContain('"calcMode": "autoNoTable"');
    expect(workbookXml).toContain('date1904="1"');
    expect(workbookXml).toContain('codeName="OfficekitBook"');
    expect(workbookXml).toContain('filterPrivacy="1"');
    expect(workbookXml).toContain('showObjects="all"');
    expect(workbookXml).toContain('calcMode="autoNoTable"');
  });

  test("preserves extended workbook settings from metadata-free Excel OOXML files", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-settings-extended-fallback-"));
    const filePath = path.join(dir, "settings-extended-fallback.xlsx");
    await writeFile(filePath, buildExternalExcelExtendedSettingsZip());

    const workbook = await runCli(["get", filePath, "/workbook", "--json"]);
    await runCli([
      "set",
      filePath,
      "/workbook",
      "--prop",
      "calc.iterate=true",
      "--prop",
      "calc.iterateCount=77",
      "--prop",
      "workbook.lockStructure=true",
    ]);
    const zip = readStoredZip(await readFile(filePath));
    const workbookXml = zip.get("xl/workbook.xml")!.toString("utf8");

    expect(workbook.stdout).toContain('"backupFile": true');
    expect(workbook.stdout).toContain('"dateCompatibility": true');
    expect(workbook.stdout).toContain('"calcMode": "autoNoTable"');
    expect(workbook.stdout).toContain('"iterateCount": 5');
    expect(workbook.stdout).toContain('"refMode": "R1C1"');
    expect(workbook.stdout).toContain('"lockStructure": true');
    expect(workbookXml).toContain('backupFile="1"');
    expect(workbookXml).toContain('dateCompatibility="1"');
    expect(workbookXml).toContain('calcMode="autoNoTable"');
    expect(workbookXml).toContain('iterateCount="77"');
    expect(workbookXml).toContain('lockStructure="1"');
  });

  test("keeps workbook settings, styles.xml, formula cells, and style ids together on styled fallback workbooks", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-styled-formula-fallback-"));
    const filePath = path.join(dir, "styled-formula-fallback.xlsx");
    await writeFile(filePath, buildExternalExcelStyledFormulaZip());

    const beforeWorkbook = await runCli(["get", filePath, "/workbook", "--json"]);
    const beforeCell = await runCli(["get", filePath, "/Sheet1/B1", "--json"]);

    await runCli([
      "set",
      filePath,
      "/workbook",
      "--prop",
      "calc.iterateCount=77",
      "--prop",
      "workbook.lockStructure=true",
    ]);
    await runCli([
      "set",
      filePath,
      "/Sheet1/B1",
      "--prop",
      "value=34",
      "--prop",
      "formula==SUM(A1:A1)",
    ]);

    const afterCell = await runCli(["get", filePath, "/Sheet1/B1", "--json"]);
    const zip = readStoredZip(await readFile(filePath));
    const workbookXml = zip.get("xl/workbook.xml")!.toString("utf8");
    const stylesXml = zip.get("xl/styles.xml")!.toString("utf8");
    const sheetXml = zip.get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(beforeWorkbook.stdout).toContain('"date1904": true');
    expect(beforeWorkbook.stdout).toContain('"calcMode": "autoNoTable"');
    expect(beforeCell.stdout).toContain('"styleId": "1"');
    expect(beforeCell.stdout).toContain('"formula": "SUM(A1:A1)"');
    expect(afterCell.stdout).toContain('"styleId": "1"');
    expect(afterCell.stdout).toContain('"formula": "SUM(A1:A1)"');
    expect(afterCell.stdout).toContain('"value": "34"');
    expect(workbookXml).toContain('calcMode="autoNoTable"');
    expect(workbookXml).toContain('iterateCount="77"');
    expect(workbookXml).toContain('lockStructure="1"');
    expect(stylesXml).toContain("cellXfs");
    expect(sheetXml).toContain(' s="1"');
    expect(sheetXml).toContain("<f>SUM(A1:A1)</f>");
    expect(sheetXml).toContain("<v>34</v>");
  });

  test("mutates a real harvested OfficeCLI styled workbook without dropping style ids", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-real-style-fixture-"));
    const fixturePath = path.resolve(
      import.meta.dir,
      "../../../fixtures/officecli-source/examples/excel/outputs/beautiful_charts.xlsx",
    );
    const filePath = path.join(dir, "beautiful_charts.xlsx");
    await writeFile(filePath, await readFile(fixturePath));

    const before = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    await runCli([
      "set",
      filePath,
      "/Sheet1/G2",
      "--prop",
      "formula==SUM(B2:E2)",
      "--prop",
      "value=303",
      "--prop",
      "styleId=10",
    ]);
    const after = await runCli(["get", filePath, "/Sheet1/G2", "--json"]);
    const zip = readStoredZip(await readFile(filePath));
    const stylesXml = zip.get("xl/styles.xml")!.toString("utf8");
    const sheetXml = zip.get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(before.stdout).toContain('"styleId": "8"');
    expect(before.stdout).toContain('"value": "120"');
    expect(after.stdout).toContain('"styleId": "10"');
    expect(after.stdout).toContain('"formula": "SUM(B2:E2)"');
    expect(after.stdout).toContain('"value": "303"');
    expect(stylesXml).toContain("cellXfs");
    expect(sheetXml).toContain('r="G2"');
    expect(sheetXml).toContain(' s="10"');
    expect(sheetXml).toContain("<f>SUM(B2:E2)</f>");
    expect(sheetXml).toContain("<v>303</v>");
  });

  test("keeps workbook settings, styles.xml, formula cells, and style ids together on a real harvested OfficeCLI workbook", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-real-combined-fixture-"));
    const fixturePath = path.resolve(
      import.meta.dir,
      "../../../fixtures/officecli-source/examples/excel/outputs/beautiful_charts.xlsx",
    );
    const filePath = path.join(dir, "beautiful_charts.xlsx");
    await writeFile(filePath, await readFile(fixturePath));

    await runCli([
      "set",
      filePath,
      "/workbook",
      "--prop",
      "date1904=true",
      "--prop",
      "codeName=OfficekitCharts",
      "--prop",
      "filterPrivacy=true",
      "--prop",
      "showObjects=all",
      "--prop",
      "calc.mode=autoNoTable",
      "--prop",
      "workbook.lockStructure=true",
    ]);
    await runCli([
      "set",
      filePath,
      "/Sheet1/G2",
      "--prop",
      "formula==SUM(B2:E2)",
      "--prop",
      "value=303",
      "--prop",
      "styleId=10",
    ]);

    const workbook = await runCli(["get", filePath, "/workbook", "--json"]);
    const cell = await runCli(["get", filePath, "/Sheet1/G2", "--json"]);
    const zip = readStoredZip(await readFile(filePath));
    const workbookXml = zip.get("xl/workbook.xml")!.toString("utf8");
    const stylesXml = zip.get("xl/styles.xml")!.toString("utf8");
    const sheetXml = zip.get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(workbook.stdout).toContain('"date1904": true');
    expect(workbook.stdout).toContain('"codeName": "OfficekitCharts"');
    expect(workbook.stdout).toContain('"filterPrivacy": true');
    expect(workbook.stdout).toContain('"showObjects": "all"');
    expect(workbook.stdout).toContain('"calcMode": "autoNoTable"');
    expect(workbook.stdout).toContain('"lockStructure": true');
    expect(cell.stdout).toContain('"styleId": "10"');
    expect(cell.stdout).toContain('"formula": "SUM(B2:E2)"');
    expect(cell.stdout).toContain('"value": "303"');
    expect(workbookXml).toContain('date1904="1"');
    expect(workbookXml).toContain('codeName="OfficekitCharts"');
    expect(workbookXml).toContain('filterPrivacy="1"');
    expect(workbookXml).toContain('showObjects="all"');
    expect(workbookXml).toContain('calcMode="autoNoTable"');
    expect(workbookXml).toContain('lockStructure="1"');
    expect(stylesXml).toContain("cellXfs");
    expect(sheetXml).toContain('r="G2"');
    expect(sheetXml).toContain(' s="10"');
    expect(sheetXml).toContain("<f>SUM(B2:E2)</f>");
    expect(sheetXml).toContain("<v>303</v>");
  });

  test("mutates a real harvested OfficeCLI formula workbook without dropping formulas", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-real-formula-fixture-"));
    const fixturePath = path.resolve(
      import.meta.dir,
      "../../../fixtures/officecli-source/examples/excel/outputs/sales_report.xlsx",
    );
    const filePath = path.join(dir, "sales_report.xlsx");
    await writeFile(filePath, await readFile(fixturePath));

    const before = await runCli(["get", filePath, "/Sheet1/F3", "--json"]);
    await runCli([
      "set",
      filePath,
      "/Sheet1/F3",
      "--prop",
      "formula==SUM(B3:E3)",
      "--prop",
      "value=412",
    ]);
    const after = await runCli(["get", filePath, "/Sheet1/F3", "--json"]);
    const zip = readStoredZip(await readFile(filePath));
    const sheetXml = zip.get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(before.stdout).toContain('"formula": "SUM(B3:E3)"');
    expect(before.stdout).toContain('"value": ""');
    expect(after.stdout).toContain('"formula": "SUM(B3:E3)"');
    expect(after.stdout).toContain('"value": "412"');
    expect(sheetXml).toContain('r="F3"');
    expect(sheetXml).toContain("<f>SUM(B3:E3)</f>");
    expect(sheetXml).toContain("<v>412</v>");
  });

  test("mutates a second real harvested styled workbook without dropping existing style ids", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-real-style-fixture-2-"));
    const fixturePath = path.resolve(
      import.meta.dir,
      "../../../fixtures/officecli-source/examples/excel/outputs/charts_demo.xlsx",
    );
    const filePath = path.join(dir, "charts_demo.xlsx");
    await writeFile(filePath, await readFile(fixturePath));

    const before = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    await runCli(["set", filePath, "/Sheet1/B2", "--prop", "value=121"]);
    const after = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    const zip = readStoredZip(await readFile(filePath));
    const stylesXml = zip.get("xl/styles.xml")!.toString("utf8");
    const sheetXml = zip.get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(before.stdout).toContain('"styleId": "7"');
    expect(before.stdout).toContain('"value": "120"');
    expect(after.stdout).toContain('"styleId": "7"');
    expect(after.stdout).toContain('"value": "121"');
    expect(stylesXml).toContain("cellXfs");
    expect(sheetXml).toContain('r="B2" s="7"');
    expect(sheetXml).toContain("<v>121</v>");
  });

  test("reads and mutates a metadata-free standard PowerPoint OOXML file", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-ppt-fallback-"));
    const filePath = path.join(dir, "fallback.pptx");
    await writeFile(filePath, buildExternalPptZip("Imported title", "Imported shape"));

    const before = await runCli(["view", filePath, "outline"]);
    expect(before.stdout).toContain("Slide 1: Imported title");
    expect(before.stdout).toContain("Shape 1: Imported shape");

    await runCli(["add", filePath, "/slide[1]", "--type", "shape", "--prop", "text=New shape"]);
    const after = await runCli(["view", filePath, "outline"]);
    expect(after.stdout).toContain("Shape 2: New shape");
  });

  test("falls back to OOXML parsing when officekit metadata is stripped", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-strip-meta-"));
    const filePath = path.join(dir, "demo.docx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/body", "--type", "paragraph", "--prop", "text=No metadata needed"]);

    const zipEntries = readStoredZip(await readFile(filePath));
    zipEntries.delete("officekit/document.json");
    await writeFile(filePath, createStoredZip([...zipEntries.entries()].map(([name, data]) => ({ name, data }))));

    const result = await runCli(["get", filePath, "/body/p[1]", "--json"]);
    expect(result.stdout).toContain("No metadata needed");
  });

  test("parses a real OfficeCLI PowerPoint fixture without officekit metadata", async () => {
    const fixturePath = path.resolve(import.meta.dir, "../../../fixtures/officecli-source/examples/Alien_Guide.pptx");
    const result = await runCli(["view", fixturePath, "outline"]);
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain("Slide 1:");
    expect(result.stdout).toContain("外星人地球");
    expect(result.stdout).toContain("Shape 1:");
  });

  test("reads a deflated metadata-free OOXML workbook with shared strings", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-xlsx-deflated-"));
    const filePath = path.join(dir, "shared.xlsx");
    await writeFile(filePath, buildDeflatedExternalExcelZip());

    const result = await runCli(["get", filePath, "/Sheet1/A1", "--json"]);
    const outline = await runCli(["view", filePath, "outline"]);

    expect(result.stdout).toContain('"Shared hello"');
    expect(outline.stdout).toContain("Sheet Sheet1");
    expect(outline.stdout).toContain("A1: Shared hello");
  });

  test("imports CSV into an Excel sheet and enables header affordances", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-import-"));
    const filePath = path.join(dir, "import.xlsx");
    const csvPath = path.join(dir, "sales.csv");
    await writeFile(csvPath, "Month,Sales\nJan,120\nFeb,135\n");
    await runCli(["create", filePath]);

    const result = await runCli([
      "import",
      filePath,
      "/Sheet1",
      csvPath,
      "--format",
      "csv",
      "--header",
      "--start-cell",
      "B2",
    ]);

    const workbook = await runCli(["view", filePath, "outline"]);
    const cell = await runCli(["get", filePath, "/Sheet1/C3", "--json"]);
    const sheetXml = readStoredZip(await readFile(filePath)).get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(result.stdout).toContain('"importedRows": 3');
    expect(result.stdout).toContain('"importedCols": 2');
    expect(result.stdout).toContain('"autoFilter": "B2:C4"');
    expect(result.stdout).toContain('"freezeTopLeftCell": "B3"');
    expect(workbook.stdout).toContain("Sheet Sheet1");
    expect(cell.stdout).toContain('"value": "120"');
    expect(sheetXml).toContain('autoFilter ref="B2:C4"');
    expect(sheetXml).toContain('topLeftCell="B3"');
    expect(sheetXml).toContain('r="B2"');
    expect(sheetXml).toContain('r="C4"');
  });

  test("imports quoted CSV fields, embedded newlines, and type inference", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-import-quoted-"));
    const filePath = path.join(dir, "quoted.xlsx");
    const csvPath = path.join(dir, "quoted.csv");
    await writeFile(
      csvPath,
      'Name,Active,Formula,Date,Notes\n"Alpha, Inc",TRUE,=SUM(A2:A2),2025-04-05,"Line 1\nLine 2"\n',
    );
    await runCli(["create", filePath]);

    await runCli([
      "import",
      filePath,
      "/Sheet1",
      csvPath,
      "--format",
      "csv",
      "--header",
      "--start-cell",
      "A1",
    ]);

    const nameCell = await runCli(["get", filePath, "/Sheet1/A2", "--json"]);
    const boolCell = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    const formulaCell = await runCli(["get", filePath, "/Sheet1/C2", "--json"]);
    const dateCell = await runCli(["get", filePath, "/Sheet1/D2", "--json"]);
    const noteCell = await runCli(["get", filePath, "/Sheet1/E2", "--json"]);
    const sheetXml = readStoredZip(await readFile(filePath)).get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(nameCell.stdout).toContain('"value": "Alpha, Inc"');
    expect(boolCell.stdout).toContain('"value": "1"');
    expect(boolCell.stdout).toContain('"type": "boolean"');
    expect(formulaCell.stdout).toContain('"formula": "SUM(A2:A2)"');
    expect(dateCell.stdout).toContain('"type": "date"');
    expect(noteCell.stdout).toContain('Line 1\\nLine 2');
    expect(sheetXml).toContain('r="A2" t="inlineStr"');
    expect(sheetXml).toContain('r="B2" t="b"');
    expect(sheetXml).toContain('<f>SUM(A2:A2)</f>');
    expect(sheetXml).toContain('r="D2"><v>');
  });

  test("imports TSV with the correct delimiter", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-import-tsv-"));
    const filePath = path.join(dir, "tsv.xlsx");
    const tsvPath = path.join(dir, "data.tsv");
    await writeFile(tsvPath, "Name\tScore\nAlice\t98\n");
    await runCli(["create", filePath]);

    await runCli([
      "import",
      filePath,
      "/Sheet1",
      tsvPath,
      "--format",
      "tsv",
      "--start-cell",
      "C3",
    ]);

    const nameCell = await runCli(["get", filePath, "/Sheet1/C4", "--json"]);
    const scoreCell = await runCli(["get", filePath, "/Sheet1/D4", "--json"]);
    expect(nameCell.stdout).toContain('"value": "Alice"');
    expect(scoreCell.stdout).toContain('"value": "98"');
  });

  test("imports CSV with escaped quotes via stdin", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-import-stdin-"));
    const filePath = path.join(dir, "stdin.xlsx");
    await runCli(["create", filePath]);

    const child = spawn(process.execPath, [
      "run",
      "packages/cli/bin/officekit",
      "import",
      filePath,
      "/Sheet1",
      "--stdin",
      "--format",
      "csv",
      "--start-cell",
      "A1",
    ], {
      cwd: path.resolve(import.meta.dir, "../../.."),
      stdio: ["pipe", "pipe", "pipe"],
    });

    child.stdin.write('Name,Quote\n"Alpha ""Prime""","He said ""hi"""\n');
    child.stdin.end();

    const stdout = await waitForOutput(child.stdout, /"importedRows":\s*2/);
    const exitCode = await new Promise<number | null>((resolve) => child.once("exit", resolve));
    expect(exitCode).toBe(0);
    expect(stdout).toContain('"importedCols": 2');

    const nameCell = await runCli(["get", filePath, "/Sheet1/A2", "--json"]);
    const quoteCell = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    expect(nameCell.stdout).toContain('Alpha \\"Prime\\"');
    expect(quoteCell.stdout).toContain('He said \\"hi\\"');
  });

  test("imports TSV with formulas, booleans, dates, and numbers via extension inference", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-import-tsv-types-"));
    const filePath = path.join(dir, "types.xlsx");
    const tsvPath = path.join(dir, "typed.tsv");
    await writeFile(tsvPath, "Flag\tFormula\tDate\tValue\nTRUE\t=SUM(A2:A2)\t2025-04-05\t42\n");
    await runCli(["create", filePath]);

    await runCli([
      "import",
      filePath,
      "/Sheet1",
      tsvPath,
      "--header",
      "--start-cell",
      "A1",
    ]);

    const boolCell = await runCli(["get", filePath, "/Sheet1/A2", "--json"]);
    const formulaCell = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    const dateCell = await runCli(["get", filePath, "/Sheet1/C2", "--json"]);
    const numberCell = await runCli(["get", filePath, "/Sheet1/D2", "--json"]);
    const sheetXml = readStoredZip(await readFile(filePath)).get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(boolCell.stdout).toContain('"type": "boolean"');
    expect(formulaCell.stdout).toContain('"formula": "SUM(A2:A2)"');
    expect(dateCell.stdout).toContain('"type": "date"');
    expect(numberCell.stdout).toContain('"type": "number"');
    expect(sheetXml).toContain('r="A2" t="b"');
    expect(sheetXml).toContain('<f>SUM(A2:A2)</f>');
    expect(sheetXml).toContain('r="D2"><v>42</v>');
  });

  test("round-trips an imported workbook without dropping header affordances or inferred cell semantics", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-import-roundtrip-"));
    const filePath = path.join(dir, "roundtrip.xlsx");
    const csvPath = path.join(dir, "roundtrip.csv");
    await writeFile(
      csvPath,
      'Name,Enabled,Formula,Date,Amount\nAlpha,TRUE,=SUM(E2:E2),2025-04-05,42\n',
    );
    await runCli(["create", filePath]);
    await runCli([
      "import",
      filePath,
      "/Sheet1",
      csvPath,
      "--format",
      "csv",
      "--header",
      "--start-cell",
      "A1",
    ]);

    const importedWorkbook = await runCli(["get", filePath, "/workbook", "--json"]);
    expect(importedWorkbook.stdout).toContain('"autoFilter": "A1:E2"');
    expect(importedWorkbook.stdout).toContain('"freezeTopLeftCell": "A2"');

    await runCli([
      "set",
      filePath,
      "/workbook",
      "--prop",
      "calc.mode=autoNoTable",
      "--prop",
      "workbook.lockStructure=true",
    ]);
    await runCli(["set", filePath, "/Sheet1/E2", "--prop", "value=84"]);

    const boolCell = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    const formulaCell = await runCli(["get", filePath, "/Sheet1/C2", "--json"]);
    const dateCell = await runCli(["get", filePath, "/Sheet1/D2", "--json"]);
    const amountCell = await runCli(["get", filePath, "/Sheet1/E2", "--json"]);
    const workbook = await runCli(["get", filePath, "/workbook", "--json"]);
    const sheetXml = readStoredZip(await readFile(filePath)).get("xl/worksheets/sheet1.xml")!.toString("utf8");
    const workbookXml = readStoredZip(await readFile(filePath)).get("xl/workbook.xml")!.toString("utf8");

    expect(boolCell.stdout).toContain('"type": "boolean"');
    expect(boolCell.stdout).toContain('"value": "1"');
    expect(formulaCell.stdout).toContain('"formula": "SUM(E2:E2)"');
    expect(dateCell.stdout).toContain('"type": "date"');
    expect(amountCell.stdout).toContain('"type": "number"');
    expect(amountCell.stdout).toContain('"value": "84"');
    expect(workbook.stdout).toContain('"autoFilter": "A1:E2"');
    expect(workbook.stdout).toContain('"freezeTopLeftCell": "A2"');
    expect(workbook.stdout).toContain('"calcMode": "autoNoTable"');
    expect(workbook.stdout).toContain('"lockStructure": true');
    expect(sheetXml).toContain('autoFilter ref="A1:E2"');
    expect(sheetXml).toContain('topLeftCell="A2"');
    expect(sheetXml).toContain('r="C2"><f>SUM(E2:E2)</f>');
    expect(workbookXml).toContain('calcMode="autoNoTable"');
    expect(workbookXml).toContain('lockStructure="1"');
  });

  test("watch keeps a preview server alive until interrupted", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-watch-"));
    const filePath = path.join(dir, "watch.docx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/body", "--type", "paragraph", "--prop", "text=Watching"]);

    const child = spawn(process.execPath, ["run", "packages/cli/bin/officekit", "watch", filePath, "--port", "0"], {
      cwd: path.resolve(import.meta.dir, "../../.."),
      stdio: ["ignore", "pipe", "pipe"],
    });

    const stdout = await waitForOutput(child.stdout, /"url":\s*"([^"]+)"/);
    const urlMatch = stdout.match(/"url":\s*"([^"]+)"/);
    expect(urlMatch).not.toBeNull();
    const url = urlMatch![1];

    const health = (await fetch(`${url}/health`).then((response) => response.json())) as { ok: boolean; version: number; clients: number };
    const html = await fetch(url).then((response) => response.text());

    expect(health.ok).toBe(true);
    expect(html).toContain("Watching");

    child.kill("SIGINT");
    const exitCode = await new Promise<number | null>((resolve) => child.once("exit", resolve));
    expect(exitCode).toBe(0);
  });
});

function buildExternalWordZip(text: string) {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: "word/document.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`),
    },
  ]);
}

function buildExternalWordTableZip(left: string, right: string) {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: "word/document.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>${left}</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>${right}</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>`),
    },
  ]);
}

function buildExternalMixedWordZip(firstParagraph: string, left: string, right: string, secondParagraph: string) {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: "word/document.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>${firstParagraph}</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>${left}</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>${right}</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>${secondParagraph}</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`),
    },
  ]);
}

function buildExternalExcelZip(ref: string, value: string) {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    },
    {
      name: "xl/workbook.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    },
    {
      name: "xl/worksheets/sheet1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c></row></sheetData>
</worksheet>`),
    },
  ]);
}

function buildExternalExcelFormulaZip() {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    },
    {
      name: "xl/workbook.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    },
    {
      name: "xl/worksheets/sheet1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>21</v></c>
      <c r="B1"><f>SUM(A1:A1)</f><v>21</v></c>
    </row>
  </sheetData>
</worksheet>`),
    },
  ]);
}

function buildExternalExcelSettingsZip() {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    },
    {
      name: "xl/workbook.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr date1904="1" codeName="WorkbookCode" filterPrivacy="1" showObjects="all"/>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    },
    {
      name: "xl/styles.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`),
    },
    {
      name: "xl/worksheets/sheet1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Styled</t></is></c></row></sheetData>
</worksheet>`),
    },
  ]);
}

function buildExternalExcelExtendedSettingsZip() {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    },
    {
      name: "xl/workbook.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr backupFile="1" dateCompatibility="1" codeName="WorkbookCode"/>
  <workbookProtection lockStructure="1" lockWindows="1"/>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="autoNoTable" iterate="1" iterateCount="5" iterateDelta="0.001" fullPrecision="1" fullCalcOnLoad="1" refMode="R1C1"/>
</workbook>`),
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    },
    {
      name: "xl/worksheets/sheet1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData></sheetData>
</worksheet>`),
    },
  ]);
}

function buildExternalExcelStyledFormulaZip() {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    },
    {
      name: "xl/workbook.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr date1904="1" codeName="WorkbookCode" filterPrivacy="1" showObjects="all" backupFile="1" dateCompatibility="1"/>
  <workbookProtection lockStructure="1" lockWindows="1"/>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="autoNoTable" iterate="1" iterateCount="5" iterateDelta="0.001" fullPrecision="1" fullCalcOnLoad="1" refMode="R1C1"/>
</workbook>`),
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    },
    {
      name: "xl/styles.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`),
    },
    {
      name: "xl/worksheets/sheet1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>34</v></c>
      <c r="B1" s="1"><f>SUM(A1:A1)</f><v>34</v></c>
    </row>
  </sheetData>
</worksheet>`),
    },
  ]);
}

function buildExternalPptZip(title: string, shape: string) {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`),
    },
    {
      name: "ppt/presentation.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`),
    },
    {
      name: "ppt/_rels/presentation.xml.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`),
    },
    {
      name: "ppt/slides/slide1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${title}</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${shape}</a:t></a:r></a:p></p:txBody></p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`),
    },
  ]);
}

function buildDeflatedExternalExcelZip() {
  return createZipWithCompression([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
      compression: 8,
    },
    {
      name: "_rels/.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
      compression: 8,
    },
    {
      name: "xl/workbook.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
      compression: 8,
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`),
      compression: 8,
    },
    {
      name: "xl/sharedStrings.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Shared hello</t></si>
</sst>`),
      compression: 8,
    },
    {
      name: "xl/worksheets/sheet1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>
</worksheet>`),
      compression: 8,
    },
  ]);
}

function createZipWithCompression(entries: Array<{ name: string; data: Buffer; compression?: 0 | 8 }>) {
  const localParts: Buffer[] = [];
  const centralParts: Buffer[] = [];
  let offset = 0;

  for (const entry of entries) {
    const name = Buffer.from(entry.name, "utf8");
    const compression = entry.compression ?? 0;
    const body = compression === 8 ? deflateRawSync(entry.data) : entry.data;

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0, 6);
    localHeader.writeUInt16LE(compression, 8);
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(0, 14);
    localHeader.writeUInt32LE(body.length, 18);
    localHeader.writeUInt32LE(entry.data.length, 22);
    localHeader.writeUInt16LE(name.length, 26);
    localHeader.writeUInt16LE(0, 28);
    localParts.push(localHeader, name, body);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0, 8);
    centralHeader.writeUInt16LE(compression, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0, 14);
    centralHeader.writeUInt32LE(0, 16);
    centralHeader.writeUInt32LE(body.length, 20);
    centralHeader.writeUInt32LE(entry.data.length, 24);
    centralHeader.writeUInt16LE(name.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    centralParts.push(centralHeader, name);

    offset += localHeader.length + name.length + body.length;
  }

  const centralDirectory = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralDirectory.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);

  return Buffer.concat([...localParts, centralDirectory, end]);
}

async function waitForOutput(stream: NodeJS.ReadableStream | null, pattern: RegExp) {
  if (!stream) {
    throw new Error("Missing child stdout stream.");
  }

  return new Promise<string>((resolve, reject) => {
    let collected = "";
    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error(`Timed out waiting for output matching ${pattern}`));
    }, 5_000);

    const onData = (chunk: Buffer | string) => {
      collected += chunk.toString();
      if (pattern.test(collected)) {
        cleanup();
        resolve(collected);
      }
    };

    const onError = (error: Error) => {
      cleanup();
      reject(error);
    };

    const cleanup = () => {
      clearTimeout(timeout);
      stream.off("data", onData);
      stream.off("error", onError);
    };

    stream.on("data", onData);
    stream.on("error", onError);
  });
}
