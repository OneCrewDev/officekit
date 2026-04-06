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

  test("adds, gets, sets, and removes named ranges", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-namedrange-"));
    const filePath = path.join(dir, "namedrange.xlsx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/", "--type", "sheet", "--prop", "name=Analysis"]);
    await runCli([
      "add",
      filePath,
      "/",
      "--type",
      "namedrange",
      "--prop",
      "name=SalesRange",
      "--prop",
      "ref=Sheet1!A1:B5",
      "--prop",
      "scope=Analysis",
      "--prop",
      "comment=Tracked sales block",
    ]);

    const namedRange = await runCli(["get", filePath, "/namedrange[SalesRange]", "--json"]);
    expect(namedRange.stdout).toContain('"name": "SalesRange"');
    expect(namedRange.stdout).toContain('"ref": "Sheet1!A1:B5"');
    expect(namedRange.stdout).toContain('"scope": "Analysis"');
    expect(namedRange.stdout).toContain('"comment": "Tracked sales block"');

    await runCli([
      "set",
      filePath,
      "/namedrange[SalesRange]",
      "--prop",
      "ref=Analysis!C1:D4",
      "--prop",
      "comment=Updated block",
    ]);
    const updated = await runCli(["get", filePath, "/namedrange[1]", "--json"]);
    expect(updated.stdout).toContain('"ref": "Analysis!C1:D4"');
    expect(updated.stdout).toContain('"comment": "Updated block"');

    const workbookXml = readStoredZip(await readFile(filePath)).get("xl/workbook.xml")!.toString("utf8");
    expect(workbookXml).toContain('<definedName name="SalesRange"');
    expect(workbookXml).toContain('localSheetId="1"');
    expect(workbookXml).toContain('comment="Updated block"');

    await runCli(["remove", filePath, "/namedrange[SalesRange]"]);
    const raw = await runCli(["raw", filePath]);
    expect(raw.stdout).not.toContain('SalesRange');
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

  test("supports Excel sheet properties, extended views, and query selectors", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-sheet-props-"));
    const filePath = path.join(dir, "sheet-props.xlsx");
    await runCli(["create", filePath]);
    await runCli([
      "set",
      filePath,
      "/Sheet1",
      "--prop",
      "freeze=B2",
      "--prop",
      "zoom=125",
      "--prop",
      "gridlines=false",
      "--prop",
      "headings=false",
      "--prop",
      "tabColor=1A2B3C",
      "--prop",
      "header=&CQuarterly Report",
      "--prop",
      "footer=&RConfidential",
      "--prop",
      "orientation=landscape",
      "--prop",
      "paperSize=9",
      "--prop",
      "fitToPage=1x2",
      "--prop",
      "protect=true",
      "--prop",
      "autoFilter=A1:B5",
      "--prop",
      "rowBreaks=5,10",
      "--prop",
      "colBreaks=3",
    ]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=12"]);
    await runCli(["set", filePath, "/Sheet1/B1", "--prop", "formula==SUM(A1:A1)", "--prop", "value=12"]);
    await runCli(["add", filePath, "/", "--type", "namedrange", "--prop", "name=Revenue", "--prop", "ref=Sheet1!A1:B1"]);

    const sheet = await runCli(["get", filePath, "/Sheet1", "--json"]);
    const row = await runCli(["get", filePath, "/Sheet1/row[1]", "--json"]);
    const column = await runCli(["get", filePath, "/Sheet1/col[A]", "--json"]);
    const range = await runCli(["get", filePath, "/Sheet1/A1:B1", "--json"]);
    const textView = await runCli(["view", filePath, "text"]);
    const annotatedView = await runCli(["view", filePath, "annotated"]);
    const statsView = await runCli(["view", filePath, "stats"]);
    const issuesView = await runCli(["view", filePath, "issues"]);
    const formulaQuery = await runCli(["query", filePath, "formula"]);
    const sheetQuery = await runCli(["query", filePath, "sheet"]);
    const namedRangeQuery = await runCli(["query", filePath, "namedrange"]);
    const rawSheet = await runCli(["raw", filePath, "/Sheet1"]);

    expect(sheet.stdout).toContain('"freezeTopLeftCell": "B2"');
    expect(sheet.stdout).toContain('"zoom": 125');
    expect(sheet.stdout).toContain('"showGridLines": false');
    expect(sheet.stdout).toContain('"showHeadings": false');
    expect(sheet.stdout).toContain('"tabColor": "FF1A2B3C"');
    expect(sheet.stdout).toContain('"orientation": "landscape"');
    expect(sheet.stdout).toContain('"paperSize": 9');
    expect(sheet.stdout).toContain('"fitToPage": "1x2"');
    expect(sheet.stdout).toContain('"protection": true');
    expect(sheet.stdout).toContain('"rowBreaks": [');
    expect(row.stdout).toContain('"ref": "A1"');
    expect(column.stdout).toContain('"ref": "A1"');
    expect(range.stdout).toContain('"ref": "A1"');
    expect(range.stdout).toContain('"ref": "B1"');
    expect(textView.stdout).toContain("[/Sheet1/row[1]] 12\t12");
    expect(annotatedView.stdout).toContain("B1: [12] <- =SUM(A1:A1)");
    expect(statsView.stdout).toContain("Formula Cells: 1");
    expect(issuesView.stdout).toContain("No issues found.");
    expect(formulaQuery.stdout).toContain('"ref": "B1"');
    expect(sheetQuery.stdout).toContain('"name": "Sheet1"');
    expect(namedRangeQuery.stdout).toContain('"name": "Revenue"');
    expect(rawSheet.stdout).toContain('zoomScale="125"');
    expect(rawSheet.stdout).toContain('showGridLines="0"');
    expect(rawSheet.stdout).toContain('showRowColHeaders="0"');
    expect(rawSheet.stdout).toContain('tabColor rgb="FF1A2B3C"');
    expect(rawSheet.stdout).toContain('orientation="landscape"');
    expect(rawSheet.stdout).toContain('paperSize="9"');
    expect(rawSheet.stdout).toContain('sheetProtection sheet="1"');
  });

  test("supports Excel raw parts and filtered raw sheet output", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-raw-parts-"));
    const filePath = path.join(dir, "raw-parts.xlsx");
    const sharedFilePath = path.join(dir, "shared-raw.xlsx");
    await writeFile(filePath, buildExternalExcelSettingsZip());
    await writeFile(sharedFilePath, buildDeflatedExternalExcelZip());

    const workbookRaw = await runCli(["raw", filePath, "/workbook"]);
    const stylesRaw = await runCli(["raw", filePath, "/styles"]);
    const sharedStringsRaw = await runCli(["raw", sharedFilePath, "/sharedstrings"]);
    const sheetRaw = await runCli(["raw", filePath, "/Sheet1", "--start-row", "1", "--end-row", "1", "--cols", "A"]);

    expect(workbookRaw.stdout).toContain("<workbook");
    expect(stylesRaw.stdout).toContain("<styleSheet");
    expect(sharedStringsRaw.stdout).toContain("Shared hello");
    expect(sheetRaw.stdout).toContain('r="A1"');
  });

  test("supports advanced Excel object mutation paths", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-advanced-"));
    const filePath = path.join(dir, "advanced.xlsx");
    await writeFile(filePath, buildExternalExcelAdvancedObjectsZip());

    const beforeValidation = await runCli(["get", filePath, "/Sheet1/validation[1]", "--json"]);
    const beforeComment = await runCli(["get", filePath, "/Sheet1/comment[1]", "--json"]);
    const beforeTable = await runCli(["get", filePath, "/Sheet1/table[1]", "--json"]);
    const beforeChart = await runCli(["get", filePath, "/Sheet1/chart[1]", "--json"]);
    const beforePivot = await runCli(["get", filePath, "/Sheet1/pivottable[1]", "--json"]);
    const beforeSparkline = await runCli(["get", filePath, "/Sheet1/sparkline[1]", "--json"]);
    const beforeShape = await runCli(["get", filePath, "/Sheet1/shape[1]", "--json"]);
    const beforePicture = await runCli(["get", filePath, "/Sheet1/picture[1]", "--json"]);

    expect(beforeValidation.stdout).toContain('"validationType": "list"');
    expect(beforeComment.stdout).toContain('"text": "Initial note"');
    expect(beforeTable.stdout).toContain('"name": "Table1"');
    expect(beforeChart.stdout).toContain('"title": "Initial Chart"');
    expect(beforePivot.stdout).toContain('"name": "PivotTable1"');
    expect(beforeSparkline.stdout).toContain('"location": "C2"');
    expect(beforeShape.stdout).toContain('"name": "Shape 1"');
    expect(beforePicture.stdout).toContain('"name": "Picture 1"');

    await runCli(["set", filePath, "/Sheet1/validation[1]", "--prop", "formula1=Yes,No", "--prop", "prompt=Pick one"]);
    await runCli(["set", filePath, "/Sheet1/comment[1]", "--prop", "text=Updated note", "--prop", "author=officekit"]);
    await runCli(["set", filePath, "/Sheet1/table[1]", "--prop", "name=SalesTable", "--prop", "ref=A1:B3", "--prop", "totalsrow=true"]);
    await runCli(["set", filePath, "/Sheet1/sparkline[1]", "--prop", "type=column", "--prop", "location=D2", "--prop", "sourceRange=A2:B2"]);
    await runCli(["set", filePath, "/Sheet1/chart[1]", "--prop", "title=Updated Chart"]);
    await runCli(["set", filePath, "/Sheet1/chart[1]/series[1]", "--prop", "name=Revenue Series"]);
    await runCli(["set", filePath, "/Sheet1/pivottable[1]", "--prop", "name=Pivot Summary", "--prop", "rowGrandTotals=false"]);
    await runCli(["set", filePath, "/Sheet1/shape[1]", "--prop", "name=Updated Shape", "--prop", "text=Shape Copy", "--prop", "x=2", "--prop", "y=3"]);
    await runCli(["set", filePath, "/Sheet1/picture[1]", "--prop", "name=Updated Picture", "--prop", "alt=Preview image"]);

    const validation = await runCli(["get", filePath, "/Sheet1/validation[1]", "--json"]);
    const comment = await runCli(["get", filePath, "/Sheet1/comment[1]", "--json"]);
    const table = await runCli(["get", filePath, "/Sheet1/table[1]", "--json"]);
    const chart = await runCli(["get", filePath, "/Sheet1/chart[1]", "--json"]);
    const pivot = await runCli(["get", filePath, "/Sheet1/pivottable[1]", "--json"]);
    const sparkline = await runCli(["get", filePath, "/Sheet1/sparkline[1]", "--json"]);
    const sparklineQuery = await runCli(["query", filePath, "sparkline"]);
    const shapeQuery = await runCli(["query", filePath, "shape"]);
    const pictureQuery = await runCli(["query", filePath, "picture"]);
    const shape = await runCli(["get", filePath, "/Sheet1/shape[1]", "--json"]);
    const picture = await runCli(["get", filePath, "/Sheet1/picture[1]", "--json"]);
    const rawChart = await runCli(["raw", filePath, "/Sheet1/chart[1]"]);
    const rawDrawing = await runCli(["raw", filePath, "/Sheet1/drawing"]);

    expect(validation.stdout).toContain('"formula1": "Yes,No"');
    expect(validation.stdout).toContain('"prompt": "Pick one"');
    expect(comment.stdout).toContain('"text": "Updated note"');
    expect(comment.stdout).toContain('"author": "officekit"');
    expect(table.stdout).toContain('"name": "SalesTable"');
    expect(table.stdout).toContain('"ref": "A1:B3"');
    expect(table.stdout).toContain('"totalsRow": true');
    expect(chart.stdout).toContain('"title": "Updated Chart"');
    expect(pivot.stdout).toContain('"name": "Pivot Summary"');
    expect(pivot.stdout).toContain('"rowGrandTotals": false');
    expect(sparkline.stdout).toContain('"location": "D2"');
    expect(sparkline.stdout).toContain('"sourceRange": "A2:B2"');
    expect(sparkline.stdout).toContain('"sparklineType": "column"');
    expect(sparklineQuery.stdout).toContain('"type": "sparkline"');
    expect(shape.stdout).toContain('"name": "Updated Shape"');
    expect(shape.stdout).toContain('"text": "Shape Copy"');
    expect(picture.stdout).toContain('"name": "Updated Picture"');
    expect(shapeQuery.stdout).toContain('"type": "shape"');
    expect(pictureQuery.stdout).toContain('"type": "picture"');
    expect(rawChart.stdout).toContain("Updated Chart");
    expect(rawChart.stdout).toContain("Revenue Series");
    expect(rawDrawing.stdout).toContain("Updated Shape");
    expect(rawDrawing.stdout).toContain('descr="Preview image"');
  });

  test("adds Excel worksheet objects through OfficeCLI-style add paths", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-add-objects-"));
    const filePath = path.join(dir, "add-objects.xlsx");
    await runCli(["create", filePath]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=Name"]);
    await runCli(["set", filePath, "/Sheet1/B1", "--prop", "value=Value"]);
    await runCli(["set", filePath, "/Sheet1/A2", "--prop", "value=Alpha"]);
    await runCli(["set", filePath, "/Sheet1/B2", "--prop", "value=10", "--prop", "type=number"]);

    await runCli(["add", filePath, "/Sheet1", "--type", "validation", "--prop", "ref=A2", "--prop", "type=list", "--prop", "formula1=Yes,No", "--prop", "prompt=Pick one"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "comment", "--prop", "ref=A2", "--prop", "text=Review this row", "--prop", "author=officekit"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "autofilter", "--prop", "range=A1:B2"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "rowbreak", "--prop", "row=5"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "colbreak", "--prop", "col=2"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "table", "--prop", "ref=A1:B2", "--prop", "name=SalesTable", "--prop", "columns=Name,Value"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "sparkline", "--prop", "cell=C2", "--prop", "range=A2:B2", "--prop", "type=column"]);

    const sheet = await runCli(["get", filePath, "/Sheet1", "--json"]);
    const validation = await runCli(["get", filePath, "/Sheet1/validation[1]", "--json"]);
    const comment = await runCli(["get", filePath, "/Sheet1/comment[1]", "--json"]);
    const table = await runCli(["get", filePath, "/Sheet1/table[1]", "--json"]);
    const sparkline = await runCli(["get", filePath, "/Sheet1/sparkline[1]", "--json"]);
    const rowBreak = await runCli(["get", filePath, "/Sheet1/rowbreak[1]", "--json"]);
    const colBreak = await runCli(["get", filePath, "/Sheet1/colbreak[1]", "--json"]);

    expect(sheet.stdout).toContain('"autoFilter": "A1:B2"');
    expect(sheet.stdout).toContain('"rowBreaks": [');
    expect(sheet.stdout).toContain('"colBreaks": [');
    expect(validation.stdout).toContain('"validationType": "list"');
    expect(validation.stdout).toContain('"prompt": "Pick one"');
    expect(comment.stdout).toContain('"text": "Review this row"');
    expect(comment.stdout).toContain('"author": "officekit"');
    expect(table.stdout).toContain('"name": "SalesTable"');
    expect(table.stdout).toContain('"ref": "A1:B2"');
    expect(sparkline.stdout).toContain('"location": "C2"');
    expect(sparkline.stdout).toContain('"sourceRange": "Sheet1!A2:B2"');
    expect(sparkline.stdout).toContain('"sparklineType": "column"');
    expect(rowBreak.stdout).toContain('"id": 5');
    expect(colBreak.stdout).toContain('"id": 2');

    const zip = readStoredZip(await readFile(filePath));
    const contentTypes = zip.get("[Content_Types].xml")!.toString("utf8");
    const sheetRels = zip.get("xl/worksheets/_rels/sheet1.xml.rels")!.toString("utf8");
    const commentsXml = zip.get("xl/comments1.xml")!.toString("utf8");
    const tableXml = zip.get("xl/tables/table1.xml")!.toString("utf8");
    const sheetXml = zip.get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(contentTypes).toContain("/xl/comments1.xml");
    expect(contentTypes).toContain("/xl/tables/table1.xml");
    expect(sheetRels).toContain("/relationships/comments");
    expect(sheetRels).toContain("/relationships/table");
    expect(commentsXml).toContain("Review this row");
    expect(tableXml).toContain('name="SalesTable"');
    expect(sheetXml).toContain('<tableParts count="1">');
    expect(sheetXml).toContain('<rowBreaks count="1" manualBreakCount="1">');
    expect(sheetXml).toContain('<colBreaks count="1" manualBreakCount="1">');
    expect(sheetXml).toContain('<x14:sparklineGroup type="column">');
  });

  test("adds Excel chart and pivottable objects through OfficeCLI-style add paths", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-add-chart-pivot-"));
    const filePath = path.join(dir, "add-chart-pivot.xlsx");
    await runCli(["create", filePath]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=Metric"]);
    await runCli(["set", filePath, "/Sheet1/B1", "--prop", "value=Jan"]);
    await runCli(["set", filePath, "/Sheet1/C1", "--prop", "value=Feb"]);
    await runCli(["set", filePath, "/Sheet1/A2", "--prop", "value=Revenue"]);
    await runCli(["set", filePath, "/Sheet1/B2", "--prop", "value=10", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/C2", "--prop", "value=20", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/A3", "--prop", "value=Cost"]);
    await runCli(["set", filePath, "/Sheet1/B3", "--prop", "value=6", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/C3", "--prop", "value=9", "--prop", "type=number"]);

    await runCli(["add", filePath, "/Sheet1", "--type", "chart", "--prop", "title=Quarterly Trend", "--prop", "type=column", "--prop", "dataRange=Sheet1!A1:C3", "--prop", "x=1", "--prop", "y=2", "--prop", "width=6", "--prop", "height=8"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "pivottable", "--prop", "source=Sheet1!A1:C3", "--prop", "name=SalesPivot", "--prop", "position=H1"]);
    await runCli(["set", filePath, "/Sheet1/pivottable[1]", "--prop", "rowGrandTotals=false"]);

    const chart = await runCli(["get", filePath, "/Sheet1/chart[1]", "--json"]);
    const pivot = await runCli(["get", filePath, "/Sheet1/pivottable[1]", "--json"]);
    const chartQuery = await runCli(["query", filePath, "chart"]);
    const pivotQuery = await runCli(["query", filePath, "pivottable"]);
    const rawChart = await runCli(["raw", filePath, "/Sheet1/chart[1]"]);
    const rawDrawing = await runCli(["raw", filePath, "/Sheet1/drawing"]);

    expect(chart.stdout).toContain('"title": "Quarterly Trend"');
    expect(chart.stdout).toContain('"chartType": "bar"');
    expect(chart.stdout).toContain('"seriesNames"');
    expect(chart.stdout).toContain('Revenue');
    expect(pivot.stdout).toContain('"name": "SalesPivot"');
    expect(pivot.stdout).toContain('"rowGrandTotals": false');
    expect(chartQuery.stdout).toContain('"type": "chart"');
    expect(pivotQuery.stdout).toContain('"type": "pivottable"');
    expect(rawChart.stdout).toContain("Quarterly Trend");
    expect(rawChart.stdout).toContain("Revenue");
    expect(rawDrawing.stdout).toContain("<c:chart");

    const zip = readStoredZip(await readFile(filePath));
    const contentTypes = zip.get("[Content_Types].xml")!.toString("utf8");
    const sheetRels = zip.get("xl/worksheets/_rels/sheet1.xml.rels")!.toString("utf8");
    const drawingRels = zip.get("xl/drawings/_rels/drawing1.xml.rels")!.toString("utf8");
    const pivotXml = zip.get("xl/pivotTables/pivotTable1.xml")!.toString("utf8");

    expect(contentTypes).toContain("/xl/charts/chart1.xml");
    expect(contentTypes).toContain("/xl/pivotTables/pivotTable1.xml");
    expect(sheetRels).toContain("/relationships/drawing");
    expect(sheetRels).toContain("/relationships/pivotTable");
    expect(drawingRels).toContain("/relationships/chart");
    expect(pivotXml).toContain('name="SalesPivot"');
    expect(pivotXml).toContain('rowGrandTotals="0"');
  });

  test("adds Excel picture and shape objects through OfficeCLI-style add paths", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-add-drawing-"));
    const filePath = path.join(dir, "add-drawing.xlsx");
    const imagePath = path.join(dir, "pixel.png");
    await writeFile(imagePath, tinyPngBuffer());
    await runCli(["create", filePath]);

    await runCli(["add", filePath, "/Sheet1", "--type", "shape", "--prop", "name=Callout", "--prop", "text=Ship now", "--prop", "x=1", "--prop", "y=2", "--prop", "width=4", "--prop", "height=2", "--prop", "fill=4472C4", "--prop", "color=FFFFFF", "--prop", "align=center"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "picture", "--prop", `path=${imagePath}`, "--prop", "name=HeroImage", "--prop", "alt=Launch artwork", "--prop", "x=5", "--prop", "y=1", "--prop", "width=3", "--prop", "height=3"]);
    await runCli(["set", filePath, "/Sheet1/shape[1]", "--prop", "text=Ship this week"]);
    await runCli(["set", filePath, "/Sheet1/picture[1]", "--prop", "alt=Updated artwork"]);

    const shape = await runCli(["get", filePath, "/Sheet1/shape[1]", "--json"]);
    const picture = await runCli(["get", filePath, "/Sheet1/picture[1]", "--json"]);
    const shapeQuery = await runCli(["query", filePath, "shape"]);
    const pictureQuery = await runCli(["query", filePath, "picture"]);
    const rawDrawing = await runCli(["raw", filePath, "/Sheet1/drawing"]);

    expect(shape.stdout).toContain('"name": "Callout"');
    expect(shape.stdout).toContain('"text": "Ship this week"');
    expect(picture.stdout).toContain('"name": "HeroImage"');
    expect(shapeQuery.stdout).toContain('"type": "shape"');
    expect(pictureQuery.stdout).toContain('"type": "picture"');
    expect(rawDrawing.stdout).toContain("HeroImage");
    expect(rawDrawing.stdout).toContain("Updated artwork");
    expect(rawDrawing.stdout).toContain("Ship this week");

    const zip = readStoredZip(await readFile(filePath));
    const contentTypes = zip.get("[Content_Types].xml")!.toString("utf8");
    const drawingRels = zip.get("xl/drawings/_rels/drawing1.xml.rels")!.toString("utf8");

    expect(contentTypes).toContain("image/png");
    expect(drawingRels).toContain("/relationships/image");
    expect(zip.get("xl/media/image1.png")).toBeDefined();
  });

  test("adds and updates Excel conditional-formatting families through OfficeCLI-style paths", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-cf-"));
    const filePath = path.join(dir, "cf.xlsx");
    await runCli(["create", filePath]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=10", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/A2", "--prop", "value=20", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/A3", "--prop", "value=30", "--prop", "type=number"]);

    await runCli(["add", filePath, "/Sheet1", "--type", "cf", "--prop", "type=databar", "--prop", "sqref=A1:A3", "--prop", "color=638EC6"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "colorscale", "--prop", "range=A1:A3", "--prop", "mincolor=F8696B", "--prop", "midcolor=FFEB84", "--prop", "maxcolor=63BE7B"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "iconset", "--prop", "range=A1:A3", "--prop", "iconset=4TrafficLights", "--prop", "reverse=true", "--prop", "showvalue=false"]);

    await runCli(["set", filePath, "/Sheet1/cf[1]", "--prop", "color=4472C4"]);
    await runCli(["set", filePath, "/Sheet1/cf[2]", "--prop", "mincolor=FF0000", "--prop", "maxcolor=00FF00"]);
    await runCli(["set", filePath, "/Sheet1/cf[3]", "--prop", "iconset=5Arrows", "--prop", "showvalue=true"]);

    const cf1 = await runCli(["get", filePath, "/Sheet1/cf[1]", "--json"]);
    const cf2 = await runCli(["get", filePath, "/Sheet1/cf[2]", "--json"]);
    const cf3 = await runCli(["get", filePath, "/Sheet1/cf[3]", "--json"]);
    const cfQuery = await runCli(["query", filePath, "cf"]);
    const outline = await runCli(["view", filePath, "outline"]);
    const rawSheet = await runCli(["raw", filePath, "/Sheet1"]);

    expect(cf1.stdout).toContain('"cfType": "databar"');
    expect(cf1.stdout).toContain('FF4472C4');
    expect(cf2.stdout).toContain('"cfType": "colorscale"');
    expect(cf2.stdout).toContain('FFFF0000');
    expect(cf2.stdout).toContain('FF00FF00');
    expect(cf3.stdout).toContain('"cfType": "iconset"');
    expect(cf3.stdout).toContain('"iconset": "5Arrows"');
    expect(cf3.stdout).toContain('"showvalue": true');
    expect(cfQuery.stdout).toContain('"type": "conditionalformatting"');
    expect(outline.stdout).toContain("CF 1: A1:A3 [databar]");
    expect(rawSheet.stdout).toContain('<cfRule type="dataBar"');
    expect(rawSheet.stdout).toContain('<cfRule type="colorScale"');
    expect(rawSheet.stdout).toContain('<cfRule type="iconSet"');
    expect(rawSheet.stdout).toContain('iconSet="5Arrows"');
  });

  test("adds advanced Excel conditional-formatting rule families with dxf-backed details", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-cf-advanced-"));
    const filePath = path.join(dir, "cf-advanced.xlsx");
    await runCli(["create", filePath]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=10", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/A2", "--prop", "value=20", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/A3", "--prop", "value=20", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/B1", "--prop", "value=Alpha"]);
    await runCli(["set", filePath, "/Sheet1/B2", "--prop", "value=Beta"]);
    await runCli(["set", filePath, "/Sheet1/B3", "--prop", "value=Alpha"]);

    await runCli(["add", filePath, "/Sheet1", "--type", "formulacf", "--prop", "sqref=A1:A3", "--prop", "formula=$A1>15", "--prop", "fill=FFF2CC", "--prop", "font.bold=true"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "topn", "--prop", "range=A1:A3", "--prop", "rank=2", "--prop", "percent=true", "--prop", "bottom=true", "--prop", "font.color=FF0000"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "aboveaverage", "--prop", "range=A1:A3", "--prop", "above=false", "--prop", "fill=DDEBF7"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "duplicatevalues", "--prop", "range=A1:A3", "--prop", "fill=F4CCCC"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "uniquevalues", "--prop", "range=A1:A3", "--prop", "font.color=00AA00"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "containstext", "--prop", "range=B1:B3", "--prop", "text=Alpha", "--prop", "fill=EAD1DC"]);
    await runCli(["add", filePath, "/Sheet1", "--type", "dateoccurring", "--prop", "range=A1:A3", "--prop", "period=nextmonth", "--prop", "font.bold=true"]);

    await runCli(["set", filePath, "/Sheet1/cf[1]", "--prop", "formula=$A1>=20"]);
    await runCli(["set", filePath, "/Sheet1/cf[2]", "--prop", "rank=1", "--prop", "percent=false"]);
    await runCli(["set", filePath, "/Sheet1/cf[6]", "--prop", "text=Beta"]);
    await runCli(["set", filePath, "/Sheet1/cf[7]", "--prop", "period=last7days"]);

    const cf1 = await runCli(["get", filePath, "/Sheet1/cf[1]", "--json"]);
    const cf2 = await runCli(["get", filePath, "/Sheet1/cf[2]", "--json"]);
    const cf3 = await runCli(["get", filePath, "/Sheet1/cf[3]", "--json"]);
    const cf4 = await runCli(["get", filePath, "/Sheet1/cf[4]", "--json"]);
    const cf5 = await runCli(["get", filePath, "/Sheet1/cf[5]", "--json"]);
    const cf6 = await runCli(["get", filePath, "/Sheet1/cf[6]", "--json"]);
    const cf7 = await runCli(["get", filePath, "/Sheet1/cf[7]", "--json"]);
    const cfQuery = await runCli(["query", filePath, "cf"]);
    const rawSheet = await runCli(["raw", filePath, "/Sheet1"]);
    const rawStyles = await runCli(["raw", filePath, "/styles"]);

    expect(cf1.stdout).toContain('"cfType": "formula"');
    expect(cf1.stdout).toContain('"formula": "$A1>=20"');
    expect(cf1.stdout).toContain('"dxfId": 0');
    expect(cf2.stdout).toContain('"cfType": "topn"');
    expect(cf2.stdout).toContain('"rank": 1');
    expect(cf2.stdout).not.toContain('"percent": true');
    expect(cf3.stdout).toContain('"cfType": "aboveaverage"');
    expect(cf3.stdout).toContain('"above": false');
    expect(cf4.stdout).toContain('"cfType": "duplicatevalues"');
    expect(cf5.stdout).toContain('"cfType": "uniquevalues"');
    expect(cf6.stdout).toContain('"cfType": "containstext"');
    expect(cf6.stdout).toContain('"text": "Beta"');
    expect(cf7.stdout).toContain('"cfType": "dateoccurring"');
    expect(cf7.stdout).toContain('"period": "last7Days"');
    expect(cfQuery.stdout).toContain('"cfType": "formula"');
    expect(cfQuery.stdout).toContain('"cfType": "duplicatevalues"');
    expect(rawSheet.stdout).toContain('<cfRule type="expression"');
    expect(rawSheet.stdout).toContain('<cfRule type="top10"');
    expect(rawSheet.stdout).toContain('<cfRule type="aboveAverage"');
    expect(rawSheet.stdout).toContain('<cfRule type="duplicateValues"');
    expect(rawSheet.stdout).toContain('<cfRule type="uniqueValues"');
    expect(rawSheet.stdout).toContain('<cfRule type="containsText"');
    expect(rawSheet.stdout).toContain('<cfRule type="timePeriod"');
    expect(rawSheet.stdout).toContain('timePeriod="last7Days"');
    expect(rawStyles.stdout).toContain('<dxfs count="7">');
  });

  test("evaluates simple formulas for display and creates styles from cell props", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-style-formula-"));
    const filePath = path.join(dir, "style-formula.xlsx");
    await runCli(["create", filePath]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=10"]);
    await runCli(["set", filePath, "/Sheet1/A2", "--prop", "value=20"]);
    await runCli([
      "set",
      filePath,
      "/Sheet1/B1",
      "--prop",
      "formula==SUM(A1:A2)",
      "--prop",
      "font.bold=true",
      "--prop",
      "font.color=FF0000",
      "--prop",
      "fill=FFFF00",
      "--prop",
      "alignment.horizontal=center",
      "--prop",
      "numFmt=0.00",
    ]);

    const cell = await runCli(["get", filePath, "/Sheet1/B1", "--json"]);
    const textView = await runCli(["view", filePath, "text"]);
    const annotatedView = await runCli(["view", filePath, "annotated"]);
    const rawStyles = await runCli(["raw", filePath, "/styles"]);
    const sheetXml = readStoredZip(await readFile(filePath)).get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(cell.stdout).toContain('"styleId":');
    expect(cell.stdout).toContain('"evaluatedValue": "30"');
    expect(textView.stdout).toContain("[/Sheet1/row[1]] 10\t30");
    expect(annotatedView.stdout).toContain("B1: [30] <- =SUM(A1:A2)");
    expect(rawStyles.stdout).toContain("<fonts count=");
    expect(rawStyles.stdout).toContain('rgb="FFFF0000"');
    expect(rawStyles.stdout).toContain('patternType="solid"');
    expect(rawStyles.stdout).toContain('formatCode="0.00"');
    expect(sheetXml).toMatch(/<c r="B1" s="\d+">/);
  });

  test("reuses equivalent styles and evaluates cross-sheet/common formulas", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-style-reuse-"));
    const filePath = path.join(dir, "style-reuse.xlsx");
    await runCli(["create", filePath]);
    await runCli(["add", filePath, "/", "--type", "sheet", "--prop", "name=Summary"]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=5"]);
    await runCli(["set", filePath, "/Sheet1/A2", "--prop", "value=7"]);
    await runCli(["set", filePath, "/Summary/A1", "--prop", "value=1"]);
    await runCli(["set", filePath, "/Summary/A2", "--prop", "value=2"]);
    await runCli(["set", filePath, "/Summary/B1", "--prop", "formula==IF(Sheet1!A1>4,1,0)"]);
    await runCli(["set", filePath, "/Summary/B2", "--prop", "formula==COUNTA(Sheet1!A1:A2,Summary!A1)"]);
    await runCli(["set", filePath, "/Summary/B3", "--prop", "formula==SUMPRODUCT(Sheet1!A1:A2,Summary!A1:A2)"]);
    await runCli(["set", filePath, "/Sheet1/B1", "--prop", "value=Styled", "--prop", "font.bold=true", "--prop", "fill=00FF00", "--prop", "alignment.horizontal=center"]);
    await runCli(["set", filePath, "/Sheet1/B2", "--prop", "value=Styled too", "--prop", "font.bold=true", "--prop", "fill=00FF00", "--prop", "alignment.horizontal=center"]);

    const b1 = await runCli(["get", filePath, "/Sheet1/B1", "--json"]);
    const b2 = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    const summaryB1 = await runCli(["get", filePath, "/Summary/B1", "--json"]);
    const summaryB2 = await runCli(["get", filePath, "/Summary/B2", "--json"]);
    const summaryB3 = await runCli(["get", filePath, "/Summary/B3", "--json"]);
    const textView = await runCli(["view", filePath, "text"]);
    const rawStyles = await runCli(["raw", filePath, "/styles"]);

    expect(b1.stdout).toContain('"styleId": "1"');
    expect(b2.stdout).toContain('"styleId": "1"');
    expect(summaryB1.stdout).toContain('"evaluatedValue": "1"');
    expect(summaryB2.stdout).toContain('"evaluatedValue": "3"');
    expect(summaryB3.stdout).toContain('"evaluatedValue": "19"');
    expect(textView.stdout).toContain("[/Summary/row[1]] 1\t1");
    expect(textView.stdout).toContain("[/Summary/row[2]] 2\t3");
    expect(textView.stdout).toContain("[/Summary/row[3]] 19");
    expect(rawStyles.stdout).toContain('<cellXfs count="2">');
  });

  test("supports richer OfficeCLI-style numeric/text formulas and style keys", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-formula-style-depth-"));
    const filePath = path.join(dir, "formula-style-depth.xlsx");
    await runCli(["create", filePath]);
    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=1.234", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/A2", "--prop", "value=-5", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/A3", "--prop", "value=3", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/A4", "--prop", "value=2", "--prop", "type=number"]);
    await runCli(["set", filePath, "/Sheet1/B1", "--prop", "value=  Hello  "]);
    await runCli(["set", filePath, "/Sheet1/B2", "--prop", "value=World"]);

    await runCli(["set", filePath, "/Sheet1/C1", "--prop", "formula==COUNT(A1:A4)"]);
    await runCli(["set", filePath, "/Sheet1/C2", "--prop", "formula==ROUND(A1,2)"]);
    await runCli(["set", filePath, "/Sheet1/C3", "--prop", "formula==ROUNDUP(A1,1)"]);
    await runCli(["set", filePath, "/Sheet1/C4", "--prop", "formula==ROUNDDOWN(A1,1)"]);
    await runCli(["set", filePath, "/Sheet1/C5", "--prop", "formula==ABS(A2)"]);
    await runCli(["set", filePath, "/Sheet1/C6", "--prop", "formula==MOD(A3,A4)"]);
    await runCli(["set", filePath, "/Sheet1/C7", "--prop", "formula==POWER(A4,3)"]);
    await runCli(["set", filePath, "/Sheet1/C8", "--prop", "formula==SQRT(A3)"]);
    await runCli(["set", filePath, "/Sheet1/C9", "--prop", "formula==LEN(B1)"]);
    await runCli(["set", filePath, "/Sheet1/C10", "--prop", "formula==TRIM(B1)"]);
    await runCli(["set", filePath, "/Sheet1/C11", "--prop", "formula==UPPER(B1)"]);
    await runCli(["set", filePath, "/Sheet1/C12", "--prop", "formula==LEFT(B2,2)"]);
    await runCli(["set", filePath, "/Sheet1/C13", "--prop", "formula==RIGHT(B2,2)"]);
    await runCli(["set", filePath, "/Sheet1/C14", "--prop", "formula==MID(B2,2,3)"]);
    await runCli(["set", filePath, "/Sheet1/C15", "--prop", "formula==CONCATENATE(\"Hello\",\"-\",\"OK\")"]);

    await runCli([
      "set",
      filePath,
      "/Sheet1/D1",
      "--prop",
      "value=Styled depth",
      "--prop",
      "font.italic=true",
      "--prop",
      "font.underline=double",
      "--prop",
      "font.strike=true",
      "--prop",
      "border=thin",
      "--prop",
      "border.color=FF0000",
      "--prop",
      "rotation=45",
      "--prop",
      "indent=2",
      "--prop",
      "shrinktofit=true",
      "--prop",
      "locked=false",
      "--prop",
      "formulahidden=true",
      "--prop",
      "wraptext=true",
    ]);

    const c1 = await runCli(["get", filePath, "/Sheet1/C1", "--json"]);
    const c2 = await runCli(["get", filePath, "/Sheet1/C2", "--json"]);
    const c3 = await runCli(["get", filePath, "/Sheet1/C3", "--json"]);
    const c4 = await runCli(["get", filePath, "/Sheet1/C4", "--json"]);
    const c5 = await runCli(["get", filePath, "/Sheet1/C5", "--json"]);
    const c6 = await runCli(["get", filePath, "/Sheet1/C6", "--json"]);
    const c7 = await runCli(["get", filePath, "/Sheet1/C7", "--json"]);
    const c8 = await runCli(["get", filePath, "/Sheet1/C8", "--json"]);
    const c9 = await runCli(["get", filePath, "/Sheet1/C9", "--json"]);
    const c10 = await runCli(["get", filePath, "/Sheet1/C10", "--json"]);
    const c11 = await runCli(["get", filePath, "/Sheet1/C11", "--json"]);
    const c12 = await runCli(["get", filePath, "/Sheet1/C12", "--json"]);
    const c13 = await runCli(["get", filePath, "/Sheet1/C13", "--json"]);
    const c14 = await runCli(["get", filePath, "/Sheet1/C14", "--json"]);
    const c15 = await runCli(["get", filePath, "/Sheet1/C15", "--json"]);
    const textView = await runCli(["view", filePath, "text"]);
    const rawStyles = await runCli(["raw", filePath, "/styles"]);

    expect(c1.stdout).toContain('"evaluatedValue": "4"');
    expect(c2.stdout).toContain('"evaluatedValue": "1.23"');
    expect(c3.stdout).toContain('"evaluatedValue": "1.3"');
    expect(c4.stdout).toContain('"evaluatedValue": "1.2"');
    expect(c5.stdout).toContain('"evaluatedValue": "5"');
    expect(c6.stdout).toContain('"evaluatedValue": "1"');
    expect(c7.stdout).toContain('"evaluatedValue": "8"');
    expect(c8.stdout).toContain('"evaluatedValue": "1.7320508076"');
    expect(c9.stdout).toContain('"evaluatedValue": "9"');
    expect(c10.stdout).toContain('"evaluatedValue": "Hello"');
    expect(c11.stdout).toContain('"evaluatedValue": "  HELLO  "');
    expect(c12.stdout).toContain('"evaluatedValue": "Wo"');
    expect(c13.stdout).toContain('"evaluatedValue": "ld"');
    expect(c14.stdout).toContain('"evaluatedValue": "orl"');
    expect(c15.stdout).toContain('"evaluatedValue": "Hello-OK"');
    expect(textView.stdout).toContain("[/Sheet1/row[15]] Hello-OK");
    expect(rawStyles.stdout).toContain("<i/>");
    expect(rawStyles.stdout).toContain('<u val="double"/>');
    expect(rawStyles.stdout).toContain("<strike/>");
    expect(rawStyles.stdout).toContain('<left style="thin">');
    expect(rawStyles.stdout).toContain('rgb="FFFF0000"');
    expect(rawStyles.stdout).toContain('textRotation="45"');
    expect(rawStyles.stdout).toContain('indent="2"');
    expect(rawStyles.stdout).toContain('shrinkToFit="1"');
    expect(rawStyles.stdout).toContain('locked="0"');
    expect(rawStyles.stdout).toContain('hidden="1"');
  });

  test("supports richer chart properties beyond title and series name", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-chart-props-"));
    const filePath = path.join(dir, "chart-props.xlsx");
    await writeFile(filePath, buildExternalExcelAdvancedObjectsZip());

    await runCli(["set", filePath, "/Sheet1/chart[1]", "--prop", "legend=top", "--prop", "datalabels=value,category", "--prop", "categoryAxisTitle=Months", "--prop", "valueAxisTitle=Revenue"]);
    const chart = await runCli(["get", filePath, "/Sheet1/chart[1]", "--json"]);
    const rawChart = await runCli(["raw", filePath, "/Sheet1/chart[1]"]);

    expect(chart.stdout).toContain('"chartType": "bar"');
    expect(chart.stdout).toContain('"legend": "t"');
    expect(chart.stdout).toContain('"dataLabels": "value"');
    expect(chart.stdout).toContain('"categoryAxisTitle": "Months"');
    expect(chart.stdout).toContain('"valueAxisTitle": "Revenue"');
    expect(rawChart.stdout).toContain('<c:legendPos val="t"');
    expect(rawChart.stdout).toContain('<c:showValue val="1"');
    expect(rawChart.stdout).toContain('<c:showCategoryName val="1"');
    expect(rawChart.stdout).toContain("Months");
    expect(rawChart.stdout).toContain("Revenue");
  });

  test("supports deeper chart styling and axis controls", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-chart-deep-"));
    const filePath = path.join(dir, "chart-deep.xlsx");
    await writeFile(filePath, buildExternalExcelAdvancedObjectsZip());

    await runCli([
      "set",
      filePath,
      "/Sheet1/chart[1]",
      "--prop",
      "axismin=0",
      "--prop",
      "axismax=50",
      "--prop",
      "majorunit=10",
      "--prop",
      "minorunit=5",
      "--prop",
      "axisnumfmt=0.0",
      "--prop",
      "colors=FF0000,00FF00",
      "--prop",
      "plotfill=F1F5F9",
      "--prop",
      "chartfill=E2E8F0",
      "--prop",
      "style=12",
    ]);

    const chart = await runCli(["get", filePath, "/Sheet1/chart[1]", "--json"]);
    const rawChart = await runCli(["raw", filePath, "/Sheet1/chart[1]"]);

    expect(chart.stdout).toContain('"axisMin": 0');
    expect(chart.stdout).toContain('"axisMax": 50');
    expect(chart.stdout).toContain('"majorUnit": 10');
    expect(chart.stdout).toContain('"minorUnit": 5');
    expect(chart.stdout).toContain('"axisNumberFormat": "0.0"');
    expect(chart.stdout).toContain('"styleId": 12');
    expect(chart.stdout).toContain('"plotAreaFill": "F1F5F9"');
    expect(chart.stdout).toContain('"chartAreaFill": "E2E8F0"');
    expect(rawChart.stdout).toContain('<c:minVal val="0"');
    expect(rawChart.stdout).toContain('<c:maxVal val="50"');
    expect(rawChart.stdout).toContain('<c:majorUnit val="10"');
    expect(rawChart.stdout).toContain('<c:minorUnit val="5"');
    expect(rawChart.stdout).toContain('formatCode="0.0"');
    expect(rawChart.stdout).toContain('<c:style val="12"');
    expect(rawChart.stdout).toContain('srgbClr val="FF0000"');
    expect(rawChart.stdout).toContain('srgbClr val="F1F5F9"');
    expect(rawChart.stdout).toContain('srgbClr val="E2E8F0"');
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

  test("imports quoted formulas that contain delimiter commas", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-import-formula-quoted-"));
    const filePath = path.join(dir, "formula-quoted.xlsx");
    const csvPath = path.join(dir, "formula-quoted.csv");
    await writeFile(csvPath, 'Label,Formula\nTotal,"=SUM(1,2)"\n');
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

    const formulaCell = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    const sheetXml = readStoredZip(await readFile(filePath)).get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(formulaCell.stdout).toContain('"formula": "SUM(1,2)"');
    expect(sheetXml).toContain('<f>SUM(1,2)</f>');
  });

  test("imports multiline quoted cells without losing neighboring columns", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-import-multiline-"));
    const filePath = path.join(dir, "multiline.xlsx");
    const csvPath = path.join(dir, "multiline.csv");
    await writeFile(csvPath, 'Name,Notes,Score\nAlpha,"Line 1\nLine 2",42\nBeta,"Single line",43\n');
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

    const noteCell = await runCli(["get", filePath, "/Sheet1/B2", "--json"]);
    const scoreCell = await runCli(["get", filePath, "/Sheet1/C3", "--json"]);
    const sheetXml = readStoredZip(await readFile(filePath)).get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(noteCell.stdout).toContain('Line 1\\nLine 2');
    expect(scoreCell.stdout).toContain('"value": "43"');
    expect(sheetXml).toContain('r="B2"');
    expect(sheetXml).toContain('r="C3"');
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

  test("preserves sheet-level extras when mutating a metadata-free workbook", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "officekit-excel-sheet-extras-"));
    const filePath = path.join(dir, "sheet-extras.xlsx");
    await writeFile(filePath, buildExternalExcelSheetExtrasZip());

    await runCli(["set", filePath, "/Sheet1/A1", "--prop", "value=Updated"]);
    const zip = readStoredZip(await readFile(filePath));
    const sheetXml = zip.get("xl/worksheets/sheet1.xml")!.toString("utf8");

    expect(sheetXml).toContain('<sheetProtection sheet="1"');
    expect(sheetXml).toContain('<pageSetup orientation="landscape"');
    expect(sheetXml).toContain('<oddHeader>&amp;LHeader Left</oddHeader>');
    expect(sheetXml).toContain('<oddFooter>&amp;RFooter Right</oddFooter>');
    expect(sheetXml).toContain('<rowBreaks count="1" manualBreakCount="1">');
    expect(sheetXml).toContain('<colBreaks count="1" manualBreakCount="1">');
    expect(sheetXml).toContain('<t>Updated</t>');
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

function buildExternalExcelSheetExtrasZip() {
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
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Original</t></is></c></row></sheetData>
  <sheetProtection sheet="1"/>
  <pageSetup orientation="landscape" paperSize="9"/>
  <headerFooter><oddHeader>&amp;LHeader Left</oddHeader><oddFooter>&amp;RFooter Right</oddFooter></headerFooter>
  <rowBreaks count="1" manualBreakCount="1"><brk id="5" man="1"/></rowBreaks>
  <colBreaks count="1" manualBreakCount="1"><brk id="2" man="1"/></colBreaks>
</worksheet>`),
    },
  ]);
}

function buildExternalExcelAdvancedObjectsZip() {
  return createStoredZip([
    {
      name: "[Content_Types].xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Value</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Alpha</t></is></c>
      <c r="B2"><v>10</v></c>
    </row>
  </sheetData>
  <dataValidations count="1">
    <dataValidation type="list" sqref="A2" allowBlank="1"><formula1>"Yes,No"</formula1></dataValidation>
  </dataValidations>
  <drawing r:id="rId3"/>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
      <x14:sparklineGroups>
        <x14:sparklineGroup type="line">
          <x14:sparklines>
            <x14:sparkline>
              <xm:f>Sheet1!A2:B2</xm:f>
              <xm:sqref>C2</xm:sqref>
            </x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>`),
    },
    {
      name: "xl/worksheets/_rels/sheet1.xml.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`),
    },
    {
      name: "xl/comments1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Author 1</author></authors>
  <commentList>
    <comment ref="A2" authorId="0"><text><r><t>Initial note</t></r></text></comment>
  </commentList>
</comments>`),
    },
    {
      name: "xl/tables/table1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2" totalsRowShown="0" headerRowCount="1">
  <autoFilter ref="A1:B2"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Value"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`),
    },
    {
      name: "xl/drawings/drawing1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:row>6</xdr:row></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:row>9</xdr:row></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="3" name="Shape 1"/><xdr:cNvSpPr/><xdr:nvPr/></xdr:nvSpPr>
      <xdr:spPr/>
      <xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Initial shape text</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>4</xdr:col><xdr:row>6</xdr:row></xdr:from>
    <xdr:to><xdr:col>6</xdr:col><xdr:row>9</xdr:row></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="4" name="Picture 1"/><xdr:cNvPicPr/><xdr:nvPr/></xdr:nvPicPr>
      <xdr:blipFill/><xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:row>4</xdr:row></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:row>12</xdr:row></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`),
    },
    {
      name: "xl/drawings/_rels/drawing1.xml.rels",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
    },
    {
      name: "xl/charts/chart1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Initial Chart</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Initial Series</c:v></c:pt></c:strCache></c:strRef></c:tx>
        </c:ser>
      </c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx><c:axId val="2"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`),
    },
    {
      name: "xl/pivotTables/pivotTable1.xml",
      data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1"/>`),
    },
  ]);
}

function tinyPngBuffer() {
  return Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wn0Yl8AAAAASUVORK5CYII=",
    "base64",
  );
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
