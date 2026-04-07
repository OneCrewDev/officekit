/**
 * Chart operations for @officekit/ppt.
 *
 * Provides functions to manage charts in PowerPoint presentations:
 * - Get chart data
 * - Add charts to slides
 * - Update chart data
 * - Change chart types
 *
 * Supported chart types: bar, column, line, pie, scatter, area
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, andThen, map, notFound, invalidInput } from "./result.js";
import type { Result } from "./types.js";
import { getSlideIndex, chartPath } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Represents a chart on a slide.
 */
export interface ChartItem {
  /** Path to the chart (e.g., "/slide[1]/chart[1]") */
  path: string;
  /** Chart title */
  title?: string;
  /** Chart type (e.g., "bar", "column", "line", "pie", "scatter", "area") */
  type?: string;
  /** Chart name */
  name?: string;
  /** Position X in EMUs */
  x?: number;
  /** Position Y in EMUs */
  y?: number;
  /** Width in EMUs */
  width?: number;
  /** Height in EMUs */
  height?: number;
  /** Series data */
  series?: ChartSeries[];
  /** Categories */
  categories?: string[];
}

/**
 * Represents a series in a chart.
 */
export interface ChartSeries {
  /** Series name */
  name?: string;
  /** Values as array of numbers */
  values?: number[];
  /** Color */
  color?: string;
}

/**
 * Position for placing a chart on a slide.
 */
export interface ChartPosition {
  /** X position in EMUs (optional, defaults to 1 inch / 914400 EMUs) */
  x?: number;
  /** Y position in EMUs (optional, defaults to 2 inches / 1828800 EMUs) */
  y?: number;
  /** Width in EMUs (optional, defaults to 6 inches / 5486400 EMUs) */
  width?: number;
  /** Height in EMUs (optional, defaults to 4 inches / 3657600 EMUs) */
  height?: number;
}

/**
 * Chart data for creating a new chart.
 */
export interface ChartData {
  /** Chart title (optional) */
  title?: string;
  /** Series data */
  series: ChartSeriesInput[];
  /** Categories (X-axis labels) */
  categories: string[];
}

/**
 * Input for a chart series.
 */
export interface ChartSeriesInput {
  /** Series name */
  name?: string;
  /** Values as array of numbers */
  values: number[];
  /** Color as hex (optional) */
  color?: string;
}

/**
 * Supported chart types.
 */
export type ChartType = "bar" | "column" | "line" | "pie" | "scatter" | "area";

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
 * Throws an error if the entry is not found.
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
 * Gets the relationships entry name for a given entry.
 */
function getRelationshipsEntryName(entryName: string): string {
  const directory = path.posix.dirname(entryName);
  const basename = path.posix.basename(entryName);
  return path.posix.join(directory, "_rels", `${basename}.rels`);
}

/**
 * Generates a unique relationship ID.
 */
function generateRelId(existingRelIds: string[]): string {
  let id = 1;
  let relId = `rId${id}`;
  while (existingRelIds.includes(relId)) {
    id++;
    relId = `rId${id}`;
  }
  return relId;
}

/**
 * Gets the slide size from presentation.xml.
 */
function getSlideSize(zip: Map<string, Buffer>): { width: number; height: number } {
  const presXml = requireEntry(zip, "ppt/presentation.xml");
  const sizeMatch = /<p:sldSz\b[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"[^>]*>/.exec(presXml);
  if (sizeMatch) {
    return {
      width: parseInt(sizeMatch[1], 10),
      height: parseInt(sizeMatch[2], 10),
    };
  }
  // Default to 16:9 aspect ratio (9144000 x 5143500 EMUs = 10 inches x 5.625 inches)
  return { width: 9144000, height: 5143500 };
}

/**
 * Extracts charts from a slide's XML.
 */
function extractChartsFromSlide(
  slideXml: string,
  slideIndex: number,
  slideRels: Array<{ id: string; target: string; type?: string }>,
  slideEntry: string,
  zip: Map<string, Buffer>
): ChartItem[] {
  const charts: ChartItem[] = [];

  // Find all graphicFrame elements that contain charts
  const graphicFramePattern = /<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g;
  const graphicFrameMatches = slideXml.match(graphicFramePattern) || [];

  let chartIdx = 0;
  for (const gfXml of graphicFrameMatches) {
    // Check if this graphicFrame contains a chart
    const chartRefMatch = /<c:chart\b[^>]*r:id="([^"]*)"[^>]*>/.exec(gfXml);
    if (!chartRefMatch) {
      continue;
    }

    chartIdx++;
    const relId = chartRefMatch[1];
    const rel = slideRels.find(r => r.id === relId);
    if (!rel) {
      continue;
    }

    // Resolve chart path
    const slideDir = path.posix.dirname(slideEntry);
    const chartPath = normalizeZipPath(slideDir, rel.target);

    // Extract chart data from the chart XML
    let chartXml: string;
    try {
      chartXml = requireEntry(zip, chartPath);
    } catch {
      continue;
    }

    // Extract name
    const nameMatch = /<p:cNvPr[^>]*name="([^"]*)"[^>]*>/.exec(gfXml);
    const name = nameMatch ? nameMatch[1] : `Chart ${chartIdx}`;

    // Extract title from chart
    const titleMatch = /<c:title[^>]*>[\s\S]*?<a:txPr>[\s\S]*?<a:t>([^<]*)<\/a:t>/.exec(chartXml);
    const title = titleMatch ? titleMatch[1] : undefined;

    // Extract chart type
    const chartType = extractChartType(chartXml);

    // Extract position from graphicFrame transform
    const xfrmMatch = /<p:xfrm>[\s\S]*?<a:off[^>]*x="([^"]*)"[^>]*y="([^"]*)"[^>]*>[\s\S]*?<\/p:xfrm>/.exec(gfXml);
    const extMatch = /<p:xfrm>[\s\S]*?<a:ext[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"[^>]*>[\s\S]*?<\/p:xfrm>/.exec(gfXml);

    // Extract series data
    const series = extractChartSeries(chartXml);
    const categories = extractChartCategories(chartXml);

    charts.push({
      path: `/slide[${slideIndex}]/chart[${chartIdx}]`,
      name,
      title,
      type: chartType,
      x: xfrmMatch ? parseInt(xfrmMatch[1], 10) : undefined,
      y: xfrmMatch ? parseInt(xfrmMatch[2], 10) : undefined,
      width: extMatch ? parseInt(extMatch[1], 10) : undefined,
      height: extMatch ? parseInt(extMatch[2], 10) : undefined,
      series,
      categories,
    });
  }

  return charts;
}

/**
 * Extracts the chart type from chart XML.
 */
function extractChartType(chartXml: string): string | undefined {
  if (/<c:barChart>/.test(chartXml)) {
    const barDirMatch = /<c:barDir[^>]*val="([^"]*)"[^>]*>/.exec(chartXml);
    const barDir = barDirMatch ? barDirMatch[1] : "bar";
    return barDir === "bar" ? "bar" : "column";
  }
  if (/<c:lineChart>/.test(chartXml)) return "line";
  if (/<c:pieChart>/.test(chartXml)) return "pie";
  if (/<c:scatterChart>/.test(chartXml)) return "scatter";
  if (/<c:areaChart>/.test(chartXml)) return "area";
  return undefined;
}

/**
 * Extracts series data from chart XML.
 */
function extractChartSeries(chartXml: string): ChartSeries[] {
  const seriesList: ChartSeries[] = [];

  // Match all series
  const serPattern = /<c:ser>[\s\S]*?<\/c:ser>/g;
  const serMatches = chartXml.match(serPattern) || [];

  for (const serXml of serMatches) {
    // Extract series name
    const nameMatch = /<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/.exec(serXml);
    const name = nameMatch ? nameMatch[1] : undefined;

    // Extract values
    const values: number[] = [];
    const valPattern = /<c:val>[\s\S]*?<c:numLit>[\s\S]*?<\/c:numLit>[\s\S]*?<\/c:val>|<c:val>[\s\S]*?<c:numRef>[\s\S]*?<\/c:numRef>[\s\S]*?<\/c:val>/.exec(serXml);
    if (valPattern) {
      const ptMatches = serXml.match(/<c:pt[^>]*idx="[^"]*"[^>]*>[\s\S]*?<c:v>([^<]*)<\/c:v>/g);
      if (ptMatches) {
        for (const pt of ptMatches) {
          const valMatch = /<c:v>([^<]*)<\/c:v>/.exec(pt);
          if (valMatch) {
            const num = parseFloat(valMatch[1]);
            if (!isNaN(num)) {
              values.push(num);
            }
          }
        }
      }
    }

    // Extract color from spPr if present
    const colorMatch = /<c:spPr>[\s\S]*?<a:srgbClr\s+val="([^"]*)"[^>]*>/.exec(serXml);
    const color = colorMatch ? colorMatch[1] : undefined;

    seriesList.push({ name, values: values.length > 0 ? values : undefined, color });
  }

  return seriesList;
}

/**
 * Extracts categories from chart XML.
 */
function extractChartCategories(chartXml: string): string[] {
  const categories: string[] = [];

  // Try strLit first
  const strLitPattern = /<c:cat>[\s\S]*?<c:strLit>[\s\S]*?<\/c:strLit>[\s\S]*?<\/c:cat>|<c:cat>[\s\S]*?<c:strRef>[\s\S]*?<\/c:strRef>[\s\S]*?<\/c:cat>/.exec(chartXml);
  if (strLitPattern) {
    const ptMatches = chartXml.match(/<c:pt[^>]*idx="[^"]*"[^>]*>[\s\S]*?<c:v>([^<]*)<\/c:v>/g);
    if (ptMatches) {
      for (const pt of ptMatches) {
        const valMatch = /<c:v>([^<]*)<\/c:v>/.exec(pt);
        if (valMatch) {
          categories.push(valMatch[1]);
        }
      }
    }
  }

  return categories;
}

/**
 * Maps chart type string to OOXML chart type element.
 */
function getChartTypeElement(chartType: ChartType): { chartElement: string; openTag: string; closeTag: string } {
  switch (chartType) {
    case "bar":
      return { chartElement: "barChart", openTag: "<c:barChart><c:varyColors val=\"1\"/><c:barDir val=\"bar\"/>", closeTag: "</c:barChart>" };
    case "column":
      return { chartElement: "barChart", openTag: "<c:barChart><c:varyColors val=\"1\"/><c:barDir val=\"col\"/>", closeTag: "</c:barChart>" };
    case "line":
      return { chartElement: "lineChart", openTag: "<c:lineChart><c:varyColors val=\"1\"/>", closeTag: "</c:lineChart>" };
    case "pie":
      return { chartElement: "pieChart", openTag: "<c:pieChart><c:varyColors val=\"1\"/>", closeTag: "</c:pieChart>" };
    case "scatter":
      return { chartElement: "scatterChart", openTag: "<c:scatterChart><c:varyColors val=\"1\"/>", closeTag: "</c:scatterChart>" };
    case "area":
      return { chartElement: "areaChart", openTag: "<c:areaChart><c:varyColors val=\"1\"/>", closeTag: "</c:areaChart>" };
    default:
      return { chartElement: "barChart", openTag: "<c:barChart><c:varyColors val=\"1\"/><c:barDir val=\"bar\"/>", closeTag: "</c:barChart>" };
  }
}

// ============================================================================
// Chart Operations
// ============================================================================

/**
 * Gets chart data from a presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param chartPathStr - Path to the chart (e.g., "/slide[1]/chart[1]")
 *
 * @example
 * const result = await getChart("/path/to/presentation.pptx", "/slide[1]/chart[1]");
 * if (result.ok) {
 *   console.log(result.data.chart);
 * }
 */
export async function getChart(
  filePath: string,
  chartPathStr: string
): Promise<Result<{ chart: ChartItem }>> {
  try {
    const slideIndex = getSlideIndex(chartPathStr);
    if (slideIndex === null) {
      return invalidInput("Invalid chart path - must include slide index");
    }

    // Extract chart index from path
    const chartIndexMatch = chartPathStr.match(/\/chart\[(\d+)\]/i);
    if (!chartIndexMatch) {
      return invalidInput("Invalid chart path - must include chart[index]");
    }
    const chartIndex = parseInt(chartIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    const relsXml = requireEntry(zip, relsEntry);
    const relationships = parseRelationshipEntries(relsXml);

    // Find all charts to determine if the requested one exists
    const charts = extractChartsFromSlide(slideXml, slideIndex, relationships, slideEntry, zip);

    if (chartIndex < 1 || chartIndex > charts.length) {
      return notFound("Chart", String(chartIndex));
    }

    return ok({ chart: charts[chartIndex - 1] });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Adds a chart to a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param chartType - Type of chart (bar, column, line, pie, scatter, area)
 * @param position - Optional position and size
 * @param data - Chart data (series and categories)
 *
 * @example
 * const result = await addChart(
 *   "/path/to/presentation.pptx",
 *   1,
 *   "bar",
 *   { x: 1000000, y: 1000000, width: 6000000, height: 4000000 },
 *   {
 *     title: "Sales Data",
 *     series: [{ name: "Q1", values: [100, 200, 150] }],
 *     categories: ["Jan", "Feb", "Mar"]
 *   }
 * );
 */
export async function addChart(
  filePath: string,
  slideIndex: number,
  chartType: ChartType,
  position: ChartPosition,
  data: ChartData
): Promise<Result<{ path: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    let relsXml = "";
    try {
      relsXml = requireEntry(zip, relsEntry);
    } catch {
      // Create empty rels if it doesn't exist
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    }
    const relationships = parseRelationshipEntries(relsXml);

    // Get slide size for default positioning
    const slideSize = getSlideSize(zip);

    // Calculate position and size
    const chartWidth = position?.width ?? 5486400; // Default 6 inches
    const chartHeight = position?.height ?? 3657600; // Default 4 inches
    const chartX = position?.x ?? 914400; // Default 1 inch
    const chartY = position?.y ?? 1828800; // Default 2 inches

    // Count existing charts for naming
    const existingCharts = extractChartsFromSlide(slideXml, slideIndex, relationships, slideEntry, zip);
    const chartCount = existingCharts.length;
    const chartName = data.title || `Chart ${chartCount + 1}`;

    // Generate unique IDs
    const existingRelIds = relationships.map(r => r.id);
    const newRelId = generateRelId(existingRelIds);

    // Find existing shape IDs to generate unique shape ID
    const existingShapeIds: number[] = [];
    for (const match of slideXml.matchAll(/<p:cNvPr[^>]*id="(\d+)"[^>]*>/g)) {
      existingShapeIds.push(parseInt(match[1], 10));
    }
    let maxId = Math.max(...existingShapeIds, 0);
    const newShapeId = maxId + 1;

    // Generate chart file name
    const chartFileName = `chart${chartCount + 1}.xml`;
    const chartEntry = `ppt/slides/charts/${chartFileName}`;

    // Build chart XML
    const chartXml = buildChartXml(chartType, chartName, data);

    // Create chart part relationship
    const newRelEntry = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../slides/charts/${chartFileName}"/>`;

    // Build the graphicFrame XML
    const graphicFrameXml = `<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="${newShapeId}" name="${chartName}"/>
    <p:cNvGraphicFramePr>
      <a:graphicFrameLocks noGrp="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
    </p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="${chartX}" y="${chartY}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
    <a:ext cx="${chartWidth}" cy="${chartHeight}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
  </p:xfrm>
  <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${newRelId}"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`;

    // Insert graphicFrame into slide XML before closing </p:spTree>
    const updatedSlideXml = slideXml.replace(
      "</p:spTree>",
      `${graphicFrameXml}</p:spTree>`
    );

    // Update relationships XML
    const updatedRelsXml = relsXml.replace(
      "</Relationships>",
      `${newRelEntry}</Relationships>`
    );

    // Build new zip with updated slide and chart
    const newEntries: Array<{ name: string; data: Buffer }> = [];

    for (const [name, dataEntry] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else if (name === relsEntry) {
        newEntries.push({ name, data: Buffer.from(updatedRelsXml, "utf8") });
      } else {
        newEntries.push({ name, data: dataEntry });
      }
    }

    // Add the chart XML
    newEntries.push({ name: chartEntry, data: Buffer.from(chartXml, "utf8") });

    await writeFile(filePath, createStoredZip(newEntries));

    return ok({ path: `/slide[${slideIndex}]/chart[${chartCount + 1}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Builds the XML for a chart.
 */
function buildChartXml(chartType: ChartType, title: string, data: ChartData): string {
  const { chartElement, openTag, closeTag } = getChartTypeElement(chartType);

  // Build series XML
  let seriesXml = "";
  for (let i = 0; i < data.series.length; i++) {
    const ser = data.series[i];
    const idx = i;
    const order = i;
    const serName = ser.name || `Series ${i + 1}`;

    // Build values XML
    let valuesXml = "";
    if (ser.values && ser.values.length > 0) {
      const ptList = ser.values.map((v, vi) => `<c:pt idx="${vi}"><c:v>${v}</c:v></c:pt>`).join("");
      valuesXml = `<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${ser.values.length}"/>${ptList}</c:numLit></c:val>`;
    }

    // Build categories XML for first series (shared categories)
    let catXml = "";
    if (i === 0 && data.categories && data.categories.length > 0) {
      const ptList = data.categories.map((c, ci) => `<c:pt idx="${ci}"><c:v>${escapeXml(c)}</c:v></c:pt>`).join("");
      catXml = `<c:cat><c:strLit><c:ptCount val="${data.categories.length}"/>${ptList}</c:strLit></c:cat>`;
    }

    seriesXml += `<c:ser><c:idx val="${idx}"/><c:order val="${order}"/><c:tx><c:v>${escapeXml(serName)}</c:v></c:tx>${catXml}${valuesXml}</c:ser>`;
  }

  // Build chart XML
  const titleXml = title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1" dirty="0"/><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : "";

  const chartContent = `${openTag}
${seriesXml}
${closeTag}`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart>
${titleXml}
<c:plotArea>
<c:layout/>
${chartContent}
</c:plotArea>
<c:legend>
<c:legendPos val="b"/>
<c:overlay val="0"/>
</c:legend>
<c:plotVisOnly val="1"/>
<c:dispBlanksAs val="gap"/>
</c:chart>
</c:chartSpace>`;
}

/**
 * Escapes special XML characters.
 */
function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Sets the data for a chart.
 *
 * @param filePath - Path to the PPTX file
 * @param chartPathStr - Path to the chart (e.g., "/slide[1]/chart[1]")
 * @param series - New series data
 * @param categories - New categories (optional)
 *
 * @example
 * const result = await setChartData(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/chart[1]",
 *   [{ name: "New Series", values: [100, 200, 150] }],
 *   ["Jan", "Feb", "Mar"]
 * );
 */
export async function setChartData(
  filePath: string,
  chartPathStr: string,
  series: ChartSeriesInput[],
  categories?: string[]
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(chartPathStr);
    if (slideIndex === null) {
      return invalidInput("Invalid chart path - must include slide index");
    }

    // Extract chart index from path
    const chartIndexMatch = chartPathStr.match(/\/chart\[(\d+)\]/i);
    if (!chartIndexMatch) {
      return invalidInput("Invalid chart path - must include chart[index]");
    }
    const chartIndex = parseInt(chartIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    const relsXml = requireEntry(zip, relsEntry);
    const relationships = parseRelationshipEntries(relsXml);

    // Find the chart
    const charts = extractChartsFromSlide(slideXml, slideIndex, relationships, slideEntry, zip);

    if (chartIndex < 1 || chartIndex > charts.length) {
      return notFound("Chart", String(chartIndex));
    }

    // Get the chart path
    const chartRelIdMatch = slideXml.match(new RegExp(`<p:graphicFrame[\\s\\S]*?<c:chart[^>]*r:id="([^"]*)"[^>]*>[\\s\\S]*?</p:graphicFrame>`));
    if (!chartRelIdMatch) {
      return invalidInput("Chart relationship not found in slide");
    }

    const relId = chartRelIdMatch[1];
    const rel = relationships.find(r => r.id === relId);
    if (!rel) {
      return invalidInput("Chart relationship not found");
    }

    const slideDir = path.posix.dirname(slideEntry);
    const chartPath = normalizeZipPath(slideDir, rel.target);

    // Read current chart XML
    const chartXml = requireEntry(zip, chartPath);

    // Extract chart type from current chart
    const chartType = extractChartType(chartXml) || "bar";

    // Build new chart XML with updated data
    const existingTitleMatch = /<c:title[^>]*>[\s\S]*?<a:t>([^<]*)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml);
    const title = existingTitleMatch ? existingTitleMatch[1] : undefined;

    const newChartXml = buildChartXml(chartType as ChartType, title || "", {
      series,
      categories: categories || [],
    });

    // Build new zip with updated chart
    const newEntries: Array<{ name: string; data: Buffer }> = [];

    for (const [name, dataEntry] of zip.entries()) {
      if (name === chartPath) {
        newEntries.push({ name, data: Buffer.from(newChartXml, "utf8") });
      } else {
        newEntries.push({ name, data: dataEntry });
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
 * Changes the type of a chart.
 *
 * @param filePath - Path to the PPTX file
 * @param chartPathStr - Path to the chart (e.g., "/slide[1]/chart[1]")
 * @param chartType - New chart type (bar, column, line, pie, scatter, area)
 *
 * @example
 * const result = await setChartType(
 *   "/path/to/presentation.pptx",
 *   "/slide[1]/chart[1]",
 *   "line"
 * );
 */
export async function setChartType(
  filePath: string,
  chartPathStr: string,
  chartType: ChartType
): Promise<Result<void>> {
  try {
    const slideIndex = getSlideIndex(chartPathStr);
    if (slideIndex === null) {
      return invalidInput("Invalid chart path - must include slide index");
    }

    // Extract chart index from path
    const chartIndexMatch = chartPathStr.match(/\/chart\[(\d+)\]/i);
    if (!chartIndexMatch) {
      return invalidInput("Invalid chart path - must include chart[index]");
    }
    const chartIndex = parseInt(chartIndexMatch[1], 10);

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);
    const relsEntry = getRelationshipsEntryName(slideEntry);
    const relsXml = requireEntry(zip, relsEntry);
    const relationships = parseRelationshipEntries(relsXml);

    // Find the chart
    const charts = extractChartsFromSlide(slideXml, slideIndex, relationships, slideEntry, zip);

    if (chartIndex < 1 || chartIndex > charts.length) {
      return notFound("Chart", String(chartIndex));
    }

    // Get the chart path
    const chartRelIdMatch = slideXml.match(new RegExp(`<p:graphicFrame[\\s\\S]*?<c:chart[^>]*r:id="([^"]*)"[^>]*>[\\s\\S]*?</p:graphicFrame>`));
    if (!chartRelIdMatch) {
      return invalidInput("Chart relationship not found in slide");
    }

    const relId = chartRelIdMatch[1];
    const rel = relationships.find(r => r.id === relId);
    if (!rel) {
      return invalidInput("Chart relationship not found");
    }

    const slideDir = path.posix.dirname(slideEntry);
    const chartPath = normalizeZipPath(slideDir, rel.target);

    // Read current chart XML
    const chartXml = requireEntry(zip, chartPath);

    // Extract current data from chart
    const currentSeries = extractChartSeries(chartXml);
    const currentCategories = extractChartCategories(chartXml);
    const titleMatch = /<c:title[^>]*>[\s\S]*?<a:t>([^<]*)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml);
    const title = titleMatch ? titleMatch[1] : undefined;

    // Build new chart XML with new type but existing data
    const newChartXml = buildChartXml(chartType, title || "", {
      series: currentSeries.map(s => ({
        name: s.name,
        values: s.values || [],
        color: s.color,
      })),
      categories: currentCategories,
    });

    // Build new zip with updated chart
    const newEntries: Array<{ name: string; data: Buffer }> = [];

    for (const [name, dataEntry] of zip.entries()) {
      if (name === chartPath) {
        newEntries.push({ name, data: Buffer.from(newChartXml, "utf8") });
      } else {
        newEntries.push({ name, data: dataEntry });
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
