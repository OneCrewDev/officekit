/**
 * CSS generation for Word HTML preview.
 */

import type { HtmlRenderContext } from "./context.js";

/**
 * Generate complete Word CSS including page layout, typography, and helpers.
 */
export function generateWordCss(pageLayout?: { widthPt: number; heightPt: number; marginTopPt: number; marginBottomPt: number; marginLeftPt: number; marginRightPt: number }): string {
  const pageWidth = pageLayout?.widthPt ?? 612; // 8.5in in points
  const pageHeight = pageLayout?.heightPt ?? 792; // 11in in points
  const marginTop = pageLayout?.marginTopPt ?? 72; // 1in
  const marginBottom = pageLayout?.marginBottomPt ?? 72; // 1in
  const marginLeft = pageLayout?.marginLeftPt ?? 72; // 1in
  const marginRight = pageLayout?.marginRightPt ?? 72; // 1in
  const contentWidth = pageWidth - marginLeft - marginRight;

  return `
/* Word HTML Preview - Generated CSS */
* {
  box-sizing: border-box;
}

body {
  font-family: 'Times New Roman', '宋体', 'SimSun', '宋体', 'Songti SC', 'STSong', 'Microsoft YaHei', -apple-system, 'PingFang SC', Calibri, Arial, sans-serif;
  font-size: 11pt;
  line-height: 1.15;
  color: #000000;
  background: #f0f0f0;
  margin: 0;
  padding: 0;
}

.page {
  width: ${pageWidth}pt;
  min-height: ${pageHeight}pt;
  margin: 0.75in auto;
  padding: ${marginTop}pt ${marginRight}pt ${marginBottom}pt ${marginLeft}pt;
  background: #ffffff;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
  position: relative;
}

/* Typography */
p {
  margin: 0;
  padding: 0;
  line-height: 1.15;
}

strong, b {
  font-weight: bold;
}

em, i {
  font-style: italic;
}

u {
  text-decoration: underline;
}

s, strike, del {
  text-decoration: line-through;
}

/* Headings */
.Heading1, h1 {
  font-size: 28pt;
  font-weight: bold;
  margin: 24pt 0 12pt 0;
  color: #2e5496;
  line-height: 1.2;
}

.Heading2, h2 {
  font-size: 24pt;
  font-weight: bold;
  margin: 20pt 0 10pt 0;
  color: #2e5496;
  line-height: 1.2;
}

.Heading3, h3 {
  font-size: 18pt;
  font-weight: bold;
  margin: 16pt 0 8pt 0;
  color: #1f4e79;
  line-height: 1.3;
}

.Heading4, h4 {
  font-size: 14pt;
  font-weight: bold;
  margin: 12pt 0 6pt 0;
  color: #2e5496;
  line-height: 1.3;
}

.Title {
  font-size: 36pt;
  font-weight: bold;
  margin: 24pt 0 12pt 0;
  text-align: center;
  color: #2e5496;
  line-height: 1.2;
}

.Subtitle {
  font-size: 18pt;
  color: #666666;
  margin: 12pt 0;
  text-align: center;
  font-style: italic;
}

/* Text formatting */
sup {
  vertical-align: super;
  font-size: 0.7em;
  line-height: 0;
}

sub {
  vertical-align: sub;
  font-size: 0.7em;
  line-height: 0;
}

/* Lists */
ol, ul {
  margin: 6pt 0;
  padding-left: 24pt;
}

li {
  margin: 3pt 0;
}

/* Tables */
table {
  border-collapse: collapse;
  width: 100%;
  table-layout: fixed;
  margin: 8pt 0;
}

td, th {
  border: 1px solid #b4b4b4;
  padding: 4px 8px;
  vertical-align: top;
  text-align: left;
  min-height: 18pt;
}

th {
  background: #f2f2f2;
  font-weight: bold;
}

tr.header-row th,
tr.header-row td {
  background: #e8e8e8;
  font-weight: bold;
}

tr.odd-row td {
  background: #fafafa;
}

tr.even-row td {
  background: #ffffff;
}

.borderless {
  border: none;
}

.borderless td, .borderless th {
  border: none;
}

.borderless-table td,
.borderless-table th {
  border: none;
}

/* Table column widths via colgroup */
colgroup {
  display: table-column-group;
}

col {
  display: table-column;
}

/* Footnotes and Endnotes */
.footnote-ref {
  font-size: 0.7em;
  vertical-align: super;
  color: #0066cc;
  text-decoration: none;
}

.endnote-ref {
  font-size: 0.7em;
  vertical-align: super;
  color: #666666;
  text-decoration: none;
}

.footnote {
  font-size: 10pt;
  border-top: 1px solid #cccccc;
  padding-top: 8pt;
  margin-top: 12pt;
  color: #333333;
}

.endnote {
  font-size: 10pt;
  border-top: 1px solid #cccccc;
  padding-top: 8pt;
  margin-top: 12pt;
  color: #333333;
}

/* Header/Footer */
.header {
  font-size: 10pt;
  color: #666666;
  text-align: center;
  padding-bottom: 8pt;
  border-bottom: 1px solid #cccccc;
  margin-bottom: 8pt;
}

.footer {
  font-size: 10pt;
  color: #666666;
  text-align: center;
  padding-top: 8pt;
  border-top: 1px solid #cccccc;
  margin-top: 8pt;
}

/* Images */
img {
  max-width: 100%;
  height: auto;
  vertical-align: middle;
}

.drawing-inline {
  display: inline-block;
  vertical-align: middle;
}

.drawing-anchor {
  display: block;
  margin: 8pt 0;
}

.drawing-float-left {
  float: left;
  margin-right: 12pt;
  margin-bottom: 8pt;
}

.drawing-float-right {
  float: right;
  margin-left: 12pt;
  margin-bottom: 8pt;
}

/* Page breaks */
.page-break {
  page-break-before: always;
  break-before: page;
}

.keep-next {
  page-break-after: avoid;
  break-after: avoid;
}

.keep-lines {
  page-break-inside: avoid;
  break-inside: avoid;
}

.column-break {
  page-break-after: always;
  break-after: column;
}

/* Section columns */
.columns {
  column-count: 1;
  column-gap: 0.5in;
}

.columns-2 {
  column-count: 2;
}

.columns-3 {
  column-count: 3;
}

/* Highlight colors */
.hl-yellow { background-color: #ffff00; }
.hl-green { background-color: #00ff00; }
.hl-cyan { background-color: #00ffff; }
.hl-magenta { background-color: #ff00ff; }
.hl-blue { background-color: #0000ff; }
.hl-red { background-color: #ff0000; }
.hl-darkblue { background-color: #000080; }
.hl-darkgreen { background-color: #008000; }
.hl-darkcyan { background-color: #008080; }
.hl-darkred { background-color: #800000; }
.hl-darkmagenta { background-color: #800080; }
.hl-darkyellow { background-color: #808000; }
.hl-darkgray { background-color: #404040; }
.hl-lightgray { background-color: #c0c0c0; }
.hl-nocolor { background-color: transparent; }

/* Vertical alignment */
.vert-superscript { vertical-align: super; }
.vert-subscript { vertical-align: sub; }
.vert-raised { vertical-align: super; }
.vert-lowered { vertical-align: sub; }

/* Small caps */
.small-caps { font-variant: small-caps; }

/* Hidden text */
.hidden, .vanish, .web-hidden {
  display: none;
}

/* Border helpers */
.border-top { border-top: 1px solid #000; }
.border-bottom { border-bottom: 1px solid #000; }
.border-left { border-left: 1px solid #000; }
.border-right { border-right: 1px solid #000; }

/* List number styles */
.list-number-lower-alpha { list-style-type: lower-alpha; }
.list-number-upper-alpha { list-style-type: upper-alpha; }
.list-number-lower-roman { list-style-type: lower-roman; }
.list-number-upper-roman { list-style-type: upper-roman; }
.list-bullet { list-style-type: disc; }

/* Tab stops */
.tab { display: inline-block; }

/* Revision marks */
.ins {
  color: #006400;
  background-color: #e6ffe6;
}

.del {
  color: #8b0000;
  background-color: #ffe6e6;
  text-decoration: line-through;
}

/* Comments */
.comment {
  background-color: #fff3cd;
  border-left: 3px solid #ffc107;
  padding-left: 8pt;
  margin: 4pt 0;
}

/* Form fields */
.form-field {
  border: 1px dashed #666;
  padding: 2pt 4pt;
  background-color: #fafafa;
}

/* Bookmark */
.bookmark {
  border-bottom: 1px dashed #999;
}

/* Hyperlink */
a {
  color: #0066cc;
  text-decoration: underline;
}

a.no-underline {
  text-decoration: none;
}

/* Text box / Shape */
.text-box {
  border: 1px solid #666;
  padding: 8pt;
  margin: 8pt 0;
}

.shape {
  padding: 8pt;
}
`;
}

/**
 * Generate inline CSS for a single run element.
 */
export function generateRunInlineCss(run: {
  font?: string;
  size?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strike?: string;
  color?: string;
  highlight?: string;
  verticalAlign?: string;
  smallCaps?: boolean;
  shading?: string;
}): string {
  const cssParts: string[] = [];

  if (run.font) {
    cssParts.push(`font-family: '${run.font}', '宋体', 'SimSun', 'Songti SC', 'STSong', 'Microsoft YaHei', Calibri, sans-serif`);
  }

  if (run.size) {
    cssParts.push(`font-size: ${run.size}`);
  } else {
    cssParts.push(`font-size: 11pt`);
  }

  if (run.bold) {
    cssParts.push(`font-weight: bold`);
  }

  if (run.italic) {
    cssParts.push(`font-style: italic`);
  }

  if (run.underline && run.underline !== "none") {
    cssParts.push(`text-decoration: underline`);
    if (run.underline === "double") {
      cssParts.push(`text-decoration: underline double`);
    } else if (run.underline === "thick") {
      cssParts.push(`text-decoration: underline thick`);
    } else if (run.underline === "dotted") {
      cssParts.push(`text-decoration: underline dotted`);
    } else if (run.underline === "dashed") {
      cssParts.push(`text-decoration: underline dashed`);
    }
    // single and other single-line styles default to underline
  }

  if (run.strike && run.strike !== "noStrike") {
    if (run.strike === "single" || run.strike === "s") {
      cssParts.push(`text-decoration: line-through`);
    } else if (run.strike === "double") {
      cssParts.push(`text-decoration: line-through double`);
    } else {
      cssParts.push(`text-decoration: line-through`);
    }
  }

  if (run.color) {
    cssParts.push(`color: ${run.color.startsWith('#') ? run.color : '#' + run.color}`);
  }

  if (run.highlight) {
    const hlColor = mapHighlightColor(run.highlight);
    cssParts.push(`background-color: ${hlColor}`);
  }

  if (run.verticalAlign) {
    if (run.verticalAlign === "superscript" || run.verticalAlign === "super") {
      cssParts.push(`vertical-align: super`);
      cssParts.push(`font-size: 0.7em`);
    } else if (run.verticalAlign === "subscript" || run.verticalAlign === "sub") {
      cssParts.push(`vertical-align: sub`);
      cssParts.push(`font-size: 0.7em`);
    }
  }

  if (run.smallCaps) {
    cssParts.push(`font-variant: small-caps`);
  }

  if (run.shading) {
    cssParts.push(`background-color: #${run.shading}`);
  }

  return cssParts.join("; ");
}

/**
 * Map Word highlight color names to CSS colors.
 */
function mapHighlightColor(highlight: string): string {
  const colorMap: Record<string, string> = {
    yellow: "#ffff00",
    green: "#00ff00",
    cyan: "#00ffff",
    magenta: "#ff00ff",
    blue: "#0000ff",
    red: "#ff0000",
    darkblue: "#000080",
    darkgreen: "#008000",
    darkcyan: "#008080",
    darkred: "#800000",
    darkmagenta: "#800080",
    darkyellow: "#808000",
    darkgray: "#404040",
    lightgray: "#c0c0c0",
    black: "#000000",
    nocolor: "transparent",
    none: "transparent",
  };

  const key = highlight.toLowerCase().replace(/\s+/g, "");
  return colorMap[key] || "#ffff00";
}

/**
 * Generate inline CSS for paragraph properties.
 */
export function generateParagraphInlineCss(para: {
  alignment?: string;
  spaceBefore?: string;
  spaceAfter?: string;
  lineSpacing?: string;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndent?: number;
  style?: string;
}): string {
  const cssParts: string[] = [];

  if (para.alignment) {
    const alignMap: Record<string, string> = {
      left: "left",
      center: "center",
      right: "right",
      justify: "justify",
      both: "justify",
      start: "left",
      end: "right",
      distribute: "justify",
      "inter-character": "justify",
    };
    const cssAlign = alignMap[para.alignment.toLowerCase()];
    if (cssAlign) {
      cssParts.push(`text-align: ${cssAlign}`);
    }
  }

  if (para.spaceBefore) {
    const val = parseTwipsOrPt(para.spaceBefore);
    cssParts.push(`margin-top: ${val}pt`);
  }

  if (para.spaceAfter) {
    const val = parseTwipsOrPt(para.spaceAfter);
    cssParts.push(`margin-bottom: ${val}pt`);
  }

  if (para.lineSpacing) {
    const val = parseLineSpacing(para.lineSpacing);
    cssParts.push(`line-height: ${val}`);
  }

  if (para.leftIndent) {
    cssParts.push(`margin-left: ${para.leftIndent}pt`);
  }

  if (para.rightIndent) {
    cssParts.push(`margin-right: ${para.rightIndent}pt`);
  }

  if (para.firstLineIndent && para.firstLineIndent !== 0) {
    cssParts.push(`text-indent: ${para.firstLineIndent}pt`);
  }

  return cssParts.join("; ");
}

/**
 * Parse a value that could be in twips (ending with 'tw') or points.
 */
function parseTwipsOrPt(value: string): number {
  const trimmed = value.trim();
  if (trimmed.endsWith("tw")) {
    return parseInt(trimmed.slice(0, -2), 10) / 20;
  }
  if (trimmed.endsWith("pt")) {
    return parseFloat(trimmed.slice(0, -2));
  }
  if (trimmed.endsWith("px")) {
    return parseFloat(trimmed.slice(0, -2)) * 0.75;
  }
  // Assume twips as default for Word values
  return parseInt(trimmed, 10) / 20;
}

/**
 * Parse line spacing value to CSS line-height.
 */
function parseLineSpacing(value: string): string {
  const trimmed = value.trim();

  // Multiplier like "1.5x" or "1.5"
  if (trimmed.endsWith("x")) {
    const mult = parseFloat(trimmed.slice(0, -1));
    return String(mult);
  }

  // Try parsing as floating point
  const numVal = parseFloat(trimmed);
  if (!isNaN(numVal)) {
    // Word uses 240 = single line, 480 = double, etc.
    // Or it could be in twips
    if (numVal > 100) {
      // Likely twips or 240-based units
      return String(numVal / 240);
    }
    return String(numVal);
  }

  return "1.15"; // Default
}

/**
 * Generate table cell inline CSS.
 */
export function generateTableCellCss(cell: {
  gridSpan?: number;
  rowSpan?: number;
  fill?: string;
  valign?: string;
  width?: number;
  borders?: TableCellBorders;
}): string {
  const cssParts: string[] = [];

  if (cell.fill) {
    cssParts.push(`background-color: #${cell.fill}`);
  }

  if (cell.valign) {
    const vAlignMap: Record<string, string> = {
      top: "top",
      center: "middle",
      bottom: "bottom",
      both: "middle",
    };
    const cssVAlign = vAlignMap[cell.valign.toLowerCase()];
    if (cssVAlign) {
      cssParts.push(`vertical-align: ${cssVAlign}`);
    }
  }

  if (cell.width) {
    cssParts.push(`width: ${cell.width}pt`);
  }

  if (cell.borders) {
    const borderParts: string[] = [];
    if (cell.borders.top) {
      borderParts.push(`border-top: ${formatBorder(cell.borders.top)}`);
    }
    if (cell.borders.bottom) {
      borderParts.push(`border-bottom: ${formatBorder(cell.borders.bottom)}`);
    }
    if (cell.borders.left) {
      borderParts.push(`border-left: ${formatBorder(cell.borders.left)}`);
    }
    if (cell.borders.right) {
      borderParts.push(`border-right: ${formatBorder(cell.borders.right)}`);
    }
    if (borderParts.length > 0) {
      cssParts.push(borderParts.join("; "));
    }
  }

  return cssParts.join("; ");
}

interface TableCellBorders {
  top?: BorderInfo;
  bottom?: BorderInfo;
  left?: BorderInfo;
  right?: BorderInfo;
}

interface BorderInfo {
  style?: string;
  size?: number;
  color?: string;
}

function formatBorder(border: BorderInfo): string {
  const style = border.style || "single";
  const size = border.size || 4;
  const color = border.color ? `#${border.color}` : "#000";
  const widthPx = Math.max(1, size / 8);

  const styleMap: Record<string, string> = {
    none: "none",
    single: "solid",
    double: "double",
    dotted: "dotted",
    dashed: "dashed",
    wavy: " wavy",
    dash: "dashed",
    dashDot: "dashed",
    dashDotDot: "dashed",
    triple: "double",
    thick: "solid",
    thin: "solid",
  };

  return `${styleMap[style.toLowerCase()] || "solid"} ${widthPx}px ${color}`;
}
