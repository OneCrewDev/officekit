/**
 * Text rendering for Word HTML preview.
 * Handles paragraphs, runs, and hyperlinks.
 */

import { generateRunInlineCss, generateParagraphInlineCss } from "./css.js";
import type { HtmlRenderContext } from "./context.js";

/**
 * Render a paragraph to HTML.
 */
export function renderParagraphHtml(
  paraXml: string,
  paraIndex: number,
  context: HtmlRenderContext,
  options: {
    includeRuns?: boolean;
    defaultStyle?: string;
  } = {}
): string {
  const { includeRuns = true, defaultStyle = "Normal" } = options;

  // Extract paragraph style
  const styleMatch = /<w:pStyle[^>]*w:val="([^"]*)"/i.exec(paraXml);
  const styleId = styleMatch ? styleMatch[1] : defaultStyle;
  const styleClass = styleId.replace(/\s+/g, "");

  // Extract paragraph properties
  const pPrMatch = /<w:pPr>([\s\S]*?)<\/w:pPr>/i.exec(paraXml);
  const paraProps = parseParagraphProperties(pPrMatch?.[1] || "");

  // Build paragraph CSS
  const paraCss = generateParagraphInlineCss({
    alignment: paraProps.alignment,
    spaceBefore: paraProps.spaceBefore,
    spaceAfter: paraProps.spaceAfter,
    lineSpacing: paraProps.lineSpacing,
    leftIndent: paraProps.leftIndent,
    rightIndent: paraProps.rightIndent,
    firstLineIndent: paraProps.firstLineIndent,
  });

  // Build class attribute
  const classes = [];
  if (styleClass && styleClass !== "Normal") {
    classes.push(styleClass);
  }
  if (paraProps.pageBreakBefore) {
    classes.push("page-break");
  }
  if (paraProps.keepNext) {
    classes.push("keep-next");
  }
  if (paraProps.keepLines) {
    classes.push("keep-lines");
  }

  const classAttr = classes.length > 0 ? ` class="${classes.join(" ")}"` : "";
  const styleAttr = paraCss ? ` style="${paraCss}"` : "";

  // Start paragraph tag
  let html = `<p${classAttr}${styleAttr}>`;

  if (includeRuns) {
    // Render runs within paragraph
    html += renderRunsHtml(paraXml, context);
  } else {
    // Just extract text
    html += escapeHtml(extractTextFromParagraph(paraXml));
  }

  html += "</p>";

  return html;
}

/**
 * Render all runs in a paragraph to HTML.
 */
export function renderRunsHtml(paraXml: string, context: HtmlRenderContext): string {
  let html = "";
  const runs = extractRuns(paraXml);

  for (const run of runs) {
    html += renderRunHtml(run, context);
  }

  return html;
}

/**
 * Render a single run to HTML.
 */
export function renderRunHtml(run: RunData, context: HtmlRenderContext): string {
  if (!run.text && !run.hasHyperlinkChild) {
    return "";
  }

  // Build run CSS
  const runCss = generateRunInlineCss({
    font: run.font,
    size: run.size,
    bold: run.bold,
    italic: run.italic,
    underline: run.underline,
    strike: run.strike,
    color: run.color,
    highlight: run.highlight,
    verticalAlign: run.verticalAlign,
    smallCaps: run.smallCaps,
    shading: run.shading,
  });

  const styleAttr = runCss ? ` style="${runCss}"` : "";

  // Handle special elements
  if (run.isHyperlink) {
    const href = run.hyperlinkHref || "#";
    const target = run.hyperlinkTarget || "_blank";
    return `<a href="${escapeHtml(href)}" target="${escapeHtml(target)}"${styleAttr}>${escapeHtml(run.text)}</a>`;
  }

  if (run.isFootnoteRef) {
    const footnoteId = run.footnoteId || 0;
    context.footnoteRefs.push(footnoteId);
    return `<span class="footnote-ref" id="fnref${footnoteId}"><a href="#fn${footnoteId}">${footnoteId}</a></span>`;
  }

  if (run.isEndnoteRef) {
    const endnoteId = run.endnoteId || 0;
    context.endnoteRefs.push(endnoteId);
    return `<span class="endnote-ref" id="enref${endnoteId}"><a href="#en${endnoteId}">${endnoteId}</a></span>`;
  }

  // Build tag name based on formatting
  let tag = "span";
  if (run.bold) tag = "strong";
  else if (run.italic) tag = "em";

  const text = escapeHtml(run.text);
  return `<${tag}${styleAttr}>${text}</${tag}>`;
}

interface RunData {
  text: string;
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
  border?: {
    top?: { style?: string; size?: number; color?: string };
    bottom?: { style?: string; size?: number; color?: string };
    left?: { style?: string; size?: number; color?: string };
    right?: { style?: string; size?: number; color?: string };
  };
  isHyperlink?: boolean;
  hyperlinkHref?: string;
  hyperlinkTarget?: string;
  hasHyperlinkChild?: boolean;
  isFootnoteRef?: boolean;
  footnoteId?: number;
  isEndnoteRef?: boolean;
  endnoteId?: number;
}

interface ParagraphProperties {
  alignment?: string;
  spaceBefore?: string;
  spaceAfter?: string;
  lineSpacing?: string;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndent?: number;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
}

/**
 * Extract runs from paragraph XML, including nested elements.
 */
function extractRuns(paraXml: string): RunData[] {
  const runs: RunData[] = [];

  // First, check for hyperlinks that wrap runs
  const hyperlinkRegex = /<w:hyperlink[^>]*>([\s\S]*?)<\/w:hyperlink>/gi;
  let hyperMatch;

  while ((hyperMatch = hyperlinkRegex.exec(paraXml)) !== null) {
    const hyperXml = hyperMatch[0];
    const hrefMatch = /r:id="([^"]*)"/i.exec(hyperXml)
      || /w:anchor="([^"]*)"/i.exec(hyperXml);
    const href = hrefMatch ? hrefMatch[1] : "#";
    const targetMatch = hyperXml.match(/w:targetMode="([^"]*)"/i);

    // Extract runs inside hyperlink
    const innerRuns = extractRunsFromXml(hyperXml, {
      isHyperlink: true,
      hyperlinkHref: href,
      hyperlinkTarget: targetMatch ? targetMatch[1] : "_blank",
    });

    runs.push(...innerRuns);
  }

  // Also check for footnote references
  const footnoteRefRegex = /<w:r><w:rPr>[\s\S]*?<\/w:rPr><w:footnoteReference[^>]*w:id="([^"]*)"[^>]*\/><\/w:r>/gi;
  let fnMatch;

  while ((fnMatch = footnoteRefRegex.exec(paraXml)) !== null) {
    runs.push({
      text: fnMatch[1],
      isFootnoteRef: true,
      footnoteId: parseInt(fnMatch[1], 10),
    });
  }

  // And endnote references
  const endnoteRefRegex = /<w:r><w:rPr>[\s\S]*?<\/w:rPr><w:endnoteReference[^>]*w:id="([^"]*)"[^>]*\/><\/w:r>/gi;
  let enMatch;

  while ((enMatch = endnoteRefRegex.exec(paraXml)) !== null) {
    runs.push({
      text: enMatch[1],
      isEndnoteRef: true,
      endnoteId: parseInt(enMatch[1], 10),
    });
  }

  // If no hyperlinks found, just extract runs directly
  if (!paraXml.includes("<w:hyperlink")) {
    const directRuns = extractRunsFromXml(paraXml, {});
    runs.push(...directRuns);
  }

  return runs;
}

/**
 * Extract runs from XML content.
 */
function extractRunsFromXml(xml: string, overrides: Partial<RunData>): RunData[] {
  const runs: RunData[] = [];

  // Match w:r elements (but not w:rPr which is run properties)
  const runRegex = /<w:r(?![^>]*w:type)[^>]*>([\s\S]*?)<\/w:r>/gi;
  let runMatch;

  while ((runMatch = runRegex.exec(xml)) !== null) {
    const runXml = runMatch[0];
    const run = parseRunXml(runXml);
    Object.assign(run, overrides);
    runs.push(run);
  }

  // Also match self-closing w:r elements (might contain footnote/endnote references)
  const runSelfClosingRegex = /<w:r[^>]*\/>/gi;
  let runSelfMatch;

  while ((runSelfMatch = runSelfClosingRegex.exec(xml)) !== null) {
    const runXml = runSelfMatch[0];
    // Check for footnoteReference
    const fnRefMatch = /<w:footnoteReference[^>]*w:id="([^"]*)"/i.exec(runXml);
    if (fnRefMatch) {
      runs.push({
        text: "",
        isFootnoteRef: true,
        footnoteId: parseInt(fnRefMatch[1], 10),
      });
    }
    // Check for endnoteReference
    const enRefMatch = /<w:endnoteReference[^>]*w:id="([^"]*)"/i.exec(runXml);
    if (enRefMatch) {
      runs.push({
        text: "",
        isEndnoteRef: true,
        endnoteId: parseInt(enRefMatch[1], 10),
      });
    }
  }

  return runs;
}

/**
 * Parse a single run element to extract properties.
 */
function parseRunXml(runXml: string): RunData {
  const run: RunData = { text: "" };

  // Extract text content
  const textMatch = /<w:t[^>]*>([^<]*)<\/w:t>/i.exec(runXml);
  if (textMatch) {
    run.text = textMatch[1];
  }

  // Check for tab characters
  if (runXml.includes("<w:tab/>")) {
    run.text += "\t";
  }

  // Check for line break
  if (runXml.includes("<w:br/>")) {
    run.text += "\n";
  }

  // Parse run properties
  const rPrMatch = /<w:rPr>([\s\S]*?)<\/w:rPr>/i.exec(runXml);
  if (rPrMatch) {
    const rPrContent = rPrMatch[1];

    // Font
    const fontMatch = /<w:rFonts[^>]*w:ascii="([^"]*)"/i.exec(rPrContent)
      || /<w:rFonts[^>]*w:hAnsi="([^"]*)"/i.exec(rPrContent);
    if (fontMatch) run.font = fontMatch[1];

    // Size (convert half-points to points)
    const sizeMatch = /<w:sz[^>]*w:val="([^"]*)"/i.exec(rPrContent);
    if (sizeMatch) {
      const halfPt = parseInt(sizeMatch[1], 10);
      run.size = `${halfPt / 2}pt`;
    }

    // Bold
    if (/<w:b(?![a-z])[^>]*>/i.test(rPrContent)) run.bold = true;

    // Italic
    if (/<w:i(?![a-z])[^>]*>/i.test(rPrContent)) run.italic = true;

    // Underline
    const underlineMatch = /<w:u[^>]*w:val="([^"]*)"/i.exec(rPrContent);
    if (underlineMatch) {
      run.underline = underlineMatch[1];
    } else if (/<w:u[^>]*>/i.test(rPrContent)) {
      run.underline = "single";
    }

    // Strike
    const strikeMatch = /<w:strike[^>]*w:val="([^"]*)"/i.exec(rPrContent)
      || /<w:dstrike[^>]*w:val="([^"]*)"/i.exec(rPrContent);
    if (strikeMatch) {
      run.strike = strikeMatch[1];
    } else if (/<w:strike[^>]*>/i.test(rPrContent) || /<w:dstrike[^>]*>/i.test(rPrContent)) {
      run.strike = "single";
    }

    // Color
    const colorMatch = /<w:color[^>]*w:val="([^"]*)"/i.exec(rPrContent);
    if (colorMatch) run.color = colorMatch[1];

    // Highlight
    const highlightMatch = /<w:highlight[^>]*w:val="([^"]*)"/i.exec(rPrContent);
    if (highlightMatch) run.highlight = highlightMatch[1];

    // Vertical align
    const vertAlignMatch = /<w:vertAlign[^>]*w:val="([^"]*)"/i.exec(rPrContent);
    if (vertAlignMatch) run.verticalAlign = vertAlignMatch[1];

    // Small caps
    const smallCapsMatch = /<w:smallCaps[^>]*w:val="([^"]*)"/i.exec(rPrContent);
    if (smallCapsMatch) {
      run.smallCaps = smallCapsMatch[1] !== "0" && smallCapsMatch[1] !== "false";
    } else if (/<w:smallCaps[^>]*>/i.test(rPrContent)) {
      run.smallCaps = true;
    }

    // Shading
    const shadingMatch = /<w:shd[^>]*w:fill="([^"]*)"/i.exec(rPrContent);
    if (shadingMatch) run.shading = shadingMatch[1];

    // Check for hyperlink child
    run.hasHyperlinkChild = rPrContent.includes("<w:hyperlink");
  }

  return run;
}

/**
 * Parse paragraph properties from w:pPr content.
 */
function parseParagraphProperties(pPrContent: string): ParagraphProperties {
  const props: ParagraphProperties = {};

  // Alignment
  const jcMatch = /<w:jc[^>]*w:val="([^"]*)"/i.exec(pPrContent);
  if (jcMatch) props.alignment = jcMatch[1];

  // Spacing
  const spacingMatch = /<w:spacing[^>]*>/i.exec(pPrContent);
  if (spacingMatch) {
    const spacingContent = spacingMatch[0];

    const beforeMatch = /w:before="([^"]*)"/.exec(spacingContent);
    if (beforeMatch) props.spaceBefore = `${parseInt(beforeMatch[1], 10) / 20}pt`;

    const afterMatch = /w:after="([^"]*)"/.exec(spacingContent);
    if (afterMatch) props.spaceAfter = `${parseInt(afterMatch[1], 10) / 20}pt`;

    const lineMatch = /w:line="([^"]*)"/.exec(spacingContent);
    if (lineMatch) {
      const lineRule = /w:lineRule="([^"]*)"/.exec(spacingContent);
      if (lineRule && lineRule[1] === "exact") {
        // Exact line height in twips
        props.lineSpacing = `${parseInt(lineMatch[1], 10) / 20}pt`;
      } else {
        // Auto line height as multiplier
        props.lineSpacing = `${parseInt(lineMatch[1], 10) / 240}`;
      }
    }
  }

  // Indentation
  const indMatch = /<w:ind[^>]*>/i.exec(pPrContent);
  if (indMatch) {
    const indContent = indMatch[0];

    const leftMatch = /w:left="([^"]*)"/.exec(indContent);
    if (leftMatch) props.leftIndent = parseInt(leftMatch[1], 10) / 20;

    const rightMatch = /w:right="([^"]*)"/.exec(indContent);
    if (rightMatch) props.rightIndent = parseInt(rightMatch[1], 10) / 20;

    const firstLineMatch = /w:firstLine="([^"]*)"/.exec(indContent);
    if (firstLineMatch) props.firstLineIndent = parseInt(firstLineMatch[1], 10) / 20;
  }

  // Page break before
  if (/<w:pageBreakBefore[^>]*>/i.test(pPrContent)) {
    props.pageBreakBefore = true;
  }

  // Keep next
  if (/<w:keepNext[^>]*>/i.test(pPrContent)) {
    props.keepNext = true;
  }

  // Keep lines together
  if (/<w:keepLines[^>]*>/i.test(pPrContent)) {
    props.keepLines = true;
  }

  return props;
}

/**
 * Extract plain text from a paragraph.
 */
function extractTextFromParagraph(paraXml: string): string {
  // Extract all w:t elements
  const texts: string[] = [];
  const regex = /<w:t[^>]*>([^<]*)<\/w:t>/gi;
  let match;

  while ((match = regex.exec(paraXml)) !== null) {
    texts.push(match[1]);
  }

  let text = texts.join("");

  // Replace tabs with actual tab character
  text = text.replace(/\t/g, "\t");

  // Replace line breaks
  if (paraXml.includes("<w:br/>")) {
    text += "\n";
  }

  return text;
}

/**
 * Escape HTML special characters.
 */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
