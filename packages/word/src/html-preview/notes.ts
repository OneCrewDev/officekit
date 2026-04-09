/**
 * Notes rendering for Word HTML preview.
 * Handles footnotes, endnotes, headers, and footers.
 */

import { renderParagraphHtml } from "./text.js";
import type { HtmlRenderContext } from "./context.js";

/**
 * Render footnotes section to HTML.
 */
export function renderFootnotesHtml(
  footnotesXml: string,
  context: HtmlRenderContext
): string {
  if (!footnotesXml || context.footnoteRefs.length === 0) {
    return "";
  }

  const footnoteIds = [...new Set(context.footnoteRefs)];
  const footnotes: Array<{ id: number; html: string }> = [];

  // Parse each footnote
  for (const id of footnoteIds) {
    const footnoteXml = extractFootnoteById(footnotesXml, id);
    if (footnoteXml) {
      footnotes.push({
        id,
        html: renderFootnoteContent(footnoteXml, id, context),
      });
    }
  }

  if (footnotes.length === 0) {
    return "";
  }

  // Build footnotes section
  let html = '<div class="footnotes">\n';
  html += '<hr>\n';
  html += '<h3>Footnotes</h3>\n';

  for (const fn of footnotes) {
    html += `<div class="footnote" id="fn${fn.id}">`;
    html += `<span class="footnote-ref">${fn.id}</span> `;
    html += fn.html;
    html += '</div>\n';
  }

  html += '</div>';

  return html;
}

/**
 * Render endnotes section to HTML.
 */
export function renderEndnotesHtml(
  endnotesXml: string,
  context: HtmlRenderContext
): string {
  if (!endnotesXml || context.endnoteRefs.length === 0) {
    return "";
  }

  const endnoteIds = [...new Set(context.endnoteRefs)];
  const endnotes: Array<{ id: number; html: string }> = [];

  // Parse each endnote
  for (const id of endnoteIds) {
    const endnoteXml = extractEndnoteById(endnotesXml, id);
    if (endnoteXml) {
      endnotes.push({
        id,
        html: renderEndnoteContent(endnoteXml, id, context),
      });
    }
  }

  if (endnotes.length === 0) {
    return "";
  }

  // Build endnotes section
  let html = '<div class="endnotes">\n';
  html += '<hr>\n';
  html += '<h3>Endnotes</h3>\n';

  for (const en of endnotes) {
    html += `<div class="endnote" id="en${en.id}">`;
    html += `<span class="endnote-ref">${en.id}</span> `;
    html += en.html;
    html += '</div>\n';
  }

  html += '</div>';

  return html;
}

/**
 * Extract and render a single footnote.
 */
function renderFootnoteContent(
  footnoteXml: string,
  _id: number,
  context: HtmlRenderContext
): string {
  // Extract the paragraph content from footnote
  const paraMatch = /<w:p[\s\S]*?<\/w:p>/i.exec(footnoteXml);
  if (!paraMatch) {
    return "";
  }

  return renderParagraphHtml(paraMatch[0], 0, context, {
    includeRuns: true,
    defaultStyle: "Normal",
  });
}

/**
 * Extract and render a single endnote.
 */
function renderEndnoteContent(
  endnoteXml: string,
  _id: number,
  context: HtmlRenderContext
): string {
  // Extract the paragraph content from endnote
  const paraMatch = /<w:p[\s\S]*?<\/w:p>/i.exec(endnoteXml);
  if (!paraMatch) {
    return "";
  }

  return renderParagraphHtml(paraMatch[0], 0, context, {
    includeRuns: true,
    defaultStyle: "Normal",
  });
}

/**
 * Extract footnote XML by ID.
 */
function extractFootnoteById(footnotesXml: string, id: number): string | null {
  // Match footnote with specific ID
  const regex = new RegExp(`<w:footnote[^>]*w:id="${id}"[^>]*>([\\s\\S]*?)<\\/w:footnote>`, "i");
  const match = regex.exec(footnotesXml);
  return match ? match[1] : null;
}

/**
 * Extract endnote XML by ID.
 */
function extractEndnoteById(endnotesXml: string, id: number): string | null {
  // Match endnote with specific ID
  const regex = new RegExp(`<w:endnote[^>]*w:id="${id}"[^>]*>([\\s\\S]*?)<\\/w:endnote>`, "i");
  const match = regex.exec(endnotesXml);
  return match ? match[1] : null;
}

/**
 * Render header content to HTML.
 */
export function renderHeaderHtml(
  headerXml: string,
  context: HtmlRenderContext
): string {
  if (!headerXml) {
    return "";
  }

  const parts: string[] = [];

  // Extract paragraphs from header
  const paraRegex = /<w:p[\s\S]*?<\/w:p>/gi;
  let paraMatch;

  while ((paraMatch = paraRegex.exec(headerXml)) !== null) {
    parts.push(renderParagraphHtml(paraMatch[0], 0, context));
  }

  if (parts.length === 0) {
    return "";
  }

  return `<div class="header">${parts.join("\n")}</div>`;
}

/**
 * Render footer content to HTML.
 */
export function renderFooterHtml(
  footerXml: string,
  context: HtmlRenderContext
): string {
  if (!footerXml) {
    return "";
  }

  const parts: string[] = [];

  // Extract paragraphs from footer
  const paraRegex = /<w:p[\s\S]*?<\/w:p>/gi;
  let paraMatch;

  while ((paraMatch = paraRegex.exec(footerXml)) !== null) {
    parts.push(renderParagraphHtml(paraMatch[0], 0, context));
  }

  if (parts.length === 0) {
    return "";
  }

  return `<div class="footer">${parts.join("\n")}</div>`;
}

/**
 * Determine header/footer type from XML.
 */
export function getHeaderFooterType(headerFooterXml: string): string {
  // Check for default header
  if (/<w:headerReference[^>]*w:type="default"/i.test(headerFooterXml)) {
    return "default";
  }
  if (/<w:headerReference[^>]*w:type="first"/i.test(headerFooterXml)) {
    return "first";
  }
  if (/<w:headerReference[^>]*w:type="even"/i.test(headerFooterXml)) {
    return "even";
  }
  if (/<w:headerReference[^>]*w:type="odd"/i.test(headerFooterXml)) {
    return "odd";
  }

  // Check for footer types
  if (/<w:footerReference[^>]*w:type="default"/i.test(headerFooterXml)) {
    return "default";
  }
  if (/<w:footerReference[^>]*w:type="first"/i.test(headerFooterXml)) {
    return "first";
  }
  if (/<w:footerReference[^>]*w:type="even"/i.test(headerFooterXml)) {
    return "even";
  }

  return "default";
}

/**
 * Extract page number field from header/footer XML.
 */
export function extractPageNumberField(headerFooterXml: string): { present: boolean; format?: string } {
  const fieldMatch = /<w:fldSimple[^>]*w:instr="PAGE"[^>]*>/i.exec(headerFooterXml);
  if (fieldMatch) {
    const formatMatch = /w:format="([^"]*)"/i.exec(fieldMatch[0]);
    return {
      present: true,
      format: formatMatch ? formatMatch[1] : undefined,
    };
  }

  // Check for complex field
  const complexMatch = /<w:fldChar[^>]*w:fldCharType="begin"[^>]*>[\s\S]*?<w:instrText[^>]*>[^<]*PAGE/i.exec(headerFooterXml);
  if (complexMatch) {
    return { present: true };
  }

  return { present: false };
}

/**
 * Extract page count field from header/footer XML.
 */
export function extractPageCountField(headerFooterXml: string): { present: boolean; format?: string } {
  const fieldMatch = /<w:fldSimple[^>]*w:instr="NUMPAGES"[^>]*>/i.exec(headerFooterXml);
  if (fieldMatch) {
    const formatMatch = /w:format="([^"]*)"/i.exec(fieldMatch[0]);
    return {
      present: true,
      format: formatMatch ? formatMatch[1] : undefined,
    };
  }

  const complexMatch = /<w:fldChar[^>]*w:fldCharType="begin"[^>]*>[\s\S]*?<w:instrText[^>]*>[^<]*NUMPAGES/i.exec(headerFooterXml);
  if (complexMatch) {
    return { present: true };
  }

  return { present: false };
}

/**
 * Render page number field as HTML.
 */
export function renderPageNumberField(format?: string): string {
  // Default format is decimal
  const pageFormat = format || "decimal";

  const formatMap: Record<string, string> = {
    decimal: "1",
    upperRoman: "I",
    lowerRoman: "i",
    upperLetter: "A",
    lowerLetter: "a",
    ordinal: "1st",
    cardinalTextNumber: "one",
    ordinalTextNumber: "First",
  };

  return `<span class="page-number" data-format="${pageFormat}">${formatMap[pageFormat] || "1"}</span>`;
}

/**
 * Render page count field as HTML.
 */
export function renderPageCountField(format?: string): string {
  const pageFormat = format || "decimal";

  const formatMap: Record<string, string> = {
    decimal: "1",
    upperRoman: "I",
    lowerRoman: "i",
    upperLetter: "A",
    lowerLetter: "a",
    ordinal: "1st",
    cardinalTextNumber: "one",
    ordinalTextNumber: "First",
  };

  return `<span class="page-count" data-format="${pageFormat}">${formatMap[pageFormat] || "1"}</span>`;
}

/**
 * Check if header/footer is linked to previous.
 */
export function isLinkedToPrevious(headerFooterXml: string): boolean {
  return /<w:footerReference[^>]*w:type="default"[^>]*r:id="[^"]*"/i.test(headerFooterXml)
    && /<w:separator>[^<]*<\/w:separator>/i.test(headerFooterXml);
}

/**
 * Render page break as HTML.
 */
export function renderPageBreak(): string {
  return '<div class="page-break"></div>';
}

/**
 * Render section break as HTML.
 */
export function renderSectionBreak(): string {
  return '<div class="section-break"></div>';
}
