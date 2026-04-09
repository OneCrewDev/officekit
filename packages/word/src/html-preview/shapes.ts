/**
 * Shape and image rendering for Word HTML preview.
 * Handles wp:inline, wp:anchor, wsp shapes, and wgp groups.
 */

export interface DrawingRenderOptions {
  /** Relationship ID to image lookup */
  relationshipId?: string;
  /** Alt text for accessibility */
  altText?: string;
  /** Title for accessibility */
  title?: string;
  /** Whether drawing is for thumbnail */
  isThumbnail?: boolean;
}

/**
 * Render a drawing element (picture) to HTML.
 */
export function renderDrawingHtml(
  drawingXml: string,
  options: DrawingRenderOptions = {}
): string {
  const { altText, title } = options;

  // Check for inline drawing (wp:inline)
  const inlineMatch = /<wp:inline[^>]*>([\s\S]*?)<\/wp:inline>/i.exec(drawingXml);
  if (inlineMatch) {
    return renderInlineDrawing(inlineMatch[0], { altText, title });
  }

  // Check for anchor drawing (wp:anchor)
  const anchorMatch = /<wp:anchor[^>]*>([\s\S]*?)<\/wp:anchor>/i.exec(drawingXml);
  if (anchorMatch) {
    return renderAnchorDrawing(anchorMatch[0], { altText, title });
  }

  // Check for docPart gallery (wp:docPart)
  const docPartMatch = /<wp:docPart[^>]*>([\s\S]*?)<\/wp:docPart>/i.exec(drawingXml);
  if (docPartMatch) {
    return renderDocPartDrawing(docPartMatch[0], { altText, title });
  }

  // Fallback: try to extract any image
  const imgMatch = /<a:blip[^>]*r:embed="([^"]*)"/i.exec(drawingXml);
  if (imgMatch) {
    return `<img src="" alt="${escapeHtml(altText || "")}" title="${escapeHtml(title || "")}" class="drawing-inline" data-embed="${imgMatch[1]}">`;
  }

  return "";
}

/**
 * Render an inline drawing to HTML.
 */
function renderInlineDrawing(
  inlineXml: string,
  options: { altText?: string; title?: string }
): string {
  const { altText = "", title = "" } = options;

  // Extract image info
  const extentMatch = /<wp:extent[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"/i.exec(inlineXml);
  const cx = extentMatch ? parseInt(extentMatch[1], 10) : 0; // EMUs
  const cy = extentMatch ? parseInt(extentMatch[2], 10) : 0;

  // Convert EMUs to pixels (1 inch = 914400 EMUs, 1 inch = 96 pixels)
  const widthPx = Math.round(cx / 914400 * 96);
  const heightPx = Math.round(cy / 914400 * 96);

  // Get relationship ID for image source
  const embedMatch = /<a:blip[^>]*r:embed="([^"]*)"/i.exec(inlineXml);
  const embedId = embedMatch ? embedMatch[1] : "";

  // Check for other effects (duotone, grayscale, etc.)
  const effectPropsMatch = /<a:effectLst>([\s\S]*?)<\/a:effectLst>/i.exec(inlineXml);

  // Build img tag
  const styleParts: string[] = [];
  if (widthPx > 0) styleParts.push(`width: ${widthPx}px`);
  if (heightPx > 0) styleParts.push(`height: ${heightPx}px`);

  const styleAttr = styleParts.length > 0 ? ` style="${styleParts.join("; ")}"` : "";

  return `<img src="" alt="${escapeHtml(altText)}" title="${escapeHtml(title)}" class="drawing-inline"${styleAttr} data-embed="${embedId}" data-width="${widthPx}" data-height="${heightPx}">`;
}

/**
 * Render an anchor (floating) drawing to HTML.
 */
function renderAnchorDrawing(
  anchorXml: string,
  options: { altText?: string; title?: string }
): string {
  const { altText = "", title = "" } = options;

  // Extract positioning information
  const positionHMatch = /<wp:positionH[^>]*>([\s\S]*?)<\/wp:positionH>/i.exec(anchorXml);
  const positionH = parsePositionH(positionHMatch?.[1] || "");

  const positionVMatch = /<wp:positionV[^>]*>([\s\S]*?)<\/wp:positionV>/i.exec(anchorXml);
  const positionV = parsePositionV(positionVMatch?.[1] || "");

  // Extract extent
  const extentMatch = /<wp:extent[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"/i.exec(anchorXml);
  const cx = extentMatch ? parseInt(extentMatch[1], 10) : 0;
  const cy = extentMatch ? parseInt(extentMatch[2], 10) : 0;

  // Convert EMUs to points (1 inch = 914400 EMUs, 72 points)
  const widthPt = Math.round(cx / 914400 * 72);
  const heightPt = Math.round(cy / 914400 * 72);

  // Get relationship ID
  const embedMatch = /<a:blip[^>]*r:embed="([^"]*)"/i.exec(anchorXml);
  const embedId = embedMatch ? embedMatch[1] : "";

  // Build styles based on positioning
  const styles: string[] = [];

  // Determine float direction
  if (positionH) {
    if (positionH.relativeTo === "page" || positionH.relativeTo === "margin") {
      if (positionH.align === "right") {
        styles.push("float: right");
      } else if (positionH.align === "center") {
        styles.push("margin-left: auto");
        styles.push("margin-right: auto");
      } else {
        styles.push("float: left");
      }
    }

    if (positionH.offset !== undefined) {
      const offsetPt = positionH.offset / 914400 * 72;
      if (positionH.align === "right") {
        styles.push(`margin-right: ${offsetPt}pt`);
      } else if (positionH.align === "center") {
        // No margin for center
      } else {
        styles.push(`margin-left: ${offsetPt}pt`);
      }
    }
  } else {
    styles.push("float: left");
  }

  // Vertical positioning
  if (positionV && positionV.posOffset !== undefined) {
    if (positionV.relativeTo === "page" || positionV.relativeTo === "margin") {
      const posPt = positionV.posOffset / 914400 * 72;
      if (positionV.align === "top") {
        styles.push(`margin-top: ${posPt}pt`);
      } else if (positionV.align === "bottom") {
        styles.push(`margin-bottom: ${posPt}pt`);
      }
    }
  }

  // Size
  if (widthPt > 0) styles.push(`width: ${widthPt}pt`);
  if (heightPt > 0) styles.push(`height: ${heightPt}pt`);

  const styleAttr = styles.length > 0 ? ` style="${styles.join("; ")}"` : "";

  return `<img src="" alt="${escapeHtml(altText)}" title="${escapeHtml(title)}" class="drawing-anchor"${styleAttr} data-embed="${embedId}" data-width="${widthPt}" data-height="${heightPt}">`;
}

/**
 * Render a docPart (clip art, etc.) drawing.
 */
function renderDocPartDrawing(
  docPartXml: string,
  options: { altText?: string; title?: string }
): string {
  const { altText = "", title = "" } = options;

  // Extract docPart properties
  const docPartPrMatch = /<wp:docPartPr>([\s\S]*?)<\/wp:docPartPr>/i.exec(docPartXml);
  const name = docPartPrMatch
    ? /<wp:name[^>]*w:val="([^"]*)"/i.exec(docPartPrMatch[1])?.[1] || "Clip Art"
    : "Clip Art";

  // For now, just render a placeholder
  return `<div class="drawing-docpart" title="${escapeHtml(title || name)}" data-docpart="${escapeHtml(name)}">${escapeHtml(altText || name)}</div>`;
}

interface PositionH {
  relativeTo?: string;
  align?: string;
  posOffset?: number;
  offset?: number;
}

interface PositionV {
  relativeTo?: string;
  align?: string;
  posOffset?: number;
}

function parsePositionH(content: string): PositionH | undefined {
  if (!content) return undefined;

  const result: PositionH = {};

  // Get relativeTo
  const relativeMatch = /<wp:relativeHorizontal[^>]*w:relativeFrom="([^"]*)"/i.exec(content);
  if (relativeMatch) result.relativeTo = relativeMatch[1];

  // Get align
  const alignMatch = /<wp:align[^>]*w:val="([^"]*)"/i.exec(content);
  if (alignMatch) result.align = alignMatch[1];

  // Get position offset
  const posOffsetMatch = /<wp:positionH[^>]*><wp:posOffset>([^<]*)<\/wp:posOffset>/i.exec(content);
  if (posOffsetMatch) {
    result.posOffset = parseInt(posOffsetMatch[1], 10);
    result.offset = result.posOffset;
  }

  return result;
}

function parsePositionV(content: string): PositionV | undefined {
  if (!content) return undefined;

  const result: PositionV = {};

  const relativeMatch = /<wp:relativeVertical[^>]*w:relativeFrom="([^"]*)"/i.exec(content);
  if (relativeMatch) result.relativeTo = relativeMatch[1];

  const alignMatch = /<wp:align[^>]*w:val="([^"]*)"/i.exec(content);
  if (alignMatch) result.align = alignMatch[1];

  const posOffsetMatch = /<wp:positionV[^>]*><wp:posOffset>([^<]*)<\/wp:posOffset>/i.exec(content);
  if (posOffsetMatch) {
    result.posOffset = parseInt(posOffsetMatch[1], 10);
  }

  return result;
}

/**
 * Render a shape (wsp) element to HTML.
 */
export function renderShapeHtml(spXml: string, options: { altText?: string } = {}): string {
  const { altText = "" } = options;

  // Extract shape properties
  const spPrMatch = /<wsp:spPr>([\s\S]*?)<\/wsp:spPr>/i.exec(spXml);
  if (!spPrMatch) return "";

  const spPrContent = spPrMatch[1];

  // Get transform
  const xfrmMatch = /<a:xfrm[^>]*>([\s\S]*?)<\/a:xfrm>/i.exec(spPrContent);
  const offMatch = /<a:off[^>]*x="([^"]*)"[^>]*y="([^"]*)"/i.exec(spPrContent || "");
  const extMatch = /<a:ext[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"/i.exec(spPrContent || "");

  const x = offMatch ? parseInt(offMatch[1], 10) : 0;
  const y = offMatch ? parseInt(offMatch[2], 10) : 0;
  const cx = extMatch ? parseInt(extMatch[1], 10) : 0;
  const cy = extMatch ? parseInt(extMatch[2], 10) : 0;

  // Convert EMUs to points
  const leftPt = x / 914400 * 72;
  const topPt = y / 914400 * 72;
  const widthPt = cx / 914400 * 72;
  const heightPt = cy / 914400 * 72;

  // Get fill color
  const fillMatch = /<wsp:solidFill[^>]*>([\s\S]*?)<\/wsp:solidFill>/i.exec(spPrContent);
  let fillColor = "#e0e0e0";
  if (fillMatch) {
    const srgbMatch = /<a:srgbClr[^>]*a:val="([^"]*)"/i.exec(fillMatch[1]);
    if (srgbMatch) fillColor = "#" + srgbMatch[1];
  }

  // Get border
  const lnMatch = /<wsp:ln[^>]*>([\s\S]*?)<\/wsp:ln>/i.exec(spPrContent);
  let borderColor = "#000000";
  let borderWidth = 1;
  if (lnMatch) {
    const borderColorMatch = /<a:srgbClr[^>]*a:val="([^"]*)"/i.exec(lnMatch[1]);
    if (borderColorMatch) borderColor = "#" + borderColorMatch[1];
    const widthMatch = /<a:ln w="([^"]*)"/i.exec(lnMatch[1]);
    if (widthMatch) borderWidth = parseInt(widthMatch[1], 10) / 12700;
  }

  // Build HTML
  const style = `position: absolute; left: ${leftPt}pt; top: ${topPt}pt; width: ${widthPt}pt; height: ${heightPt}pt; background-color: ${fillColor}; border: ${borderWidth}px solid ${borderColor};`;

  // Extract text content
  const txbxMatch = /<wsp:txbx[^>]*>([\s\S]*?)<\/wsp:txbx>/i.exec(spXml);
  let content = altText;
  if (txbxMatch) {
    content = extractTextFromTextBox(txbxMatch[1]);
  }

  return `<div class="shape" style="${style}">${escapeHtml(content)}</div>`;
}

/**
 * Render a text box content.
 */
function extractTextFromTextBox(txbxContent: string): string {
  const texts: string[] = [];
  const textRegex = /<w:t[^>]*>([^<]*)<\/w:t>/gi;
  let match;

  while ((match = textRegex.exec(txbxContent)) !== null) {
    texts.push(match[1]);
  }

  return texts.join("");
}

/**
 * Render a group shape (wgp) to HTML.
 */
export function renderGroupHtml(groupXml: string, options: { altText?: string } = {}): string {
  const { altText = "" } = options;

  // Get group extent
  const grpSpPrMatch = /<wgp:grpSpPr>([\s\S]*?)<\/wgp:grpSpPr>/i.exec(groupXml);
  let widthPt = 200;
  let heightPt = 100;

  if (grpSpPrMatch) {
    const extMatch = /<a:ext[^>]*cx="([^"]*)"[^>]*cy="([^"]*)"/i.exec(grpSpPrMatch[1]);
    if (extMatch) {
      widthPt = parseInt(extMatch[1], 10) / 914400 * 72;
      heightPt = parseInt(extMatch[2], 10) / 914400 * 72;
    }
  }

  // Extract child shapes
  const childShapes: string[] = [];

  // Look for wsp (shape) elements
  const spRegex = /<wsp:sp[^>]*>([\s\S]*?)<\/wsp:sp>/gi;
  let spMatch;

  while ((spMatch = spRegex.exec(groupXml)) !== null) {
    childShapes.push(renderShapeHtml(spMatch[0], { altText }));
  }

  // Look for nested groups
  const nestedGroupRegex = /<wgp:grpSp[^>]*>([\s\S]*?)<\/wgp:grpSp>/gi;
  let nestedMatch;

  while ((nestedMatch = nestedGroupRegex.exec(groupXml)) !== null) {
    childShapes.push(renderGroupHtml(nestedMatch[0], { altText }));
  }

  const style = `position: relative; width: ${widthPt}pt; height: ${heightPt}pt;`;

  return `<div class="shape-group" style="${style}">${childShapes.join("\n")}</div>`;
}

/**
 * Resolve image data from relationship ID.
 * Returns base64 encoded image data.
 */
export async function resolveImageFromRelId(
  relId: string,
  zip: { file: (path: string) => { async: (type: string) => Promise<Uint8Array | string> } | null },
  documentRelationships: string
): Promise<string | null> {
  // Find the relationship target
  const relMatch = new RegExp(`Id="${relId}"[^>]*Target="([^"]*)"`, "i").exec(documentRelationships);
  if (!relMatch) return null;

  let imagePath = relMatch[1];
  // Make path absolute if relative
  if (!imagePath.startsWith("/") && !imagePath.includes(":")) {
    imagePath = "word/" + imagePath;
  }

  const imageEntry = zip.file(imagePath);
  if (!imageEntry) return null;

  const imageData = await imageEntry.async("base64");
  const imageType = getImageType(imagePath);

  return `data:${imageType};base64,${imageData}`;
}

/**
 * Determine image MIME type from extension.
 */
function getImageType(path: string): string {
  const ext = path.split(".").pop()?.toLowerCase();
  const typeMap: Record<string, string> = {
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    gif: "image/gif",
    bmp: "image/bmp",
    tiff: "image/tiff",
   tif: "image/tiff",
    svg: "image/svg+xml",
    emf: "image/x-emf",
    wmf: "image/x-wmf",
  };
  return typeMap[ext || ""] || "image/png";
}

/**
 * Check if a drawing contains an image.
 */
export function isDrawingImage(drawingXml: string): boolean {
  return /<a:blip[^>]*r:embed="([^"]*)"/i.test(drawingXml)
    || /<v:imagedata[^>]*>/i.test(drawingXml);
}

/**
 * Check if a drawing contains a shape.
 */
export function isDrawingShape(drawingXml: string): boolean {
  return /<wsp:sp[^>]*>/i.test(drawingXml)
    || /<wsp:grpSp[^>]*>/i.test(drawingXml);
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
