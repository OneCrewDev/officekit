/**
 * Equation operations for @officekit/ppt.
 *
 * Provides functions to add, get, set, and remove mathematical equations
 * from PowerPoint slides using LaTeX or Office Math Markup Language (OMML).
 *
 * Equations are stored as `<a14:m>` elements inside shape elements (`<p:sp>`).
 * The package supports:
 * - LaTeX input (converted to OMML)
 * - Office Math Markup Language (OMML) directly
 * - Common symbols: fractions, roots, subscripts, superscripts, integrals, Greek letters
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput, notFound } from "./result.js";
import type { Result } from "./types.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Represents an equation on a slide.
 */
export interface EquationModel {
  /** 1-based path index (e.g., "/slide[1]/shape[1]") */
  path: string;
  /** Equation name */
  name?: string;
  /** Position X in EMUs */
  x?: number;
  /** Position Y in EMUs */
  y?: number;
  /** Width in EMUs */
  width?: number;
  /** Height in EMUs */
  height?: number;
  /** The equation content (LaTeX or OMML) */
  equation: string;
  /** Format of the equation: "latex" or "omml" */
  format: "latex" | "omml";
}

/**
 * Position for placing an equation on a slide.
 */
export interface EquationPosition {
  /** X position in EMUs (optional, defaults to 1 inch / 914400 EMUs) */
  x?: number;
  /** Y position in EMUs (optional, defaults to 2 inches / 1828800 EMUs) */
  y?: number;
  /** Width in EMUs (optional, defaults to 6 inches / 5486400 EMUs) */
  width?: number;
  /** Height in EMUs (optional, defaults to 1 inch / 914400 EMUs) */
  height?: number;
}

/**
 * Input equation with optional format specification.
 */
export interface EquationInput {
  /** LaTeX string or OMML markup */
  equation: string;
  /** Format: "latex" (default) or "omml" */
  format?: "latex" | "omml";
}

// ============================================================================
// Constants
// ============================================================================

/** Default position for equation */
const DEFAULT_POSITION: EquationPosition = {
  x: 914400, // 1 inch
  y: 1828800, // 2 inches
  width: 5486400, // 6 inches
  height: 914400, // 1 inch
};

/** Namespace for Office Math (OMML) */
const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
/** Namespace for Drawing ML */
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
/** Office 2010 extension namespace for math */
const A14_NS = "http://schemas.microsoft.com/office/drawing/2010/math";

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
 * Gets the relationships entry name for a given entry.
 */
function getRelationshipsEntryName(entryName: string): string {
  const directory = path.posix.dirname(entryName);
  const basename = path.posix.basename(entryName);
  return path.posix.join(directory, "_rels", `${basename}.rels`);
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
 * Reads an entry from the zip as a string.
 */
function requireEntry(zip: Map<string, Buffer>, entryName: string): string {
  const buffer = zip.get(entryName);
  if (!buffer) {
    throw new Error(`OOXML entry '${entryName}' is missing`);
  }
  return buffer.toString("utf8");
}

/**
 * Escapes special XML characters.
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
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
 * Finds the highest shape ID in a slide.
 */
function findMaxShapeId(spTreeXml: string): number {
  let maxId = 1;
  for (const match of spTreeXml.matchAll(/<p:sp\b[^>]*\bid="(\d+)"[^>]*>/g)) {
    const id = parseInt(match[1], 10);
    if (id > maxId) maxId = id;
  }
  for (const match of spTreeXml.matchAll(/<p:sp\b[^>]*>/g)) {
    const idMatch = match[0].match(/\bid="(\d+)"/);
    if (idMatch) {
      const id = parseInt(idMatch[1], 10);
      if (id > maxId) maxId = id;
    }
  }
  return maxId;
}

/**
 * Counts the number of equation shapes on a slide.
 */
function countEquationShapes(spTreeXml: string): number {
  let count = 0;
  for (const match of spTreeXml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<a14:m\b[\s\S]*?<\/a14:m>[\s\S]*?<\/p:sp>/g)) {
    count++;
  }
  // Also count shapes with oMath elements (OMML)
  for (const match of spTreeXml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<o:math\b[\s\S]*?<\/o:math>[\s\S]*?<\/p:sp>/g)) {
    count++;
  }
  return count;
}

// ============================================================================
// LaTeX to OMML Conversion
// ============================================================================

/**
 * Converts LaTeX to OMML (Office Math Markup Language).
 * Supports common mathematical structures.
 */
function latexToOmml(latex: string): string {
  // Remove surrounding whitespace
  latex = latex.trim();

  // Handle common LaTeX commands
  let omml = latex;

  // Fractions: \frac{a}{b} -> <m:frac><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:frac>
  omml = omml.replace(/\\frac\{([^}]*)\}\{([^}]*)\}/g, (_, num, den) => {
    return `<m:frac><m:num>${latexToOmml(num)}</m:num><m:den>${latexToOmml(den)}</m:den></m:frac>`;
  });

  // Square roots: \sqrt{x} -> <m:rad><m:deg></m:deg><m:base><m:r><m:t>x</m:t></m:r></m:base></m:rad>
  omml = omml.replace(/\\sqrt\{([^}]*)\}/g, (_, content) => {
    return `<m:rad><m:deg></m:deg><m:base>${latexToOmml(content)}</m:base></m:rad>`;
  });

  // Nth roots: \sqrt[n]{x} -> <m:rad><m:deg><m:r><m:t>n</m:t></m:r></m:deg><m:base>...</m:base></m:rad>
  omml = omml.replace(/\\sqrt\[([^\]]*)\]\{([^}]*)\}/g, (_, index, content) => {
    return `<m:rad><m:deg>${latexToOmml(index)}</m:deg><m:base>${latexToOmml(content)}</m:base></m:rad>`;
  });

  // Subscripts: x_{n} -> <m:sub><m:r><m:t>x</m:t></m:r><m:r><m:t>n</m:t></m:r></m:sub>
  omml = omml.replace(/([a-zA-Z0-9])\_{([^}]*)}/g, (_, base, sub) => {
    return `<m:sub>${latexToOmml(base)}${latexToOmml(sub)}</m:sub>`;
  });
  // Handle braces around subscripts
  omml = omml.replace(/_\{([^}]*)\}/g, (_, sub) => {
    return `<m:sub><m:r><m:t></m:t></m:r>${latexToOmml(sub)}</m:sub>`;
  });

  // Superscripts (exponents): x^{n} -> <m:sup><m:r><m:t>x</m:t></m:r><m:r><m:t>n</m:t></m:r></m:sup>
  omml = omml.replace(/([a-zA-Z0-9])\s*\^\{([^}]*)\}/g, (_, base, sup) => {
    return `<m:sup>${latexToOmml(base)}${latexToOmml(sup)}</m:sup>`;
  });
  // Handle ^ alone with braces
  omml = omml.replace(/\^\{([^}]*)\}/g, (_, sup) => {
    return `<m:sup><m:r><m:t></m:t></m:r>${latexToOmml(sup)}</m:sup>`;
  });

  // Greek letters
  const greekLetters: Record<string, string> = {
    "\\alpha": "α", "\\beta": "β", "\\gamma": "γ", "\\delta": "δ",
    "\\epsilon": "ε", "\\zeta": "ζ", "\\eta": "η", "\\theta": "θ",
    "\\iota": "ι", "\\kappa": "κ", "\\lambda": "λ", "\\mu": "μ",
    "\\nu": "ν", "\\xi": "ξ", "\\pi": "π", "\\rho": "ρ",
    "\\sigma": "σ", "\\tau": "τ", "\\upsilon": "υ", "\\phi": "φ",
    "\\chi": "χ", "\\psi": "ψ", "\\omega": "ω",
    "\\Alpha": "Α", "\\Beta": "Β", "\\Gamma": "Γ", "\\Delta": "Δ",
    "\\Epsilon": "Ε", "\\Zeta": "Ζ", "\\Eta": "Η", "\\Theta": "Θ",
    "\\Iota": "Ι", "\\Kappa": "Κ", "\\Lambda": "Λ", "\\Mu": "Μ",
    "\\Nu": "Ν", "\\Xi": "Ξ", "\\Pi": "Π", "\\Rho": "Ρ",
    "\\Sigma": "Σ", "\\Tau": "Τ", "\\Upsilon": "Υ", "\\Phi": "Φ",
    "\\Chi": "Χ", "\\Psi": "Ψ", "\\Omega": "Ω",
  };

  for (const [latex, char] of Object.entries(greekLetters)) {
    omml = omml.replace(new RegExp(latex.replace("\\", "\\\\"), "g"), char);
  }

  // Special symbols
  const specialSymbols: Record<string, string> = {
    "\\infty": "∞", "\\pm": "±", "\\mp": "∓", "\\times": "×",
    "\\div": "÷", "\\cdot": "·", "\\leq": "≤", "\\geq": "≥",
    "\\neq": "≠", "\\approx": "≈", "\\equiv": "≡", "\\sum": "∑",
    "\\prod": "∏", "\\int": "∫", "\\partial": "∂", "\\nabla": "∇",
    "\\forall": "∀", "\\exists": "∃", "\\in": "∈", "\\notin": "∉",
    "\\subset": "⊂", "\\supset": "⊃", "\\cup": "∪", "\\cap": "∩",
    "\\emptyset": "∅", "\\rightarrow": "→", "\\leftarrow": "←",
    "\\Rightarrow": "⇒", "\\Leftarrow": "⇐", "\\leftrightarrow": "↔",
    "\\langle": "⟨", "\\rangle": "⟩",
  };

  for (const [latex, char] of Object.entries(specialSymbols)) {
    omml = omml.replace(new RegExp(latex.replace("\\", "\\\\"), "g"), char);
  }

  // Integrals: \int_{a}^{b} -> <m:lim><m:sub>a</m:sub><m:sup>b</m:sup></m:lim>
  omml = omml.replace(/\\int_\{([^}]*)\}^\([^}]*\)/g, (_, _sub, sup) => {
    return `<m:lim>${latexToOmml(sup)}</m:lim>`;
  });

  // Handle plain text runs (wrapped in <m:r><m:t>)
  // If the content is not already wrapped in OMML tags, treat it as text
  if (!omml.startsWith("<m:") && !omml.startsWith("<o:")) {
    omml = `<m:r><m:t>${escapeXml(omml)}</m:t></m:r>`;
  }

  return omml;
}

/**
 * Wraps OMML content in the proper Office Math container.
 */
function wrapInMathContainer(ommlContent: string): string {
  return `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${ommlContent}</m:oMath>`;
}

// ============================================================================
// Equation Shape XML Building
// ============================================================================

/**
 * Creates a shape XML with an equation.
 */
function createEquationShapeXml(
  shapeId: number,
  name: string,
  position: EquationPosition,
  equationContent: string,
  isLatex: boolean
): string {
  const x = position.x ?? DEFAULT_POSITION.x;
  const y = position.y ?? DEFAULT_POSITION.y;
  const width = position.width ?? DEFAULT_POSITION.width;
  const height = position.height ?? DEFAULT_POSITION.height;

  let mathXml: string;
  if (isLatex) {
    const omml = latexToOmml(equationContent);
    mathXml = wrapInMathContainer(omml);
  } else {
    // Assume OMML - wrap if not already wrapped
    if (equationContent.includes("<m:oMath")) {
      mathXml = equationContent;
    } else {
      mathXml = wrapInMathContainer(equationContent);
    }
  }

  // Escape the math XML for embedding in the shape
  const escapedMath = mathXml.replace(/"/g, "&quot;");

  return `      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="${shapeId}" name="${escapeXml(name)}"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${x}" y="${y}"/>
            <a:ext cx="${width}" cy="${height}"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/math" content="${escapedMath}">${mathXml}</a14:m>
          </a:p>
        </p:txBody>
      </p:sp>
`;
}

/**
 * Extracts equation information from a shape XML.
 */
function extractEquationFromShape(shapeXml: string, shapePath: string): EquationModel | null {
  // Look for a14:m element
  const a14Match = shapeXml.match(/<a14:m\b[^>]*>([\s\S]*?)<\/a14:m>/);
  if (!a14Match) {
    // Look for o:math element
    const oMathMatch = shapeXml.match(/<o:math\b[^>]*>([\s\S]*?)<\/o:math>/);
    if (!oMathMatch) {
      return null;
    }
  }

  const mathContent = a14Match ? a14Match[1] : shapeXml.match(/<o:math\b[^>]*>([\s\S]*?)<\/o:math>/)?.[1] ?? "";

  // Extract position
  const xfrmMatch = shapeXml.match(/<a:xfrm>[\s\S]*?<a:off x="(\d+)" y="(\d+)"[\s\S]*?<\/a:xfrm>/);
  const extMatch = shapeXml.match(/<a:ext cx="(\d+)" cy="(\d+)"/);

  // Extract name
  const nameMatch = shapeXml.match(/<p:cNvPr\b[^>]*name="([^"]*)"[^>]*>/);
  const name = nameMatch ? nameMatch[1] : undefined;

  // Extract shape ID
  const idMatch = shapeXml.match(/<p:cNvPr\b[^>]*\bid="(\d+)"[^>]*>/);

  // Determine format (OMML if contains o:math or m:oMath)
  const format: "latex" | "omml" = mathContent.includes("<m:") || mathContent.includes("<o:") ? "omml" : "latex";

  return {
    path: shapePath,
    name,
    x: xfrmMatch ? parseInt(xfrmMatch[1], 10) : undefined,
    y: xfrmMatch ? parseInt(xfrmMatch[2], 10) : undefined,
    width: extMatch ? parseInt(extMatch[1], 10) : undefined,
    height: extMatch ? parseInt(extMatch[2], 10) : undefined,
    equation: mathContent,
    format,
  };
}

// ============================================================================
// Public API
// ============================================================================

/**
 * Adds an equation to a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param equation - LaTeX string or OMML markup
 * @param position - Position and size in EMUs (optional)
 * @returns Result with path to the new equation shape
 *
 * @example
 * // Add a LaTeX fraction
 * const result = await addEquation("/path/to/presentation.pptx", 1, "\\frac{a}{b}");
 *
 * @example
 * // Add with custom position
 * const result = await addEquation("/path/to/presentation.pptx", 1, "x^2 + y^2 = z^2", {
 *   x: 1000000,
 *   y: 2000000,
 *   width: 4000000,
 *   height: 500000
 * });
 */
export async function addEquation(
  filePath: string,
  slideIndex: number,
  equation: string,
  position: EquationPosition = DEFAULT_POSITION
): Promise<Result<{ path: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<never>;
    }

    const slideEntryPath = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntryPath);
    const slideRelsPath = getRelationshipsEntryName(slideEntryPath);
    const slideRelsXml = zip.get(slideRelsPath)?.toString("utf8") ?? "";
    const relationships = parseRelationshipEntries(slideRelsXml);
    const existingRelIds = relationships.map(r => r.id);

    // Find spTree and count all shapes
    const spTreeMatch = slideXml.match(/<p:spTree>([\s\S]*?)<\/p:spTree>/);
    if (!spTreeMatch) {
      return err("operation_failed", "Slide does not contain a shape tree (spTree)");
    }

    const spTreeContent = spTreeMatch[1];
    const allShapes = spTreeContent.match(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g) || [];
    const totalShapes = allShapes.length;
    const maxShapeId = findMaxShapeId(spTreeContent);
    const newShapeId = maxShapeId + 1;

    // Determine if input is LaTeX or OMML
    const isLatex = !equation.includes("<m:") && !equation.includes("<o:");

    // Create the equation shape
    const shapeName = `Equation ${totalShapes + 1}`;
    const newShapeXml = createEquationShapeXml(newShapeId, shapeName, position, equation, isLatex);

    // Insert the new shape before </p:spTree>
    const updatedSlideXml = slideXml.replace(/<\/p:spTree>/, `${newShapeXml}</p:spTree>`);

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntryPath) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));

    // The new shape is inserted at the end, so its index is totalShapes + 1
    const shapeIndex = totalShapes + 1;

    return ok({ path: `/slide[${slideIndex}]/shape[${shapeIndex}]` });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets all equations on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with array of equations
 *
 * @example
 * const result = await getEquations("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   for (const eq of result.data.equations) {
 *     console.log(eq.equation, eq.path);
 *   }
 * }
 */
export async function getEquations(
  filePath: string,
  slideIndex: number
): Promise<Result<{ equations: EquationModel[] }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<never>;
    }

    const slideEntryPath = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntryPath);

    // Find all shapes that contain equations
    const equations: EquationModel[] = [];

    // Find the spTree
    const spTreeMatch = slideXml.match(/<p:spTree>([\s\S]*?)<\/p:spTree>/);
    if (!spTreeMatch) {
      return ok({ equations: [] });
    }

    const spTreeContent = spTreeMatch[1];

    // Find all shapes
    const shapeMatches = spTreeContent.match(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g) || [];

    let shapeIndex = 0;
    for (const shapeXml of shapeMatches) {
      shapeIndex++;

      // Check if this shape has an equation
      if (!shapeXml.includes("<a14:m") && !shapeXml.includes("<o:math") && !shapeXml.includes("<m:oMath")) {
        continue;
      }

      const equation = extractEquationFromShape(shapeXml, `/slide[${slideIndex}]/shape[${shapeIndex}]`);
      if (equation) {
        equations.push(equation);
      }
    }

    return ok({ equations });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Updates an equation on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param path - Path to the equation shape (e.g., "/slide[1]/shape[3]")
 * @param equation - New LaTeX string or OMML markup
 * @returns Result with updated path
 *
 * @example
 * const result = await setEquation("/path/to/presentation.pptx", "/slide[1]/shape[1]", "\\frac{x}{y}");
 */
export async function setEquation(
  filePath: string,
  path: string,
  equation: string
): Promise<Result<{ path: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Parse the path to get slide index and shape index
    const slideMatch = path.match(/^\/slide\[(\d+)\]/);
    if (!slideMatch) {
      return invalidInput("Invalid equation path. Path must start with /slide[N]");
    }

    const slideIndex = parseInt(slideMatch[1], 10);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<never>;
    }

    const slideEntryPath = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntryPath);

    // Determine if input is LaTeX or OMML
    const isLatex = !equation.includes("<m:") && !equation.includes("<o:");

    let mathXml: string;
    if (isLatex) {
      const omml = latexToOmml(equation);
      mathXml = wrapInMathContainer(omml);
    } else {
      if (equation.includes("<m:oMath")) {
        mathXml = equation;
      } else {
        mathXml = wrapInMathContainer(equation);
      }
    }

    // Find the shape in the slide
    // We need to find the shape by path
    const shapeIndexMatch = path.match(/shape\[(\d+)\]/);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid equation path. Must include shape index");
    }

    const targetShapeIndex = parseInt(shapeIndexMatch[1], 10);

    // Find all shapes and locate the target
    const spTreeMatch = slideXml.match(/<p:spTree>([\s\S]*?)<\/p:spTree>/);
    if (!spTreeMatch) {
      return notFound("Equation", path, "Shape not found on slide");
    }

    const spTreeContent = spTreeMatch[1];
    const shapeMatches = spTreeContent.match(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g) || [];

    if (targetShapeIndex < 1 || targetShapeIndex > shapeMatches.length) {
      return notFound("Equation", path, `Shape index ${targetShapeIndex} is out of range`);
    }

    const targetShapeXml = shapeMatches[targetShapeIndex - 1];

    // Check if this shape has an equation
    if (!targetShapeXml.includes("<a14:m") && !targetShapeXml.includes("<o:math") && !targetShapeXml.includes("<m:oMath")) {
      return notFound("Equation", path, "Shape does not contain an equation");
    }

    // Extract the shape ID
    const idMatch = targetShapeXml.match(/<p:cNvPr\b[^>]*\bid="(\d+)"[^>]*>/);
    const shapeId = idMatch ? idMatch[1] : "1";

    // Extract the name
    const nameMatch = targetShapeXml.match(/<p:cNvPr\b[^>]*name="([^"]*)"[^>]*>/);
    const shapeName = nameMatch ? nameMatch[1] : "Equation";

    // Extract position from the shape
    const xfrmMatch = targetShapeXml.match(/<a:xfrm>[\s\S]*?<a:off x="(\d+)" y="(\d+)"[\s\S]*?<\/a:xfrm>[\s\S]*?<a:ext cx="(\d+)" cy="(\d+)"/);
    const position: EquationPosition = xfrmMatch ? {
      x: parseInt(xfrmMatch[1], 10),
      y: parseInt(xfrmMatch[2], 10),
      width: parseInt(xfrmMatch[3], 10),
      height: parseInt(xfrmMatch[4], 10),
    } : DEFAULT_POSITION;

    // Create updated shape
    const updatedShapeXml = createEquationShapeXml(parseInt(shapeId, 10), shapeName, position, equation, isLatex);

    // Replace the old shape with the new one
    const updatedSpTreeContent = spTreeContent.replace(targetShapeXml, updatedShapeXml);
    const updatedSlideXml = slideXml.replace(/<p:spTree>[\s\S]*?<\/p:spTree>/, `<p:spTree>${updatedSpTreeContent}</p:spTree>`);

    // Build new zip
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntryPath) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));

    return ok({ path });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Removes an equation from a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param path - Path to the equation shape (e.g., "/slide[1]/shape[3]")
 * @returns Result
 *
 * @example
 * const result = await removeEquation("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 */
export async function removeEquation(
  filePath: string,
  path: string
): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Parse the path to get slide index and shape index
    const slideMatch = path.match(/^\/slide\[(\d+)\]/);
    if (!slideMatch) {
      return invalidInput("Invalid equation path. Path must start with /slide[N]");
    }

    const slideIndex = parseInt(slideMatch[1], 10);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return slidePathResult as Result<never>;
    }

    const slideEntryPath = slidePathResult.data;
    const slideXml = requireEntry(zip, slideEntryPath);

    // Find the shape index
    const shapeIndexMatch = path.match(/shape\[(\d+)\]/);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid equation path. Must include shape index");
    }

    const targetShapeIndex = parseInt(shapeIndexMatch[1], 10);

    // Find all shapes
    const spTreeMatch = slideXml.match(/<p:spTree>([\s\S]*?)<\/p:spTree>/);
    if (!spTreeMatch) {
      return notFound("Equation", path, "Shape not found on slide");
    }

    const spTreeContent = spTreeMatch[1];
    const shapeMatches = spTreeContent.match(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g) || [];

    if (targetShapeIndex < 1 || targetShapeIndex > shapeMatches.length) {
      return notFound("Equation", path, `Shape index ${targetShapeIndex} is out of range`);
    }

    const targetShapeXml = shapeMatches[targetShapeIndex - 1];

    // Check if this shape has an equation
    if (!targetShapeXml.includes("<a14:m") && !targetShapeXml.includes("<o:math") && !targetShapeXml.includes("<m:oMath")) {
      return notFound("Equation", path, "Shape does not contain an equation");
    }

    // Remove the shape
    const updatedSpTreeContent = spTreeContent.replace(targetShapeXml, "");
    const updatedSlideXml = slideXml.replace(/<p:spTree>[\s\S]*?<\/p:spTree>/, `<p:spTree>${updatedSpTreeContent}</p:spTree>`);

    // Build new zip
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntryPath) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
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
