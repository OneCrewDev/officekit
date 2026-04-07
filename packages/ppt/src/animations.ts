/**
 * Animation operations for @officekit/ppt.
 *
 * Provides functions to manage animations on slide elements:
 * - getAnimations: Get all animations on a slide
 * - setAnimation: Add or update an animation on an element
 * - removeAnimation: Remove an animation from an element
 *
 * Animation types: entrance, emphasis, exit, motion path
 * Animation triggers: onClick, afterPrev, withPrev, onLoad
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput, notFound } from "./result.js";
import type { Result, AnimationModel } from "./types.js";
import { getSlideIndex } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Animation trigger types.
 */
export type AnimationTrigger = "onClick" | "afterPrev" | "withPrev" | "onLoad";

/**
 * Animation class/types.
 */
export type AnimationClass = "entrance" | "exit" | "emphasis" | "motionPath";

/**
 * Specification for setting an animation.
 */
export interface AnimationSpec {
  /** Animation effect name (e.g., "fade", "fly", "wipe") */
  effect: string;
  /** Animation class: entrance, exit, emphasis, motionPath */
  class?: AnimationClass;
  /** Animation trigger */
  trigger?: AnimationTrigger;
  /** Duration in milliseconds */
  duration?: number;
  /** Delay in milliseconds */
  delay?: number;
}

/**
 * Result from getting animations on a slide.
 */
export interface SlideAnimationsResult {
  /** Slide index (1-based) */
  slideIndex: number;
  /** Slide path */
  path: string;
  /** All animations on the slide */
  animations: AnimationModel[];
  /** Count of animations */
  count: number;
}

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
 * Gets the slide entry path from the zip by slide index.
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
 * Loads a presentation and returns its zip contents.
 */
async function loadPresentation(filePath: string): Promise<Result<Map<string, Buffer>>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);
    return ok(zip);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Extracts animations from a slide's timing XML.
 */
function extractAnimationsFromTiming(timingXml: string, slideIndex: number): AnimationModel[] {
  const animations: AnimationModel[] = [];

  // Find all p:par (parallel time nodes) which contain animations
  const parPattern = /<p:par>[\s\S]*?<\/p:par>/g;
  const parMatches = timingXml.match(parPattern) || [];

  for (const parXml of parMatches) {
    // Extract the effect name
    const effectMatch = parXml.match(/<p:effect[^>]*name="([^"]*)"[^>]*>/);
    const effect = effectMatch ? effectMatch[1] : undefined;

    // Extract preset class
    const presetClassMatch = parXml.match(/<a:presetClass val="([^"]*)"[^>]*>/);
    const presetClass = presetClassMatch ? presetClassMatch[1] : undefined;

    // Map preset class to animation class
    let animClass: AnimationModel["class"] = undefined;
    if (presetClass === "entr") {
      animClass = "entrance";
    } else if (presetClass === "exit") {
      animClass = "exit";
    } else if (presetClass === "emph") {
      animClass = "emphasis";
    } else if (presetClass === "motionPath") {
      animClass = "motionPath";
    }

    // Extract duration
    const durMatch = parXml.match(/<p:cTn[^>]*dur="([^"]*)"[^>]*>/);
    let duration: number | undefined;
    if (durMatch) {
      // Duration in OOXML is in milliseconds
      duration = parseInt(durMatch[1], 10);
    }

    // Extract delay
    const delayMatch = parXml.match(/<a:stCondLst>[\s\S]*?<a:cond[^>]*delay="([^"]*)"[^>]*>/);
    let delay: number | undefined;
    if (delayMatch) {
      delay = parseInt(delayMatch[1], 10);
    }

    // Extract preset ID
    const presetIdMatch = parXml.match(/<a:presetID val="([^"]*)"[^>]*>/);
    let presetId: number | undefined;
    if (presetIdMatch) {
      presetId = parseInt(presetIdMatch[1], 10);
    }

    // Find the target shape reference
    const targetMatch = parXml.match(/<p:tar[^>]*spid="([^"]*)"[^>]*>/);
    let path: string | undefined;
    if (targetMatch) {
      // spid is the shape ID, we need to convert to path
      // For now, we'll construct a placeholder path
      path = `/slide[${slideIndex}]/shape[?]`; // Shape ID lookup would require additional processing
    }

    // Extract trigger type
    let trigger: AnimationTrigger | undefined;
    const triggerMatch = parXml.match(/<p:trigger[^>]*type="([^"]*)"[^>]*>/);
    if (triggerMatch) {
      const triggerType = triggerMatch[1];
      switch (triggerType) {
        case "click":
          trigger = "onClick";
          break;
        case "afterEffect":
          trigger = "afterPrev";
          break;
        case "withEffect":
          trigger = "withPrev";
          break;
        case "onLoad":
          trigger = "onLoad";
          break;
      }
    }

    // Only add if we have an effect name
    if (effect) {
      animations.push({
        path: path || `/slide[${slideIndex}]`,
        effect,
        class: animClass,
        presetId,
        duration,
        delay,
      });
    }
  }

  return animations;
}

/**
 * Gets animations from a slide.
 */
function getAnimationsFromSlideXml(slideXml: string, slideIndex: number): AnimationModel[] {
  // Find the timing element
  const timingMatch = slideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);
  if (!timingMatch) {
    return [];
  }

  const timingXml = timingMatch[0];
  return extractAnimationsFromTiming(timingXml, slideIndex);
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

// ============================================================================
// Public API
// ============================================================================

/**
 * Gets all animations on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with all animations on the slide
 *
 * @example
 * const result = await getAnimations("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(`Found ${result.data.count} animations`);
 *   for (const anim of result.data.animations) {
 *     console.log(`  ${anim.effect} (${anim.class})`);
 *   }
 * }
 */
export async function getAnimations(
  filePath: string,
  slideIndex: number,
): Promise<Result<SlideAnimationsResult>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code ?? "load_failed", zipResult.error?.message ?? "Failed to load presentation");
  }
  const zip = zipResult.data;
  if (!zip) {
    return err("operation_failed", "Failed to load presentation");
  }

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
  }

  const slideEntry = slidePathResult.data;
  if (!slideEntry) {
    return err("slide_not_found", "Slide entry not found");
  }
  const slideXml = requireEntry(zip, slideEntry);

  const animations = getAnimationsFromSlideXml(slideXml, slideIndex);

  return ok({
    slideIndex,
    path: `/slide[${slideIndex}]`,
    animations,
    count: animations.length,
  });
}

/**
 * Removes an animation from an element.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @returns Result indicating success
 *
 * @example
 * const result = await removeAnimation("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 */
export async function removeAnimation(
  filePath: string,
  pptPath: string,
): Promise<Result<void>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("removeAnimation requires a slide path");
  }

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

    // Extract shape index from path
    const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid shape path");
    }

    // Find the timing element and remove animations for this shape
    const updatedSlideXml = removeAnimationsForShape(slideXml, shapeIndexMatch[1]);

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
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

/**
 * Removes animations for a specific shape from slide XML.
 */
function removeAnimationsForShape(slideXml: string, shapeId: string): string {
  // Find timing element
  const timingMatch = slideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);
  if (!timingMatch) {
    return slideXml; // No timing element, nothing to remove
  }

  const timingXml = timingMatch[0];
  const timingIndex = slideXml.indexOf(timingXml);

  // Find all p:par elements containing this shape's animations
  const parPattern = /<p:par>[\s\S]*?<\/p:par>/g;
  let updatedTiming = timingXml;
  let parMatch;

  while ((parMatch = parPattern.exec(timingXml)) !== null) {
    const parXml = parMatch[0];
    const targetMatch = parXml.match(/<p:tar[^>]*spid="([^"]*)"[^>]*>/);

    if (targetMatch && targetMatch[1] === shapeId) {
      // Remove this p:par element from timing
      updatedTiming = updatedTiming.replace(parXml, "");
    }
  }

  // Clean up empty timing element
  const emptyTimingPattern = /<p:timing>[\s\S]*?<\/p:timing>/;
  const cleanTiming = updatedTiming.replace(/\s*<p:par>\s*<\/p:par>\s*/g, "");

  if (cleanTiming.match(/<p:timing>\s*<\/p:timing>/)) {
    // Timing is now empty, remove it entirely
    return slideXml.slice(0, timingIndex) + slideXml.slice(timingIndex + timingXml.length);
  }

  return slideXml.replace(timingXml, cleanTiming);
}

/**
 * Sets an animation on an element.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param animation - Animation specification
 * @returns Result indicating success
 *
 * @example
 * // Add a fade entrance animation
 * const result = await setAnimation("/path/to/presentation.pptx", "/slide[1]/shape[1]", {
 *   effect: "fade",
 *   class: "entrance",
 *   trigger: "onClick",
 *   duration: 500
 * });
 *
 * // Add a fly emphasis animation
 * const result = await setAnimation("/path/to/presentation.pptx", "/slide[1]/shape[2]", {
 *   effect: "fly",
 *   class: "emphasis",
 *   duration: 300,
 *   delay: 1000
 * });
 */
export async function setAnimation(
  filePath: string,
  pptPath: string,
  animation: AnimationSpec,
): Promise<Result<void>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("setAnimation requires a slide path");
  }

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

    // Extract shape index from path
    const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
    if (!shapeIndexMatch) {
      return invalidInput("Invalid shape path");
    }

    // First remove any existing animations on this shape
    let updatedSlideXml = removeAnimationsForShape(slideXml, shapeIndexMatch[1]);

    // Then add the new animation
    updatedSlideXml = addAnimationToSlide(updatedSlideXml, shapeIndexMatch[1], animation);

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
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

/**
 * Adds an animation to a slide's XML.
 */
function addAnimationToSlide(
  slideXml: string,
  shapeId: string,
  animation: AnimationSpec,
): string {
  // Build the animation XML
  const animXml = buildAnimationXml(shapeId, animation);

  // Check if timing element exists
  const timingMatch = slideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);

  if (timingMatch) {
    // Insert the new animation into existing timing
    const timingXml = timingMatch[0];
    const insertIndex = timingXml.indexOf("</p:timing>");
    const updatedTiming = timingXml.slice(0, insertIndex) + animXml + timingXml.slice(insertIndex);
    return slideXml.replace(timingMatch[0], updatedTiming);
  }

  // Create new timing element
  const newTiming = `<p:timing>${animXml}
  </p:timing>`;

  // Insert before </p:sld> or </p:spTree> or at end
  if (slideXml.includes("</p:spTree>")) {
    return slideXml.replace("</p:spTree>", `${newTiming}
  </p:spTree>`);
  }
  if (slideXml.includes("</p:sld>")) {
    return slideXml.replace("</p:sld>", `${newTiming}
  </p:sld>`);
  }

  return slideXml + newTiming;
}

/**
 * Builds animation XML for a shape.
 */
function buildAnimationXml(shapeId: string, animation: AnimationSpec): string {
  // Map animation class to OOXML preset class
  let presetClass = "entr";
  if (animation.class === "exit") {
    presetClass = "exit";
  } else if (animation.class === "emphasis") {
    presetClass = "emph";
  } else if (animation.class === "motionPath") {
    presetClass = "motionPath";
  }

  // Map trigger type
  let triggerType = "click";
  let triggerDelay = "0";
  if (animation.trigger === "afterPrev") {
    triggerType = "afterEffect";
    triggerDelay = "0";
  } else if (animation.trigger === "withPrev") {
    triggerType = "withEffect";
    triggerDelay = "0";
  } else if (animation.trigger === "onLoad") {
    triggerType = "onLoad";
    triggerDelay = "0";
  }

  const duration = animation.duration || 500;
  const delay = animation.delay || 0;

  // Build the animation XML structure
  // Using a basic fade animation as default
  const effectName = animation.effect || "fade";

  return `
    <p:par>
      <p:cnvPr id="2" name=""/>
      <p:cnvSpPr/>
      <p:stCxnSz/>
      <p:chOff/>
      <p:chSz/>
      <p:spPr/>
      <p:shape培/>
      <p:tmPrms/>
      <p:par>
        <p:pPr/>
        <p:extLst/>
        <p:anim calcmode="lin">
          <p:cBhvr>
            <p:cTn id="1" dur="${duration}" fill="hold">
              <p:stCondLst>
                <p:cond delay="${delay}">
                  <p:tn cfg="onStopAudio"/>
                </p:cond>
              </p:stCondLst>
            </p:cTn>
            <p:tavLst>
              <p:ta>
                <p:stAttr name="style.visibility" display="0" to="1"/>
              </p:ta>
            </p:tavLst>
          </p:cBhvr>
        </p:anim>
      </p:par>
    </p:par>`;
}