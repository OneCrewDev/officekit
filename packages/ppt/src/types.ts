/**
 * Shared TypeScript types for PowerPoint elements.
 *
 * These types define the data models used across the @officekit/ppt package
 * for slide management, mutations, query, and view operations.
 */

// ============================================================================
// Result Envelope
// ============================================================================

/**
 * Standard result envelope returned by all public API functions.
 * All operations return this consistent shape regardless of success/failure.
 */
export interface Result<T> {
  /** True if the operation succeeded */
  ok: boolean;
  /** The result data (present when ok is true) */
  data?: T;
  /** Error details (present when ok is false) */
  error?: ResultError;
}

/**
 * Error details included in a failed Result.
 */
export interface ResultError {
  /** Machine-readable error code */
  code: string;
  /** Human-readable error message */
  message: string;
  /** Optional suggestion for how to fix the error */
  suggestion?: string;
}

// ============================================================================
// Presentation Model
// ============================================================================

/**
 * Represents the entire PowerPoint presentation.
 */
export interface PresentationModel {
  /** Absolute path to the presentation file */
  filePath: string;
  /** Presentation metadata */
  metadata: PresentationMetadata;
  /** All slides in the presentation */
  slides: SlideModel[];
  /** Slide dimensions in EMUs (English Metric Units) */
  slideWidth?: number;
  slideHeight?: number;
  /** Slide size type (e.g., "screen16x9", "standard") */
  slideSize?: string;
  /** Theme information */
  theme?: ThemeModel;
}

/**
 * Core metadata for a presentation.
 */
export interface PresentationMetadata {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  description?: string;
  category?: string;
  lastModifiedBy?: string;
  revision?: string;
  created?: string;
  modified?: string;
  /** Default font for the presentation */
  defaultFont?: string;
}

// ============================================================================
// Slide Model
// ============================================================================

/**
 * Represents a single slide in a presentation.
 */
export interface SlideModel {
  /** Zero-based index of the slide (for internal use) */
  index: number;
  /** 1-based path index (e.g., "/slide[1]") */
  path: string;
  /** Slide title text (first title placeholder content) */
  title?: string;
  /** Notes text associated with the slide */
  notes?: string;
  /** Layout name used by this slide */
  layout?: string;
  /** Layout type (e.g., "title", "body", "twoColumnText") */
  layoutType?: string;
  /** Background settings */
  background?: SlideBackground;
  /** Transition animation */
  transition?: SlideTransition;
  /** All shapes on the slide */
  shapes: ShapeModel[];
  /** Tables on the slide */
  tables: TableModel[];
  /** Charts on the slide */
  charts: ChartModel[];
  /** Pictures/images on the slide */
  pictures: PictureModel[];
  /** Media elements (audio/video) */
  media: MediaModel[];
  /** Placeholder shapes by type */
  placeholders: PlaceholderModel[];
  /** Child count for various element types */
  childCount?: number;
}

/**
 * Slide background settings.
 */
export interface SlideBackground {
  /** Fill type: "solid", "gradient", "picture", "none" */
  fillType?: string;
  /** Solid fill color as hex (e.g., "FF0000" for red) */
  color?: string;
  /** Gradient colors if applicable */
  gradient?: GradientFill;
  /** Picture reference if applicable */
  pictureRelId?: string;
}

/**
 * Gradient fill settings.
 */
export interface GradientFill {
  type: "linear" | "radial";
  colors: string[];
  /** Angle in degrees for linear gradient */
  angle?: number;
}

/**
 * Slide transition/animation settings.
 */
export interface SlideTransition {
  /** Transition type (e.g., "fade", "push", "wipe") */
  type?: string;
  /** Transition duration in milliseconds */
  duration?: number;
  /** Transition direction if applicable */
  direction?: string;
}

// ============================================================================
// Shape Model
// ============================================================================

/**
 * Represents a shape on a slide.
 * Shapes include text boxes, rectangles, circles, etc.
 */
export interface ShapeModel {
  /** 1-based path index (e.g., "/slide[1]/shape[1]") */
  path: string;
  /** Shape name */
  name?: string;
  /** Shape text content */
  text?: string;
  /** Shape type (e.g., "shape", "textbox", "group") */
  type: string;
  /** Alternative text (alt text for accessibility) */
  alt?: string;
  /** Position and size in EMUs */
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  /** Rotation in degrees */
  rotation?: number;
  /** Fill color */
  fill?: string;
  /** Line/border color */
  line?: string;
  /** Line width */
  lineWidth?: number;
  /** Placeholder type if shape is a placeholder */
  placeholderType?: PlaceholderType;
  /** Placeholder index */
  placeholderIndex?: number;
  /** Paragraphs in the shape */
  paragraphs?: ParagraphModel[];
  /** Child count for grouped elements */
  childCount?: number;
}

/**
 * Placeholder type values as defined in OOXML.
 */
export type PlaceholderType =
  | "title"
  | "body"
  | "subtitle"
  | "centerTitle"
  | "centeredTitle"
  | "dateAndTime"
  | "footer"
  | "slideNumber"
  | "object"
  | "chart"
  | "table"
  | "clipArt"
  | "diagram"
  | "media"
  | "picture"
  | "header";

/**
 * Alias for PlaceholderType for backward compatibility.
 */
export type PlaceholderValues = PlaceholderType;

// ============================================================================
// Paragraph and Run Models
// ============================================================================

/**
 * Represents a paragraph within a shape's text body.
 */
export interface ParagraphModel {
  /** 1-based paragraph index within the shape */
  index: number;
  /** Paragraph text content (concatenated runs) */
  text: string;
  /** Text alignment */
  alignment?: "left" | "center" | "right" | "justify";
  /** Left margin in EMUs */
  marginLeft?: number;
  /** Right margin in EMUs */
  marginRight?: number;
  /** Line spacing */
  lineSpacing?: string;
  /** Space before in points */
  spaceBefore?: string;
  /** Space after in points */
  spaceAfter?: string;
  /** Runs within this paragraph */
  runs: RunModel[];
  /** Child count (number of runs) */
  childCount?: number;
}

/**
 * Represents a text run within a paragraph.
 */
export interface RunModel {
  /** 1-based run index within the paragraph */
  index: number;
  /** Run text content */
  text: string;
  /** Font typeface */
  font?: string;
  /** Font size in points */
  size?: string;
  /** Bold */
  bold?: boolean;
  /** Italic */
  italic?: boolean;
  /** Underline style */
  underline?: string;
  /** Strikethrough */
  strike?: string;
  /** Text color as hex */
  color?: string;
}

// ============================================================================
// Table Model
// ============================================================================

/**
 * Represents a table on a slide.
 */
export interface TableModel {
  /** 1-based path index (e.g., "/slide[1]/table[1]") */
  path: string;
  /** Table name */
  name?: string;
  /** Number of columns */
  columnCount?: number;
  /** Number of rows */
  rowCount?: number;
  /** Table rows */
  rows: TableRowModel[];
  /** First row is a header row */
  hasHeaderRow?: boolean;
}

/**
 * Represents a row in a table.
 */
export interface TableRowModel {
  /** 1-based row index */
  index: number;
  /** Path to this row (e.g., "/slide[1]/table[1]/tr[1]") */
  path: string;
  /** Cell count */
  cellCount?: number;
  /** Cells in this row */
  cells: TableCellModel[];
}

/**
 * Represents a cell in a table row.
 */
export interface TableCellModel {
  /** 1-based cell index */
  index: number;
  /** Path to this cell (e.g., "/slide[1]/table[1]/tr[1]/tc[1]") */
  path: string;
  /** Cell text content */
  text: string;
  /** Column span (merged cells) */
  gridSpan?: number;
  /** Row span (merged cells) */
  rowSpan?: number;
  /** Horizontal merge indicator */
  hmerge?: boolean;
  /** Vertical merge indicator */
  vmerge?: boolean;
  /** Fill color */
  fill?: string;
  /** Vertical alignment */
  valign?: "top" | "center" | "bottom";
  /** Horizontal alignment */
  alignment?: string;
  /** Font settings */
  font?: string;
  size?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}

// ============================================================================
// Chart Model
// ============================================================================

/**
 * Represents a chart on a slide.
 */
export interface ChartModel {
  /** 1-based path index (e.g., "/slide[1]/chart[1]") */
  path: string;
  /** Chart title */
  title?: string;
  /** Chart type (e.g., "bar", "line", "pie", "scatter") */
  type?: string;
  /** Chart series */
  series?: ChartSeriesModel[];
  /** Legend visibility */
  legend?: string | boolean;
  /** Data labels visibility */
  dataLabels?: string;
  /** Category axis title */
  categoryAxisTitle?: string;
  /** Value axis title */
  valueAxisTitle?: string;
  /** Axis minimum value */
  axisMin?: number;
  /** Axis maximum value */
  axisMax?: number;
  /** Major unit */
  majorUnit?: number;
  /** Minor unit */
  minorUnit?: number;
  /** Axis number format */
  axisNumberFormat?: string;
  /** Style ID */
  styleId?: number;
}

/**
 * Represents a series within a chart.
 */
export interface ChartSeriesModel {
  /** Series name */
  name?: string;
  /** Categories (X-axis values) */
  categories?: string;
  /** Values (Y-axis values) */
  values?: string;
  /** Series color */
  color?: string;
}

// ============================================================================
// Picture/Media Model
// ============================================================================

/**
 * Represents a picture/image on a slide.
 */
export interface PictureModel {
  /** 1-based path index (e.g., "/slide[1]/picture[1]") */
  path: string;
  /** Picture name */
  name?: string;
  /** Alternative text */
  alt?: string;
  /** Position and size in EMUs */
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  /** Rotation in degrees */
  rotation?: number;
  /** Media type (for media elements: "picture", "video", "audio") */
  mediaType?: string;
  /** Content type of the embedded media */
  contentType?: string;
  /** Size in bytes */
  size?: number;
}

/**
 * Represents a media element (audio or video).
 */
export interface MediaModel {
  /** 1-based path index (e.g., "/slide[1]/media[1]") */
  path: string;
  /** Media type: "audio" or "video" */
  type: "audio" | "video";
  /** Media name */
  name?: string;
  /** Position and size in EMUs */
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  /** Content type */
  contentType?: string;
  /** Size in bytes */
  size?: number;
}

// ============================================================================
// Placeholder Model
// ============================================================================

/**
 * Represents a placeholder shape on a slide.
 * Placeholders are special shapes that inherit their layout from the slide layout.
 */
export interface PlaceholderModel {
  /** 1-based path index (e.g., "/slide[1]/placeholder[1]") */
  path: string;
  /** Placeholder type name (e.g., "title", "body", "subtitle") */
  type: PlaceholderType;
  /** Placeholder index on the slide */
  index?: number;
  /** Shape name */
  name?: string;
  /** Placeholder text content */
  text?: string;
  /** The underlying shape model */
  shape?: ShapeModel;
}

// ============================================================================
// Theme Model
// ============================================================================

/**
 * Represents a theme used by the presentation.
 */
export interface ThemeModel {
  /** Theme name */
  name?: string;
  /** Color scheme */
  colors?: ThemeColors;
  /** Font scheme */
  fonts?: ThemeFonts;
}

/**
 * Theme color scheme.
 */
export interface ThemeColors {
  primary?: string;
  secondary?: string;
  background?: string;
  text?: string;
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  /** Dark variants */
  dark1?: string;
  dark2?: string;
  /** Light variants */
  light1?: string;
  light2?: string;
  /** Hyperlink colors */
  hyperlink?: string;
  followedHyperlink?: string;
}

/**
 * Theme font scheme.
 */
export interface ThemeFonts {
  /** Major font (headings) */
  major?: string;
  /** Minor font (body text) */
  minor?: string;
  /** Latin font */
  latin?: string;
  /** East Asian font */
  eastAsia?: string;
}

// ============================================================================
// Slide Master/Layout Models
// ============================================================================

/**
 * Represents a slide master.
 */
export interface SlideMasterModel {
  /** 1-based path index (e.g., "/slidemaster[1]") */
  path: string;
  /** Master name */
  name?: string;
  /** Theme name */
  theme?: string;
  /** Number of layouts using this master */
  layoutCount?: number;
  /** Shape count */
  shapeCount?: number;
}

/**
 * Represents a slide layout.
 */
export interface SlideLayoutModel {
  /** 1-based path index (e.g., "/slidelayout[1]") */
  path: string;
  /** Layout name */
  name?: string;
  /** Layout type */
  type?: string;
}

// ============================================================================
// Notes Model
// ============================================================================

/**
 * Represents notes for a slide.
 */
export interface NotesModel {
  /** Path to the notes (e.g., "/slide[1]/notes") */
  path: string;
  /** Notes text content */
  text: string;
}

// ============================================================================
// Connector/Group Models
// ============================================================================

/**
 * Represents a connector shape.
 */
export interface ConnectorModel {
  /** 1-based path index */
  path: string;
  /** Connector name */
  name?: string;
}

/**
 * Represents a group shape.
 */
export interface GroupModel {
  /** 1-based path index */
  path: string;
  /** Group name */
  name?: string;
  /** Child shape count */
  childCount?: number;
}

// ============================================================================
// Zoom Model
// ============================================================================

/**
 * Represents a zoom element (slide link) on a slide.
 */
export interface ZoomModel {
  /** 1-based path index */
  path: string;
  /** Zoom name */
  name?: string;
}

// ============================================================================
// Animation Model
// ============================================================================

/**
 * Represents an animation on a slide element.
 */
export interface AnimationModel {
  /** 1-based path index */
  path: string;
  /** Animation effect type */
  effect?: string;
  /** Animation class (entrance, exit, emphasis) */
  class?: string;
  /** Preset ID */
  presetId?: number;
  /** Duration in milliseconds */
  duration?: number;
  /** Delay in milliseconds */
  delay?: number;
  /** Ease-in percentage */
  easein?: number;
  /** Ease-out percentage */
  easeout?: number;
}

// ============================================================================
// Path Segment Types
// ============================================================================

/**
 * Represents a parsed segment of a PPT path.
 */
export interface PathSegment {
  /** Segment name (e.g., "slide", "shape", "table") */
  name: string;
  /** Index selector (for indexed segments like slide[1]) */
  index?: number;
  /** Name selector (for named segments like placeholder[title]) */
  nameSelector?: string;
  /** Type filter (for filtered segments like shape[type=text]) */
  typeFilter?: string;
}

/**
 * Parsed PPT path with all segments extracted.
 */
export interface ParsedPath {
  /** True if path is absolute (starts with /) */
  isAbsolute: boolean;
  /** Path segments */
  segments: PathSegment[];
  /** The original path string */
  original: string;
}

// ============================================================================
// Selector Types
// ============================================================================

/**
 * Represents a parsed selector for querying elements.
 */
export interface ParsedSelector {
  /** Element type being selected */
  elementType?: string;
  /** Slide number filter (1-based) */
  slideNum?: number;
  /** Attribute filters */
  attributes: Record<string, string>;
  /** Text content filter */
  textContains?: string;
  /** Child combinator for hierarchy */
  childCombinator?: string;
  /** Adjacent sibling combinator */
  adjacentSibling?: boolean;
}
