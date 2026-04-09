/**
 * HTML Preview rendering context types.
 */

export interface HtmlRenderContext {
  /** Footnote reference IDs encountered during rendering */
  footnoteRefs: number[];
  /** Endnote reference IDs encountered during rendering */
  endnoteRefs: number[];
  /** Top-anchored images that need to be rendered separately */
  topAnchoredImages: Array<{ markerId: string; imgHtml: string }>;
  /** Whether we're currently rendering the document body */
  renderingBody: boolean;
  /** Available width for current line (in points) */
  lineWidthPt: number;
  /** Accumulated width on current line (in points) */
  lineAccumPt: number;
  /** Whether line-break tracking is active */
  lineBreakEnabled: boolean;
  /** Default font size for width estimation (in points) */
  defaultFontSizePt: number;
  /** Current paragraph index */
  paraIndex: number;
  /** Document settings for rPr default fallback */
  rPrDefaultFontSize?: number;
  /** Cached page layout info */
  pageLayout?: PageLayoutInfo;
}

export interface PageLayoutInfo {
  pageWidthTwips: number;
  pageHeightTwips: number;
  marginTopTwips: number;
  marginBottomTwips: number;
  marginLeftTwips: number;
  marginRightTwips: number;
  orientation?: string;
  columns?: number;
  columnSpaceTwips?: number;
}

export interface HtmlPreviewOptions {
  /** Page filter string (e.g., "1", "2-5", "1,3,5") */
  pageFilter?: string;
  /** Whether to include full styles */
  includeStyles?: boolean;
  /** Custom CSS to inject */
  customCss?: string;
  /** Render in single page mode (no page breaks, content flows naturally) */
  singlePage?: boolean;
}

export function createHtmlRenderContext(): HtmlRenderContext {
  return {
    footnoteRefs: [],
    endnoteRefs: [],
    topAnchoredImages: [],
    renderingBody: false,
    lineWidthPt: 612, // Default 8.5in - 1.5in margins = 7in = 504pt content
    lineAccumPt: 0,
    lineBreakEnabled: false,
    defaultFontSizePt: 11,
    paraIndex: 0,
  };
}
