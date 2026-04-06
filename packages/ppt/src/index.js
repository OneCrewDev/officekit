// Manifest and contract exports
export {
  getPptAdapterContract,
  getPptAdapterManifest,
  pptAdapterManifest,
  summarizePptAdapterContract,
  summarizePptAdapter
} from "./manifest.js";

// Types
export * from "./types.js";

// Result envelope helpers
export * from "./result.js";

// Document handle and registry (resident/in-memory mode)
export * from "./document-handle.js";
export * from "./registry.js";

// Path parsing and resolution
export * from "./path.js";

// Selector grammar parser
export * from "./selectors.js";

// Slide management
export * from "./slides.js";

// Layout operations
export * from "./layouts.js";

// Notes management
export * from "./notes.js";

// Core mutations (Set, Remove, Swap, CopyFrom, Raw, Batch)
export * from "./mutations.js";

// Shape mutations
export * from "./shapes.js";

// Connector operations (addConnector, getConnectors, setConnectorEndpoints, removeConnector, setConnectorStyle)
export * from "./connectors.js";

// Table mutations
export * from "./tables.js";

// Text content operations
export * from "./text.js";

// Template merge operations (merge with {{key}} placeholders)
export * from "./merge.js";

// Background and fill effects
export * from "./background.js";

// Query and get operations
export * from "./query.js";

// View operations (ViewAsText, ViewAsAnnotated, ViewAsOutline, ViewAsStats, ViewAsIssues)
export * from "./views.js";

// Media operations (getMedia, addPicture, removeMedia, replacePicture, getMediaData)
export * from "./media.js";

// Advanced media operations (addVideo, addAudio, getMediaElements, removeMediaElement, setMediaOptions)
export * from "./media-advanced.js";

// Hyperlink operations (getHyperlink, setHyperlink, removeHyperlink, setExternalHyperlink, setInternalHyperlink)
export * from "./hyperlinks.js";

// Chart operations (getChart, addChart, setChartData, setChartType)
export * from "./charts.js";

// Theme operations (getTheme, getThemeColor, setThemeColor, getThemeFont, applyTheme)
export * from "./theme.js";

// Preview operations (viewAsHtml, viewAsSvg, generatePreview)
export * from "./preview-html.js";
export * from "./preview-svg.js";

// Overflow checking operations (checkShapeTextOverflow, checkSlideOverflow, getOverflowIssues)
export * from "./overflow.js";

// Animation operations (getAnimations, setAnimation, removeAnimation)
export * from "./animations.js";

// Watch operations (watch - live preview with auto-refresh)
export * from "./watch.js";

// 3D Model operations (get3DModels, add3DModel, remove3DModel, set3DModelRotation)
export * from "./models-3d.js";

// MCP Server for AI assistant integration
export * from "./mcp-server.js";
export * from "./mcp-tools.js";

// Equation operations (addEquation, getEquations, setEquation, removeEquation)
export * from "./equations.js";
