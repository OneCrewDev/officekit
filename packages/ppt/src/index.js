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

// Table mutations
export * from "./tables.js";

// Query and get operations
export * from "./query.js";

// View operations (ViewAsText, ViewAsAnnotated, ViewAsOutline, ViewAsStats, ViewAsIssues)
export * from "./views.js";
