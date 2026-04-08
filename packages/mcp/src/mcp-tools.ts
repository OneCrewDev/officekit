/**
 * MCP Tool Definitions for @officekit/mcp.
 *
 * This module defines all the tools that the MCP server exposes to AI assistants
 * for operating on Office documents (Word, Excel, PowerPoint).
 */

import type { McpTool, McpToolInputSchema } from "./types.js";

// ============================================================================
// Tool Input Schemas
// ============================================================================

/**
 * Common filePath schema required by most tools.
 */
const filePathSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document (.docx, .xlsx, .pptx)",
    },
  },
  required: ["filePath"],
};

/**
 * Schema for path-based operations.
 */
const pathSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    path: {
      type: "string",
      description: "Path to the element (e.g., '/slide[1]/shape[1]' for PPT, '/body/p[1]' for Word, '/sheet[1]/cell[A1]' for Excel)",
    },
  },
  required: ["filePath", "path"],
};

/**
 * Schema for view operations with optional slide/sheet index.
 */
const viewOptionsSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    index: {
      type: "number",
      description: "Optional 1-based index for specific slide/sheet",
    },
  },
  required: ["filePath"],
};

/**
 * Schema for adding elements.
 */
const addSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    path: {
      type: "string",
      description: "Parent path where to add the element",
    },
    type: {
      type: "string",
      description: "Type of element to add (e.g., 'slide', 'shape', 'sheet', 'paragraph', 'cell')",
    },
    properties: {
      type: "object",
      description: "Properties for the new element",
    },
  },
  required: ["filePath", "path", "type"],
};

/**
 * Schema for setting element properties.
 */
const setSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    path: {
      type: "string",
      description: "Path to the element",
    },
    properties: {
      type: "object",
      description: "Properties to set",
    },
  },
  required: ["filePath", "path", "properties"],
};

/**
 * Schema for query operations.
 */
const querySchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    selector: {
      type: "string",
      description: "Query selector string",
    },
  },
  required: ["filePath"],
};

/**
 * Schema for batch operations.
 */
const batchSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    operations: {
      type: "array",
      description: "Array of operations to execute",
      items: {
        type: "object",
        properties: {
          op: {
            type: "string",
            description: "Operation type (set, add, remove, move, swap)",
          },
          path: {
            type: "string",
            description: "Path to the element",
          },
          properties: {
            type: "object",
            description: "Operation properties",
          },
        },
      },
    },
  },
  required: ["filePath", "operations"],
};

/**
 * Schema for merge operations.
 */
const mergeSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    templatePath: {
      type: "string",
      description: "Path to the template document with {{placeholders}}",
    },
    outputPath: {
      type: "string",
      description: "Path for the output document",
    },
    data: {
      type: "string",
      description: "JSON data to merge into the template",
    },
  },
  required: ["templatePath", "outputPath", "data"],
};

/**
 * Schema for raw XML operations.
 */
const rawSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    path: {
      type: "string",
      description: "Path to the element",
    },
  },
  required: ["filePath", "path"],
};

/**
 * Schema for raw-set operations.
 */
const rawSetSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    path: {
      type: "string",
      description: "Path to the element",
    },
    xml: {
      type: "string",
      description: "Raw XML to set",
    },
  },
  required: ["filePath", "path", "xml"],
};

/**
 * Schema for move operations.
 */
const moveSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    path: {
      type: "string",
      description: "Path to the element to move",
    },
    toPath: {
      type: "string",
      description: "Destination parent path",
    },
    index: {
      type: "number",
      description: "Position in destination",
    },
  },
  required: ["filePath", "path", "toPath"],
};

/**
 * Schema for swap operations.
 */
const swapSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the Office document",
    },
    path1: {
      type: "string",
      description: "First element path",
    },
    path2: {
      type: "string",
      description: "Second element path",
    },
  },
  required: ["filePath", "path1", "path2"],
};

// ============================================================================
// Tool Definitions
// ============================================================================

/**
 * All MCP tools exposed by the officekit MCP server.
 */
export const mcpTools: McpTool[] = [
  // ========== Document Creation ==========
  {
    name: "Create",
    description: "Creates a new blank Office document. Automatically detects format from file extension (.docx, .xlsx, .pptx).",
    inputSchema: {
      type: "object",
      properties: {
        filePath: {
          type: "string",
          description: "Path where to create the document",
        },
      },
      required: ["filePath"],
    },
  },

  // ========== View Operations ==========
  {
    name: "ViewAsText",
    description: "Extracts plain text content from the document. Use index for specific slide/sheet.",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsAnnotated",
    description: "Gets an annotated view showing element types, positions, names, and properties.",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsOutline",
    description: "Gets an outline/summary view of the document structure.",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsStats",
    description: "Gets statistics about the document (element counts, text lengths, etc.).",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsIssues",
    description: "Finds potential problems in the document (missing content, formatting issues, etc.).",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsHtml",
    description: "Renders the document as a self-contained HTML document (for PPT and Word).",
    inputSchema: viewOptionsSchema,
  },

  // ========== Query Operations ==========
  {
    name: "Get",
    description: "Gets detailed information about a specific element at the given path.",
    inputSchema: pathSchema,
  },
  {
    name: "Query",
    description: "Queries the document for elements matching a selector criteria.",
    inputSchema: querySchema,
  },

  // ========== Mutation Operations ==========
  {
    name: "Add",
    description: "Adds a new element to the document. Specify type and properties.",
    inputSchema: addSchema,
  },
  {
    name: "Set",
    description: "Sets properties on an existing element.",
    inputSchema: setSchema,
  },
  {
    name: "Remove",
    description: "Removes an element from the document.",
    inputSchema: pathSchema,
  },
  {
    name: "Move",
    description: "Moves an element from one location to another.",
    inputSchema: moveSchema,
  },
  {
    name: "Swap",
    description: "Swaps two elements in the document.",
    inputSchema: swapSchema,
  },

  // ========== Batch Operations ==========
  {
    name: "Batch",
    description: "Executes multiple mutations in a single batch operation for efficiency.",
    inputSchema: batchSchema,
  },

  // ========== Template Operations ==========
  {
    name: "Merge",
    description: "Merges data into a template document with {{placeholder}} keys.",
    inputSchema: mergeSchema,
  },

  // ========== Raw XML Operations ==========
  {
    name: "Raw",
    description: "Gets the raw XML for an element at the given path.",
    inputSchema: rawSchema,
  },
  {
    name: "RawSet",
    description: "Sets the raw XML for an element. Use with caution as this bypasses safety checks.",
    inputSchema: rawSetSchema,
  },

  // ========== Validation ==========
  {
    name: "Validate",
    description: "Validates the document against OpenXML schema and reports issues.",
    inputSchema: filePathSchema,
  },

  // ========== Document Info ==========
  {
    name: "DocumentInfo",
    description: "Gets basic information about the document (format, size, element counts).",
    inputSchema: filePathSchema,
  },
];

/**
 * Map of tool names to their definitions.
 */
export const toolByName = new Map<string, McpTool>(
  mcpTools.map((tool) => [tool.name, tool])
);
