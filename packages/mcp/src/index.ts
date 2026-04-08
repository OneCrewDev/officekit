/**
 * @officekit/mcp - MCP Server for Office Documents
 *
 * This package provides a Model Context Protocol (MCP) server that exposes
 * officekit operations as tools for AI assistants.
 *
 * @example
 * // Create and start an MCP server
 * import { createMcpServer, startMcpServer } from "@officekit/mcp";
 *
 * const server = createMcpServer();
 * await startMcpServer(server);
 */

// MCP Server
export {
  McpServer,
  createMcpServer,
  startMcpServer,
} from "./mcp-server.js";

// MCP Tools
export {
  mcpTools,
  toolByName,
} from "./mcp-tools.js";

// Types (from types.ts)
export {
  type McpServerOptions,
  type McpTransport,
  type McpTool,
  type McpToolInputSchema,
  type McpRequest,
  type McpResponse,
  type Result,
  type MaybeResult,
  type ResultError,
  type ResultErr,
  type DocumentFormat,
  type ToolResult,
  type ToolSuccessResult,
  type ToolErrorResult,
  isOk,
  isErr,
} from "./types.js";
