/**
 * Type definitions for @officekit/mcp.
 *
 * This module provides TypeScript types for the MCP server implementation.
 */

// ============================================================================
// Result Types (mirroring core package)
// ============================================================================

/**
 * Standardized Result envelope for MCP operations.
 */
export interface Result<T> {
  ok: true;
  data: T;
}

/**
 * Error result with code and message.
 */
export interface ResultError {
  code: string;
  message: string;
  suggestion?: string;
}

/**
 * Failed Result type.
 */
export interface ResultErr {
  ok: false;
  error: ResultError;
}

/**
 * Union type for Result.
 */
export type MaybeResult<T> = Result<T> | ResultErr;

/**
 * Checks if a result is successful.
 */
export function isOk<T>(result: MaybeResult<T>): result is Result<T> {
  return result.ok === true;
}

/**
 * Checks if a result is an error.
 */
export function isErr<T>(result: MaybeResult<T>): result is ResultErr {
  return result.ok === false;
}

// ============================================================================
// MCP Protocol Types
// ============================================================================

/**
 * MCP JSON-RPC request.
 */
export interface McpRequest {
  jsonrpc: "2.0";
  id: string | number | null;
  method: string;
  params?: Record<string, unknown>;
}

/**
 * MCP JSON-RPC response.
 */
export interface McpResponse {
  jsonrpc: "2.0";
  id: string | number | null;
  result?: unknown;
  error?: {
    code: number;
    message: string;
    data?: unknown;
  };
}

/**
 * MCP tool definition.
 */
export interface McpTool {
  name: string;
  description: string;
  inputSchema: McpToolInputSchema;
}

/**
 * JSON Schema for tool input.
 */
export interface McpToolInputSchema {
  type: "object";
  properties: Record<string, unknown>;
  required?: string[];
}

/**
 * MCP server options.
 */
export interface McpServerOptions {
  /** Transport type: "stdio" or "http" (default: "stdio") */
  transport?: McpTransport;
  /** Port to listen on for HTTP transport (default: 3100) */
  port?: number;
  /** Host to bind to (default: 'localhost') */
  host?: string;
  /** Callback when server starts */
  onStart?: (port: number) => void;
  /** Callback when server stops */
  onStop?: () => void;
}

/**
 * Transport type for MCP communication.
 */
export type McpTransport = "stdio" | "http";

// ============================================================================
// Office Document Types
// ============================================================================

/**
 * Supported Office document formats.
 */
export type DocumentFormat = "word" | "excel" | "powerpoint";

/**
 * Result of executing an MCP tool.
 */
export interface ToolResult {
  content: Array<{ type: string; text: string }>;
  isError?: boolean;
}

/**
 * Success response from a tool call.
 */
export interface ToolSuccessResult extends ToolResult {
  content: Array<{ type: string; text: string }>;
  isError?: false;
}

/**
 * Error response from a tool call.
 */
export interface ToolErrorResult extends ToolResult {
  content: Array<{ type: string; text: string }>;
  isError: true;
}
