/**
 * MCP Server for @officekit/ppt.
 *
 * This module provides a Model Context Protocol (MCP) server that exposes
 * PPTX operations as tools for AI assistants.
 */

import { Server } from "node:http";
import { readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import {
  addSlide,
  removeSlide,
  moveSlide,
  duplicateSlide,
  getSlides,
} from "./slides.js";
import {
  setShapeText,
  addShape,
  removeShape,
  swapShapes,
  setShapeProperty,
} from "./shapes.js";
import {
  swapSlides,
  copyShape,
  copySlide,
  rawGet,
  rawSet,
  batch,
} from "./mutations.js";
import { get, getSlide, getShape, getTable, getChart, querySlides, queryShapes } from "./query.js";
import {
  viewAsText,
  viewAsAnnotated,
  viewAsOutline,
  viewAsStats,
  viewAsIssues,
} from "./views.js";
import { viewAsHtml } from "./preview-html.js";
import { viewAsSvg } from "./preview-svg.js";
import { checkShapeTextOverflow } from "./views.js";
import { err, ok, isOk, isErr } from "./result.js";
import type { Result } from "./types.js";

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
// MCP Server Implementation
// ============================================================================

/**
 * MCP Server for PPTX operations.
 */
export class McpServer {
  private options: Required<McpServerOptions>;
  private server: Server | null = null;
  private transport: McpTransport;

  constructor(options: McpServerOptions = {}) {
    this.transport = options.transport ?? "stdio";
    this.options = {
      transport: this.transport,
      port: options.port ?? 3100,
      host: options.host ?? "localhost",
      onStart: options.onStart ?? (() => {}),
      onStop: options.onStop ?? (() => {}),
    };
  }

  /**
   * Gets the transport type for this server.
   */
  get transportType(): McpTransport {
    return this.transport;
  }

  /**
   * Handles an MCP JSON-RPC request.
   */
  private async handleRequest(request: McpRequest): Promise<McpResponse> {
    const { method, id, params } = request;

    try {
      switch (method) {
        case "initialize":
          return this.handleInitialize(id);

        case "tools/list":
          return this.handleToolsList(id);

        case "tools/call":
          return this.handleToolsCall(id, params as Record<string, unknown>);

        default:
          return {
            jsonrpc: "2.0",
            id,
            error: {
              code: -32601,
              message: `Method not found: ${method}`,
            },
          };
      }
    } catch (error) {
      return {
        jsonrpc: "2.0",
        id,
        error: {
          code: -32603,
          message: error instanceof Error ? error.message : "Internal error",
        },
      };
    }
  }

  /**
   * Handles the initialize request.
   */
  private handleInitialize(id: string | number | null): McpResponse {
    return {
      jsonrpc: "2.0",
      id,
      result: {
        protocolVersion: "2024-11-05",
        capabilities: {
          tools: {},
        },
        serverInfo: {
          name: "@officekit/ppt",
          version: "0.0.0",
        },
      },
    };
  }

  /**
   * Handles the tools/list request.
   */
  private handleToolsList(id: string | number | null): McpResponse {
    // Import tools dynamically to avoid circular dependencies
    const { pptTools } = require("./mcp-tools.js");

    return {
      jsonrpc: "2.0",
      id,
      result: {
        tools: pptTools,
      },
    };
  }

  /**
   * Handles the tools/call request.
   */
  private async handleToolsCall(
    id: string | number | null,
    params: Record<string, unknown>
  ): Promise<McpResponse> {
    const { name, arguments: args } = params as { name: string; arguments?: Record<string, unknown> };

    if (!name) {
      return {
        jsonrpc: "2.0",
        id,
        error: {
          code: -32602,
          message: "Missing tool name",
        },
      };
    }

    const toolArgs = args ?? {};

    try {
      const result = await this.executeTool(name, toolArgs);

      if (isErr(result)) {
        return {
          jsonrpc: "2.0",
          id,
          result: {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  ok: false,
                  error: result.error,
                }),
              },
            ],
            isError: true,
          },
        };
      }

      return {
        jsonrpc: "2.0",
        id,
        result: {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                ok: true,
                data: result.data,
              }),
            },
          ],
        },
      };
    } catch (error) {
      return {
        jsonrpc: "2.0",
        id,
        result: {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                ok: false,
                error: {
                  code: "execution_error",
                  message: error instanceof Error ? error.message : String(error),
                },
              }),
            },
          ],
          isError: true,
        },
      };
    }
  }

  /**
   * Executes a tool by name with the given arguments.
   */
  private async executeTool(name: string, args: Record<string, unknown>): Promise<Result<unknown>> {
    const { filePath } = args as { filePath?: string };

    if (!filePath) {
      return err("invalid_input", "filePath is required for all tools");
    }

    switch (name) {
      // ========== Slide Management ==========
      case "Add": {
        const { layoutId } = args as { layoutId?: number };
        return await addSlide(filePath, layoutId);
      }

      case "Remove": {
        const { index } = args as { index: number };
        return await removeSlide(filePath, index);
      }

      case "Move": {
        const { fromIndex, toIndex } = args as { fromIndex: number; toIndex: number };
        return await moveSlide(filePath, fromIndex, toIndex);
      }

      case "Swap": {
        const { index1, index2 } = args as { index1: number; index2: number };
        return await swapSlides(filePath, index1, index2);
      }

      case "CopyFrom": {
        const { sourceIndex, targetIndex, sourcePath, targetSlideIndex } = args as {
          sourceIndex?: number;
          targetIndex?: number;
          sourcePath?: string;
          targetSlideIndex?: number;
        };
        if (sourceIndex !== undefined && targetIndex !== undefined) {
          return await duplicateSlide(filePath, sourceIndex);
        }
        if (sourcePath !== undefined && targetSlideIndex !== undefined) {
          return await copyShape(filePath, sourcePath, targetSlideIndex);
        }
        return err("invalid_input", "Either (sourceIndex, targetIndex) or (sourcePath, targetSlideIndex) required");
      }

      // ========== Query Operations ==========
      case "Get": {
        const { pptPath } = args as { pptPath: string };
        return await get(filePath, pptPath);
      }

      case "Query": {
        const { selector } = args as { selector?: string };
        if (selector) {
          return await querySlides(filePath, selector);
        }
        return await querySlides(filePath);
      }

      // ========== Mutation Operations ==========
      case "Set": {
        const { pptPath, text } = args as { pptPath: string; text: string };
        return await setShapeText(filePath, pptPath, text);
      }

      case "AddPart": {
        const { slideIndex, shapeType, x, y, width, height } = args as {
          slideIndex: number;
          shapeType: string;
          x: number;
          y: number;
          width: number;
          height: number;
        };
        return await addShape(filePath, slideIndex, shapeType as any, { x, y }, { width, height });
      }

      case "Raw": {
        const { pptPath } = args as { pptPath: string };
        return await rawGet(filePath, pptPath);
      }

      case "RawSet": {
        const { pptPath, xml } = args as { pptPath: string; xml: string };
        return await rawSet(filePath, pptPath, xml);
      }

      case "Batch": {
        const { operations } = args as { operations: Array<{ op: string; params: Record<string, unknown> }> };
        return await batch(filePath, operations as any);
      }

      // ========== View Operations ==========
      case "ViewAsText": {
        const { slideIndex } = args as { slideIndex?: number };
        return await viewAsText(filePath, slideIndex);
      }

      case "ViewAsAnnotated": {
        const { slideIndex } = args as { slideIndex?: number };
        return await viewAsAnnotated(filePath, slideIndex);
      }

      case "ViewAsOutline": {
        const { slideIndex } = args as { slideIndex?: number };
        return await viewAsOutline(filePath, slideIndex);
      }

      case "ViewAsStats": {
        const { slideIndex } = args as { slideIndex?: number };
        return await viewAsStats(filePath, slideIndex);
      }

      case "ViewAsIssues": {
        const { slideIndex } = args as { slideIndex?: number };
        return await viewAsIssues(filePath, slideIndex);
      }

      case "ViewAsHtml": {
        const { slideIndex } = args as { slideIndex?: number };
        return await viewAsHtml(filePath, slideIndex);
      }

      case "ViewAsSvg": {
        const { slideIndex } = args as { slideIndex?: number };
        return await viewAsSvg(filePath, slideIndex);
      }

      // ========== Validation Operations ==========
      case "CheckShapeTextOverflow": {
        const { pptPath } = args as { pptPath: string };
        return await checkShapeTextOverflow(filePath, pptPath);
      }

      default:
        return err("not_found", `Unknown tool: ${name}`);
    }
  }

  /**
   * Starts the MCP server.
   */
  async start(): Promise<void> {
    if (this.transport === "stdio") {
      await this.startStdioServer();
    } else {
      await this.startHttpServer();
    }
  }

  /**
   * Stops the MCP server.
   */
  async stop(): Promise<void> {
    if (this.server) {
      await new Promise<void>((resolve) => {
        this.server!.close(() => resolve());
      });
      this.server = null;
    }
    this.options.onStop();
  }

  /**
   * Starts an HTTP server for MCP communication.
   */
  private async startHttpServer(): Promise<void> {
    this.server = new Server(async (req, res) => {
      if (req.method === "POST" && req.url === "/mcp") {
        let body = "";
        for await (const chunk of req) {
          body += chunk;
        }

        try {
          const request = JSON.parse(body) as McpRequest;
          const response = await this.handleRequest(request);
          res.writeHead(200, { "Content-Type": "application/json" });
          res.end(JSON.stringify(response));
        } catch (error) {
          res.writeHead(400, { "Content-Type": "application/json" });
          res.end(
            JSON.stringify({
              jsonrpc: "2.0",
              id: null,
              error: {
                code: -32700,
                message: "Parse error",
              },
            })
          );
        }
      } else if (req.method === "GET" && req.url === "/health") {
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ status: "ok" }));
      } else {
        res.writeHead(404);
        res.end();
      }
    });

    await new Promise<void>((resolve) => {
      this.server!.listen(this.options.port, this.options.host, () => {
        this.options.onStart(this.options.port);
        resolve();
      });
    });
  }

  /**
   * Starts a stdio server for MCP communication.
   */
  private async startStdioServer(): Promise<void> {
    const { readFileSync } = await import("node:fs");

    // Read requests from stdin
    let buffer = "";

    process.stdin.setEncoding("utf8");

    process.stdin.on("data", async (chunk: string) => {
      buffer += chunk;

      // Process complete JSON messages (newline-delimited)
      const lines = buffer.split("\n");
      buffer = lines.pop() ?? "";

      for (const line of lines) {
        if (line.trim()) {
          try {
            const request = JSON.parse(line) as McpRequest;
            const response = await this.handleRequest(request);
            process.stdout.write(JSON.stringify(response) + "\n");
          } catch (error) {
            const errorResponse: McpResponse = {
              jsonrpc: "2.0",
              id: null,
              error: {
                code: -32700,
                message: error instanceof Error ? error.message : "Parse error",
              },
            };
            process.stdout.write(JSON.stringify(errorResponse) + "\n");
          }
        }
      }
    });

    process.stdin.on("end", () => {
      this.stop();
    });
  }
}

// ============================================================================
// Factory Functions
// ============================================================================

/**
 * Creates an MCP server instance.
 *
 * @param options - Server options (includes transport type)
 * @returns MCP server instance
 *
 * @example
 * // Create a stdio-based MCP server (default)
 * const server = createMcpServer();
 *
 * // Create an HTTP MCP server on a specific port
 * const server = createMcpServer({ transport: "http", port: 3100 });
 */
export function createMcpServer(options: McpServerOptions = {}): McpServer {
  return new McpServer(options);
}

/**
 * Starts an MCP server with the specified transport.
 *
 * @param server - MCP server instance
 * @param transport - Transport type: "stdio" or "http"
 * @returns The started server instance
 *
 * @example
 * const server = createMcpServer();
 * await startMcpServer(server, "stdio");
 *
 * // Or start with HTTP transport
 * await startMcpServer(server, "http");
 */
export async function startMcpServer(
  server: McpServer,
  transport?: McpTransport
): Promise<McpServer> {
  // If transport is specified and differs from current, create new server
  if (transport !== undefined) {
    const currentTransport = server.transportType;
    if (transport !== currentTransport) {
      const newServer = createMcpServer({ transport });
      await newServer.start();
      return newServer;
    }
  }
  await server.start();
  return server;
}

// ============================================================================
// Main Entry Point (for stdio mode)
// ============================================================================

/**
 * Main entry point for stdio MCP server.
 * Run with: node --loader ts-node/ppt/src/mcp-server.ts
 */
async function main(): Promise<void> {
  const server = createMcpServer();
  await startMcpServer(server);
}

// Run main if this is the entry point
const isMainModule = import.meta.url === `file://${process.argv[1]}`;
if (isMainModule) {
  main().catch(console.error);
}
