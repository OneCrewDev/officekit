/**
 * MCP Server for @officekit/mcp.
 *
 * This module provides a Model Context Protocol (MCP) server that exposes
 * officekit operations as tools for AI assistants.
 */

import { Server } from "node:http";
import { createRequire } from "node:module";
import type {
  McpRequest,
  McpResponse,
  McpServerOptions,
  McpTransport,
  MaybeResult,
} from "./types.js";
import { isOk, isErr } from "./types.js";

// Create a require function for workspace packages
const require = createRequire(import.meta.url);

// ============================================================================
// MCP Server Implementation
// ============================================================================

/**
 * MCP Server for officekit document operations.
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
  async handleRequest(request: McpRequest): Promise<McpResponse> {
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
          name: "@officekit/mcp",
          version: "0.1.0",
        },
      },
    };
  }

  /**
   * Handles the tools/list request.
   */
  private handleToolsList(id: string | number | null): McpResponse {
    const { mcpTools } = require("./mcp-tools.js");

    return {
      jsonrpc: "2.0",
      id,
      result: {
        tools: mcpTools,
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
   * Detects the document format from file extension.
   */
  private detectFormat(filePath: string): string {
    const ext = filePath.toLowerCase().split(".").pop();
    switch (ext) {
      case "docx":
        return "word";
      case "xlsx":
        return "excel";
      case "pptx":
        return "powerpoint";
      default:
        return "unknown";
    }
  }

  /**
   * Gets the appropriate package loader for the document format.
   */
  private async getPackageForFormat(format: string): Promise<any> {
    switch (format) {
      case "powerpoint":
        return require("@officekit/ppt");
      case "word":
        return require("@officekit/word");
      case "excel":
        return require("@officekit/excel");
      default:
        throw new Error(`Unsupported document format: ${format}`);
    }
  }

  /**
   * Executes a tool by name with the given arguments.
   */
  private async executeTool(name: string, args: Record<string, unknown>): Promise<MaybeResult<unknown>> {
    const { filePath } = args as { filePath?: string };

    // All tools require a filePath except Create and Merge
    if (!filePath && name !== "Create" && name !== "Merge") {
      return {
        ok: false,
        error: {
          code: "invalid_input",
          message: "filePath is required for all tools except Create and Merge",
        },
      };
    }

    const format = filePath ? this.detectFormat(filePath) : "unknown";

    // Route to format-specific package
    let pkg: any;
    try {
      pkg = await this.getPackageForFormat(format);
    } catch {
      // Fall back to ppt for unknown formats
      pkg = require("@officekit/ppt");
    }

    switch (name) {
      // ========== Document Creation ==========
      case "Create": {
        const { filePath: createPath } = args as { filePath: string };
        // Use createExcelDocument, createWordDocument, etc. based on extension
        const createFn = pkg.createExcelDocument || pkg.createWordDocument;
        if (!createFn) {
          return { ok: false, error: { code: "not_supported", message: "Create not supported for this format" } };
        }
        const result = await createFn(createPath);
        return { ok: true, data: result };
      }

      // ========== View Operations ==========
      case "ViewAsText": {
        const { filePath: viewPath, index } = args as { filePath: string; index?: number };
        const viewFn = pkg.viewAsText || pkg.viewExcelDocument || pkg.viewWordDocument;
        if (viewFn) {
          return await viewFn(viewPath, index);
        }
        return { ok: false, error: { code: "not_supported", message: "ViewAsText not supported for this format" } };
      }

      case "ViewAsAnnotated": {
        const { filePath: viewPath, index } = args as { filePath: string; index?: number };
        // PowerPoint has viewAsAnnotated
        if (pkg.viewAsAnnotated) {
          return await pkg.viewAsAnnotated(viewPath, index);
        }
        return { ok: false, error: { code: "not_supported", message: "ViewAsAnnotated not supported for this format" } };
      }

      case "ViewAsOutline": {
        const { filePath: viewPath } = args as { filePath: string };
        if (pkg.viewWordOutlineJson) {
          return await pkg.viewWordOutlineJson(viewPath);
        }
        if (pkg.viewAsOutline) {
          return await pkg.viewAsOutline(viewPath);
        }
        return { ok: false, error: { code: "not_supported", message: "ViewAsOutline not supported for this format" } };
      }

      case "ViewAsStats": {
        const { filePath: viewPath } = args as { filePath: string };
        if (pkg.viewWordStatsJson) {
          return await pkg.viewWordStatsJson(viewPath);
        }
        if (pkg.viewAsStats) {
          return await pkg.viewAsStats(viewPath);
        }
        return { ok: false, error: { code: "not_supported", message: "ViewAsStats not supported for this format" } };
      }

      case "ViewAsIssues": {
        const { filePath: viewPath } = args as { filePath: string };
        if (pkg.viewWordIssuesJson) {
          return await pkg.viewWordIssuesJson(viewPath);
        }
        if (pkg.viewAsIssues) {
          return await pkg.viewAsIssues(viewPath);
        }
        return { ok: false, error: { code: "not_supported", message: "ViewAsIssues not supported for this format" } };
      }

      case "ViewAsHtml": {
        const { filePath: viewPath, index } = args as { filePath: string; index?: number };
        if (pkg.viewAsHtml) {
          return await pkg.viewAsHtml(viewPath, index);
        }
        if (pkg.renderExcelHtmlFromRoot) {
          return await pkg.renderExcelHtmlFromRoot(viewPath);
        }
        return { ok: false, error: { code: "not_supported", message: "ViewAsHtml not supported for this format" } };
      }

      // ========== Query Operations ==========
      case "Get": {
        const { filePath: getPath, path: elementPath } = args as { filePath: string; path: string };
        // Try format-specific get functions
        if (pkg.getPptNode) {
          return await pkg.getPptNode(getPath, elementPath);
        }
        if (pkg.getWordNode) {
          return await pkg.getWordNode(getPath, elementPath);
        }
        if (pkg.getExcelNode) {
          return await pkg.getExcelNode(getPath, elementPath);
        }
        return { ok: false, error: { code: "not_supported", message: "Get not supported for this format" } };
      }

      case "Query": {
        const { filePath: queryPath, selector } = args as { filePath: string; selector?: string };
        if (pkg.queryPptNodes) {
          return await pkg.queryPptNodes(queryPath, selector);
        }
        if (pkg.queryWordNodes) {
          return await pkg.queryWordNodes(queryPath, selector);
        }
        if (pkg.queryExcelNodes) {
          return await pkg.queryExcelNodes(queryPath, selector);
        }
        return { ok: false, error: { code: "not_supported", message: "Query not supported for this format" } };
      }

      // ========== Mutation Operations ==========
      case "Add": {
        const { filePath: addPath, path: addElementPath, type, properties } = args as {
          filePath: string;
          path: string;
          type: string;
          properties?: Record<string, unknown>;
        };
        if (pkg.addPptNode) {
          return await pkg.addPptNode(addPath, addElementPath, type, properties);
        }
        if (pkg.addWordNode) {
          return await pkg.addWordNode(addPath, addElementPath, type, properties);
        }
        if (pkg.addExcelNode) {
          return await pkg.addExcelNode(addPath, addElementPath, type, properties);
        }
        return { ok: false, error: { code: "not_supported", message: "Add not supported for this format" } };
      }

      case "Set": {
        const { filePath: setPath, path: setElementPath, properties } = args as {
          filePath: string;
          path: string;
          properties: Record<string, unknown>;
        };
        if (pkg.setPptNode) {
          return await pkg.setPptNode(setPath, setElementPath, properties);
        }
        if (pkg.setWordNode) {
          return await pkg.setWordNode(setPath, setElementPath, properties);
        }
        if (pkg.setExcelNode) {
          return await pkg.setExcelNode(setPath, setElementPath, properties);
        }
        return { ok: false, error: { code: "not_supported", message: "Set not supported for this format" } };
      }

      case "Remove": {
        const { filePath: removePath, path: removeElementPath } = args as { filePath: string; path: string };
        if (pkg.removePptNode) {
          return await pkg.removePptNode(removePath, removeElementPath);
        }
        if (pkg.removeWordNode) {
          return await pkg.removeWordNode(removePath, removeElementPath);
        }
        if (pkg.removeExcelNode) {
          return await pkg.removeExcelNode(removePath, removeElementPath);
        }
        return { ok: false, error: { code: "not_supported", message: "Remove not supported for this format" } };
      }

      case "Move": {
        const { filePath: movePath, path: moveElementPath, toPath, index } = args as {
          filePath: string;
          path: string;
          toPath: string;
          index?: number;
        };
        if (pkg.movePptNode) {
          return await pkg.movePptNode(movePath, moveElementPath, toPath, index);
        }
        if (pkg.moveWordNode) {
          return await pkg.moveWordNode(movePath, moveElementPath, toPath, index);
        }
        if (pkg.moveExcelNode) {
          return await pkg.moveExcelNode(movePath, moveElementPath, toPath, index);
        }
        return { ok: false, error: { code: "not_supported", message: "Move not supported for this format" } };
      }

      case "Swap": {
        const { filePath: swapPath, path1, path2 } = args as { filePath: string; path1: string; path2: string };
        if (pkg.swapPptNodes) {
          return await pkg.swapPptNodes(swapPath, path1, path2);
        }
        if (pkg.swapWordNodes) {
          return await pkg.swapWordNodes(swapPath, path1, path2);
        }
        if (pkg.swapExcelNodes) {
          return await pkg.swapExcelNodes(swapPath, path1, path2);
        }
        return { ok: false, error: { code: "not_supported", message: "Swap not supported for this format" } };
      }

      // ========== Batch Operations ==========
      case "Batch": {
        const { filePath: batchPath, operations } = args as {
          filePath: string;
          operations: Array<{ op: string; path?: string; properties?: Record<string, unknown> }>;
        };
        if (pkg.batchPptNodes) {
          return await pkg.batchPptNodes(batchPath, operations);
        }
        if (pkg.batchWordNodes) {
          return await pkg.batchWordNodes(batchPath, operations);
        }
        if (pkg.batchExcelNodes) {
          return await pkg.batchExcelNodes(batchPath, operations);
        }
        return { ok: false, error: { code: "not_supported", message: "Batch not supported for this format" } };
      }

      // ========== Template Operations ==========
      case "Merge": {
        const { templatePath, outputPath, data } = args as {
          templatePath: string;
          outputPath: string;
          data: string;
        };
        // Merge is format-specific based on template extension
        const mergePkg = await this.getPackageForFormat(this.detectFormat(templatePath));
        if (mergePkg.mergePptDocument) {
          return await mergePkg.mergePptDocument(templatePath, outputPath, JSON.parse(data));
        }
        if (mergePkg.mergeWordDocument) {
          return await mergePkg.mergeWordDocument(templatePath, outputPath, JSON.parse(data));
        }
        if (mergePkg.mergeExcelDocument) {
          return await mergePkg.mergeExcelDocument(templatePath, outputPath, JSON.parse(data));
        }
        return { ok: false, error: { code: "not_supported", message: "Merge not supported for this format" } };
      }

      // ========== Raw XML Operations ==========
      case "Raw": {
        const { filePath: rawPath, path: rawElementPath } = args as { filePath: string; path: string };
        if (pkg.rawPptDocument) {
          return await pkg.rawPptDocument(rawPath, rawElementPath);
        }
        if (pkg.rawWordDocument) {
          return await pkg.rawWordDocument(rawPath, rawElementPath);
        }
        if (pkg.rawExcelDocument) {
          return await pkg.rawExcelDocument(rawPath, rawElementPath);
        }
        return { ok: false, error: { code: "not_supported", message: "Raw not supported for this format" } };
      }

      case "RawSet": {
        const { filePath: rawSetPath, path: rawSetElementPath, xml } = args as {
          filePath: string;
          path: string;
          xml: string;
        };
        if (pkg.rawSetPptNode) {
          return await pkg.rawSetPptNode(rawSetPath, rawSetElementPath, xml);
        }
        if (pkg.rawSetWordDocument) {
          return await pkg.rawSetWordDocument(rawSetPath, rawSetElementPath, xml);
        }
        if (pkg.rawSetExcelNode) {
          return await pkg.rawSetExcelNode(rawSetPath, rawSetElementPath, xml);
        }
        return { ok: false, error: { code: "not_supported", message: "RawSet not supported for this format" } };
      }

      // ========== Validation ==========
      case "Validate": {
        const { filePath: validatePath } = args as { filePath: string };
        if (pkg.validatePptDocument) {
          return await pkg.validatePptDocument(validatePath);
        }
        if (pkg.validateWordDocument) {
          return await pkg.validateWordDocument(validatePath);
        }
        if (pkg.validateExcelDocument) {
          return await pkg.validateExcelDocument(validatePath);
        }
        return { ok: false, error: { code: "not_supported", message: "Validate not supported for this format" } };
      }

      // ========== Document Info ==========
      case "DocumentInfo": {
        const { filePath: infoPath } = args as { filePath: string };
        // Return basic info based on format
        return {
          ok: true,
          data: {
            format,
            path: infoPath,
            message: `Document info for ${format} format`,
          },
        };
      }

      default:
        return { ok: false, error: { code: "not_found", message: `Unknown tool: ${name}` } };
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
 * Run with: node --loader ts-node/mcp/src/mcp-server.ts
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
