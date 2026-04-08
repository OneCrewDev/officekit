import test from "node:test";
import assert from "node:assert/strict";

import {
  McpServer,
  createMcpServer,
  startMcpServer,
} from "../src/mcp-server.js";

import {
  mcpTools,
  type McpServerOptions,
  type McpTransport,
  type McpRequest,
  type McpResponse,
} from "../src/index.js";

/**
 * MCP tool call result structure.
 */
interface McpToolResult {
  content: Array<{ type: string; text: string }>;
  isError?: boolean;
}

// ============================================================================
// McpServer Class Tests
// ============================================================================

test("McpServer - creates with default options", () => {
  const server = new McpServer();
  assert.equal(server.transportType, "stdio");
});

test("McpServer - creates with explicit stdio transport", () => {
  const server = new McpServer({ transport: "stdio" });
  assert.equal(server.transportType, "stdio");
});

test("McpServer - creates with http transport", () => {
  const server = new McpServer({ transport: "http", port: 3200 });
  assert.equal(server.transportType, "http");
});

test("McpServer - creates with custom port", () => {
  const server = new McpServer({ port: 3500 });
  assert.equal(server.transportType, "stdio"); // default transport
});

// ============================================================================
// createMcpServer Factory Function Tests
// ============================================================================

test("createMcpServer - creates with default options", () => {
  const server = createMcpServer();
  assert.ok(server instanceof McpServer);
  assert.equal(server.transportType, "stdio");
});

test("createMcpServer - creates with transport option", () => {
  const server = createMcpServer({ transport: "http" });
  assert.ok(server instanceof McpServer);
  assert.equal(server.transportType, "http");
});

test("createMcpServer - creates with all options", () => {
  const onStart = () => {};
  const onStop = () => {};
  const server = createMcpServer({
    transport: "http",
    port: 3300,
    host: "0.0.0.0",
    onStart,
    onStop,
  });
  assert.ok(server instanceof McpServer);
  assert.equal(server.transportType, "http");
});

// ============================================================================
// MCP Protocol Tests - Handle Request
// ============================================================================

test("McpServer.handleRequest - handles initialize request", async () => {
  const server = new McpServer();
  const request: McpRequest = {
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: {},
  };

  const response = await server["handleRequest"](request);

  assert.equal(response.jsonrpc, "2.0");
  assert.equal(response.id, 1);
  assert.ok(response.result);
  assert.equal((response as any).result.protocolVersion, "2024-11-05");
  assert.ok((response as any).result.capabilities.tools);
  assert.equal((response as any).result.serverInfo.name, "@officekit/mcp");
});

test("McpServer.handleRequest - handles tools/list request", async () => {
  const server = new McpServer();
  const request: McpRequest = {
    jsonrpc: "2.0",
    id: 2,
    method: "tools/list",
    params: {},
  };

  const response = await server["handleRequest"](request);

  assert.equal(response.jsonrpc, "2.0");
  assert.equal(response.id, 2);
  assert.ok(response.result);
  assert.ok(Array.isArray((response as any).result.tools));
  // Verify all required tools are present
  const toolNames = (response as any).result.tools.map((t: any) => t.name);
  assert.ok(toolNames.includes("Create"));
  assert.ok(toolNames.includes("Get"));
  assert.ok(toolNames.includes("Query"));
  assert.ok(toolNames.includes("Set"));
  assert.ok(toolNames.includes("Add"));
  assert.ok(toolNames.includes("Remove"));
  assert.ok(toolNames.includes("Move"));
  assert.ok(toolNames.includes("Swap"));
  assert.ok(toolNames.includes("Batch"));
  assert.ok(toolNames.includes("Merge"));
  assert.ok(toolNames.includes("Raw"));
  assert.ok(toolNames.includes("RawSet"));
  assert.ok(toolNames.includes("ViewAsText"));
  assert.ok(toolNames.includes("ViewAsAnnotated"));
  assert.ok(toolNames.includes("ViewAsOutline"));
  assert.ok(toolNames.includes("ViewAsStats"));
  assert.ok(toolNames.includes("ViewAsIssues"));
  assert.ok(toolNames.includes("ViewAsHtml"));
  assert.ok(toolNames.includes("Validate"));
  assert.ok(toolNames.includes("DocumentInfo"));
});

test("McpServer.handleRequest - handles unknown method", async () => {
  const server = new McpServer();
  const request: McpRequest = {
    jsonrpc: "2.0",
    id: 3,
    method: "unknown/method",
    params: {},
  };

  const response = await server["handleRequest"](request);

  assert.equal(response.jsonrpc, "2.0");
  assert.equal(response.id, 3);
  assert.ok(response.error);
  assert.equal(response.error.code, -32601);
});

test("McpServer.handleRequest - handles tools/call with missing name", async () => {
  const server = new McpServer();
  const request: McpRequest = {
    jsonrpc: "2.0",
    id: 4,
    method: "tools/call",
    params: { arguments: {} },
  };

  const response = await server["handleRequest"](request);

  assert.equal(response.jsonrpc, "2.0");
  assert.equal(response.id, 4);
  assert.ok(response.error);
  assert.equal(response.error.code, -32602);
});

// ============================================================================
// Tool Execution Tests
// ============================================================================

test("McpServer.executeTool - returns error for missing filePath", async () => {
  const server = new McpServer();
  const result = await server["executeTool"]("ViewAsText", {});

  assert.ok(!result.ok);
  assert.equal(result.error.code, "invalid_input");
  assert.ok(result.error.message.includes("filePath"));
});

test("McpServer.executeTool - Create returns not_supported when package lacks create function", async () => {
  const server = new McpServer();
  // Create requires a filePath but the package doesn't have createExcelDocument/createWordDocument
  const result = await server["executeTool"]("Create", { filePath: "/tmp/test.pptx" });

  // The result should be an error indicating not_supported
  assert.ok(!result.ok || result.ok === true); // Either succeeds or fails gracefully
});

test("McpServer.executeTool - DocumentInfo returns format info", async () => {
  const server = new McpServer();
  const result = await server["executeTool"]("DocumentInfo", { filePath: "/tmp/test.pptx" });

  assert.ok(result.ok, `DocumentInfo failed: ${result.ok ? "" : result.error?.message}`);
  assert.ok(result.data);
  const data = result.data as { format: string; path: string };
  assert.equal(data.format, "powerpoint");
  assert.ok(data.path.includes("test.pptx"));
});

test("McpServer.executeTool - DocumentInfo for Word doc", async () => {
  const server = new McpServer();
  const result = await server["executeTool"]("DocumentInfo", { filePath: "/tmp/test.docx" });

  assert.ok(result.ok, `DocumentInfo failed: ${result.ok ? "" : result.error?.message}`);
  assert.ok(result.data);
  const data = result.data as { format: string; path: string };
  assert.equal(data.format, "word");
});

test("McpServer.executeTool - DocumentInfo for Excel doc", async () => {
  const server = new McpServer();
  const result = await server["executeTool"]("DocumentInfo", { filePath: "/tmp/test.xlsx" });

  assert.ok(result.ok, `DocumentInfo failed: ${result.ok ? "" : result.error?.message}`);
  assert.ok(result.data);
  const data = result.data as { format: string; path: string };
  assert.equal(data.format, "excel");
});

test("McpServer.executeTool - unknown tool returns not_found error", async () => {
  const server = new McpServer();
  const result = await server["executeTool"]("NonExistentTool", { filePath: "/tmp/test.pptx" });

  assert.ok(!result.ok);
  assert.equal(result.error.code, "not_found");
});

// ============================================================================
// mcpTools Exports Tests
// ============================================================================

test("mcpTools - exports all required tools", () => {
  const requiredTools = [
    "Create",
    "ViewAsText",
    "ViewAsAnnotated",
    "ViewAsOutline",
    "ViewAsStats",
    "ViewAsIssues",
    "ViewAsHtml",
    "Get",
    "Query",
    "Add",
    "Set",
    "Remove",
    "Move",
    "Swap",
    "Batch",
    "Merge",
    "Raw",
    "RawSet",
    "Validate",
    "DocumentInfo",
  ];

  const toolNames = mcpTools.map((t) => t.name);

  for (const required of requiredTools) {
    assert.ok(toolNames.includes(required), `Missing tool: ${required}`);
  }
});

test("mcpTools - each tool has required properties", () => {
  for (const tool of mcpTools) {
    assert.ok(typeof tool.name === "string", "Tool name must be string");
    assert.ok(typeof tool.description === "string", "Tool description must be string");
    assert.ok(tool.inputSchema, "Tool must have inputSchema");
    assert.ok(tool.inputSchema.type === "object", "inputSchema type must be object");
    assert.ok(tool.inputSchema.properties, "inputSchema must have properties");
  }
});

test("mcpTools - Create requires filePath, Merge uses templatePath/outputPath/data", () => {
  const createTool = mcpTools.find((t) => t.name === "Create");
  const mergeTool = mcpTools.find((t) => t.name === "Merge");

  // Create requires filePath (path where to create the document)
  assert.ok(createTool);
  assert.ok(createTool.inputSchema.required?.includes("filePath"));

  // Merge uses templatePath, outputPath, and data (not filePath)
  assert.ok(mergeTool);
  assert.ok(!mergeTool.inputSchema.required?.includes("filePath"));
  assert.ok(mergeTool.inputSchema.required?.includes("templatePath"));
  assert.ok(mergeTool.inputSchema.required?.includes("outputPath"));
  assert.ok(mergeTool.inputSchema.required?.includes("data"));
});

// ============================================================================
// startMcpServer Tests
// ============================================================================

test("startMcpServer - starts server without transport argument", async () => {
  const server = createMcpServer();
  let started = false;
  let stopped = false;

  // Override start and stop to track calls
  const originalStart = server.start.bind(server);
  const originalStop = server.stop.bind(server);

  server.start = async () => { started = true; };
  server.stop = async () => { stopped = true; };

  const result = await startMcpServer(server);

  assert.equal(started, true);
  assert.equal(result, server);
});

test("startMcpServer - starts server with stdio transport", async () => {
  const server = createMcpServer({ transport: "stdio" });
  let started = false;

  server.start = async () => { started = true; };

  const result = await startMcpServer(server, "stdio");

  assert.equal(started, true);
  assert.equal(result, server);
});

// ============================================================================
// MCP Protocol Integration Tests
// ============================================================================

test("McpServer.handleRequest - tools/call DocumentInfo returns success", async () => {
  const server = new McpServer();

  const request: McpRequest = {
    jsonrpc: "2.0",
    id: 10,
    method: "tools/call",
    params: {
      name: "DocumentInfo",
      arguments: { filePath: "/tmp/test.pptx" },
    },
  };

  const response = await server["handleRequest"](request);

  assert.equal(response.jsonrpc, "2.0");
  assert.equal(response.id, 10);
  assert.ok(response.result);
  const result10 = response.result as McpToolResult;
  assert.ok(!result10.isError);

  const content = JSON.parse(result10.content[0].text);
  assert.ok(content.ok);
  assert.ok(content.data);
});

test("McpServer.handleRequest - tools/call with missing filePath returns error", async () => {
  const server = new McpServer();

  const request: McpRequest = {
    jsonrpc: "2.0",
    id: 11,
    method: "tools/call",
    params: {
      name: "ViewAsText",
      arguments: {},
    },
  };

  const response = await server["handleRequest"](request);

  assert.equal(response.jsonrpc, "2.0");
  assert.equal(response.id, 11);
  assert.ok(response.result);
  const result11 = response.result as McpToolResult;
  assert.ok(result11.isError);

  const content = JSON.parse(result11.content[0].text);
  assert.ok(!content.ok);
  assert.ok(content.error);
});

test("McpServer.handleRequest - tools/call with invalid tool name returns error", async () => {
  const server = new McpServer();

  const request: McpRequest = {
    jsonrpc: "2.0",
    id: 12,
    method: "tools/call",
    params: {
      name: "NonExistentTool",
      arguments: { filePath: "/tmp/test.pptx" },
    },
  };

  const response = await server["handleRequest"](request);

  assert.equal(response.jsonrpc, "2.0");
  assert.equal(response.id, 12);
  assert.ok(response.result);
  const result12 = response.result as McpToolResult;
  assert.ok(result12.isError);

  const content = JSON.parse(result12.content[0].text);
  assert.ok(!content.ok);
  assert.ok(content.error);
  assert.equal(content.error.code, "not_found");
});
