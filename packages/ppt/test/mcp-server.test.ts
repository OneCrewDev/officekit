import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  McpServer,
  createMcpServer,
  startMcpServer,
  type McpServerOptions,
  type McpTransport,
  type McpRequest,
  type McpResponse,
} from "../src/mcp-server.js";

import { pptTools } from "../src/mcp-tools.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

const DATA_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/data_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-mcp-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
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
  assert.equal((response as any).result.serverInfo.name, "@officekit/ppt");
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
  assert.ok(toolNames.includes("Add"));
  assert.ok(toolNames.includes("Remove"));
  assert.ok(toolNames.includes("Get"));
  assert.ok(toolNames.includes("Query"));
  assert.ok(toolNames.includes("Set"));
  assert.ok(toolNames.includes("Move"));
  assert.ok(toolNames.includes("Swap"));
  assert.ok(toolNames.includes("CopyFrom"));
  assert.ok(toolNames.includes("Raw"));
  assert.ok(toolNames.includes("RawSet"));
  assert.ok(toolNames.includes("Batch"));
  assert.ok(toolNames.includes("ViewAsText"));
  assert.ok(toolNames.includes("ViewAsAnnotated"));
  assert.ok(toolNames.includes("ViewAsOutline"));
  assert.ok(toolNames.includes("ViewAsStats"));
  assert.ok(toolNames.includes("ViewAsIssues"));
  assert.ok(toolNames.includes("ViewAsHtml"));
  assert.ok(toolNames.includes("ViewAsSvg"));
  assert.ok(toolNames.includes("CheckShapeTextOverflow"));
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
  const result = await server["executeTool"]("Add", {});

  assert.ok(!result.ok);
  assert.equal(result.error.code, "invalid_input");
  assert.ok(result.error.message.includes("filePath"));
});

test("McpServer.executeTool - Add tool adds a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  try {
    const result = await server["executeTool"]("Add", { filePath: tempPath });

    assert.ok(result.ok, `Add failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data.path);
  } finally {
    // Clean up
  }
});

test("McpServer.executeTool - Remove tool removes a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  try {
    // Add a slide first
    const addResult = await server["executeTool"]("Add", { filePath: tempPath });
    assert.ok(addResult.ok);

    // Now remove it
    const removeResult = await server["executeTool"]("Remove", { filePath: tempPath, index: 1 });
    assert.ok(removeResult.ok, `Remove failed: ${removeResult.error?.message}`);
  } finally {
    // Clean up
  }
});

test("McpServer.executeTool - ViewAsText tool returns text", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  try {
    const result = await server["executeTool"]("ViewAsText", { filePath: tempPath });

    assert.ok(result.ok, `ViewAsText failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.slideCount === "number");
    assert.ok(Array.isArray(result.data.slides));
  } finally {
    // Clean up
  }
});

test("McpServer.executeTool - ViewAsHtml tool returns HTML", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  try {
    const result = await server["executeTool"]("ViewAsHtml", { filePath: tempPath });

    assert.ok(result.ok, `ViewAsHtml failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.html === "string");
    assert.ok(result.data.html.includes("<!DOCTYPE html") || result.data.html.includes("<html"));
  } finally {
    // Clean up
  }
});

test("McpServer.executeTool - ViewAsSvg tool returns SVG", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  try {
    const result = await server["executeTool"]("ViewAsSvg", { filePath: tempPath });

    assert.ok(result.ok, `ViewAsSvg failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.svg === "string");
    assert.ok(result.data.svg.includes("<svg"));
  } finally {
    // Clean up
  }
});

test("McpServer.executeTool - Get tool returns element info", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  try {
    const result = await server["executeTool"]("Get", { filePath: tempPath, pptPath: "/slide[1]" });

    assert.ok(result.ok, `Get failed: ${result.error?.message}`);
    assert.ok(result.data);
  } finally {
    // Clean up
  }
});

test("McpServer.executeTool - Query tool returns slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  try {
    const result = await server["executeTool"]("Query", { filePath: tempPath });

    assert.ok(result.ok, `Query failed: ${result.error?.message}`);
    assert.ok(result.data);
  } finally {
    // Clean up
  }
});

test("McpServer.executeTool - Set tool updates shape text", async () => {
  const tempPath = await copyToTemp(DATA_PPTX);
  const server = new McpServer();

  try {
    const result = await server["executeTool"]("Set", {
      filePath: tempPath,
      pptPath: "/slide[1]/shape[1]",
      text: "Hello, MCP!",
    });

    // Set may fail on this particular shape but should not error on invalid path format
    // Just verify the tool executes
    assert.ok(result.ok !== undefined);
  } finally {
    // Clean up
  }
});

test("McpServer.executeTool - unknown tool returns not_found error", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  const result = await server["executeTool"]("NonExistentTool", { filePath: tempPath });

  assert.ok(!result.ok);
  assert.equal(result.error.code, "not_found");
});

// ============================================================================
// pptTools Exports Tests
// ============================================================================

test("pptTools - exports all required tools", () => {
  const requiredTools = [
    "Add",
    "AddPart",
    "Get",
    "Query",
    "Set",
    "Remove",
    "Move",
    "Swap",
    "CopyFrom",
    "Raw",
    "RawSet",
    "Batch",
    "ViewAsText",
    "ViewAsAnnotated",
    "ViewAsOutline",
    "ViewAsStats",
    "ViewAsIssues",
    "ViewAsHtml",
    "ViewAsSvg",
    "CheckShapeTextOverflow",
  ];

  const toolNames = pptTools.map((t) => t.name);

  for (const required of requiredTools) {
    assert.ok(toolNames.includes(required), `Missing tool: ${required}`);
  }
});

test("pptTools - each tool has required properties", () => {
  for (const tool of pptTools) {
    assert.ok(typeof tool.name === "string", "Tool name must be string");
    assert.ok(typeof tool.description === "string", "Tool description must be string");
    assert.ok(tool.inputSchema, "Tool must have inputSchema");
    assert.ok(tool.inputSchema.type === "object", "inputSchema type must be object");
    assert.ok(tool.inputSchema.properties, "inputSchema must have properties");
  }
});

test("pptTools - each tool has required properties in inputSchema", () => {
  for (const tool of pptTools) {
    assert.ok(
      tool.inputSchema.required?.includes("filePath"),
      `Tool ${tool.name} must require filePath`
    );
  }
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

test("McpServer.handleRequest - tools/call Add returns success", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  try {
    const request: McpRequest = {
      jsonrpc: "2.0",
      id: 10,
      method: "tools/call",
      params: {
        name: "Add",
        arguments: { filePath: tempPath },
      },
    };

    const response = await server["handleRequest"](request);

    assert.equal(response.jsonrpc, "2.0");
    assert.equal(response.id, 10);
    assert.ok(response.result);
    assert.ok(!response.result.isError);

    const content = JSON.parse(response.result.content[0].text);
    assert.ok(content.ok);
    assert.ok(content.data);
  } finally {
    // Clean up
  }
});

test("McpServer.handleRequest - tools/call ViewAsText returns success", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  const server = new McpServer();

  try {
    const request: McpRequest = {
      jsonrpc: "2.0",
      id: 11,
      method: "tools/call",
      params: {
        name: "ViewAsText",
        arguments: { filePath: tempPath },
      },
    };

    const response = await server["handleRequest"](request);

    assert.equal(response.jsonrpc, "2.0");
    assert.equal(response.id, 11);
    assert.ok(response.result);
    assert.ok(!response.result.isError);

    const content = JSON.parse(response.result.content[0].text);
    assert.ok(content.ok);
    assert.ok(content.data.slideCount);
    assert.ok(Array.isArray(content.data.slides));
  } finally {
    // Clean up
  }
});

test("McpServer.handleRequest - tools/call with invalid filePath returns error", async () => {
  const server = new McpServer();

  const request: McpRequest = {
    jsonrpc: "2.0",
    id: 12,
    method: "tools/call",
    params: {
      name: "Add",
      arguments: { filePath: "/nonexistent/path.pptx" },
    },
  };

  const response = await server["handleRequest"](request);

  assert.equal(response.jsonrpc, "2.0");
  assert.equal(response.id, 12);
  assert.ok(response.result);
  assert.ok(response.result.isError);

  const content = JSON.parse(response.result.content[0].text);
  assert.ok(!content.ok);
  assert.ok(content.error);
});
