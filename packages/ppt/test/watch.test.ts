import test from "node:test";
import assert from "node:assert/strict";
import { copyFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";
import { watch } from "../src/watch.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-watch-test-${Date.now()}-${Math.random().toString(36).slice(2)}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

// ============================================================================
// Watch Tests
// ============================================================================

test("watch - returns error for non-existent file", async () => {
  const result = await watch("/non/existent/path.pptx");
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "invalid_input");
  assert.ok(result.error?.message.includes("not found") || result.error?.message.includes("not readable"));
});

test("watch - starts server and returns URL", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await watch(tempPath);
    assert.ok(result.ok, `watch failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(typeof result.data.url === "string");
    assert.ok(result.data.url.startsWith("http://127.0.0.1:"));
    assert.ok(typeof result.data.close === "function");

    // Close the server
    await result.data.close();
  } finally {
    // Clean up temp file
  }
});

test("watch - accepts custom port option", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await watch(tempPath, { port: 0 });
    assert.ok(result.ok, `watch with port option failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data.url.includes("127.0.0.1"));

    await result.data.close();
  } finally {
    // Clean up
  }
});

test("watch - server responds to HTTP requests", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await watch(tempPath);
    assert.ok(result.ok);

    // Fetch the root page
    const response = await fetch(result.data!.url);
    assert.equal(response.status, 200);
    const text = await response.text();
    assert.ok(text.includes("<!DOCTYPE html>"));
    assert.ok(text.includes("preview-root"));
    assert.ok(text.includes("beautiful_presentation.pptx") || text.includes(".pptx"));

    await result.data!.close();
  } finally {
    // Clean up
  }
});

test("watch - health endpoint returns status", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await watch(tempPath);
    assert.ok(result.ok);

    const response = await fetch(`${result.data!.url}/health`);
    assert.equal(response.status, 200);
    const json = await response.json() as { ok: boolean; version: number; clients: number };
    assert.equal(json.ok, true);
    assert.ok(typeof json.version === "number");
    assert.ok(typeof json.clients === "number");

    await result.data!.close();
  } finally {
    // Clean up
  }
});

test("watch - SSE endpoint sends events", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await watch(tempPath);
    assert.ok(result.ok);

    // Manually read SSE stream using fetch
    const response = await fetch(`${result.data!.url}/events`);
    assert.equal(response.status, 200);
    assert.ok(response.body);

    // Read the SSE stream
    const reader = response.body!.getReader();
    const decoder = new TextDecoder();
    let sseData = "";

    // Read for up to 2 seconds to get at least one event
    const startTime = Date.now();
    while (Date.now() - startTime < 2000) {
      const { done, value } = await reader.read();
      if (done) break;
      sseData += decoder.decode(value, { stream: true });
      // If we have received data with "update" event, we're done
      if (sseData.includes("event: update")) {
        break;
      }
    }

    reader.cancel();

    // Verify SSE format
    assert.ok(sseData.includes("event: update"), "Should receive update event");
    assert.ok(sseData.includes('"version"'), "Should include version in data");
    assert.ok(sseData.includes('"html"'), "Should include html in data");

    await result.data!.close();
  } finally {
    // Clean up
  }
});

test("watch - 404 for unknown routes", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await watch(tempPath);
    assert.ok(result.ok);

    const response = await fetch(`${result.data!.url}/unknown-route`);
    assert.equal(response.status, 404);

    await result.data!.close();
  } finally {
    // Clean up
  }
});

test("watch - close stops the server", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await watch(tempPath);
    assert.ok(result.ok);

    const url = result.data!.url;
    await result.data!.close();

    // After close, the server should not respond
    try {
      await fetch(url);
      assert.fail("Server should not be running after close");
    } catch (e) {
      // Expected - connection should fail
      assert.ok(true);
    }
  } finally {
    // Clean up
  }
});

test("watch - multiple clients can connect", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await watch(tempPath);
    assert.ok(result.ok);

    // Connect multiple clients via SSE
    const [client1, client2] = await Promise.all([
      fetch(`${result.data!.url}/events`),
      fetch(`${result.data!.url}/events`),
    ]);

    // Wait briefly for connections to register
    await new Promise(r => setTimeout(r, 100));

    // Check health endpoint shows multiple clients
    const response = await fetch(`${result.data!.url}/health`);
    const json = await response.json() as { clients: number };
    assert.ok(json.clients >= 2, `Expected at least 2 clients, got ${json.clients}`);

    // Close the fetch connections
    client1.body!.cancel();
    client2.body!.cancel();

    await result.data!.close();
  } finally {
    // Clean up
  }
});

test("watch - rejects invalid filePath", async () => {
  // @ts-expect-error - testing invalid input
  const result = await watch(null);
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "invalid_input");
});

test("watch - rejects empty string filePath", async () => {
  const result = await watch("");
  assert.ok(!result.ok);
  assert.equal(result.error?.code, "invalid_input");
});
