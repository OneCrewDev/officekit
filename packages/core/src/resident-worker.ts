import { createServer } from "node:net";
import { rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";
import { loadDocument, persistDocument } from "./document-store.js";
import { removeSessionRecord, writeSessionRecord } from "./session-registry.js";

const filePath = process.argv[2];

if (!filePath) {
  console.error("resident worker requires <file>.");
  process.exit(1);
}

interface ResidentDocument {
  document: Awaited<ReturnType<typeof loadDocument>>;
  dirty: boolean;
}

let resident: ResidentDocument | null = null;
let listenPort: number | null = null;

function getListenPort(filePath: string): number {
  // Generate a port number based on file path hash
  // Use a range that's less likely to conflict (49152-65535)
  const hash = createHash("sha1").update(path.resolve(filePath)).digest("hex");
  const num = parseInt(hash.slice(0, 4), 16);
  return 49152 + (num % 16383);
}

async function main() {
  const doc = await loadDocument(filePath);
  resident = { document: doc, dirty: false };
  listenPort = getListenPort(filePath);

  // Create TCP server for IPC
  const server = createServer((socket) => {
    let buffer = "";

    socket.on("data", async (chunk) => {
      buffer += chunk.toString();
      const lines = buffer.split("\n");
      buffer = lines.pop() ?? "";

      for (const line of lines) {
        if (!line.trim()) continue;
        try {
          const request = JSON.parse(line);
          const response = await handleRequest(request);
          socket.write(JSON.stringify(response) + "\n");
        } catch (e) {
          const error = { id: null, ok: false, error: e instanceof Error ? e.message : String(e) };
          socket.write(JSON.stringify(error) + "\n");
        }
      }
    });

    socket.on("error", (err) => {
      console.error("Socket error:", err.message);
    });
  });

  await new Promise<void>((resolve, reject) => {
    server.listen(listenPort!, "127.0.0.1", (err?: Error) => {
      if (err) reject(err);
      else resolve();
    });
  });

  // Write session record with port
  await writeSessionRecord("resident", filePath, {
    kind: "resident",
    filePath,
    pid: process.pid,
    startedAt: new Date().toISOString(),
    format: doc.format,
    socketPath: `tcp://127.0.0.1:${listenPort}`,
  });

  console.error(`Resident worker ready: ${filePath} (port: ${listenPort})`);
}

interface Request {
  id: string;
  command: "get" | "set" | "remove" | "query" | "view" | "materialize" | "persist";
  targetPath?: string;
  options?: Record<string, unknown>;
  props?: Record<string, string>;
  type?: string;
  mode?: string;
}

async function handleRequest(request: Request) {
  if (!resident) {
    return { id: request.id, ok: false, error: "No resident document loaded" };
  }

  const { document } = resident;

  try {
    switch (request.command) {
      case "get": {
        const { materializePath } = await import("./document-store.js");
        const result = materializePath(document, request.targetPath ?? "/");
        return { id: request.id, ok: true, data: result };
      }

      case "materialize": {
        const { materializePath } = await import("./document-store.js");
        const result = materializePath(document, request.targetPath ?? "/");
        return { id: request.id, ok: true, data: result };
      }

      case "query": {
        const { queryDocumentNodes } = await import("./document-store.js");
        const result = await queryDocumentNodes(filePath, request.targetPath ?? "/");
        return { id: request.id, ok: true, data: result };
      }

      case "view": {
        const { viewDocument } = await import("./document-store.js");
        const result = await viewDocument(filePath, request.mode ?? "outline");
        return { id: request.id, ok: true, data: { output: result.output } };
      }

      case "persist": {
        if (resident.dirty) {
          await persistDocument(filePath, document);
          resident.dirty = false;
        }
        return { id: request.id, ok: true };
      }

      default:
        return { id: request.id, ok: false, error: `Unknown command: ${request.command}` };
    }
  } catch (e) {
    return { id: request.id, ok: false, error: e instanceof Error ? e.message : String(e) };
  }
}

async function shutdown(exitCode = 0) {
  // Persist if dirty
  if (resident?.dirty) {
    try {
      await persistDocument(filePath, resident.document);
    } catch (e) {
      console.error("Failed to persist on shutdown:", e);
    }
  }

  // Remove session record
  await removeSessionRecord("resident", filePath);

  resident = null;
  listenPort = null;

  process.exit(exitCode);
}

process.on("SIGINT", () => {
  shutdown(0).catch(() => process.exit(1));
});

process.on("SIGTERM", () => {
  shutdown(0).catch(() => process.exit(1));
});

main()
  .then(() => {
    // Keep running
  })
  .catch(async (error) => {
    console.error(error instanceof Error ? error.message : String(error));
    await removeSessionRecord("resident", filePath);
    process.exit(1);
  });
