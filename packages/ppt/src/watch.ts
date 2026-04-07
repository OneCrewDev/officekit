/**
 * Watch mode for @officekit/ppt.
 *
 * Monitors a PPTX file and serves live HTML previews with auto-refresh.
 * When the PPTX file changes on disk, the HTML preview is automatically regenerated
 * and pushed to all connected clients via Server-Sent Events (SSE).
 */

import fs from "node:fs";
import http from "node:http";
import path from "node:path";
import { err, ok, invalidInput } from "./result.js";
import { viewAsHtml } from "./preview-html.js";
import type { Result } from "./types.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Options for the watch function.
 */
export interface WatchOptions {
  /** Port to serve the preview server on (default: 0 = random available port) */
  port?: number;
  /** Whether to auto-open the browser (default: false) */
  autoOpen?: boolean;
}

/**
 * The watch server instance returned by the watch function.
 */
export interface WatchServer {
  /** URL where the preview server is running */
  url: string;
  /** Close the watch server and stop watching */
  close: () => Promise<void>;
}

/**
 * Internal state for the watch server.
 */
interface ServerState {
  html: string;
  version: number;
  clients: Set<http.ServerResponse>;
}

// ============================================================================
// Constants
// ============================================================================

/** Default debounce delay in milliseconds */
const DEFAULT_DEBOUNCE_MS = 300;

/** Default port (0 = random available) */
const DEFAULT_PORT = 0;

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Wraps body HTML in the preview shell document.
 */
function wrapBodyHtml(bodyHtml: string, title: string): string {
  const escapeHtml = (value: string): string =>
    value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(title)}</title>
  <style>
    :root { color-scheme: light dark; }
    html, body { margin: 0; min-height: 100vh; font-family: Inter, system-ui, sans-serif; background: #0b1020; color: #e2e8f0; }
    body { display: grid; grid-template-rows: auto 1fr; }
    header { padding: 12px 16px; border-bottom: 1px solid rgba(255,255,255,0.12); background: rgba(9,13,28,0.92); position: sticky; top: 0; display: flex; align-items: center; gap: 12px; }
    header strong { font-size: 14px; }
    .meta { font-size: 11px; opacity: 0.6; }
    main { padding: 20px; }
    .waiting { display: grid; place-items: center; min-height: 50vh; opacity: 0.75; border: 1px dashed rgba(255,255,255,0.18); border-radius: 16px; }
    .slide-container { background: #fff; border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3); margin: 20px auto; overflow: hidden; max-width: 100%; }
    .slide { position: relative; }
    .slide-label { background: #333; color: #fff; padding: 8px 16px; font-size: 12px; font-weight: 500; }
    .shape { box-sizing: border-box; }
    .shape-text { padding: 4pt; line-height: 1.2; color: #000; }
    .table { overflow: hidden; }
    .table table { font-size: 10pt; text-align: left; border-collapse: collapse; }
    .picture { text-align: center; }
  </style>
</head>
<body>
  <header>
    <strong>${escapeHtml(title)}</strong>
    <span class="meta">Live preview (watching for changes)</span>
  </header>
  <main id="preview-root">
    ${bodyHtml}
  </main>
  <script>
    (function() {
      var source = new EventSource('/events');
      source.addEventListener('update', function(event) {
        var data = JSON.parse(event.data);
        var root = document.getElementById('preview-root');
        if (root && data.html) {
          root.innerHTML = data.html;
        }
        if (data.scrollTo) {
          var target = document.querySelector(data.scrollTo);
          if (target) target.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      });
    })();
  </script>
</body>
</html>`;
}

/**
 * Extracts the body content from the full HTML document.
 */
function extractBodyHtml(fullHtml: string): string {
  const match = fullHtml.match(/<main id="preview-root">([\s\S]*)<\/main>\s*<\/body>\s*<\/html>/i);
  return match ? match[1] : fullHtml;
}

/**
 * Creates a waiting state HTML.
 */
function createWaitingHtml(): string {
  return '<div class="waiting"><p>Waiting for first render...</p></div>';
}

// ============================================================================
// Watch Function
// ============================================================================

/**
 * Starts watching a PPTX file and serves live HTML previews.
 *
 * @param filePath - Path to the PPTX file to watch
 * @param options - Watch options (port, autoOpen)
 * @returns Result containing the WatchServer on success
 *
 * @example
 * // Start watching a presentation
 * const result = await watch("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(`Preview server running at ${result.data.url}`);
 *   // When done:
 *   await result.data.close();
 * }
 *
 * @example
 * // Start on a specific port
 * const result = await watch("/path/to/presentation.pptx", { port: 3000 });
 */
export async function watch(
  filePath: string,
  options?: WatchOptions,
): Promise<Result<WatchServer>> {
  // Validate filePath
  if (!filePath || typeof filePath !== "string") {
    return invalidInput("filePath must be a non-empty string");
  }

  // Check if file exists initially
  try {
    await fs.promises.access(filePath, fs.constants.R_OK);
  } catch {
    return invalidInput(`File not found or not readable: ${filePath}`);
  }

  const port = options?.port ?? DEFAULT_PORT;
  const debounceMs = DEFAULT_DEBOUNCE_MS;

  // State
  const clients = new Set<http.ServerResponse>();
  let state: ServerState = {
    html: wrapBodyHtml(createWaitingHtml(), path.basename(filePath)),
    version: 0,
    clients,
  };
  let watcher: fs.FSWatcher | null = null;
  let debounceTimer: ReturnType<typeof setTimeout> | null = null;
  let server: http.Server | null = null;
  let actualPort = 0;

  /**
   * Regenerates the HTML preview and notifies all clients.
   */
  async function regenerateHtml(trigger: "initial" | "change" | "manual"): Promise<void> {
    const htmlResult = await viewAsHtml(filePath);
    if (!htmlResult.ok) {
      // If we can't generate HTML, show error state
      state = {
        ...state,
        html: wrapBodyHtml(
          `<div class="waiting"><p>Error: ${htmlResult.error?.message ?? "Failed to render"}</p></div>`,
          path.basename(filePath)
        ),
        version: state.version + 1,
      };
    } else {
      // Extract just the slide content (not the full document)
      const slideHtml = extractBodyHtml(htmlResult.data!.html);
      state = {
        ...state,
        html: wrapBodyHtml(slideHtml, path.basename(filePath)),
        version: state.version + 1,
      };
    }

    // Notify all connected clients
    const event = `event: update\ndata: ${JSON.stringify({
      version: state.version,
      action: trigger,
      html: extractBodyHtml(state.html),
    })}\n\n`;

    for (const client of clients) {
      client.write(event);
    }
  }

  // Initial render
  await regenerateHtml("initial");

  // Create HTTP server
  return new Promise((resolve, reject) => {
    server = http.createServer(async (request, response) => {
      const { method = "GET", url = "/" } = request;

      // Root path - serve the HTML
      if (method === "GET" && url === "/") {
        response.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
        response.end(state.html);
        return;
      }

      // Health check endpoint
      if (method === "GET" && url === "/health") {
        response.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        response.end(JSON.stringify({
          ok: true,
          version: state.version,
          clients: clients.size,
          watching: filePath,
        }));
        return;
      }

      // SSE events endpoint
      if (method === "GET" && url === "/events") {
        response.writeHead(200, {
          "Cache-Control": "no-cache, no-transform",
          "Connection": "keep-alive",
          "Content-Type": "text/event-stream; charset=utf-8",
        });
        // Send initial state
        response.write(`event: update\ndata: ${JSON.stringify({
          version: state.version,
          action: "initial",
          html: extractBodyHtml(state.html),
        })}\n\n`);
        clients.add(response);

        request.on("close", () => {
          clients.delete(response);
        });
        return;
      }

      // 404 for unknown routes
      response.writeHead(404, { "Content-Type": "application/json; charset=utf-8" });
      response.end(JSON.stringify({ ok: false, error: "not_found" }));
    });

    server.on("error", (err) => {
      reject(err);
    });

    server.listen(port, "127.0.0.1", () => {
      const address = server!.address();
      if (!address || typeof address === "string") {
        reject(new Error("Unable to determine server address"));
        return;
      }
      actualPort = address.port;

      // Start watching the file
      watcher = fs.watch(filePath, { persistent: false }, (eventType) => {
        if (eventType === "change") {
          // Debounce rapid changes
          if (debounceTimer) clearTimeout(debounceTimer);
          debounceTimer = setTimeout(() => {
            regenerateHtml("change").catch(() => {});
          }, debounceMs);
        }
      });

      resolve(ok({
        url: `http://127.0.0.1:${actualPort}`,
        close: async () => {
          // Clean up
          if (debounceTimer) clearTimeout(debounceTimer);
          if (watcher) watcher!.close();
          if (server) {
            for (const client of clients) {
              client.end();
            }
            clients.clear();
            await new Promise<void>((res, rej) => {
              server!.close((err) => err ? rej(err) : res());
            });
          }
        },
      }));
    });
  });
}
