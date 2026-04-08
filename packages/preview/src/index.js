import fs from "node:fs";
import http from "node:http";

const DEFAULT_WAITING_MESSAGE = "Waiting for first update...";
const PREVIEW_PORT_MIN = 20000;
const PREVIEW_PORT_SPAN = 20000;

/**
 * Build a lightweight preview HTML shell that can be swapped live by SSE updates.
 * @param {object} options
 * @param {string} [options.title]
 * @param {string} [options.bodyHtml]
 * @returns {string}
 */
export function buildPreviewHtml({
  title = "officekit preview",
  bodyHtml = `<div class="waiting">${DEFAULT_WAITING_MESSAGE}</div>`,
} = {}) {
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${escapeHtml(title)}</title>
    <style>
      :root { color-scheme: light dark; }
      html, body { margin: 0; min-height: 100%; font-family: Inter, system-ui, sans-serif; background: #0b1020; color: #f6f8fc; }
      body { display: grid; grid-template-rows: auto 1fr; }
      header { padding: 12px 16px; border-bottom: 1px solid rgba(255,255,255,0.12); background: rgba(9,13,28,0.92); position: sticky; top: 0; }
      main { padding: 20px; }
      .waiting { display: grid; place-items: center; min-height: 50vh; opacity: 0.75; border: 1px dashed rgba(255,255,255,0.18); border-radius: 16px; }
      .meta { font-size: 12px; opacity: 0.72; }
    </style>
  </head>
  <body>
    <header>
      <strong>${escapeHtml(title)}</strong>
      <div class="meta">Live preview powered by officekit/packages/preview</div>
    </header>
    <main id="preview-root">${bodyHtml}</main>
    <script>
      (() => {
        const root = document.getElementById("preview-root");
        const source = new EventSource("/events");
        source.addEventListener("update", (event) => {
          const payload = JSON.parse(event.data);
          if (typeof payload.html === "string") {
            root.innerHTML = payload.html;
          }
          if (payload.scrollTo) {
            const target = document.querySelector(payload.scrollTo);
            if (target) target.scrollIntoView({ behavior: "smooth", block: "center" });
          }
        });
      })();
    </script>
  </body>
</html>`;
}

/**
 * @param {object} options
 * @param {number} [options.port]
 * @param {string} [options.initialHtml]
 * @returns {Promise<{port:number,url:string,publish:(message: PreviewMessageInput) => PreviewState, close:() => Promise<void>, snapshot:() => PreviewState}>}
 */
export async function startPreviewServer({ port = 0, initialHtml } = {}) {
  /** @type {PreviewState} */
  let state = {
    version: 0,
    html: initialHtml ?? buildPreviewHtml(),
    lastAction: "full",
  };

  if (typeof Bun !== "undefined" && typeof Bun.serve === "function") {
    /** @type {Set<ReadableStreamDefaultController<Uint8Array>>} */
    const clients = new Set();
    const encoder = new TextEncoder();
    const server = startBunPreviewServer(port, async (request) => {
      const url = new URL(request.url);

      if (request.method === "GET" && url.pathname === "/") {
        return new Response(state.html, {
          headers: { "content-type": "text/html; charset=utf-8" },
        });
      }

      if (request.method === "GET" && url.pathname === "/health") {
        return Response.json({ ok: true, version: state.version, clients: clients.size });
      }

      if (request.method === "GET" && url.pathname === "/events") {
        const stream = new ReadableStream({
          start(controller) {
            clients.add(controller);
            controller.enqueue(encoder.encode(`event: update\ndata: ${JSON.stringify({ version: state.version, html: extractBodyHtml(state.html) })}\n\n`));
          },
          cancel() {
            for (const controller of clients) {
              if (controller.desiredSize === null) {
                clients.delete(controller);
              }
            }
          },
        });

        return new Response(stream, {
          headers: {
            "cache-control": "no-cache, no-transform",
            connection: "keep-alive",
            "content-type": "text/event-stream; charset=utf-8",
          },
        });
      }

      if (request.method === "POST" && url.pathname === "/message") {
        const body = await request.text();
        const payload = body.length > 0 ? JSON.parse(body) : {};
        return Response.json(publish(payload));
      }

      return Response.json({ ok: false, error: "not_found" }, { status: 404 });
    });

    function publish(message = {}) {
      state = {
        version: typeof message.version === "number" ? message.version : state.version + 1,
        html: message.fullHtml ?? wrapBodyHtml(message.html ?? extractBodyHtml(state.html)),
        scrollTo: message.scrollTo,
        lastAction: message.action ?? "full",
      };

      const event = encoder.encode(`event: update\ndata: ${JSON.stringify({
        version: state.version,
        action: state.lastAction,
        html: extractBodyHtml(state.html),
        scrollTo: state.scrollTo,
      })}\n\n`);

      for (const controller of clients) {
        try {
          controller.enqueue(event);
        } catch {
          clients.delete(controller);
        }
      }

      return state;
    }

    return {
      port: server.port,
      url: `http://127.0.0.1:${server.port}`,
      publish,
      snapshot: () => state,
      close: async () => {
        for (const controller of clients) {
          try {
            controller.close();
          } catch {
            // Ignore already-closed streams.
          }
        }
        clients.clear();
        server.stop(true);
      },
    };
  }

  /** @type {Set<import("node:http").ServerResponse>} */
  const clients = new Set();

  const server = http.createServer(async (request, response) => {
    const { method = "GET", url = "/" } = request;

    if (method === "GET" && url === "/") {
      response.writeHead(200, { "content-type": "text/html; charset=utf-8" });
      response.end(state.html);
      return;
    }

    if (method === "GET" && url === "/health") {
      response.writeHead(200, { "content-type": "application/json; charset=utf-8" });
      response.end(JSON.stringify({ ok: true, version: state.version, clients: clients.size }));
      return;
    }

    if (method === "GET" && url === "/events") {
      response.writeHead(200, {
        "cache-control": "no-cache, no-transform",
        connection: "keep-alive",
        "content-type": "text/event-stream; charset=utf-8",
      });
      response.write(`event: update\ndata: ${JSON.stringify({ version: state.version, html: extractBodyHtml(state.html) })}\n\n`);
      clients.add(response);
      request.on("close", () => clients.delete(response));
      return;
    }

    if (method === "POST" && url === "/message") {
      const body = await readRequestBody(request);
      const payload = body.length > 0 ? JSON.parse(body) : {};
      const nextState = publish(payload);
      response.writeHead(200, { "content-type": "application/json; charset=utf-8" });
      response.end(JSON.stringify(nextState));
      return;
    }

    response.writeHead(404, { "content-type": "application/json; charset=utf-8" });
    response.end(JSON.stringify({ ok: false, error: "not_found" }));
  });

  await listenPreviewServer(server, port);

  const address = server.address();
  if (!address || typeof address === "string") {
    throw new Error("Unable to determine preview server address.");
  }

  function publish(message = {}) {
    state = {
      version: typeof message.version === "number" ? message.version : state.version + 1,
      html: message.fullHtml ?? wrapBodyHtml(message.html ?? extractBodyHtml(state.html)),
      scrollTo: message.scrollTo,
      lastAction: message.action ?? "full",
    };

    const event = `event: update\ndata: ${JSON.stringify({
      version: state.version,
      action: state.lastAction,
      html: extractBodyHtml(state.html),
      scrollTo: state.scrollTo,
    })}\n\n`;

    for (const client of clients) {
      client.write(event);
    }

    return state;
  }

  return {
    port: address.port,
    url: `http://127.0.0.1:${address.port}`,
    publish,
    snapshot: () => state,
    close: async () => {
      for (const client of clients) {
        client.end();
      }
      clients.clear();
      await new Promise((resolve, reject) => server.close((error) => (error ? reject(error) : resolve())));
    },
  };
}

/**
 * @param {import("node:http").Server} server
 * @param {number} requestedPort
 * @returns {Promise<void>}
 */
async function listenPreviewServer(server, requestedPort) {
  const attempts = requestedPort === 0 ? 12 : 1;

  for (let attempt = 0; attempt < attempts; attempt += 1) {
    const candidatePort = requestedPort === 0 ? randomPreviewPort() : requestedPort;

    try {
      await new Promise((resolve, reject) => {
        const onError = (error) => {
          server.off("listening", onListening);
          reject(error);
        };
        const onListening = () => {
          server.off("error", onError);
          resolve();
        };

        server.once("error", onError);
        server.once("listening", onListening);
        server.listen(candidatePort, "127.0.0.1");
      });
      return;
    } catch (error) {
      if (!(requestedPort === 0 && error && typeof error === "object" && "code" in error && error.code === "EADDRINUSE")) {
        throw error;
      }
    }
  }

  throw new Error("Unable to find an open preview port.");
}

/**
 * @returns {number}
 */
function randomPreviewPort() {
  return PREVIEW_PORT_MIN + Math.floor(Math.random() * PREVIEW_PORT_SPAN);
}

/**
 * @param {number} requestedPort
 * @param {(request: Request) => Response | Promise<Response>} fetch
 * @returns {ReturnType<typeof Bun.serve>}
 */
function startBunPreviewServer(requestedPort, fetch) {
  const attempts = requestedPort === 0 ? 12 : 1;
  let lastError;

  for (let attempt = 0; attempt < attempts; attempt += 1) {
    const candidatePort = requestedPort === 0 ? randomPreviewPort() : requestedPort;
    try {
      return Bun.serve({
        port: candidatePort,
        hostname: "127.0.0.1",
        fetch,
      });
    } catch (error) {
      lastError = error;
      if (!(requestedPort === 0 && error && typeof error === "object" && "code" in error && error.code === "EADDRINUSE")) {
        throw error;
      }
    }
  }

  throw lastError ?? new Error("Unable to find an open preview port.");
}

/**
 * Start a preview session that renders once and re-renders when the target file changes.
 * The watcher never opens or interprets the office file itself; the caller supplies the render callback.
 *
 * @param {object} options
 * @param {string} options.filePath
 * @param {(filePath:string, trigger:"initial"|"change"|"manual") => Promise<string> | string} options.render
 * @param {number} [options.port]
 * @param {number} [options.debounceMs]
 */
export async function startPreviewSession({
  filePath,
  render,
  port = 0,
  debounceMs = 75,
}) {
  const server = await startPreviewServer({ port, initialHtml: buildPreviewHtml() });
  let timer = null;

  const rerender = async (trigger) => {
    const html = await render(filePath, trigger);
    server.publish({ action: "full", fullHtml: wrapBodyHtml(html) });
  };

  await rerender("initial");

  const watcher = fs.watch(filePath, { persistent: false }, () => {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => {
      rerender("change").catch(() => {});
    }, debounceMs);
  });

  return {
    ...server,
    refresh: () => rerender("manual"),
    close: async () => {
      watcher.close();
      if (timer) clearTimeout(timer);
      await server.close();
    },
  };
}

/**
 * @param {string} value
 * @returns {string}
 */
export function wrapBodyHtml(value) {
  return buildPreviewHtml({ bodyHtml: value });
}

/**
 * @param {string} html
 * @returns {string}
 */
export function extractBodyHtml(html) {
  const match = html.match(/<main id="preview-root">([\s\S]*)<\/main>/i);
  return match ? match[1] : html;
}

/**
 * @param {import("node:http").IncomingMessage} request
 * @returns {Promise<string>}
 */
function readRequestBody(request) {
  return new Promise((resolve, reject) => {
    let body = "";
    request.setEncoding("utf8");
    request.on("data", (chunk) => {
      body += chunk;
    });
    request.on("end", () => resolve(body));
    request.on("error", reject);
  });
}

/**
 * @param {string} value
 * @returns {string}
 */
function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

/**
 * @typedef {object} PreviewMessageInput
 * @property {"full"|"replace"|"add"|"remove"} [action]
 * @property {number} [version]
 * @property {string} [html]
 * @property {string} [fullHtml]
 * @property {string} [scrollTo]
 */

/**
 * @typedef {object} PreviewState
 * @property {number} version
 * @property {string} html
 * @property {string} [scrollTo]
 * @property {string} lastAction
 */
