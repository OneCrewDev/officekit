import { createDocument, addDocumentNode, addDocumentPart, checkDocument, getDocumentNode, getResidentDocument, hasResidentSession, importDelimitedData, loadDocument, mergeDocument, parseProps, queryDocumentNodes, queryResidentDocument, rawDocument, rawSetDocument, removeDocumentNode, renderDocumentHtml, setDocumentNode, viewDocument, viewResidentDocument, moveDocumentNode, swapDocumentNodes, copyDocumentNode, validateDocument } from "./document-store.js";
import { UnsupportedCapabilityError, UsageError } from "./errors.js";
import { summarizeParity } from "./parity.js";
import { readSessionRecord, removeSessionRecord, waitForSessionRecord } from "./session-registry.js";

export interface CommandResult {
  exitCode: number;
  stdout?: string;
  stderr?: string;
  waitUntilClose?: Promise<void>;
}

export async function executeCommand(argv: string[]): Promise<CommandResult> {
  const [command, ...rest] = argv;
  if (!command) {
    throw new UsageError("No command provided.");
  }

  if (command === "mcp") {
    throw new UnsupportedCapabilityError(command);
  }

  if (command === "create") {
    const filePath = rest[0];
    if (!filePath) throw new UsageError("create requires a target file path.");
    const created = await createDocument(filePath);
    return { exitCode: 0, stdout: JSON.stringify({ ok: true, created: created.filePath, format: created.format }, null, 2) };
  }

  if (command === "add") {
    const filePath = rest[0];
    const targetPath = rest[1];
    if (!filePath || !targetPath) throw new UsageError("add requires <file> <path>.");
    const parsed = parseProps(rest.slice(2));
    const result = await addDocumentNode(filePath, targetPath, parsed);
    return { exitCode: 0, stdout: parsed.json ? JSON.stringify(result, null, 2) : summarizeResult(result) };
  }

  if (command === "set") {
    const filePath = rest[0];
    const targetPath = rest[1];
    if (!filePath || !targetPath) throw new UsageError("set requires <file> <path>.");
    const parsed = parseProps(rest.slice(2));
    const result = await setDocumentNode(filePath, targetPath, parsed);
    return { exitCode: 0, stdout: parsed.json ? JSON.stringify(result, null, 2) : summarizeResult(result) };
  }

  if (command === "remove") {
    const filePath = rest[0];
    const targetPath = rest[1];
    if (!filePath || !targetPath) throw new UsageError("remove requires <file> <path>.");
    const result = await removeDocumentNode(filePath, targetPath);
    return { exitCode: 0, stdout: JSON.stringify(result, null, 2) };
  }

  if (command === "get" || command === "query") {
    const filePath = rest[0];
    const targetPath = rest[1] ?? "/";
    if (!filePath) throw new UsageError(`${command} requires <file> [path].`);
    const parsed = parseProps(rest.slice(2));

    // Try resident session first if available
    const hasResident = await hasResidentSession(filePath);
    let result: unknown;
    if (hasResident) {
      result = command === "query"
        ? await queryResidentDocument(filePath, targetPath)
        : await getResidentDocument(filePath, targetPath);
    } else {
      result = command === "query"
        ? await queryDocumentNodes(filePath, targetPath)
        : await getDocumentNode(filePath, targetPath);
    }
    return { exitCode: 0, stdout: parsed.json || command === "query" ? JSON.stringify(result, null, 2) : summarizeResult(result) };
  }

  if (command === "view") {
    const filePath = rest[0];
    const mode = rest[1] ?? "outline";
    if (!filePath) throw new UsageError("view requires <file> <mode?>.");

    // Try resident session first if available
    const hasResident = await hasResidentSession(filePath);
    let output: string;
    if (hasResident) {
      output = (await viewResidentDocument(filePath, mode)).output;
    } else {
      output = (await viewDocument(filePath, mode)).output;
    }
    return { exitCode: 0, stdout: output };
  }

  if (command === "raw") {
    const filePath = rest[0];
    if (!filePath) throw new UsageError("raw requires <file>.");
    const partPath = rest[1] && !rest[1].startsWith("--") ? rest[1] : undefined;
    let startRow: number | undefined;
    let endRow: number | undefined;
    let cols: string[] | undefined;
    for (let index = partPath ? 2 : 1; index < rest.length; index += 1) {
      const token = rest[index];
      if (token === "--start-row") {
        startRow = Number(rest[index + 1]);
        index += 1;
        continue;
      }
      if (token === "--end-row") {
        endRow = Number(rest[index + 1]);
        index += 1;
        continue;
      }
      if (token === "--cols") {
        cols = (rest[index + 1] ?? "")
          .split(",")
          .map((value) => value.trim().toUpperCase())
          .filter(Boolean);
        index += 1;
      }
    }
    return {
      exitCode: 0,
      stdout: await rawDocument(filePath, {
        ...(partPath ? { partPath } : {}),
        ...(startRow !== undefined ? { startRow } : {}),
        ...(endRow !== undefined ? { endRow } : {}),
        ...(cols ? { cols } : {}),
      }),
    };
  }

  if (command === "check") {
    const filePath = rest[0];
    if (!filePath) throw new UsageError("check requires <file>.");
    return { exitCode: 0, stdout: JSON.stringify(await checkDocument(filePath), null, 2) };
  }

  if (command === "validate") {
    const filePath = rest[0];
    if (!filePath) throw new UsageError("validate requires <file>.");
    const result = await validateDocument(filePath);
    return {
      exitCode: result.valid ? 0 : 1,
      stdout: JSON.stringify(result, null, 2),
    };
  }

  if (command === "import") {
    const filePath = rest[0];
    const parentPath = rest[1];
    if (!filePath || !parentPath) {
      throw new UsageError("import requires <file.xlsx> <parent-path> and a source file or --file.");
    }

    let sourceFile: string | undefined;
    let useStdin = false;
    let delimiter = ",";
    let hasHeader = false;
    let startCell = "A1";

    for (let index = 2; index < rest.length; index += 1) {
      const token = rest[index];
      if (token === "--file") {
        sourceFile = rest[index + 1];
        index += 1;
        continue;
      }
      if (token === "--format") {
        const format = (rest[index + 1] ?? "csv").toLowerCase();
        delimiter = format === "tsv" ? "\t" : ",";
        index += 1;
        continue;
      }
      if (token === "--stdin") {
        useStdin = true;
        continue;
      }
      if (token === "--header") {
        hasHeader = true;
        continue;
      }
      if (token === "--start-cell") {
        startCell = rest[index + 1] ?? "A1";
        index += 1;
        continue;
      }
      if (!token.startsWith("--") && !sourceFile) {
        sourceFile = token;
      }
    }

    const fs = await import("node:fs/promises");
    let content: string;
    if (useStdin) {
      content = await new Promise<string>((resolve, reject) => {
        let buffer = "";
        process.stdin.setEncoding("utf8");
        process.stdin.on("data", (chunk) => {
          buffer += chunk;
        });
        process.stdin.once("end", () => resolve(buffer));
        process.stdin.once("error", reject);
      });
    } else {
      if (!sourceFile) {
        throw new UsageError("import currently requires a source CSV/TSV file or --stdin.", "Use: officekit import book.xlsx /Sheet1 data.csv --format csv");
      }
      if (delimiter === "," && (sourceFile.endsWith(".tsv") || sourceFile.endsWith(".tab"))) {
        delimiter = "\t";
      }
      content = await fs.readFile(sourceFile, "utf8");
    }
    const result = await importDelimitedData(filePath, parentPath, content, {
      delimiter,
      hasHeader,
      startCell,
    });
    return { exitCode: 0, stdout: JSON.stringify(result, null, 2) };
  }

  if (command === "watch") {
    const filePath = rest[0];
    if (!filePath) throw new UsageError("watch requires <file>.");
    const portIndex = rest.indexOf("--port");
    const port = portIndex >= 0 ? Number(rest[portIndex + 1]) : 0;
    // @ts-expect-error workspace JS runtime module
    const preview = await import("../../preview/src/index.js");
    const session = await preview.startPreviewSession({
      filePath,
      port,
      render: async () => renderDocumentHtml(await loadDocument(filePath)),
    });
    const waitUntilClose = new Promise<void>((resolve) => {
      const shutdown = async () => {
        process.off("SIGINT", onSignal);
        process.off("SIGTERM", onSignal);
        await session.close();
        resolve();
      };
      const onSignal = () => {
        shutdown().finally(() => process.exit(0));
      };
      process.on("SIGINT", onSignal);
      process.on("SIGTERM", onSignal);
    });
    return {
      exitCode: 0,
      stdout: JSON.stringify({ ok: true, url: session.url, port: session.port }, null, 2),
      waitUntilClose,
    };
  }

  if (command === "move") {
    const filePath = rest[0];
    const sourcePath = rest[1];
    const targetPath = rest[2];
    if (!filePath || !sourcePath || !targetPath) {
      throw new UsageError("move requires <file> <source-path> <target-path>.");
    }
    let after: string | undefined;
    let before: string | undefined;
    for (let index = 3; index < rest.length; index += 1) {
      const token = rest[index];
      if (token === "--after") {
        after = rest[index + 1];
        index += 1;
        continue;
      }
      if (token === "--before") {
        before = rest[index + 1];
        index += 1;
        continue;
      }
    }
    const result = await moveDocumentNode(filePath, sourcePath, targetPath, { after, before });
    return { exitCode: 0, stdout: JSON.stringify(result, null, 2) };
  }

  if (command === "swap") {
    const filePath = rest[0];
    const path1 = rest[1];
    const path2 = rest[2];
    if (!filePath || !path1 || !path2) {
      throw new UsageError("swap requires <file> <path1> <path2>.");
    }
    const result = await swapDocumentNodes(filePath, path1, path2);
    return { exitCode: 0, stdout: JSON.stringify(result, null, 2) };
  }

  if (command === "copy") {
    const filePath = rest[0];
    const sourcePath = rest[1];
    const targetPath = rest[2];
    if (!filePath || !sourcePath || !targetPath) {
      throw new UsageError("copy requires <file> <source-path> <target-path>.");
    }
    let index: number | undefined;
    let after: string | undefined;
    let before: string | undefined;
    for (let i = 3; i < rest.length; i += 1) {
      const token = rest[i];
      if (token === "--index") {
        index = Number(rest[i + 1]);
        i += 1;
        continue;
      }
      if (token === "--after") {
        after = rest[i + 1];
        i += 1;
        continue;
      }
      if (token === "--before") {
        before = rest[i + 1];
        i += 1;
        continue;
      }
    }
    const result = await copyDocumentNode(filePath, sourcePath, targetPath, { index, after, before });
    return { exitCode: 0, stdout: JSON.stringify(result, null, 2) };
  }

  if (command === "batch") {
    const filePath = rest[0];
    if (!filePath) throw new UsageError("batch requires <file>.");
    let operationsJson = "";
    let useStdin = false;
    for (let index = 1; index < rest.length; index += 1) {
      const token = rest[index];
      if (token === "--stdin") {
        useStdin = true;
        continue;
      }
      if (!token.startsWith("--")) {
        operationsJson = token;
      }
    }
    let operations: Array<{ action: string; target: string; options?: Record<string, unknown> }>;
    if (useStdin) {
      operationsJson = await new Promise<string>((resolve, reject) => {
        let buffer = "";
        process.stdin.setEncoding("utf8");
        process.stdin.on("data", (chunk) => {
          buffer += chunk;
        });
        process.stdin.once("end", () => resolve(buffer));
        process.stdin.once("error", reject);
      });
    }
    try {
      operations = JSON.parse(operationsJson);
    } catch {
      return { exitCode: 1, stderr: "Failed to parse operations JSON" };
    }
    const results = [];
    for (const op of operations) {
      const { action, target, options = {} } = op;
      try {
        let result;
        switch (action.toLowerCase()) {
          case "add": {
            const cmdOpts = { type: options.type as string | undefined, props: (options.props as Record<string, string>) ?? {}, json: false };
            result = await addDocumentNode(filePath, target, cmdOpts);
            break;
          }
          case "set": {
            const cmdOpts = { type: options.type as string | undefined, props: (options.props as Record<string, string>) ?? {}, json: false };
            result = await setDocumentNode(filePath, target, cmdOpts);
            break;
          }
          case "remove":
            result = await removeDocumentNode(filePath, target);
            break;
          case "get":
            result = await getDocumentNode(filePath, target);
            break;
          case "query":
            result = await queryDocumentNodes(filePath, target);
            break;
          default:
            result = { ok: false, error: { code: "unknown_action", message: `Unknown action: ${action}` } };
        }
        results.push({ action, target, status: (result as { ok?: boolean }).ok ? "success" : "failed", result });
      } catch (e) {
        results.push({ action, target, status: "error", error: e instanceof Error ? e.message : String(e) });
      }
    }
    return { exitCode: 0, stdout: JSON.stringify({ ok: true, results }, null, 2) };
  }

  if (command === "raw-set") {
    const filePath = rest[0];
    const partPath = rest[1];
    if (!filePath || !partPath) {
      throw new UsageError("raw-set requires <file> <part-path>.");
    }
    let xpath = "";
    let action = "";
    let xml: string | undefined;
    for (let index = 2; index < rest.length; index += 1) {
      const token = rest[index];
      if (token === "--xpath") {
        xpath = rest[index + 1] ?? "";
        index += 1;
        continue;
      }
      if (token === "--action") {
        action = rest[index + 1] ?? "";
        index += 1;
        continue;
      }
      if (token === "--xml") {
        xml = rest[index + 1];
        index += 1;
        continue;
      }
    }
    if (!xpath || !action) {
      throw new UsageError("raw-set requires --xpath <xpath-expr> --action <action> [--xml <xml>].");
    }
    const result = await rawSetDocument(filePath, partPath, xpath, action, xml);
    return { exitCode: 0, stdout: JSON.stringify(result, null, 2) };
  }

  if (command === "add-part") {
    const filePath = rest[0];
    const parentPath = rest[1];
    if (!filePath || !parentPath) {
      throw new UsageError("add-part requires <file> <parent-path>.");
    }
    const parsed = parseProps(rest.slice(2));
    const partType = parsed.type ?? "chart";
    if (!partType) {
      throw new UsageError("add-part requires --type <part-type>.");
    }
    const result = await addDocumentPart(filePath, parentPath, partType, parsed.props);
    return { exitCode: 0, stdout: JSON.stringify(result, null, 2) };
  }

  if (command === "merge") {
    const templatePath = rest[0];
    const outputPath = rest[1];
    if (!templatePath || !outputPath) {
      throw new UsageError("merge requires <template> <output>.");
    }
    const dataIndex = rest.indexOf("--data");
    const dataArg = dataIndex >= 0 ? rest[dataIndex + 1] : undefined;
    if (!dataArg) {
      throw new UsageError("merge requires --data <json-or-file-path>.");
    }

    // Parse the data - could be JSON string or path to a .json file
    let data: Record<string, unknown>;
    if (dataArg.startsWith("{")) {
      try {
        data = JSON.parse(dataArg);
      } catch {
        return { exitCode: 1, stderr: "Failed to parse --data as JSON" };
      }
    } else {
      // Treat as file path
      try {
        const fs = await import("node:fs/promises");
        const content = await fs.readFile(dataArg, "utf8");
        data = JSON.parse(content);
      } catch {
        return { exitCode: 1, stderr: `Failed to read or parse data file: ${dataArg}` };
      }
    }

    const result = await mergeDocument(templatePath, data, outputPath);
    return { exitCode: 0, stdout: JSON.stringify(result, null, 2) };
  }

  if (command === "unwatch") {
    const filePath = rest[0];
    if (!filePath) {
      throw new UsageError("unwatch requires <file>.");
    }
    const session = await readSessionRecord("watch", filePath);
    if (!session?.pid) {
      return {
        exitCode: 1,
        stderr: JSON.stringify({ ok: false, message: `No active watch session for ${filePath}` }, null, 2),
      };
    }
    process.kill(session.pid, "SIGTERM");
    return {
      exitCode: 0,
      stdout: JSON.stringify({ ok: true, message: `Watch stopped for ${filePath}`, pid: session.pid }, null, 2),
    };
  }

  if (command === "open") {
    const filePath = rest[0];
    if (!filePath) {
      throw new UsageError("open requires <file>.");
    }
    try {
      const existing = await readSessionRecord("resident", filePath);
      if (existing?.pid) {
        return {
          exitCode: 0,
          stdout: JSON.stringify({ ok: true, filePath, pid: existing.pid, reused: true }, null, 2),
        };
      }
      const { spawn } = await import("node:child_process");
      const { fileURLToPath } = await import("node:url");
      const workerPath = fileURLToPath(new URL("./resident-worker.ts", import.meta.url));
      const child = spawn(process.execPath, ["run", workerPath, filePath], {
        cwd: process.cwd(),
        detached: true,
        stdio: "ignore",
      });
      child.unref();
      const session = await waitForSessionRecord("resident", filePath, 2000);
      if (!session?.pid) {
        try {
          process.kill(child.pid!, "SIGTERM");
        } catch {}
        return { exitCode: 1, stderr: `Failed to open ${filePath}: resident session did not start` };
      }
      return {
        exitCode: 0,
        stdout: JSON.stringify({ ok: true, filePath, pid: session.pid, reused: false }, null, 2),
      };
    } catch (e) {
      return { exitCode: 1, stderr: `Failed to open ${filePath}: ${e instanceof Error ? e.message : String(e)}` };
    }
  }

  if (command === "close") {
    const filePath = rest[0];
    if (!filePath) {
      throw new UsageError("close requires <file>.");
    }
    const session = await readSessionRecord("resident", filePath);
    if (!session?.pid) {
      return {
        exitCode: 1,
        stderr: JSON.stringify({ ok: false, message: `No active resident session for ${filePath}` }, null, 2),
      };
    }
    process.kill(session.pid, "SIGTERM");
    await removeSessionRecord("resident", filePath);
    return {
      exitCode: 0,
      stdout: JSON.stringify({ ok: true, filePath, pid: session.pid }, null, 2),
    };
  }

  if (command === "about") {
    return {
      exitCode: 0,
      stdout: JSON.stringify({ product: "officekit", version: "1.0.0" }, null, 2),
    };
  }

  if (command === "contracts") {
    const format = rest[0];
    if (format === "--format" && rest[1] === "json") {
      return { exitCode: 0, stdout: JSON.stringify(summarizeParity(), null, 2) };
    }
    return { exitCode: 0, stdout: JSON.stringify(summarizeParity(), null, 2) };
  }

  return { exitCode: 2, stderr: `Command '${command}' is not implemented yet in the current vertical slice.` };
}

function summarizeResult(result: unknown): string {
  if (typeof result === "string") return result;
  return JSON.stringify(result, null, 2);
}
