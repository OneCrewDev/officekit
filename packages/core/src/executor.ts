import { createDocument, addDocumentNode, checkDocument, getDocumentNode, importDelimitedData, parseProps, queryDocumentNodes, rawDocument, removeDocumentNode, renderDocumentHtml, setDocumentNode, viewDocument } from "./document-store.js";
import { UnsupportedCapabilityError, UsageError } from "./errors.js";

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
    const result = command === "query"
      ? await queryDocumentNodes(filePath, targetPath)
      : await getDocumentNode(filePath, targetPath);
    return { exitCode: 0, stdout: parsed.json || command === "query" ? JSON.stringify(result, null, 2) : summarizeResult(result) };
  }

  if (command === "view") {
    const filePath = rest[0];
    const mode = rest[1] ?? "outline";
    if (!filePath) throw new UsageError("view requires <file> <mode?>.");
    const result = await viewDocument(filePath, mode);
    return { exitCode: 0, stdout: result.output };
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
      render: async () => renderDocumentHtml(await getDocumentNode(filePath, "/") as never),
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

  return { exitCode: 2, stderr: `Command '${command}' is not implemented yet in the current vertical slice.` };
}

function summarizeResult(result: unknown): string {
  if (typeof result === "string") return result;
  return JSON.stringify(result, null, 2);
}
