import { createDocument, addDocumentNode, checkDocument, getDocumentNode, parseProps, rawDocument, removeDocumentNode, renderDocumentHtml, setDocumentNode, viewDocument } from "./document-store.js";
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
    const result = await getDocumentNode(filePath, targetPath);
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
    return { exitCode: 0, stdout: await rawDocument(filePath) };
  }

  if (command === "check") {
    const filePath = rest[0];
    if (!filePath) throw new UsageError("check requires <file>.");
    return { exitCode: 0, stdout: JSON.stringify(await checkDocument(filePath), null, 2) };
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
