/**
 * CSV/TSV Import Module
 *
 * Provides unified parsing and type inference for delimited text data
 * that can be imported into Excel or Word documents.
 */

export interface ImportResult {
  importedRows: number;
  importedCols: number;
  sheet?: string;
  startCell: string;
  path?: string;
  autoFilter?: string;
  freezeTopLeftCell?: string;
}

export interface ParsedCell {
  value: string;
  type?: "string" | "number" | "boolean" | "date";
  formula?: string;
}

/**
 * Parse delimited text content into rows of cells.
 * Handles:
 * - Quoted fields with embedded delimiters
 * - Escaped quotes within quoted fields (double quotes)
 * - Different line endings (LF and CRLF)
 * - BOM character removal
 */
export function parseDelimitedRows(content: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  if (!content) return rows;
  if (content.charCodeAt(0) === 0xfeff) content = content.slice(1);
  const currentRow: string[] = [];
  let field = "";
  let inQuotes = false;
  for (let index = 0; index < content.length; index += 1) {
    const char = content[index];
    if (inQuotes) {
      if (char === '"') {
        if (content[index + 1] === '"') {
          field += '"';
          index += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += char;
      }
      continue;
    }
    if (char === '"') {
      inQuotes = true;
      continue;
    }
    if (char === delimiter) {
      currentRow.push(field);
      field = "";
      continue;
    }
    if (char === "\n" || char === "\r") {
      if (char === "\r" && content[index + 1] === "\n") index += 1;
      currentRow.push(field);
      field = "";
      if (!(currentRow.length === 1 && currentRow[0] === "")) rows.push([...currentRow]);
      currentRow.length = 0;
      continue;
    }
    field += char;
  }
  if (field.length > 0 || currentRow.length > 0) {
    currentRow.push(field);
    if (!(currentRow.length === 1 && currentRow[0] === "")) rows.push([...currentRow]);
  }
  return rows;
}

/**
 * Infer the cell type and value from a raw string.
 * Handles:
 * - Empty strings
 * - Formulas (starting with =)
 * - Booleans (true/false)
 * - ISO date strings
 * - Numbers
 * - Plain strings
 */
export function inferCellType(rawValue: string): ParsedCell {
  if (rawValue === "") return { value: "" };
  if (rawValue.startsWith("=")) return { value: "", formula: rawValue.slice(1) };
  if (/^(true|false)$/i.test(rawValue)) return { value: rawValue.toUpperCase() === "TRUE" ? "1" : "0", type: "boolean" };
  const isoDate = tryParseIsoDate(rawValue);
  if (isoDate) return { value: isoDate, type: "date" };
  if (!Number.isNaN(Number(rawValue))) return { value: rawValue, type: "number" };
  return { value: rawValue, type: "string" };
}

/**
 * Try to parse an ISO date string and return Excel serial date number.
 */
function tryParseIsoDate(rawValue: string): string | null {
  const date = new Date(rawValue);
  if (Number.isNaN(date.getTime()) || !/^\d{4}-\d{2}-\d{2}/.test(rawValue)) {
    return null;
  }
  return String((date.getTime() - Date.UTC(1899, 11, 30)) / (24 * 60 * 60 * 1000));
}

/**
 * Detect the delimiter based on file extension or format hint.
 */
export function detectDelimiter(fileNameOrFormat: string): string {
  const lower = fileNameOrFormat.toLowerCase();
  if (lower === "tsv" || lower === "tab") return "\t";
  if (lower === "csv") return ",";
  if (lower.endsWith(".tsv") || lower.endsWith(".tab")) return "\t";
  if (lower.endsWith(".csv")) return ",";
  return ",";
}

/**
 * Parse import command line options from raw args array.
 */
export interface ImportCliOptions {
  delimiter: string;
  hasHeader: boolean;
  startCell: string;
  sourceFile?: string;
  useStdin: boolean;
}

export function parseImportArgs(args: string[]): {
  parentPath: string;
  options: ImportCliOptions;
  sourceFile?: string;
} {
  const filePath = args[0];
  const parentPath = args[1];
  if (!filePath || !parentPath) {
    throw new Error("import requires <file> <parent-path> and a source file or --file.");
  }

  let delimiter = ",";
  let hasHeader = false;
  let startCell = "A1";
  let sourceFile: string | undefined;
  let useStdin = false;

  for (let index = 2; index < args.length; index += 1) {
    const token = args[index];
    if (token === "--file") {
      sourceFile = args[index + 1];
      index += 1;
      continue;
    }
    if (token === "--format") {
      const format = (args[index + 1] ?? "csv").toLowerCase();
      delimiter = format === "tsv" || format === "tab" ? "\t" : ",";
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
      startCell = args[index + 1] ?? "A1";
      index += 1;
      continue;
    }
    if (!token.startsWith("--") && !sourceFile) {
      sourceFile = token;
    }
  }

  // Auto-detect delimiter from file extension
  if (sourceFile && delimiter === ",") {
    const autoDelim = detectDelimiter(sourceFile);
    if (autoDelim !== ",") delimiter = autoDelim;
  }

  return {
    parentPath,
    options: {
      delimiter,
      hasHeader,
      startCell,
      sourceFile,
      useStdin,
    },
    sourceFile,
  };
}
