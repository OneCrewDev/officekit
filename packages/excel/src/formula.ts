// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

// Excel Formula Evaluation Engine
// Ported from OfficeCLI FormulaEvaluator.cs

import type { ExcelCellModel } from "./adapter.ts";

/**
 * Result of a formula evaluation. Can be numeric, string, boolean, or error.
 */
export class FormulaResult {
  public constructor(
    public readonly numericValue?: number,
    public readonly stringValue?: string,
    public readonly boolValue?: boolean,
    public readonly errorValue?: string,
    public readonly arrayValue?: number[],
  ) {}

  public get isNumeric(): boolean {
    return this.numericValue !== undefined;
  }

  public get isString(): boolean {
    return this.stringValue !== undefined;
  }

  public get isBool(): boolean {
    return this.boolValue !== undefined;
  }

  public get isError(): boolean {
    return this.errorValue !== undefined;
  }

  public get isArray(): boolean {
    return this.arrayValue !== undefined;
  }

  public static Number(v: number): FormulaResult {
    return new FormulaResult(v);
  }

  public static Str(v: string): FormulaResult {
    return new FormulaResult(undefined, v);
  }

  public static Bool(v: boolean): FormulaResult {
    return new FormulaResult(undefined, undefined, v);
  }

  public static Error(v: string): FormulaResult {
    return new FormulaResult(undefined, undefined, undefined, v);
  }

  public static Array(v: number[]): FormulaResult {
    return new FormulaResult(undefined, undefined, undefined, undefined, v);
  }

  public asNumber(): number {
    if (this.numericValue !== undefined) return this.numericValue;
    if (this.boolValue !== undefined) return this.boolValue ? 1 : 0;
    return 0;
  }

  public asString(): string {
    if (this.stringValue !== undefined) return this.stringValue;
    if (this.numericValue !== undefined) return this.numericValue.toString();
    if (this.boolValue !== undefined) return this.boolValue ? "TRUE" : "FALSE";
    if (this.errorValue !== undefined) return this.errorValue;
    return "";
  }

  public toCellValueText(): string {
    if (this.numericValue !== undefined) {
      let v = this.numericValue;
      if (v !== 0) {
        const digits = 15 - Math.floor(Math.log10(Math.abs(v))) - 1;
        if (digits >= 0 && digits <= 15) {
          v = Math.round(v * Math.pow(10, digits)) / Math.pow(10, digits);
        }
      }
      return v.toString();
    }
    if (this.boolValue !== undefined) return this.boolValue ? "1" : "0";
    if (this.stringValue !== undefined) return this.stringValue;
    return "";
  }
}

/**
 * 2D range data for lookup functions (VLOOKUP, HLOOKUP, INDEX).
 */
export class RangeData {
  public constructor(
    public readonly cells: (FormulaResult | null)[][],
    public readonly rows: number,
    public readonly cols: number,
  ) {}

  public toDoubleArray(): number[] {
    const values: number[] = [];
    for (let r = 0; r < this.rows; r++) {
      for (let c = 0; c < this.cols; c++) {
        const cell = this.cells[r][c];
        if (cell?.isNumeric) values.push(cell.numericValue!);
        else if (cell?.isBool) values.push(cell.boolValue! ? 1 : 0);
      }
    }
    return values;
  }

  /** Flatten all cells into a flat list (preserving nulls for ISERROR etc.) */
  public toFlatResults(): (FormulaResult | null)[] {
    const results: (FormulaResult | null)[] = [];
    for (let r = 0; r < this.rows; r++) {
      for (let c = 0; c < this.cols; c++) {
        results.push(this.cells[r][c]);
      }
    }
    return results;
  }

  /** Returns the first error found in the range, or null if none. */
  public firstError(): FormulaResult | null {
    for (let r = 0; r < this.rows; r++) {
      for (let c = 0; c < this.cols; c++) {
        const cell = this.cells[r][c];
        if (cell?.isError) return cell;
      }
    }
    return null;
  }
}

type CellMap = Record<string, ExcelCellModel>;

// ==================== Helper Functions ====================

function parseRef(r: string): { col: string; row: number } {
  const match = r.match(/^([A-Z]+)(\d+)$/i);
  if (match) {
    return { col: match[1].toUpperCase(), row: parseInt(match[2], 10) };
  }
  return { col: "A", row: 1 };
}

function colToIndex(col: string): number {
  let r = 0;
  for (const c of col.toUpperCase()) {
    r = r * 26 + (c.charCodeAt(0) - "A".charCodeAt(0) + 1);
  }
  return r;
}

export function indexToCol(i: number): string {
  let r = "";
  while (i > 0) {
    i--;
    r = String.fromCharCode(65 + (i % 26)) + r;
    i = Math.floor(i / 26);
  }
  return r;
}

function stripDollar(s: string): string {
  return s.replace(/\$/g, "");
}

function isCellRef(s: string): boolean {
  return /^[A-Z]{1,3}\d+$/i.test(s);
}

function parseNumber(s: string, i: { value: number }): string | null {
  const start = i.value;
  if (i.value < s.length && (s[i.value] === "-" || s[i.value] === "+")) i.value++;
  let hasDigits = false;
  while (i.value < s.length && /\d/.test(s[i.value])) {
    i.value++;
    hasDigits = true;
  }
  if (i.value < s.length && s[i.value] === ".") {
    i.value++;
    while (i.value < s.length && /\d/.test(s[i.value])) {
      i.value++;
      hasDigits = true;
    }
  }
  if (i.value < s.length && (s[i.value] === "e" || s[i.value] === "E")) {
    i.value++;
    if (i.value < s.length && (s[i.value] === "+" || s[i.value] === "-")) i.value++;
    while (i.value < s.length && /\d/.test(s[i.value])) i.value++;
  }
  if (!hasDigits) {
    i.value = start;
    return null;
  }
  return s.substring(start, i.value);
}

// ==================== Tokenizer ====================

enum TokenType {
  Number,
  String,
  CellRef,
  Range,
  Op,
  LParen,
  RParen,
  Comma,
  Func,
  Bool,
  Compare,
  SheetCellRef,
  SheetRange,
}

interface Token {
  type: TokenType;
  value: string;
}

function tokenize(formula: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  formula = formula.trim();

  while (i < formula.length) {
    const ch = formula[i];

    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    if (ch === ">" || ch === "<" || ch === "=") {
      if (ch === "=" && i === 0) {
        i++;
        continue;
      }
      if (i + 1 < formula.length && (formula[i + 1] === "=" || formula[i + 1] === ">")) {
        tokens.push({ type: TokenType.Compare, value: formula.substring(i, i + 2) });
        i += 2;
      } else {
        tokens.push({ type: TokenType.Compare, value: ch });
        i++;
      }
      continue;
    }

    if (ch === "+" || ch === "-" || ch === "*" || ch === "/" || ch === "^" || ch === "%") {
      if ((ch === "-" || ch === "+") && (tokens.length === 0 ||
          tokens[tokens.length - 1].type === TokenType.Op ||
          tokens[tokens.length - 1].type === TokenType.LParen ||
          tokens[tokens.length - 1].type === TokenType.Comma ||
          tokens[tokens.length - 1].type === TokenType.Compare)) {
        const num = parseNumber(formula, { value: i });
        if (num !== null) {
          i = { value: i }.value;
          // Need to manually advance since we can't modify i directly
          const numResult = parseNumber(formula, { value: i });
          if (numResult !== null) {
            // Re-parse correctly
            const startIdx = i;
            if (formula[i] === "-" || formula[i] === "+") i++;
            while (i < formula.length && /\d/.test(formula[i])) i++;
            if (i < formula.length && formula[i] === ".") {
              i++;
              while (i < formula.length && /\d/.test(formula[i])) i++;
            }
            if (i < formula.length && (formula[i] === "e" || formula[i] === "E")) {
              i++;
              if (i < formula.length && (formula[i] === "+" || formula[i] === "-")) i++;
              while (i < formula.length && /\d/.test(formula[i])) i++;
            }
            tokens.push({ type: TokenType.Number, value: formula.substring(startIdx, i) });
            continue;
          }
        }
      }
      if (ch === "%") {
        tokens.push({ type: TokenType.Op, value: "%" });
        i++;
        continue;
      }
      tokens.push({ type: TokenType.Op, value: ch });
      i++;
      continue;
    }

    if (ch === "(") { tokens.push({ type: TokenType.LParen, value: "(" }); i++; continue; }
    if (ch === ")") { tokens.push({ type: TokenType.RParen, value: ")" }); i++; continue; }
    if (ch === ",") { tokens.push({ type: TokenType.Comma, value: "," }); i++; continue; }
    if (ch === "&") { tokens.push({ type: TokenType.Op, value: "&" }); i++; continue; }

    if (ch === '"') {
      i++;
      const sb: string[] = [];
      while (i < formula.length) {
        if (formula[i] === '"') {
          if (i + 1 < formula.length && formula[i + 1] === '"') {
            sb.push('"');
            i += 2;
          } else {
            i++;
            break;
          }
        } else {
          sb.push(formula[i]);
          i++;
        }
      }
      tokens.push({ type: TokenType.String, value: sb.join("") });
      continue;
    }

    // Quoted sheet reference: 'Sheet Name'!CellRef or 'Sheet Name'!Range
    if (ch === "'") {
      const si = i + 1;
      const ei = formula.indexOf("'", si);
      if (ei > si && ei + 1 < formula.length && formula[ei + 1] === "!") {
        const sheetName = formula.substring(si, ei);
        i = ei + 2;
        const refStart = i;
        while (i < formula.length && (/[A-Z0-9\$]/.test(formula[i]) || formula[i] === ":")) i++;
        const refPart = stripDollar(formula.substring(refStart, i));
        if (refPart.includes(":"))
          tokens.push({ type: TokenType.SheetRange, value: `${sheetName}!${refPart}` });
        else
          tokens.push({ type: TokenType.SheetCellRef, value: `${sheetName}!${refPart.toUpperCase()}` });
        continue;
      }
    }

    if (/\d/.test(ch) || ch === ".") {
      const num = parseNumber(formula, { value: i });
      if (num !== null) {
        tokens.push({ type: TokenType.Number, value: num });
        continue;
      }
    }

    if (/[A-Z_$]/.test(ch) || ch === "_" || ch === "$") {
      const start = i;
      while (i < formula.length && (/[A-Z0-9_$.]/.test(formula[i]))) i++;
      const word = formula.substring(start, i);
      const stripped = stripDollar(word);

      if (stripped.toUpperCase() === "TRUE") { tokens.push({ type: TokenType.Bool, value: "TRUE" }); continue; }
      if (stripped.toUpperCase() === "FALSE") { tokens.push({ type: TokenType.Bool, value: "FALSE" }); continue; }

      // Unquoted sheet reference: SheetName!CellRef or SheetName!Range
      if (i < formula.length && formula[i] === "!") {
        const sheetName = word;
        i++;
        const refStart = i;
        while (i < formula.length && (/[A-Z0-9$]/.test(formula[i]) || formula[i] === ":")) i++;
        const refPart = stripDollar(formula.substring(refStart, i));
        if (refPart.includes(":"))
          tokens.push({ type: TokenType.SheetRange, value: `${sheetName}!${refPart}` });
        else
          tokens.push({ type: TokenType.SheetCellRef, value: `${sheetName}!${refPart.toUpperCase()}` });
        continue;
      }

      if (i < formula.length && formula[i] === ":" && isCellRef(stripped)) {
        i++;
        const s2 = i;
        while (i < formula.length && (/[A-Z0-9$]/.test(formula[i]))) i++;
        tokens.push({ type: TokenType.Range, value: `${stripped}:${stripDollar(formula.substring(s2, i))}` });
        continue;
      }

      if (i < formula.length && formula[i] === "(" && !isCellRef(stripped)) {
        tokens.push({ type: TokenType.Func, value: word.replace(/\./g, "_").toUpperCase() });
        continue;
      }

      if (isCellRef(stripped)) {
        tokens.push({ type: TokenType.CellRef, value: stripped.toUpperCase() });
        continue;
      }

      throw new Error(`Unknown: ${word}`);
    }

    throw new Error(`Unexpected: ${ch}`);
  }

  return tokens;
}

// ==================== Formula Evaluator ====================

export class FormulaEvaluator {
  protected cells: CellMap;
  protected readonly sheetData: { cells: CellMap }[];
  protected readonly workbookPart?: unknown;
  protected visiting: Set<string>;
  protected readonly depth: number;
  protected readonly sheetKey: string;
  protected cellIndex?: Map<string, ExcelCellModel>;

  constructor(cells: CellMap, sheetDataAll?: { cells: CellMap }[], workbookPart?: unknown) {
    this.cells = cells;
    this.sheetData = sheetDataAll ?? [{ cells }];
    this.workbookPart = workbookPart;
    this.visiting = new Set();
    this.depth = 0;
    this.sheetKey = "";
  }

  public tryEvaluate(formula: string): number | null {
    const result = this.tryEvaluateFull(formula);
    if (result?.numericValue !== undefined) return result.numericValue;
    if (result?.boolValue === true) return 1;
    if (result?.boolValue === false) return 0;
    return null;
  }

  public tryEvaluateFull(formula: string): FormulaResult | null {
    try {
      if (this.depth === 0) this.visiting.clear();
      return this.evaluateFormula(formula);
    } catch {
      return null;
    }
  }

  protected evaluateFormula(formula: string): FormulaResult | null {
    const tokens = tokenize(formula);
    let pos = 0;
    const result = this.parseExpression(tokens, pos);
    if (typeof result === "number") {
      pos = result;
      return null;
    }
    const [res, newPos] = result;
    if (newPos === tokens.length) return res;
    return null;
  }

  // ==================== Recursive Descent Parser ====================

  protected parseExpression(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    const result = this.parseComparison(tokens, pos);
    if (typeof result === "number") return result;
    const [left, newPos] = result;
    if (left === null) return [null, newPos];
    return [left, newPos];
  }

  protected parseComparison(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    const result = this.parseConcat(tokens, pos);
    if (typeof result === "number") return result;
    let [left, newPos] = result;
    if (left === null) return [null, newPos];

    while (newPos < tokens.length && tokens[newPos].type === TokenType.Compare) {
      const op = tokens[newPos].value;
      newPos++;
      const rightResult = this.parseConcat(tokens, newPos);
      if (typeof rightResult === "number") return rightResult;
      let [right, posAfterRight] = rightResult;
      if (right === null) return [null, posAfterRight];
      newPos = posAfterRight;

      if (left.isError) return [left, newPos];
      if (right.isError) return [right, newPos];

      const cmp = this.compareValues(left, right);
      left = op === "=" ? FormulaResult.Bool(cmp === 0)
        : op === "<>" ? FormulaResult.Bool(cmp !== 0)
        : op === "<" ? FormulaResult.Bool(cmp < 0)
        : op === ">" ? FormulaResult.Bool(cmp > 0)
        : op === "<=" ? FormulaResult.Bool(cmp <= 0)
        : op === ">=" ? FormulaResult.Bool(cmp >= 0)
        : FormulaResult.Bool(false);
    }
    return [left, newPos];
  }

  protected parseConcat(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    const result = this.parseAddSub(tokens, pos);
    if (typeof result === "number") return result;
    let [left, newPos] = result;
    if (left === null) return [null, newPos];

    while (newPos < tokens.length && tokens[newPos].type === TokenType.Op && tokens[newPos].value === "&") {
      newPos++;
      const rightResult = this.parseAddSub(tokens, newPos);
      if (typeof rightResult === "number") return rightResult;
      let [right, posAfterRight] = rightResult;
      if (right === null) return [null, posAfterRight];
      newPos = posAfterRight;

      if (left.isError) return [left, newPos];
      if (right.isError) return [right, newPos];
      left = FormulaResult.Str(left.asString() + right.asString());
    }
    return [left, newPos];
  }

  protected parseAddSub(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    const result = this.parseMulDiv(tokens, pos);
    if (typeof result === "number") return result;
    let [left, newPos] = result;
    if (left === null) return [null, newPos];

    while (newPos < tokens.length && tokens[newPos].type === TokenType.Op && (tokens[newPos].value === "+" || tokens[newPos].value === "-")) {
      const op = tokens[newPos].value;
      newPos++;
      const rightResult = this.parseMulDiv(tokens, newPos);
      if (typeof rightResult === "number") return rightResult;
      let [right, posAfterRight] = rightResult;
      if (right === null) return [null, posAfterRight];
      newPos = posAfterRight;

      if (left.isError) return [left, newPos];
      if (right.isError) return [right, newPos];
      left = FormulaResult.Number(op === "+" ? left.asNumber() + right.asNumber() : left.asNumber() - right.asNumber());
    }
    return [left, newPos];
  }

  protected parseMulDiv(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    const result = this.parsePower(tokens, pos);
    if (typeof result === "number") return result;
    let [left, newPos] = result;
    if (left === null) return [null, newPos];

    while (newPos < tokens.length && tokens[newPos].type === TokenType.Op && (tokens[newPos].value === "*" || tokens[newPos].value === "/")) {
      const op = tokens[newPos].value;
      newPos++;
      const rightResult = this.parsePower(tokens, newPos);
      if (typeof rightResult === "number") return rightResult;
      let [right, posAfterRight] = rightResult;
      if (right === null) return [null, posAfterRight];
      newPos = posAfterRight;

      if (left.isError) return [left, newPos];
      if (right.isError) return [right, newPos];
      if (op === "/" && right.asNumber() === 0) return [FormulaResult.Error("#DIV/0!"), newPos];
      left = FormulaResult.Number(op === "*" ? left.asNumber() * right.asNumber() : left.asNumber() / right.asNumber());
    }
    return [left, newPos];
  }

  protected parsePower(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    const result = this.parseUnary(tokens, pos);
    if (typeof result === "number") return result;
    let [b, newPos] = result;
    if (b === null) return [null, newPos];

    while (newPos < tokens.length && tokens[newPos].type === TokenType.Op && tokens[newPos].value === "^") {
      newPos++;
      const expResult = this.parseUnary(tokens, newPos);
      if (typeof expResult === "number") return expResult;
      let [e, posAfterExp] = expResult;
      if (e === null) return [null, posAfterExp];
      newPos = posAfterExp;

      if (b.isError) return [b, newPos];
      if (e.isError) return [e, newPos];
      b = FormulaResult.Number(Math.pow(b.asNumber(), e.asNumber()));
    }
    return [b, newPos];
  }

  protected parseUnary(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    if (pos < tokens.length && tokens[pos].type === TokenType.Op) {
      if (tokens[pos].value === "-") {
        pos++;
        const result = this.parseUnary(tokens, pos);
        if (typeof result === "number") return result;
        const [v, newPos] = result;
        if (v === null) return [null, newPos];
        if (v.isError) return [v, newPos];
        if (v.isArray) return [FormulaResult.Array(v.arrayValue!.map(x => -x)), newPos];
        return [FormulaResult.Number(-v.asNumber()), newPos];
      }
      if (tokens[pos].value === "+") {
        pos++;
        return this.parseUnary(tokens, pos);
      }
    }
    return this.parsePostfix(tokens, pos);
  }

  protected parsePostfix(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    const result = this.parseAtom(tokens, pos);
    if (typeof result === "number") return result;
    let [v, newPos] = result;
    if (v === null) return [null, newPos];

    while (newPos < tokens.length && tokens[newPos].type === TokenType.Op && tokens[newPos].value === "%") {
      newPos++;
      v = FormulaResult.Number(v.asNumber() / 100.0);
    }
    return [v, newPos];
  }

  protected parseAtom(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    if (pos >= tokens.length) return [null, pos];
    const tok = tokens[pos];

    switch (tok.type) {
      case TokenType.Number: {
        const n = parseFloat(tok.value);
        return [isNaN(n) ? null : FormulaResult.Number(n), pos + 1];
      }
      case TokenType.String:
        return [FormulaResult.Str(tok.value), pos + 1];
      case TokenType.Bool:
        return [FormulaResult.Bool(tok.value === "TRUE"), pos + 1];
      case TokenType.CellRef:
        return [this.resolveCellResult(tok.value), pos + 1];
      case TokenType.SheetCellRef:
        return [this.resolveSheetCellResult(tok.value), pos + 1];
      case TokenType.Range:
        return [FormulaResult.Number(0), pos + 1];
      case TokenType.SheetRange:
        return [FormulaResult.Number(0), pos + 1];
      case TokenType.LParen: {
        pos++;
        const innerResult = this.parseExpression(tokens, pos);
        if (typeof innerResult === "number") {
          pos = innerResult;
        } else {
          const [inner, newPos] = innerResult;
          pos = newPos;
          if (pos < tokens.length && tokens[pos].type === TokenType.RParen) pos++;
          return [inner, pos];
        }
        if (pos < tokens.length && tokens[pos].type === TokenType.RParen) pos++;
        return [null, pos];
      }
      case TokenType.Func:
        return this.parseFunction(tokens, pos);
      default:
        return [null, pos + 1];
    }
  }

  protected parseFunction(tokens: Token[], pos: number): [FormulaResult | null, number] | number {
    const name = tokens[pos].value;
    pos++;
    if (pos >= tokens.length || tokens[pos].type !== TokenType.LParen) return [null, pos];
    pos++;

    const args: unknown[] = [];
    if (pos < tokens.length && tokens[pos].type !== TokenType.RParen) {
      while (true) {
        if (pos < tokens.length && (tokens[pos].type === TokenType.Range || tokens[pos].type === TokenType.SheetRange)) {
          args.push(this.expand2DRange(tokens[pos].value));
          pos++;
        } else {
          const exprResult = this.parseExpression(tokens, pos);
          if (typeof exprResult === "number") return exprResult;
          const [expr, newPos] = exprResult;
          if (expr === null) return [null, newPos];
          args.push(expr);
          pos = newPos;
        }
        if (pos >= tokens.length || tokens[pos].type !== TokenType.Comma) break;
        pos++;
      }
    }
    if (pos < tokens.length && tokens[pos].type === TokenType.RParen) pos++;
    const result = this.evalFunction(name, args);
    return [result, pos];
  }

  // ==================== Cell & Range Resolution ====================

  protected resolveCellResult(cellRef: string): FormulaResult {
    cellRef = stripDollar(cellRef).toUpperCase();
    const qualifiedRef = this.sheetKey === "" ? cellRef : `${this.sheetKey}!${cellRef}`;

    if (!this.visiting.add(qualifiedRef)) return FormulaResult.Number(0);

    try {
      const cell = this.findCell(cellRef);
      if (!cell) return FormulaResult.Number(0);

      if (cell.formula) {
        try {
          const subEval = new FormulaEvaluator(this.cells, this.sheetData, this.workbookPart);
          subEval.visiting = this.visiting;
          const evaluated = subEval.tryEvaluateFull(cell.formula);
          if (evaluated) return evaluated;
        } catch { /* fall through to cached value */ }
      }

      const cached = cell.value;
      if (cached !== undefined && cached !== "") {
        if (cell.type === "boolean") return FormulaResult.Bool(cached === "1");
        if (cell.type === "number") {
          const n = parseFloat(cached);
          return isNaN(n) ? FormulaResult.Str(cached) : FormulaResult.Number(n);
        }
        const n = parseFloat(cached);
        return isNaN(n) ? FormulaResult.Str(cached) : FormulaResult.Number(n);
      }

      return FormulaResult.Number(0);
    } finally {
      this.visiting.delete(qualifiedRef);
    }
  }

  protected resolveSheetCellResult(sheetCellRef: string): FormulaResult {
    if (this.depth > 20) return FormulaResult.Number(0);

    const bangIdx = sheetCellRef.indexOf("!");
    if (bangIdx < 0) return FormulaResult.Number(0);

    const sheetName = sheetCellRef.substring(0, bangIdx);
    const cellRef = sheetCellRef.substring(bangIdx + 1);

    const targetSheet = this.sheetData.find((_, idx) => {
      return true;
    });

    if (!targetSheet) return FormulaResult.Number(0);

    const subEval = new FormulaEvaluator(targetSheet.cells, this.sheetData, this.workbookPart);
    subEval.visiting = this.visiting;
    (subEval as unknown as { depth: number }).depth = this.depth + 1;
    (subEval as unknown as { sheetKey: string }).sheetKey = sheetName;
    return subEval.resolveCellResult(cellRef);
  }

  protected findCell(cellRef: string): ExcelCellModel | null {
    if (!this.cellIndex) {
      this.cellIndex = new Map();
      for (const [ref, cell] of Object.entries(this.cells)) {
        this.cellIndex.set(ref.toUpperCase(), cell);
      }
    }
    return this.cellIndex.get(cellRef.toUpperCase()) ?? null;
  }

  protected expand2DRange(rangeExpr: string): RangeData {
    let sheetPrefix: string | null = null;
    let expr = rangeExpr;
    const bangIdx = rangeExpr.indexOf("!");
    if (bangIdx >= 0) {
      sheetPrefix = rangeExpr.substring(0, bangIdx);
      expr = rangeExpr.substring(bangIdx + 1);
    }

    const parts = expr.split(":");
    if (parts.length !== 2) return new RangeData([], 0, 0);

    const { col: col1, row: row1 } = parseRef(stripDollar(parts[0]));
    const { col: col2, row: row2 } = parseRef(stripDollar(parts[1]));

    const c1 = colToIndex(col1);
    const c2 = colToIndex(col2);
    const r1 = Math.min(row1, row2);
    const r2 = Math.max(row1, row2);
    const cMin = Math.min(c1, c2);
    const cMax = Math.max(c1, c2);

    const rows = r2 - r1 + 1;
    const cols = cMax - cMin + 1;
    const cells: (FormulaResult | null)[][] = [];

    for (let r = 0; r < rows; r++) {
      const row: (FormulaResult | null)[] = [];
      for (let c = 0; c < cols; c++) {
        const cellRef = `${indexToCol(cMin + c)}${r1 + r}`;
        if (sheetPrefix !== null) {
          row.push(this.resolveSheetCellResult(`${sheetPrefix}!${cellRef}`));
        } else {
          row.push(this.resolveCellResult(cellRef));
        }
      }
      cells.push(row);
    }

    return new RangeData(cells, rows, cols);
  }

  // ==================== Comparison ====================

  protected compareValues(a: FormulaResult, b: FormulaResult): number {
    if (a.isNumeric && b.isNumeric) return (a.numericValue ?? 0) - (b.numericValue ?? 0);
    if (a.isString && b.isString) return a.stringValue!.localeCompare(b.stringValue!, undefined, { sensitivity: "base" });
    return a.asNumber() - b.asNumber();
  }

  protected allArgs(args: unknown[]): FormulaResult[] {
    const results: FormulaResult[] = [];
    for (const a of args) {
      if (a instanceof RangeData) {
        for (let r = 0; r < a.rows; r++) {
          for (let c = 0; c < a.cols; c++) {
            results.push(a.cells[r][c] ?? FormulaResult.Number(0));
          }
        }
      } else if (Array.isArray(a)) {
        for (const v of a) {
          results.push(FormulaResult.Number(v as number));
        }
      } else if (a instanceof FormulaResult) {
        results.push(a);
      }
    }
    return results;
  }

  protected checkRangeErrors(args: unknown[]): FormulaResult | null {
    for (const a of args) {
      if (a instanceof RangeData) {
        const err = a.firstError();
        if (err) return err;
      } else if (a instanceof FormulaResult && a.isError) {
        return a;
      }
    }
    return null;
  }

  protected flattenNumbers(args: unknown[]): number[] {
    const result: number[] = [];
    for (const a of args) {
      if (a instanceof RangeData) {
        result.push(...a.toDoubleArray());
      } else if (Array.isArray(a)) {
        result.push(...(a as number[]));
      } else if (a instanceof FormulaResult) {
        if (a.isNumeric) result.push(a.numericValue!);
        else if (a.isBool) result.push(a.boolValue! ? 1 : 0);
      }
    }
    return result;
  }

  // ==================== Function Dispatch ====================

  protected evalFunction(name: string, args: unknown[]): FormulaResult | null {
    const nums = (): number[] => this.flattenNumbers(args);
    const arg = (i: number): FormulaResult | null => args[i] instanceof FormulaResult ? args[i] as FormulaResult : null;
    const num = (i: number): number => arg(i)?.asNumber() ?? 0;
    const str = (i: number): string => arg(i)?.asString() ?? "";

    switch (name) {
      // ===== Math & Aggregation =====
      case "SUM": {
        const err = this.checkRangeErrors(args);
        if (err) return err;
        return FormulaResult.Number(nums().reduce((a, b) => a + b, 0));
      }
      case "SUMPRODUCT": return this.evalSumProduct(args);
      case "AVERAGE": {
        const a = nums();
        return a.length > 0 ? FormulaResult.Number(a.reduce((x, y) => x + y, 0) / a.length) : null;
      }
      case "COUNT": return FormulaResult.Number(nums().length);
      case "COUNTA": {
        let count = 0;
        for (const a of args) {
          if (a instanceof FormulaResult && !a.isError && a.asString() !== "") count++;
          else if (Array.isArray(a)) count += (a as unknown[]).length;
        }
        return FormulaResult.Number(count);
      }
      case "COUNTBLANK": return FormulaResult.Number(0);
      case "MIN": {
        const a = nums();
        return a.length > 0 ? FormulaResult.Number(Math.min(...a)) : FormulaResult.Number(0);
      }
      case "MAX": {
        const a = nums();
        return a.length > 0 ? FormulaResult.Number(Math.max(...a)) : FormulaResult.Number(0);
      }
      case "ABS": return FormulaResult.Number(Math.abs(num(0)));
      case "SIGN": return FormulaResult.Number(Math.sign(num(0)));
      case "INT": return FormulaResult.Number(Math.floor(num(0)));
      case "TRUNC": {
        if (args.length >= 2) {
          const mult = Math.pow(10, num(1));
          return FormulaResult.Number(this.trunc(num(0) * mult) / mult);
        }
        return FormulaResult.Number(this.trunc(num(0)));
      }
      case "ROUND": return FormulaResult.Number(Math.round(num(0) * Math.pow(10, num(1))) / Math.pow(10, num(1)));
      case "ROUNDUP": return FormulaResult.Number(this.roundUp(num(0), num(1)));
      case "ROUNDDOWN": return FormulaResult.Number(this.roundDown(num(0), num(1)));
      case "CEILING":
      case "CEILING_MATH": return FormulaResult.Number(this.ceilingF(num(0), args.length >= 2 ? num(1) : 1));
      case "FLOOR":
      case "FLOOR_MATH": return FormulaResult.Number(this.floorF(num(0), args.length >= 2 ? num(1) : 1));
      case "MOD": {
        const divisor = num(1);
        return divisor !== 0 ? FormulaResult.Number(num(0) - num(1) * Math.floor(num(0) / num(1))) : FormulaResult.Error("#DIV/0!");
      }
      case "POWER": return FormulaResult.Number(Math.pow(num(0), num(1)));
      case "SQRT": return num(0) >= 0 ? FormulaResult.Number(Math.sqrt(num(0))) : FormulaResult.Error("#NUM!");
      case "FACT": return FormulaResult.Number(this.factorial(num(0)));
      case "COMBIN": return FormulaResult.Number(this.combin(num(0), num(1)));
      case "PERMUT": return FormulaResult.Number(this.permut(num(0), num(1)));
      case "GCD": return FormulaResult.Number(nums().reduce((a, b) => this.gcd(Math.round(a), Math.round(b)), 0));
      case "LCM": return FormulaResult.Number(nums().reduce((a, b) => this.lcm(Math.round(a), Math.round(b)), 1));
      case "RAND": return FormulaResult.Number(Math.random());
      case "RANDBETWEEN": return FormulaResult.Number(Math.floor(Math.random() * (num(1) - num(0) + 1)) + num(0));
      case "EVEN": return FormulaResult.Number(this.evenF(num(0)));
      case "ODD": return FormulaResult.Number(this.oddF(num(0)));
      case "PRODUCT": return FormulaResult.Number(nums().reduce((a, b) => a * b, 1));
      case "QUOTIENT": {
        const divisor = num(1);
        return divisor !== 0 ? FormulaResult.Number(this.trunc(num(0) / divisor)) : FormulaResult.Error("#DIV/0!");
      }
      case "MROUND": {
        const mult = num(1);
        return mult !== 0 ? FormulaResult.Number(Math.round(num(0) / mult) * mult) : FormulaResult.Error("#NUM!");
      }
      case "ROMAN": return FormulaResult.Str(this.toRoman(num(0)));
      case "ARABIC": return FormulaResult.Number(this.fromRoman(str(0)));
      case "BASE": return FormulaResult.Str(num(0).toString(num(1)).toUpperCase());
      case "DECIMAL": return FormulaResult.Number(parseInt(str(0), num(1)));
      case "LOG": return args.length >= 2 ? FormulaResult.Number(Math.log(num(0)) / Math.log(num(1))) : FormulaResult.Number(Math.log10(num(0)));
      case "LOG10": return FormulaResult.Number(Math.log10(num(0)));
      case "LN": return FormulaResult.Number(Math.log(num(0)));
      case "EXP": return FormulaResult.Number(Math.exp(num(0)));

      // ===== Trigonometry =====
      case "PI": return FormulaResult.Number(Math.PI);
      case "SIN": return FormulaResult.Number(Math.sin(num(0)));
      case "COS": return FormulaResult.Number(Math.cos(num(0)));
      case "TAN": return FormulaResult.Number(Math.tan(num(0)));
      case "ASIN": return FormulaResult.Number(Math.asin(num(0)));
      case "ACOS": return FormulaResult.Number(Math.acos(num(0)));
      case "ATAN": return FormulaResult.Number(Math.atan(num(0)));
      case "ATAN2": return FormulaResult.Number(Math.atan2(num(0), num(1)));
      case "SINH": return FormulaResult.Number(Math.sinh(num(0)));
      case "COSH": return FormulaResult.Number(Math.cosh(num(0)));
      case "TANH": return FormulaResult.Number(Math.tanh(num(0)));
      case "ASINH": return FormulaResult.Number(Math.asinh(num(0)));
      case "ACOSH": return FormulaResult.Number(Math.acosh(num(0)));
      case "ATANH": return FormulaResult.Number(Math.atanh(num(0)));
      case "DEGREES": return FormulaResult.Number(num(0) * 180.0 / Math.PI);
      case "RADIANS": return FormulaResult.Number(num(0) * Math.PI / 180.0);

      // ===== Statistical =====
      case "MEDIAN": return this.evalMedian(nums());
      case "MODE":
      case "MODE_SNGL": return this.evalMode(nums());
      case "LARGE": return this.evalLarge(args);
      case "SMALL": return this.evalSmall(args);
      case "RANK":
      case "RANK_EQ": return this.evalRank(args);
      case "PERCENTILE":
      case "PERCENTILE_INC": return this.evalPercentile(args);
      case "PERCENTRANK":
      case "PERCENTRANK_INC": return this.evalPercentRank(args);
      case "STDEV":
      case "STDEV_S": return this.evalStdev(nums(), true);
      case "STDEVP":
      case "STDEV_P": return this.evalStdev(nums(), false);
      case "VAR":
      case "VAR_S": return this.evalVar(nums(), true);
      case "VARP":
      case "VAR_P": return this.evalVar(nums(), false);
      case "GEOMEAN": {
        const a = nums();
        return a.length > 0 ? FormulaResult.Number(Math.pow(a.reduce((x, y) => x * y, 1), 1 / a.length)) : null;
      }
      case "HARMEAN": {
        const a = nums();
        return a.length > 0 ? FormulaResult.Number(a.length / a.reduce((x, y) => x + 1 / y, 0)) : null;
      }

      // ===== Logical =====
      case "IF": return this.evalIf(args);
      case "IFS": return this.evalIfs(args);
      case "AND": return FormulaResult.Bool(this.allArgs(args).every(r => r.asNumber() !== 0));
      case "OR": return FormulaResult.Bool(this.allArgs(args).some(r => r.asNumber() !== 0));
      case "NOT": return FormulaResult.Bool(num(0) === 0);
      case "XOR": return FormulaResult.Bool(this.allArgs(args).filter(r => r.asNumber() !== 0).length % 2 === 1);
      case "TRUE": return FormulaResult.Bool(true);
      case "FALSE": return FormulaResult.Bool(false);
      case "IFERROR":
      case "IFNA": return arg(0)?.isError ? arg(1) : arg(0);
      case "SWITCH": return this.evalSwitch(args);
      case "CHOOSE": return this.evalChoose(args);

      // ===== Text =====
      case "CONCATENATE":
      case "CONCAT": return FormulaResult.Str(this.allArgs(args).map(r => r.asString()).join(""));
      case "TEXTJOIN": return this.evalTextJoin(args);
      case "LEFT": {
        const s = str(0);
        const n = Math.floor(num(1));
        return FormulaResult.Str(s.length >= n ? s.substring(0, n) : s);
      }
      case "RIGHT": {
        const s = str(0);
        const n = Math.floor(num(1));
        return FormulaResult.Str(s.length >= n ? s.substring(s.length - n) : s);
      }
      case "MID": return this.evalMid(args);
      case "LEN": return FormulaResult.Number(str(0).length);
      case "TRIM": return FormulaResult.Str(str(0).trim().replace(/\s+/g, " "));
      case "CLEAN": return FormulaResult.Str(str(0).replace(/[\x00-\x1F]/g, ""));
      case "UPPER": return FormulaResult.Str(str(0).toUpperCase());
      case "LOWER": return FormulaResult.Str(str(0).toLowerCase());
      case "PROPER": return FormulaResult.Str(str(0).toLowerCase().replace(/\b\w/g, c => c.toUpperCase()));
      case "REPT": return FormulaResult.Str(str(0).repeat(Math.floor(num(1))));
      case "CHAR": return FormulaResult.Str(String.fromCharCode(Math.floor(num(0))));
      case "CODE": {
        const s = str(0);
        return FormulaResult.Number(s.length > 0 ? s.charCodeAt(0) : 0);
      }
      case "FIND": return this.evalFind(args, true);
      case "SEARCH": return this.evalFind(args, false);
      case "REPLACE": return this.evalReplace(args);
      case "SUBSTITUTE": return this.evalSubstitute(args);
      case "EXACT": return FormulaResult.Bool(str(0) === str(1));
      case "VALUE": {
        const v = parseFloat(str(0));
        return isNaN(v) ? FormulaResult.Error("#VALUE!") : FormulaResult.Number(v);
      }
      case "TEXT": return this.evalText(args);
      case "T": return arg(0)?.isString ? arg(0) : FormulaResult.Str("");
      case "N": return FormulaResult.Number(num(0));
      case "FIXED": return this.evalFixed(args);
      case "NUMBERVALUE": return this.evalNumberValue(args);
      case "DOLLAR":
      case "YEN": return FormulaResult.Str(num(0).toLocaleString("en-US", { style: "currency", currency: name === "YEN" ? "JPY" : "USD" }));

      // ===== Lookup & Reference =====
      case "INDEX": return this.evalIndex(args);
      case "MATCH": return this.evalMatch(args);
      case "ROW": return this.evalRowCol(args, true);
      case "COLUMN": return this.evalRowCol(args, false);
      case "ROWS": {
        if (args.length > 0 && args[0] instanceof RangeData) return FormulaResult.Number(args[0].rows);
        if (args.length > 0 && Array.isArray(args[0])) return FormulaResult.Number((args[0] as unknown[]).length);
        return FormulaResult.Number(1);
      }
      case "COLUMNS": {
        if (args.length > 0 && args[0] instanceof RangeData) return FormulaResult.Number(args[0].cols);
        if (args.length > 0 && Array.isArray(args[0])) return FormulaResult.Number(1);
        return FormulaResult.Number(1);
      }
      case "ADDRESS": return this.evalAddress(args);
      case "VLOOKUP": return this.evalVlookup(args);
      case "HLOOKUP": return this.evalHlookup(args);

      // ===== Date & Time =====
      case "TODAY": return FormulaResult.Number(this.dateToOaDate(new Date()));
      case "NOW": return FormulaResult.Number(this.dateToOaDate(new Date()));
      case "DATE": return FormulaResult.Number(this.dateToOaDate(new Date(Math.floor(num(0)), Math.floor(num(1)) - 1, Math.floor(num(2)))));
      case "YEAR": return FormulaResult.Number(this.oaDateToDate(num(0)).getFullYear());
      case "MONTH": return FormulaResult.Number(this.oaDateToDate(num(0)).getMonth() + 1);
      case "DAY": return FormulaResult.Number(this.oaDateToDate(num(0)).getDate());
      case "HOUR": return FormulaResult.Number(this.oaDateToDate(num(0)).getHours());
      case "MINUTE": return FormulaResult.Number(this.oaDateToDate(num(0)).getMinutes());
      case "SECOND": return FormulaResult.Number(this.oaDateToDate(num(0)).getSeconds());
      case "WEEKDAY": return FormulaResult.Number(this.oaDateToDate(num(0)).getDay() + 1);
      case "DATEVALUE": {
        const d = new Date(str(0));
        return isNaN(d.getTime()) ? FormulaResult.Error("#VALUE!") : FormulaResult.Number(this.dateToOaDate(d));
      }
      case "TIMEVALUE": {
        const d = new Date(`1970-01-01T${str(0)}`);
        return isNaN(d.getTime()) ? FormulaResult.Error("#VALUE!") : FormulaResult.Number(this.dateToOaDate(d));
      }
      case "EDATE": {
        const d = this.oaDateToDate(num(0));
        d.setMonth(d.getMonth() + Math.floor(num(1)));
        return FormulaResult.Number(this.dateToOaDate(d));
      }
      case "EOMONTH": return this.evalEomonth(args);
      case "DAYS": return FormulaResult.Number(num(0) - num(1));
      case "DATEDIF": return this.evalDateDif(args);
      case "NETWORKDAYS":
      case "NETWORKDAYS_INTL": return this.evalNetworkDays(args);
      case "WORKDAY":
      case "WORKDAY_INTL": return this.evalWorkDay(args);
      case "ISOWEEKNUM": {
        const d = this.oaDateToDate(num(0));
        const jan1 = new Date(d.getFullYear(), 0, 1);
        const days = Math.floor((d.getTime() - jan1.getTime()) / 86400000);
        return FormulaResult.Number(Math.ceil((days + jan1.getDay() + 1) / 7));
      }
      case "YEARFRAC": return this.evalYearFrac(args);

      // ===== Info =====
      case "ISNUMBER": return FormulaResult.Bool(arg(0)?.isNumeric === true);
      case "ISTEXT": return FormulaResult.Bool(arg(0)?.isString === true);
      case "ISBLANK": return FormulaResult.Bool(arg(0) === null || arg(0) === undefined || (arg(0)?.asString() === "" && !arg(0)?.isNumeric));
      case "ISERROR":
      case "ISERR": {
        if (args.length > 0 && args[0] instanceof RangeData) {
          return FormulaResult.Array(args[0].toFlatResults().map(r => r?.isError ? 1.0 : 0.0));
        }
        return FormulaResult.Bool(arg(0)?.isError === true);
      }
      case "ISNA": return FormulaResult.Bool(arg(0)?.errorValue === "#N/A");
      case "ISLOGICAL": return FormulaResult.Bool(arg(0)?.isBool === true);
      case "ISEVEN": return FormulaResult.Bool(Math.floor(num(0)) % 2 === 0);
      case "ISODD": return FormulaResult.Bool(Math.floor(num(0)) % 2 !== 0);
      case "ISNONTEXT": return FormulaResult.Bool(arg(0)?.isString !== true);
      case "TYPE": {
        const a = arg(0);
        if (a?.isNumeric) return FormulaResult.Number(1);
        if (a?.isString) return FormulaResult.Number(2);
        if (a?.isBool) return FormulaResult.Number(4);
        if (a?.isError) return FormulaResult.Number(16);
        return FormulaResult.Number(1);
      }
      case "NA": return FormulaResult.Error("#N/A");
      case "ERROR_TYPE": {
        const err = arg(0)?.errorValue;
        const typeMap: Record<string, number> = { "#NULL!": 1, "#DIV/0!": 2, "#VALUE!": 3, "#REF!": 4, "#NAME?": 5, "#NUM!": 6, "#N/A": 7 };
        return FormulaResult.Number(err ? (typeMap[err] ?? 0) : 0);
      }

      // ===== Conditional Aggregation =====
      case "SUMIF": return this.evalSumIf(args);
      case "SUMIFS": return this.evalSumIfs(args);
      case "COUNTIF": return this.evalCountIf(args);
      case "COUNTIFS": return this.evalCountIfs(args);
      case "AVERAGEIF": return this.evalAverageIf(args);
      case "AVERAGEIFS": return this.evalAverageIfs(args);
      case "MAXIFS": return this.evalMaxMinIfs(args, true);
      case "MINIFS": return this.evalMaxMinIfs(args, false);

      // ===== Financial =====
      case "PMT": return this.evalPmt(args);
      case "FV": return this.evalFv(args);
      case "PV": return this.evalPv(args);
      case "NPER": return this.evalNper(args);
      case "NPV": return this.evalNpv(args);
      case "IPMT": return this.evalIpmt(args);
      case "PPMT": return this.evalPpmt(args);
      case "SLN": return args.length >= 3 ? FormulaResult.Number((num(0) - num(1)) / num(2)) : null;
      case "SYD": return this.evalSyd(args);
      case "DB": return this.evalDb(args);
      case "DDB": return this.evalDdb(args);

      // ===== Conversion =====
      case "BIN2DEC": return FormulaResult.Number(parseInt(str(0), 2));
      case "DEC2BIN": return FormulaResult.Str(num(0).toString(2));
      case "HEX2DEC": return FormulaResult.Number(parseInt(str(0), 16));
      case "DEC2HEX": return FormulaResult.Str(Math.floor(num(0)).toString(16).toUpperCase());
      case "OCT2DEC": return FormulaResult.Number(parseInt(str(0), 8));
      case "DEC2OCT": return FormulaResult.Str(Math.floor(num(0)).toString(8));
      case "BIN2HEX": return FormulaResult.Str(parseInt(str(0), 2).toString(16).toUpperCase());
      case "BIN2OCT": return FormulaResult.Str(parseInt(str(0), 2).toString(8));
      case "HEX2BIN": return FormulaResult.Str(parseInt(str(0), 16).toString(2));
      case "HEX2OCT": return FormulaResult.Str(parseInt(str(0), 16).toString(8));
      case "OCT2BIN": return FormulaResult.Str(parseInt(str(0), 8).toString(2));
      case "OCT2HEX": return FormulaResult.Str(parseInt(str(0), 8).toString(16).toUpperCase());

      default:
        return null;
    }
  }

  // ==================== Logical ====================

  private evalIf(args: unknown[]): FormulaResult | null {
    const c = args[0] instanceof FormulaResult ? args[0] : null;
    if (!c) return null;
    const isTrue = c.isNumeric ? c.numericValue !== 0 : c.boolValue === true;
    if (isTrue) return args.length > 1 && args[1] instanceof FormulaResult ? args[1] : FormulaResult.Number(0);
    return args.length > 2 && args[2] instanceof FormulaResult ? args[2] : FormulaResult.Bool(false);
  }

  private evalIfs(args: unknown[]): FormulaResult | null {
    for (let i = 0; i + 1 < args.length; i += 2) {
      const c = args[i] instanceof FormulaResult ? args[i] as FormulaResult : null;
      if (c && c.asNumber() !== 0) return args[i + 1] instanceof FormulaResult ? args[i + 1] as FormulaResult : null;
    }
    return FormulaResult.Error("#N/A");
  }

  private evalSwitch(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const val = args[0] instanceof FormulaResult ? args[0] as FormulaResult : null;
    if (!val) return null;
    for (let i = 1; i + 1 < args.length; i += 2) {
      const cv = args[i] instanceof FormulaResult ? args[i] as FormulaResult : null;
      if (cv && this.compareValues(val, cv) === 0) return args[i + 1] instanceof FormulaResult ? args[i + 1] as FormulaResult : null;
    }
    return args.length % 2 === 0 && args[args.length - 1] instanceof FormulaResult ? args[args.length - 1] as FormulaResult : FormulaResult.Error("#N/A");
  }

  private evalChoose(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const idx = args[0] instanceof FormulaResult ? Math.floor(args[0].asNumber()) : 1;
    return idx >= 1 && idx < args.length && args[idx] instanceof FormulaResult ? args[idx] : FormulaResult.Error("#VALUE!");
  }

  // ==================== Text ====================

  private evalMid(args: unknown[]): FormulaResult | null {
    const s = args.length > 0 && args[0] instanceof FormulaResult ? args[0].asString() : "";
    const start = args.length > 1 && args[1] instanceof FormulaResult ? Math.floor(args[1].asNumber()) - 1 : 0;
    const len = args.length > 2 && args[2] instanceof FormulaResult ? Math.floor(args[2].asNumber()) : 0;
    if (start < 0 || start >= s.length) return FormulaResult.Str("");
    return FormulaResult.Str(s.substring(start, Math.min(start + len, s.length)));
  }

  private evalFind(args: unknown[], caseSensitive: boolean): FormulaResult | null {
    const find = args.length > 0 && args[0] instanceof FormulaResult ? args[0].asString() : "";
    const within = args.length > 1 && args[1] instanceof FormulaResult ? args[1].asString() : "";
    const startPos = args.length > 2 && args[2] instanceof FormulaResult ? Math.floor(args[2].asNumber()) - 1 : 0;
    const idx = within.indexOf(find, startPos);
    return idx >= 0 ? FormulaResult.Number(idx + 1) : FormulaResult.Error("#VALUE!");
  }

  private evalReplace(args: unknown[]): FormulaResult | null {
    const s = args.length > 0 && args[0] instanceof FormulaResult ? args[0].asString() : "";
    const start = args.length > 1 && args[1] instanceof FormulaResult ? Math.floor(args[1].asNumber()) - 1 : 0;
    const len = args.length > 2 && args[2] instanceof FormulaResult ? Math.floor(args[2].asNumber()) : 0;
    const rep = args.length > 3 && args[3] instanceof FormulaResult ? args[3].asString() : "";
    if (start < 0 || start > s.length) return FormulaResult.Error("#VALUE!");
    return FormulaResult.Str(s.substring(0, start) + rep + s.substring(Math.min(start + len, s.length)));
  }

  private evalSubstitute(args: unknown[]): FormulaResult | null {
    const s = args.length > 0 && args[0] instanceof FormulaResult ? args[0].asString() : "";
    const old = args.length > 1 && args[1] instanceof FormulaResult ? args[1].asString() : "";
    const neo = args.length > 2 && args[2] instanceof FormulaResult ? args[2].asString() : "";
    if (args.length > 3 && args[3] instanceof FormulaResult) {
      const n = Math.floor(args[3].asNumber());
      let idx = -1;
      for (let i = 0; i < n; i++) {
        idx = s.indexOf(old, idx + 1);
        if (idx < 0) return FormulaResult.Str(s);
      }
      return FormulaResult.Str(s.substring(0, idx) + neo + s.substring(idx + old.length));
    }
    return FormulaResult.Str(s.split(old).join(neo));
  }

  private evalText(args: unknown[]): FormulaResult | null {
    const val = args.length > 0 && args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const fmt = args.length > 1 && args[1] instanceof FormulaResult ? args[1].asString() : "0";
    try {
      return FormulaResult.Str(val.toFixed(fmt.replace(/[^0-9]/g, "").length || 0));
    } catch {
      return FormulaResult.Str(val.toString());
    }
  }

  private evalFixed(args: unknown[]): FormulaResult | null {
    const v = args.length > 0 && args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const d = args.length > 1 && args[1] instanceof FormulaResult ? Math.floor(args[1].asNumber()) : 2;
    return FormulaResult.Str(v.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }));
  }

  private evalNumberValue(args: unknown[]): FormulaResult | null {
    let s = args.length > 0 && args[0] instanceof FormulaResult ? args[0].asString() : "";
    s = s.replace(",", "").replace(" ", "").trim();
    const v = parseFloat(s);
    return isNaN(v) ? FormulaResult.Error("#VALUE!") : FormulaResult.Number(v);
  }

  private evalTextJoin(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const delim = args[0] instanceof FormulaResult ? args[0].asString() : "";
    const ignoreEmpty = args[1] instanceof FormulaResult && args[1].asNumber() !== 0;
    const parts: string[] = [];
    for (let i = 2; i < args.length; i++) {
      const arg = args[i];
      if (arg instanceof RangeData) {
        const rd = arg as RangeData;
        for (let row = 0; row < rd.rows; row++) {
          for (let col = 0; col < rd.cols; col++) {
            const cv = rd.cells[row][col];
            if (cv) {
              const s = cv.asString();
              if (!ignoreEmpty || s !== "") parts.push(s);
            }
          }
        }
      } else if (Array.isArray(arg)) {
        for (const v of arg as unknown[]) parts.push(String(v));
      } else if (arg instanceof FormulaResult) {
        const fr = arg as FormulaResult;
        const s = fr.asString();
        if (!ignoreEmpty || s !== "") parts.push(s);
      }
    }
    return FormulaResult.Str(parts.join(delim));
  }

  // ==================== Lookup ====================

  private evalIndex(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    if (args[0] instanceof RangeData) {
      const rowIdx = args[1] instanceof FormulaResult ? Math.floor(args[1].asNumber()) : 0;
      const colIdx = args.length > 2 && args[2] instanceof FormulaResult ? Math.floor(args[2].asNumber()) : 1;
      if (rowIdx < 1 || rowIdx > args[0].rows || colIdx < 1 || colIdx > args[0].cols) return FormulaResult.Error("#REF!");
      return args[0].cells[rowIdx - 1][colIdx - 1] ?? FormulaResult.Number(0);
    }
    if (Array.isArray(args[0])) {
      const idx = args[1] instanceof FormulaResult ? Math.floor(args[1].asNumber()) - 1 : 0;
      const arr = args[0] as number[];
      return idx >= 0 && idx < arr.length ? FormulaResult.Number(arr[idx]) : FormulaResult.Error("#REF!");
    }
    return null;
  }

  private evalMatch(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const lookup = args[0] instanceof FormulaResult ? args[0] : null;
    if (!lookup) return null;
    if (args[1] instanceof RangeData) {
      if (args[1].cols === 1) {
        for (let i = 0; i < args[1].rows; i++) {
          const cell = args[1].cells[i][0];
          if (cell && this.compareValues(cell, lookup) === 0) return FormulaResult.Number(i + 1);
        }
      } else if (args[1].rows === 1) {
        for (let i = 0; i < args[1].cols; i++) {
          const cell = args[1].cells[0][i];
          if (cell && this.compareValues(cell, lookup) === 0) return FormulaResult.Number(i + 1);
        }
      }
    }
    if (Array.isArray(args[1])) {
      const arr = args[1] as number[];
      for (let i = 0; i < arr.length; i++) {
        if (Math.abs(arr[i] - lookup.asNumber()) < 1e-10) return FormulaResult.Number(i + 1);
      }
    }
    return FormulaResult.Error("#N/A");
  }

  private evalRowCol(args: unknown[], isRow: boolean): FormulaResult | null {
    if (args.length === 0) return null;
    if (args[0] instanceof FormulaResult) {
      const ref = args[0].asString().match(/([A-Z]+)(\d+)/i);
      if (ref) {
        return FormulaResult.Number(isRow ? parseInt(ref[2], 10) : colToIndex(ref[1]));
      }
    }
    return null;
  }

  private evalAddress(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const row = Math.floor(args[0] instanceof FormulaResult ? args[0].asNumber() : 1);
    const col = Math.floor(args[1] instanceof FormulaResult ? args[1].asNumber() : 1);
    const abs = args.length > 2 && args[2] instanceof FormulaResult ? Math.floor(args[2].asNumber()) : 1;
    const cs = indexToCol(col);
    switch (abs) {
      case 1: return FormulaResult.Str(`$${cs}$${row}`);
      case 2: return FormulaResult.Str(`${cs}$${row}`);
      case 3: return FormulaResult.Str(`$${cs}${row}`);
      default: return FormulaResult.Str(`${cs}${row}`);
    }
  }

  private evalVlookup(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const lookupVal = args[0] instanceof FormulaResult ? args[0] : null;
    if (!lookupVal) return FormulaResult.Error("#N/A");
    const table = args[1] instanceof RangeData ? args[1] : null;
    if (!table) return FormulaResult.Error("#N/A");
    const colIndex = args[2] instanceof FormulaResult ? Math.floor(args[2].asNumber()) : 0;
    if (colIndex < 1 || colIndex > table.cols) return FormulaResult.Error("#REF!");
    const exactMatch = args.length > 3 && args[3] instanceof FormulaResult && (args[3].asNumber() === 0 || args[3].asString().toUpperCase() === "FALSE");

    let foundRow = -1;
    if (exactMatch) {
      for (let i = 0; i < table.rows; i++) {
        const cell = table.cells[i][0];
        if (cell && this.compareValues(cell, lookupVal) === 0) { foundRow = i; break; }
      }
    } else {
      for (let i = 0; i < table.rows; i++) {
        const cell = table.cells[i][0];
        if (cell === null) continue;
        if (this.compareValues(cell, lookupVal) <= 0) foundRow = i;
        else break;
      }
    }

    return foundRow >= 0 ? (table.cells[foundRow][colIndex - 1] ?? FormulaResult.Number(0)) : FormulaResult.Error("#N/A");
  }

  private evalHlookup(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const lookupVal = args[0] instanceof FormulaResult ? args[0] : null;
    if (!lookupVal) return FormulaResult.Error("#N/A");
    const table = args[1] instanceof RangeData ? args[1] : null;
    if (!table) return FormulaResult.Error("#N/A");
    const rowIndex = args[2] instanceof FormulaResult ? Math.floor(args[2].asNumber()) : 0;
    if (rowIndex < 1 || rowIndex > table.rows) return FormulaResult.Error("#REF!");
    const exactMatch = args.length > 3 && args[3] instanceof FormulaResult && (args[3].asNumber() === 0 || args[3].asString().toUpperCase() === "FALSE");

    let foundCol = -1;
    if (exactMatch) {
      for (let i = 0; i < table.cols; i++) {
        const cell = table.cells[0][i];
        if (cell && this.compareValues(cell, lookupVal) === 0) { foundCol = i; break; }
      }
    } else {
      for (let i = 0; i < table.cols; i++) {
        const cell = table.cells[0][i];
        if (cell === null) continue;
        if (this.compareValues(cell, lookupVal) <= 0) foundCol = i;
        else break;
      }
    }

    return foundCol >= 0 ? (table.cells[rowIndex - 1][foundCol] ?? FormulaResult.Number(0)) : FormulaResult.Error("#N/A");
  }

  // ==================== Statistical ====================

  private evalMedian(v: number[]): FormulaResult | null {
    if (v.length === 0) return null;
    const s = v.slice().sort((a, b) => a - b);
    return FormulaResult.Number(s.length % 2 === 1 ? s[Math.floor(s.length / 2)] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2.0);
  }

  private evalMode(v: number[]): FormulaResult | null {
    if (v.length === 0) return null;
    const freq = new Map<number, number>();
    for (const x of v) freq.set(x, (freq.get(x) ?? 0) + 1);
    const top = [...freq.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0])[0];
    return top[1] > 1 ? FormulaResult.Number(top[0]) : FormulaResult.Error("#N/A");
  }

  private evalLarge(args: unknown[]): FormulaResult | null {
    const arr = args.length > 0 && Array.isArray(args[0]) ? args[0] as number[] : null;
    const k = args.length > 1 && args[1] instanceof FormulaResult ? Math.floor(args[1].asNumber()) : 1;
    if (!arr || k < 1 || k > arr.length) return FormulaResult.Error("#NUM!");
    const sorted = arr.slice().sort((a, b) => b - a);
    return FormulaResult.Number(sorted[k - 1]);
  }

  private evalSmall(args: unknown[]): FormulaResult | null {
    const arr = args.length > 0 && Array.isArray(args[0]) ? args[0] as number[] : null;
    const k = args.length > 1 && args[1] instanceof FormulaResult ? Math.floor(args[1].asNumber()) : 1;
    if (!arr || k < 1 || k > arr.length) return FormulaResult.Error("#NUM!");
    const sorted = arr.slice().sort((a, b) => a - b);
    return FormulaResult.Number(sorted[k - 1]);
  }

  private evalRank(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const val = args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const arr = args[1] && Array.isArray(args[1]) ? args[1] as number[] : null;
    if (!arr) return null;
    const order = args.length > 2 && args[2] instanceof FormulaResult ? Math.floor(args[2].asNumber()) : 0;
    const sorted = order === 0 ? arr.slice().sort((a, b) => b - a) : arr.slice().sort((a, b) => a - b);
    for (let i = 0; i < sorted.length; i++) if (Math.abs(sorted[i] - val) < 1e-10) return FormulaResult.Number(i + 1);
    return FormulaResult.Error("#N/A");
  }

  private evalPercentile(args: unknown[]): FormulaResult | null {
    const arr = args.length > 0 && Array.isArray(args[0]) ? args[0] as number[] : null;
    const k = args.length > 1 && args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    if (!arr || arr.length === 0 || k < 0 || k > 1) return FormulaResult.Error("#NUM!");
    const sorted = arr.slice().sort((a, b) => a - b);
    const idx = k * (sorted.length - 1);
    const lower = Math.floor(idx);
    const upper = Math.min(lower + 1, sorted.length - 1);
    return FormulaResult.Number(sorted[lower] + (idx - lower) * (sorted[upper] - sorted[lower]));
  }

  private evalPercentRank(args: unknown[]): FormulaResult | null {
    const arr = args.length > 0 && Array.isArray(args[0]) ? args[0] as number[] : null;
    const val = args.length > 1 && args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    if (!arr || arr.length === 0) return FormulaResult.Error("#NUM!");
    return FormulaResult.Number(arr.filter(x => x < val).length / (arr.length - 1));
  }

  private evalStdev(v: number[], sample: boolean): FormulaResult | null {
    if (v.length < (sample ? 2 : 1)) return FormulaResult.Error("#DIV/0!");
    const mean = v.reduce((a, b) => a + b, 0) / v.length;
    const sumSq = v.reduce((a, b) => a + (b - mean) * (b - mean), 0);
    return FormulaResult.Number(Math.sqrt(sumSq / (sample ? v.length - 1 : v.length)));
  }

  private evalVar(v: number[], sample: boolean): FormulaResult | null {
    if (v.length < (sample ? 2 : 1)) return FormulaResult.Error("#DIV/0!");
    const mean = v.reduce((a, b) => a + b, 0) / v.length;
    return FormulaResult.Number(v.reduce((a, b) => a + (b - mean) * (b - mean), 0) / (sample ? v.length - 1 : v.length));
  }

  // ==================== Conditional Aggregation ====================

  private asDoubles(a: unknown): number[] | null {
    if (a instanceof RangeData) return a.toDoubleArray();
    if (Array.isArray(a)) return a as number[];
    return null;
  }

  private matchesCriteria(value: number, criteria: string): boolean {
    return this.matchesCriteriaInternal(FormulaResult.Number(value), criteria);
  }

  private matchesCriteriaInternal(cellValue: FormulaResult | null, criteria: string): boolean {
    criteria = criteria.trim();
    if (criteria === "") return true;

    const numVal = cellValue?.asNumber() ?? 0;

    if (criteria.startsWith(">=")) {
      const n = parseFloat(criteria.substring(2));
      if (!isNaN(n)) return numVal >= n;
    }
    if (criteria.startsWith("<=")) {
      const n = parseFloat(criteria.substring(2));
      if (!isNaN(n)) return numVal <= n;
    }
    if (criteria.startsWith("<>")) {
      const operand = criteria.substring(2);
      const n = parseFloat(operand);
      if (!isNaN(n)) return Math.abs(numVal - n) > 1e-10;
      return cellValue?.asString().toLowerCase() !== operand.toLowerCase();
    }
    if (criteria.startsWith(">")) {
      const n = parseFloat(criteria.substring(1));
      if (!isNaN(n)) return numVal > n;
    }
    if (criteria.startsWith("<")) {
      const n = parseFloat(criteria.substring(1));
      if (!isNaN(n)) return numVal < n;
    }
    if (criteria.startsWith("=")) {
      const operand = criteria.substring(1);
      const n = parseFloat(operand);
      if (!isNaN(n)) return Math.abs(numVal - n) < 1e-10;
      return cellValue?.asString().toLowerCase() === operand.toLowerCase();
    }

    const plain = parseFloat(criteria);
    if (!isNaN(plain)) return Math.abs(numVal - plain) < 1e-10;

    const cellStr = cellValue?.asString() ?? "";
    if (criteria.includes("*") || criteria.includes("?")) {
      const pattern = criteria.replace(/[*?]/g, (c) => c === "*" ? ".*" : ".").replace(/~\*/g, "*").replace(/~\?/g, "?");
      return new RegExp(`^${pattern}$`, "i").test(cellStr);
    }

    return cellStr.toLowerCase() === criteria.toLowerCase();
  }

  private evalSumIf(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const range = this.asDoubles(args[0]);
    const criteria = args[1] instanceof FormulaResult ? args[1].asString() : "";
    const sumRange = args.length > 2 ? (this.asDoubles(args[2]) ?? range) : range;
    if (!range || !sumRange) return null;
    let sum = 0;
    for (let i = 0; i < range.length && i < sumRange.length; i++) {
      if (this.matchesCriteria(range[i], criteria)) sum += sumRange[i];
    }
    return FormulaResult.Number(sum);
  }

  private evalSumIfs(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const sumRange = this.asDoubles(args[0]);
    if (!sumRange) return null;
    let sum = 0;
    for (let i = 0; i < sumRange.length; i++) {
      let match = true;
      for (let c = 1; c + 1 < args.length; c += 2) {
        const cr = this.asDoubles(args[c]);
        const critArg = args[c + 1];
        const crit = critArg instanceof FormulaResult ? (critArg as FormulaResult).asString() : "";
        if (!cr || i >= cr.length || !this.matchesCriteria(cr[i], crit)) { match = false; break; }
      }
      if (match) sum += sumRange[i];
    }
    return FormulaResult.Number(sum);
  }

  private evalCountIf(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const range = this.asDoubles(args[0]);
    const criteriaArg = args[1];
    const criteria = criteriaArg instanceof FormulaResult ? (criteriaArg as FormulaResult).asString() : "";
    return range ? FormulaResult.Number(range.filter(v => this.matchesCriteria(v, criteria)).length) : null;
  }

  private evalCountIfs(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const first = this.asDoubles(args[0]);
    if (!first) return null;
    let count = 0;
    for (let i = 0; i < first.length; i++) {
      let match = true;
      for (let c = 0; c + 1 < args.length; c += 2) {
        const cr = this.asDoubles(args[c]);
        const critArg = args[c + 1];
        const crit = critArg instanceof FormulaResult ? (critArg as FormulaResult).asString() : "";
        if (!cr || i >= cr.length || !this.matchesCriteria(cr[i], crit)) { match = false; break; }
      }
      if (match) count++;
    }
    return FormulaResult.Number(count);
  }

  private evalAverageIf(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const range = this.asDoubles(args[0]);
    const criteriaArg = args[1];
    const criteria = criteriaArg instanceof FormulaResult ? (criteriaArg as FormulaResult).asString() : "";
    const avgRange = args.length > 2 ? (this.asDoubles(args[2]) ?? range) : range;
    if (!range || !avgRange) return null;
    const vals: number[] = [];
    for (let i = 0; i < range.length && i < avgRange.length; i++) {
      if (this.matchesCriteria(range[i], criteria)) vals.push(avgRange[i]);
    }
    return vals.length > 0 ? FormulaResult.Number(vals.reduce((a, b) => a + b, 0) / vals.length) : FormulaResult.Error("#DIV/0!");
  }

  private evalAverageIfs(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const avgRange = this.asDoubles(args[0]);
    if (!avgRange) return null;
    const vals: number[] = [];
    for (let i = 0; i < avgRange.length; i++) {
      let match = true;
      for (let c = 1; c + 1 < args.length; c += 2) {
        const cr = this.asDoubles(args[c]);
        const critArg = args[c + 1];
        const crit = critArg instanceof FormulaResult ? (critArg as FormulaResult).asString() : "";
        if (!cr || i >= cr.length || !this.matchesCriteria(cr[i], crit)) { match = false; break; }
      }
      if (match) vals.push(avgRange[i]);
    }
    return vals.length > 0 ? FormulaResult.Number(vals.reduce((a, b) => a + b, 0) / vals.length) : FormulaResult.Error("#DIV/0!");
  }

  private evalMaxMinIfs(args: unknown[], isMax: boolean): FormulaResult | null {
    if (args.length < 3) return null;
    const valRange = this.asDoubles(args[0]);
    if (!valRange) return null;
    const vals: number[] = [];
    for (let i = 0; i < valRange.length; i++) {
      let match = true;
      for (let c = 1; c + 1 < args.length; c += 2) {
        const cr = this.asDoubles(args[c]);
        const critArg = args[c + 1];
        const crit = critArg instanceof FormulaResult ? (critArg as FormulaResult).asString() : "";
        if (!cr || i >= cr.length || !this.matchesCriteria(cr[i], crit)) { match = false; break; }
      }
      if (match) vals.push(valRange[i]);
    }
    return vals.length > 0 ? FormulaResult.Number(isMax ? Math.max(...vals) : Math.min(...vals)) : FormulaResult.Number(0);
  }

  private evalSumProduct(args: unknown[]): FormulaResult | null {
    if (args.length === 0) return FormulaResult.Number(0);
    const arrays = args.map(a => {
      if (a instanceof RangeData) return a.toDoubleArray();
      if (Array.isArray(a)) return a;
      if (a instanceof FormulaResult && a.isNumeric) return [a.numericValue!];
      return null;
    });
    if (arrays.every(a => a === null) && args.length === 1 && args[0] instanceof FormulaResult && args[0].isNumeric) {
      return args[0];
    }
    if (arrays.some(a => a === null)) return null;
    const len = Math.min(...arrays.map(a => a!.length));
    let sum = 0;
    for (let i = 0; i < len; i++) {
      let p = 1;
      for (const arr of arrays) p *= arr![i];
      sum += p;
    }
    return FormulaResult.Number(sum);
  }

  // ==================== Date ====================

  private dateToOaDate(d: Date): number {
    const start = new Date(1899, 11, 30);
    return (d.getTime() - start.getTime()) / 86400000;
  }

  private oaDateToDate(oa: number): Date {
    const start = new Date(1899, 11, 30);
    return new Date(start.getTime() + oa * 86400000);
  }

  private evalEomonth(args: unknown[]): FormulaResult | null {
    const d = args.length > 0 && args[0] instanceof FormulaResult ? this.oaDateToDate(args[0].asNumber()) : new Date();
    const months = args.length > 1 && args[1] instanceof FormulaResult ? Math.floor(args[1].asNumber()) : 0;
    d.setMonth(d.getMonth() + months);
    const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
    d.setDate(lastDay);
    return FormulaResult.Number(this.dateToOaDate(d));
  }

  private evalDateDif(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const d1 = args[0] instanceof FormulaResult ? this.oaDateToDate(args[0].asNumber()) : new Date();
    const d2 = args[1] instanceof FormulaResult ? this.oaDateToDate(args[1].asNumber()) : new Date();
    const unit = args[2] instanceof FormulaResult ? args[2].asString().toUpperCase() : "D";
    switch (unit) {
      case "D": return FormulaResult.Number(Math.floor((d2.getTime() - d1.getTime()) / 86400000));
      case "M": return FormulaResult.Number((d2.getFullYear() - d1.getFullYear()) * 12 + d2.getMonth() - d1.getMonth());
      case "Y": return FormulaResult.Number(d2.getFullYear() - d1.getFullYear());
      default: return null;
    }
  }

  private evalNetworkDays(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const start = args[0] instanceof FormulaResult ? this.oaDateToDate(args[0].asNumber()) : new Date();
    const end = args[1] instanceof FormulaResult ? this.oaDateToDate(args[1].asNumber()) : new Date();
    let count = 0;
    const d = new Date(start);
    while (d <= end) {
      const day = d.getDay();
      if (day !== 0 && day !== 6) count++;
      d.setDate(d.getDate() + 1);
    }
    return FormulaResult.Number(count);
  }

  private evalWorkDay(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const start = args[0] instanceof FormulaResult ? this.oaDateToDate(args[0].asNumber()) : new Date();
    const days = args[1] instanceof FormulaResult ? Math.floor(args[1].asNumber()) : 0;
    let d = new Date(start);
    const step = days > 0 ? 1 : -1;
    let rem = Math.abs(days);
    while (rem > 0) {
      d.setDate(d.getDate() + step);
      const day = d.getDay();
      if (day !== 0 && day !== 6) rem--;
    }
    return FormulaResult.Number(this.dateToOaDate(d));
  }

  private evalYearFrac(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const d1 = args[0] instanceof FormulaResult ? this.oaDateToDate(args[0].asNumber()) : new Date();
    const d2 = args[1] instanceof FormulaResult ? this.oaDateToDate(args[1].asNumber()) : new Date();
    return FormulaResult.Number(Math.abs((d2.getTime() - d1.getTime()) / 31557600000));
  }

  // ==================== Financial ====================

  private evalPmt(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const rate = args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const nper = args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    const pv = args[2] instanceof FormulaResult ? args[2].asNumber() : 0;
    const fv = args.length > 3 && args[3] instanceof FormulaResult ? args[3].asNumber() : 0;
    if (rate === 0) return FormulaResult.Number(-(pv + fv) / nper);
    return FormulaResult.Number(-(rate * (pv * Math.pow(1 + rate, nper) + fv)) / (Math.pow(1 + rate, nper) - 1));
  }

  private evalFv(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const rate = args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const nper = args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    const pmt = args[2] instanceof FormulaResult ? args[2].asNumber() : 0;
    const pv = args.length > 3 && args[3] instanceof FormulaResult ? args[3].asNumber() : 0;
    if (rate === 0) return FormulaResult.Number(-(pv + pmt * nper));
    return FormulaResult.Number(-(pv * Math.pow(1 + rate, nper) + pmt * (Math.pow(1 + rate, nper) - 1) / rate));
  }

  private evalPv(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const rate = args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const nper = args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    const pmt = args[2] instanceof FormulaResult ? args[2].asNumber() : 0;
    const fv = args.length > 3 && args[3] instanceof FormulaResult ? args[3].asNumber() : 0;
    if (rate === 0) return FormulaResult.Number(-(fv + pmt * nper));
    return FormulaResult.Number(-(fv / Math.pow(1 + rate, nper) + pmt * (1 - Math.pow(1 + rate, -nper)) / rate));
  }

  private evalNper(args: unknown[]): FormulaResult | null {
    if (args.length < 3) return null;
    const rate = args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const pmt = args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    const pv = args[2] instanceof FormulaResult ? args[2].asNumber() : 0;
    const fv = args.length > 3 && args[3] instanceof FormulaResult ? args[3].asNumber() : 0;
    if (rate === 0) return pmt !== 0 ? FormulaResult.Number(-(pv + fv) / pmt) : null;
    return FormulaResult.Number(Math.log((-fv * rate + pmt) / (pv * rate + pmt)) / Math.log(1 + rate));
  }

  private evalNpv(args: unknown[]): FormulaResult | null {
    if (args.length < 2) return null;
    const rateArg = args[0];
    const rate = rateArg instanceof FormulaResult ? (rateArg as FormulaResult).asNumber() : 0;
    const values: number[] = [];
    for (let i = 1; i < args.length; i++) {
      const arg = args[i];
      if (Array.isArray(arg)) values.push(...(arg as number[]));
      else if (arg instanceof FormulaResult) values.push((arg as FormulaResult).asNumber());
    }
    let npv = 0;
    for (let i = 0; i < values.length; i++) npv += values[i] / Math.pow(1 + rate, i + 1);
    return FormulaResult.Number(npv);
  }

  private evalIpmt(args: unknown[]): FormulaResult | null {
    if (args.length < 4) return null;
    const rate = args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const per = args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    const nper = args[2] instanceof FormulaResult ? args[2].asNumber() : 0;
    const pv = args[3] instanceof FormulaResult ? args[3].asNumber() : 0;
    if (rate === 0) return FormulaResult.Number(0);
    const pmt = rate * (pv * Math.pow(1 + rate, nper)) / (Math.pow(1 + rate, nper) - 1);
    const fvBefore = pv * Math.pow(1 + rate, per - 1) + pmt * (Math.pow(1 + rate, per - 1) - 1) / rate;
    return FormulaResult.Number(-(fvBefore * rate));
  }

  private evalPpmt(args: unknown[]): FormulaResult | null {
    if (args.length < 4) return null;
    const pmtResult = this.evalPmt(args);
    const ipmtResult = this.evalIpmt(args);
    if (!pmtResult || !ipmtResult) return null;
    return FormulaResult.Number(pmtResult.asNumber() - ipmtResult.asNumber());
  }

  private evalSyd(args: unknown[]): FormulaResult | null {
    if (args.length < 4) return null;
    const cost = args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const salvage = args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    const life = args[2] instanceof FormulaResult ? args[2].asNumber() : 0;
    const per = args[3] instanceof FormulaResult ? args[3].asNumber() : 0;
    return FormulaResult.Number((cost - salvage) * (life - per + 1) * 2 / (life * (life + 1)));
  }

  private evalDb(args: unknown[]): FormulaResult | null {
    if (args.length < 4) return null;
    const cost = args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const salvage = args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    const life = args[2] instanceof FormulaResult ? args[2].asNumber() : 0;
    const period = args[3] instanceof FormulaResult ? Math.floor(args[3].asNumber()) : 1;
    const rate = Math.round(1 - Math.pow(salvage / cost, 1.0 / life) * 1000) / 1000;
    let total = 0;
    for (let p = 1; p <= period; p++) {
      const dep = (cost - total) * rate;
      total += dep;
      if (p === period) return FormulaResult.Number(dep);
    }
    return FormulaResult.Number(0);
  }

  private evalDdb(args: unknown[]): FormulaResult | null {
    if (args.length < 4) return null;
    const cost = args[0] instanceof FormulaResult ? args[0].asNumber() : 0;
    const salvage = args[1] instanceof FormulaResult ? args[1].asNumber() : 0;
    const life = args[2] instanceof FormulaResult ? args[2].asNumber() : 0;
    const period = args[3] instanceof FormulaResult ? Math.floor(args[3].asNumber()) : 1;
    const factor = args.length > 4 && args[4] instanceof FormulaResult ? args[4].asNumber() : 2;
    let bv = cost;
    for (let p = 1; p <= period; p++) {
      const dep = Math.min(bv * factor / life, Math.max(bv - salvage, 0));
      bv -= dep;
      if (p === period) return FormulaResult.Number(dep);
    }
    return FormulaResult.Number(0);
  }

  // ==================== Math Utilities ====================

  private trunc(v: number): number {
    return v >= 0 ? Math.floor(v) : Math.ceil(v);
  }

  private roundUp(v: number, d: number): number {
    const f = Math.pow(10, d);
    return Math.ceil(Math.abs(v) * f) / f * Math.sign(v);
  }

  private roundDown(v: number, d: number): number {
    const f = Math.pow(10, d);
    return Math.floor(Math.abs(v) * f) / f * Math.sign(v);
  }

  private ceilingF(v: number, s: number): number {
    return s === 0 ? 0 : Math.ceil(v / s) * s;
  }

  private floorF(v: number, s: number): number {
    return s === 0 ? 0 : Math.floor(v / s) * s;
  }

  private evenF(v: number): number {
    const c = Math.ceil(Math.abs(v));
    return ((c % 2 === 0 ? c : c + 1)) * Math.sign(v);
  }

  private oddF(v: number): number {
    const c = Math.ceil(Math.abs(v));
    return ((c % 2 === 1 ? c : c + 1)) * Math.sign(v);
  }

  private factorial(n: number): number {
    let r = 1;
    for (let i = 2; i <= Math.floor(n); i++) r *= i;
    return r;
  }

  private combin(n: number, k: number): number {
    if (k < 0 || k > n) return 0;
    return this.factorial(n) / (this.factorial(k) * this.factorial(n - k));
  }

  private permut(n: number, k: number): number {
    if (k < 0 || k > n) return 0;
    return this.factorial(n) / this.factorial(n - k);
  }

  private gcd(a: number, b: number): number {
    a = Math.abs(a);
    b = Math.abs(b);
    while (b !== 0) {
      const t = b;
      b = a % b;
      a = t;
    }
    return a;
  }

  private lcm(a: number, b: number): number {
    if (a === 0 || b === 0) return 0;
    return Math.abs(a / this.gcd(a, b) * b);
  }

  private toRoman(n: number): string {
    const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
    const syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
    let sb = "";
    for (let i = 0; i < vals.length; i++) {
      while (n >= vals[i]) { sb += syms[i]; n -= vals[i]; }
    }
    return sb;
  }

  private fromRoman(s: string): number {
    const map: Record<string, number> = { "M": 1000, "D": 500, "C": 100, "L": 50, "X": 10, "V": 5, "I": 1 };
    let result = 0;
    for (let i = 0; i < s.length; i++) {
      const val = map[s[i].toUpperCase()] ?? 0;
      if (i + 1 < s.length && val < (map[s[i + 1].toUpperCase()] ?? 0)) result -= val;
      else result += val;
    }
    return result;
  }
}
