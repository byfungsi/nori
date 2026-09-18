import {
  err,
  ok,
  parseAddress,
  type Result,
  type Position,
  type Scalar,
} from "@nori-internal/model";
/** Platform-neutral formula syntax, retaining absolute reference flags for future copy/fill. */
export type FormulaAst =
  | { readonly kind: "literal"; readonly value: Scalar }
  | {
      readonly kind: "reference";
      readonly sheet: string | undefined;
      readonly position: Position;
      readonly absoluteRow: boolean;
      readonly absoluteColumn: boolean;
    }
  | {
      readonly kind: "range";
      readonly start: ReferenceAst;
      readonly end: ReferenceAst;
    }
  | {
      readonly kind: "unary";
      readonly operator: "+" | "-" | "%";
      readonly operand: FormulaAst;
    }
  | {
      readonly kind: "binary";
      readonly operator: string;
      readonly left: FormulaAst;
      readonly right: FormulaAst;
    }
  | {
      readonly kind: "call";
      readonly name: string;
      readonly args: readonly FormulaAst[];
    };
/** A single reference, optionally qualified with a sheet name. */
export type ReferenceAst = Extract<FormulaAst, { kind: "reference" }>;
/** Invalid or unsupported formula syntax, with a source offset. */
export class FormulaSyntaxError extends Error {
  /** Stable error discriminator. */
  readonly tag = "FormulaSyntaxError";
  /** Source offset at which parsing stopped. */
  constructor(
    readonly offset: number,
    message: string,
  ) {
    super(`Invalid formula at ${offset}: ${message}`);
  }
}
type Token = { text: string; offset: number };
const lex =
  /\s+|"(?:[^"]|"")*"|'(?:[^']|'')*'|(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?|\$?[A-Za-z_][A-Za-z0-9_.$]*|<=|>=|<>|[+\-*/^&=<>(),:!%]/gy;
/** Parse a bounded Excel-style expression; unsupported tokens fail explicitly. */
export function parseFormula(
  source: string,
): Result<FormulaAst, FormulaSyntaxError> {
  try {
    if (source.length > 8192)
      throw new FormulaSyntaxError(8192, "Formula exceeds 8192 characters");
    const text = source.startsWith("=") ? source.slice(1) : source;
    const tokens: Token[] = [];
    let offset = 0;
    const scanner = new RegExp(lex);
    while (offset < text.length) {
      scanner.lastIndex = offset;
      const match = scanner.exec(text);
      if (!match) throw new FormulaSyntaxError(offset, "Unsupported token");
      if (match[0].trim()) tokens.push({ text: match[0], offset });
      offset = scanner.lastIndex;
    }
    let cursor = 0;
    let depth = 0;
    const peek = () => tokens[cursor]?.text;
    const fail = (message: string): never => {
      throw new FormulaSyntaxError(
        tokens[cursor]?.offset ?? text.length,
        message,
      );
    };
    const take = (): string => {
      const t = tokens[cursor++];
      return t?.text ?? fail("Expected expression");
    };
    const expect = (value: string) => {
      if (take() !== value) fail(`Expected ${value}`);
    };
    const ref = (value: string, sheet: string | undefined): ReferenceAst => {
      const parsed = parseAddress(value);
      if (!parsed.ok) return fail("Invalid reference");
      return {
        kind: "reference",
        sheet,
        position: parsed.value,
        absoluteColumn: value.startsWith("$"),
        absoluteRow: /\$\d/.test(value),
      };
    };
    const primary = (): FormulaAst => {
      const t = take();
      if (t === "+" || t === "-")
        return { kind: "unary", operator: t, operand: expression(6) };
      if (t === "(") {
        const result = expression(0);
        expect(")");
        return result;
      }
      if (t.startsWith('"'))
        return { kind: "literal", value: t.slice(1, -1).replaceAll('""', '"') };
      if (/^(?:\d|\.)/.test(t)) {
        const n = Number(t);
        if (!Number.isFinite(n)) return fail("Non-finite number");
        return { kind: "literal", value: n };
      }
      if (peek() === "(") {
        take();
        const args: FormulaAst[] = [];
        if (peek() !== ")") {
          do {
            args.push(expression(0));
            if (peek() !== ",") break;
            take();
          } while (true);
        }
        expect(")");
        return { kind: "call", name: t.toUpperCase(), args };
      }
      if (peek() === "!") {
        take();
        return ref(
          take(),
          t.startsWith("'") ? t.slice(1, -1).replaceAll("''", "'") : t,
        );
      }
      if (t.toUpperCase() === "TRUE" || t.toUpperCase() === "FALSE")
        return { kind: "literal", value: t.toUpperCase() === "TRUE" };
      return ref(t, undefined);
    };
    const precedence: Readonly<Record<string, number>> = {
      "=": 1,
      "<>": 1,
      "<": 1,
      ">": 1,
      "<=": 1,
      ">=": 1,
      "&": 2,
      "+": 3,
      "-": 3,
      "*": 4,
      "/": 4,
      "^": 5,
    };
    const expression = (minimum: number): FormulaAst => {
      if (++depth > 128) fail("Expression nesting exceeds 128");
      let left = primary();
      if (peek() === ":") {
        take();
        const right = primary();
        if (left.kind !== "reference" || right.kind !== "reference")
          return fail("Ranges require cell references");
        if (right.sheet && right.sheet !== left.sheet)
          return fail("Three-dimensional ranges are unsupported");
        left = {
          kind: "range",
          start: left,
          end: { ...right, sheet: right.sheet ?? left.sheet },
        };
      }
      while (peek() === "%") {
        take();
        left = { kind: "unary", operator: "%", operand: left };
      }
      while (true) {
        const operator = peek();
        const p = operator ? precedence[operator] : undefined;
        if (operator === undefined || p === undefined || p < minimum) break;
        take();
        const right = expression(p + 1);
        left = { kind: "binary", operator, left, right };
      }
      depth--;
      return left;
    };
    const ast = expression(0);
    if (cursor !== tokens.length) fail("Unexpected trailing token");
    return ok(ast);
  } catch (error) {
    if (error instanceof FormulaSyntaxError) return err(error);
    throw error;
  }
}
/** Syntactic precedents include both branches of lazy conditionals. Ranges stay compact. */
export function collectReferences(
  ast: FormulaAst,
): readonly (ReferenceAst | Extract<FormulaAst, { kind: "range" }>)[] {
  switch (ast.kind) {
    case "reference":
    case "range":
      return [ast];
    case "literal":
      return [];
    case "unary":
      return collectReferences(ast.operand);
    case "binary":
      return [...collectReferences(ast.left), ...collectReferences(ast.right)];
    case "call":
      return ast.args.flatMap(collectReferences);
  }
}
