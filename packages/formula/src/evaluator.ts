import { cellError, isCellError, type Scalar } from "@nori-internal/model";
import type { FormulaAst, ReferenceAst } from "./parser";
/** Rectangular results reserve a first-class path for future dynamic-array spilling. */
export type FormulaResult =
  | Scalar
  | { readonly kind: "array"; readonly values: readonly (readonly Scalar[])[] };
/** A custom function is pure and returns spreadsheet errors as values. */
export type FormulaFunction = (args: readonly FormulaResult[]) => FormulaResult;
/** Per-runtime registry; callers may extend a copy before creating a workbook. */
export type FunctionRegistry = ReadonlyMap<string, FormulaFunction>;
/** Core provides reference resolution; formula has no knowledge of workbook storage. */
export interface FormulaResolver {
  readonly cell: (reference: ReferenceAst) => FormulaResult;
  readonly range: (start: ReferenceAst, end: ReferenceAst) => FormulaResult;
}
/** Identify an array result without treating spreadsheet errors as arrays. */
export function isArrayResult(
  value: FormulaResult,
): value is Extract<FormulaResult, { kind: "array" }> {
  return typeof value === "object" && value !== null && value.kind === "array";
}
function flat(args: readonly FormulaResult[]): Scalar[] {
  return args.flatMap((a) =>
    isArrayResult(a) ? a.values.flatMap((row) => [...row]) : [a],
  );
}
function number(value: Scalar): number | undefined {
  if (value === null) return 0;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return Number(value);
  if (
    typeof value === "string" &&
    value.trim() &&
    Number.isFinite(Number(value))
  )
    return Number(value);
  return undefined;
}
function numeric(value: number): Scalar {
  return Number.isFinite(value) ? value : cellError("#NUM!");
}
function aggregate(
  args: readonly FormulaResult[],
  operation: string,
): FormulaResult {
  const values = flat(args);
  const nums = values.filter((v): v is number => typeof v === "number");
  if (operation === "COUNT") return nums.length;
  if (operation === "COUNTA") return values.filter((v) => v !== null).length;
  const error = values.find(isCellError);
  if (error) return error;
  if (operation === "AVERAGE" && !nums.length) return cellError("#DIV/0!");
  if (operation === "MIN")
    return nums.length ? nums.reduce((a, b) => Math.min(a, b)) : 0;
  if (operation === "MAX")
    return nums.length ? nums.reduce((a, b) => Math.max(a, b)) : 0;
  const sum = nums.reduce((a, b) => a + b, 0);
  return numeric(operation === "AVERAGE" ? sum / nums.length : sum);
}
/** Common nonvolatile functions; IF and IFERROR are evaluator-owned for lazy branches. */
export function createFunctionRegistry(): Map<string, FormulaFunction> {
  const registry = new Map<string, FormulaFunction>();
  for (const name of ["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA"])
    registry.set(name, (args) => aggregate(args, name));
  for (const name of ["ABS", "ROUND"])
    registry.set(name, (args) => {
      if (args.length < 1 || args.length > (name === "ROUND" ? 2 : 1))
        return cellError("#VALUE!");
      const values = flat(args);
      const error = values.find(isCellError);
      if (error) return error;
      const n = number(values[0] ?? null);
      const digits = number(values[1] ?? 0);
      if (n === undefined || digits === undefined) return cellError("#VALUE!");
      const factor = 10 ** Math.trunc(digits);
      return numeric(
        name === "ABS"
          ? Math.abs(n)
          : (Math.sign(n) * Math.round(Math.abs(n) * factor)) / factor,
      );
    });
  for (const name of ["AND", "OR", "NOT"])
    registry.set(name, (args) => {
      const values = flat(args);
      const error = values.find(isCellError);
      if (error) return error;
      if (!values.length || (name === "NOT" && values.length !== 1))
        return cellError("#VALUE!");
      if (name === "NOT") return !values[0];
      return name === "AND" ? values.every(Boolean) : values.some(Boolean);
    });
  registry.set("CONCAT", (args) => {
    const values = flat(args);
    return values.find(isCellError) ?? values.map((v) => v ?? "").join("");
  });
  registry.set("LEN", (args) => {
    const value = args[0] ?? null;
    if (args.length !== 1 || isArrayResult(value)) return cellError("#VALUE!");
    return isCellError(value) ? value : String(value ?? "").length;
  });
  return registry;
}
/** Evaluate scalar expressions and ranges; unresolved or invalid operations return cell errors. */
export function evaluateFormula(
  ast: FormulaAst,
  resolver: FormulaResolver,
  registry: FunctionRegistry = createFunctionRegistry(),
): FormulaResult {
  const run = (node: FormulaAst): FormulaResult => {
    switch (node.kind) {
      case "literal":
        return node.value;
      case "reference":
        return resolver.cell(node);
      case "range":
        return resolver.range(node.start, node.end);
      case "unary": {
        const v = run(node.operand);
        if (isArrayResult(v)) return cellError("#VALUE!");
        if (isCellError(v)) return v;
        const n = number(v);
        return n === undefined
          ? cellError("#VALUE!")
          : numeric(
              node.operator === "-" ? -n : node.operator === "%" ? n / 100 : n,
            );
      }
      case "binary": {
        const a = run(node.left);
        const b = run(node.right);
        if (isArrayResult(a) || isArrayResult(b)) return cellError("#VALUE!");
        if (isCellError(a)) return a;
        if (isCellError(b)) return b;
        if (node.operator === "&") return String(a ?? "") + String(b ?? "");
        const x = number(a);
        const y = number(b);
        if (["=", "<>", "<", ">", "<=", ">="].includes(node.operator)) {
          const left = typeof a === "string" ? a.toLowerCase() : (a ?? 0);
          const right = typeof b === "string" ? b.toLowerCase() : (b ?? 0);
          switch (node.operator) {
            case "=":
              return left === right;
            case "<>":
              return left !== right;
            case "<":
              return left < right;
            case ">":
              return left > right;
            case "<=":
              return left <= right;
            default:
              return left >= right;
          }
        }
        if (x === undefined || y === undefined) return cellError("#VALUE!");
        switch (node.operator) {
          case "+":
            return numeric(x + y);
          case "-":
            return numeric(x - y);
          case "*":
            return numeric(x * y);
          case "/":
            return y === 0 ? cellError("#DIV/0!") : numeric(x / y);
          case "^":
            return numeric(x ** y);
          default:
            return cellError("#ERROR!");
        }
      }
      case "call": {
        if (node.name === "IF" || node.name === "IFERROR") {
          const first = node.args[0];
          if (
            !first ||
            node.args.length < 2 ||
            node.args.length > (node.name === "IF" ? 3 : 2)
          )
            return cellError("#VALUE!");
          const value = run(first);
          if (isArrayResult(value)) return cellError("#VALUE!");
          if (node.name === "IFERROR")
            return isCellError(value)
              ? run(node.args[1] ?? { kind: "literal", value: null })
              : value;
          if (isCellError(value)) return value;
          const branch = node.args[value ? 1 : 2];
          return branch ? run(branch) : false;
        }
        const fn = registry.get(node.name);
        return fn ? fn(node.args.map(run)) : cellError("#NAME?");
      }
    }
  };
  return run(ast);
}
