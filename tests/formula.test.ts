import { describe, it, expect } from "vitest";
import { cellError } from "@nori-internal/model";
import {
  collectReferences,
  createFunctionRegistry,
  DependencyGraph,
  evaluateFormula,
  parseFormula,
} from "@nori-internal/formula";
import { unwrap } from "./helpers";
const resolver = {
  cell: () => 4,
  range: () => ({
    kind: "array" as const,
    values: [
      [1, 2],
      [3, null],
    ],
  }),
};
const calculate = (formula: string) =>
  evaluateFormula(unwrap(parseFormula(formula)), resolver);
describe("formula public API", () => {
  it.each([
    ["-3^2", 9],
    ["2^-2", 0.25],
    ["1+2*3", 7],
    ["(1+2)*3", 9],
    ["2^3^2", 64],
    ["50%*8", 4],
    ["SUM(A1:B2)", 6],
    ["AVERAGE(A1:B2)", 2],
    ["MIN(A1:B2)", 1],
    ["MAX(A1:B2)", 3],
    ["COUNT(A1:B2)", 3],
    ["COUNTA(A1:B2)", 3],
    ["ABS(-12)", 12],
    ["ROUND(-1.25,1)", -1.3],
    ["IF(FALSE,1/0,9)", 9],
    ["IFERROR(1/0,42)", 42],
    ["AND(TRUE,1)", true],
    ["OR(FALSE,0)", false],
    ["NOT(FALSE)", true],
    ['CONCAT("hi"," ""there""")', 'hi "there"'],
    ['LEN("hello")', 5],
    ['"A"="a"', true],
    ["A1+1", 5],
  ])("evaluates %s", (source, expected) =>
    expect(calculate(source)).toEqual(expected),
  );
  it.each(["SUM(", "A0", "1+", "Sheet!", "1 2", "@A1", "A1:Other!A3"])(
    "rejects malformed or unsupported formula %s",
    (source) => expect(parseFormula(source).ok).toBe(false),
  );
  it("preserves absolute references and quoted sheet names", () => {
    const ast = unwrap(parseFormula("'Bob''s data'!$B$2:$C3"));
    expect(collectReferences(ast)).toMatchObject([
      {
        kind: "range",
        start: { sheet: "Bob's data", absoluteRow: true, absoluteColumn: true },
        end: { sheet: "Bob's data", absoluteRow: false },
      },
    ]);
  });
  it("keeps array results first class", () =>
    expect(calculate("A1:B2")).toEqual({
      kind: "array",
      values: [
        [1, 2],
        [3, null],
      ],
    }));
  it("propagates errors and rejects unsupported functions", () => {
    expect(calculate("1/0")).toEqual(cellError("#DIV/0!"));
    expect(calculate("BOGUS(1)")).toEqual(cellError("#NAME?"));
    expect(calculate('AVERAGE("text")')).toEqual(cellError("#DIV/0!"));
    expect(calculate("1e308*10")).toEqual(cellError("#NUM!"));
  });
  it("allows pure custom functions", () => {
    const functions = createFunctionRegistry();
    functions.set("DOUBLE", (args) =>
      typeof args[0] === "number" ? args[0] * 2 : cellError("#VALUE!"),
    );
    expect(
      evaluateFormula(unwrap(parseFormula("DOUBLE(4)")), resolver, functions),
    ).toBe(8);
  });
  it("bounds nested expressions", () =>
    expect(parseFormula("(".repeat(200) + "1" + ")".repeat(200)).ok).toBe(
      false,
    ));
  it("updates dependency edges and handles cycles without infinite traversal", () => {
    const graph = new DependencyGraph();
    graph.set("B", ["A"]);
    graph.set("C", ["B"]);
    expect([...graph.affectedBy("A")].sort()).toEqual(["B", "C"]);
    graph.set("B", ["D"]);
    expect(graph.affectedBy("A").size).toBe(0);
    graph.set("D", ["C"]);
    expect(graph.affectedBy("D").size).toBe(3);
    graph.remove("B");
    expect(graph.affectedBy("D").size).toBe(0);
  });
});

it("COUNT skips error cells while COUNTA counts them", () => {
  const r = {
    cell: () => cellError("#N/A"),
    range: () => ({
      kind: "array" as const,
      values: [[1, cellError("#N/A"), null]],
    }),
  };
  expect(evaluateFormula(unwrap(parseFormula("COUNT(A1:C1)")), r)).toBe(1);
  expect(evaluateFormula(unwrap(parseFormula("COUNTA(A1:C1)")), r)).toBe(2);
});
