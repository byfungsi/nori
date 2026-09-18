import { it, expect } from "vitest";
import { createPivot } from "@nori-internal/pivot";
import { unwrap } from "./helpers";
it("groups typed keys and computes numeric aggregates and totals", () => {
  const result = unwrap(
    createPivot(
      [
        { region: "A", value: 10 },
        { region: "A", value: 20 },
        { region: "B", value: null },
        { region: 1, value: 4 },
        { region: "1", value: 6 },
      ],
      {
        rows: ["region"],
        values: ["sum", "count", "average", "min", "max"].map((aggregate) => ({
          id: aggregate,
          field: "value",
          aggregate: aggregation(aggregate),
        })),
      },
    ),
  );
  expect(result.groups).toHaveLength(4);
  expect(result.groups[0]?.values).toEqual({
    sum: 30,
    count: 2,
    average: 15,
    min: 10,
    max: 20,
  });
  expect(result.groups[1]?.values).toEqual({
    sum: 0,
    count: 0,
    average: null,
    min: null,
    max: null,
  });
  expect(result.totals).toEqual({
    sum: 40,
    count: 4,
    average: 10,
    min: 4,
    max: 20,
  });
});
function aggregation(
  input: string,
): "sum" | "count" | "average" | "min" | "max" {
  switch (input) {
    case "sum":
    case "count":
    case "average":
    case "min":
    case "max":
      return input;
    default:
      throw Error("fixture");
  }
}
it("handles empty input and rejects missing fields or duplicate measure IDs", () => {
  const definition = {
    rows: ["region"],
    values: [{ id: "total", field: "value", aggregate: "sum" as const }],
  };
  expect(unwrap(createPivot([], definition))).toEqual({
    groups: [],
    totals: { total: 0 },
  });
  expect(createPivot([{ region: "A" }], definition).ok).toBe(false);
  expect(
    createPivot([], {
      ...definition,
      values: [...definition.values, ...definition.values],
    }).ok,
  ).toBe(false);
});
