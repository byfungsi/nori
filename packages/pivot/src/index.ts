import {
  err,
  isCellError,
  ok,
  type Result,
  type Scalar,
} from "@nori-internal/model";
/** Rows are semantic named fields, independent of sheet coordinates or renderer. */
export type PivotRecord = Readonly<Record<string, Scalar>>;
/** Explicit measure names make results stable when field labels change. */
export interface PivotDefinition {
  readonly rows: readonly string[];
  readonly values: readonly {
    readonly id: string;
    readonly field: string;
    readonly aggregate: "sum" | "count" | "average" | "min" | "max";
  }[];
}
/** Group keys retain scalar types; measure values use their declared IDs. */
export interface PivotResult {
  readonly groups: readonly {
    readonly key: readonly Scalar[];
    readonly values: Readonly<Record<string, number | null>>;
  }[];
  readonly totals: Readonly<Record<string, number | null>>;
}
/** Invalid pivot configuration or missing source fields. */
export class PivotError extends Error {
  /** Stable error discriminator. */
  readonly tag = "PivotError";
  /** Explain which source field or measure was rejected. */
  constructor(message: string) {
    super(`Invalid pivot: ${message}`);
  }
}
/** Group records in first-seen order; numeric aggregations ignore blanks/text. Count counts nonblanks. */
export function createPivot(
  records: readonly PivotRecord[],
  definition: PivotDefinition,
): Result<PivotResult, PivotError> {
  if (
    !definition.values.length ||
    new Set(definition.values.map((v) => v.id)).size !==
      definition.values.length ||
    definition.values.some(
      (v) =>
        !v.id ||
        !v.field ||
        !["sum", "count", "average", "min", "max"].includes(v.aggregate),
    )
  )
    return err(
      new PivotError("Provide unique measures with supported aggregations"),
    );
  const fields = [...definition.rows, ...definition.values.map((v) => v.field)];
  for (const record of records)
    for (const field of fields)
      if (!Object.hasOwn(record, field))
        return err(new PivotError(`Missing field ${field}`));
  const groups = new Map<string, { key: Scalar[]; records: PivotRecord[] }>();
  for (const record of records) {
    const key = definition.rows.map((f) => record[f] ?? null);
    const encoded = JSON.stringify(key);
    const group = groups.get(encoded) ?? { key, records: [] };
    group.records.push(record);
    groups.set(encoded, group);
  }
  const summarize = (
    rows: readonly PivotRecord[],
  ): Readonly<Record<string, number | null>> =>
    Object.freeze(
      Object.fromEntries(
        definition.values.map((measure) => {
          const values = rows.map((r) => r[measure.field] ?? null);
          const numbers = values.filter(
            (v): v is number => typeof v === "number",
          );
          let value: number | null;
          switch (measure.aggregate) {
            case "count":
              value = values.filter(
                (v) => v !== null && !isCellError(v),
              ).length;
              break;
            case "sum":
              value = numbers.reduce((a, b) => a + b, 0);
              break;
            case "average":
              value = numbers.length
                ? numbers.reduce((a, b) => a + b, 0) / numbers.length
                : null;
              break;
            case "min":
              value = numbers.length
                ? numbers.reduce((a, b) => Math.min(a, b))
                : null;
              break;
            case "max":
              value = numbers.length
                ? numbers.reduce((a, b) => Math.max(a, b))
                : null;
              break;
          }
          return [measure.id, value];
        }),
      ),
    );
  return ok(
    Object.freeze({
      groups: Object.freeze(
        [...groups.values()].map((g) =>
          Object.freeze({
            key: Object.freeze(g.key),
            values: summarize(g.records),
          }),
        ),
      ),
      totals: summarize(records),
    }),
  );
}
