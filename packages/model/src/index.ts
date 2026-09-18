import { normalizeRange, rangesIntersect, rangeContains } from "./cell-ranges";
export {
  normalizeRange,
  rangesIntersect,
  rangeContains,
  mergedRangeAt,
  expandRangeToMerges,
} from "./cell-ranges";
import { z } from "zod";
import type {
  SheetId,
  CellError,
  Scalar,
  CellStyle,
  Cell,
  CellRange,
  Position,
  SheetSnapshot,
  SheetFilter,
  SheetSort,
  ColumnFilter,
  FilterComparison,
  WorkbookSnapshot,
} from "./workbook-types";
export type {
  SheetId,
  CellError,
  Scalar,
  CellStyle,
  Cell,
  Position,
  CellRange,
  SheetSnapshot,
  SheetFilter,
  SheetSort,
  ColumnFilter,
  FilterComparison,
  WorkbookSnapshot,
} from "./workbook-types";

/** Expected failures are returned without throwing. */
export type Result<T, E extends Error> =
  | { readonly ok: true; readonly value: T }
  | { readonly ok: false; readonly error: E };
/** Construct a successful result. */
export function ok<T>(value: T): Result<T, never> {
  return { ok: true, value };
}
/** Construct a failed result. */
export function err<E extends Error>(error: E): Result<never, E> {
  return { ok: false, error };
}
/** Invalid canonical data, including unsupported snapshot versions. */
export class ModelError extends Error {
  /** Stable error discriminator. */
  readonly tag = "ModelError";
  /** Human-readable validation context. */
  constructor(message: string) {
    super(`Invalid workbook data: ${message}`);
  }
}
const sheetIdSchema = z
  .string()
  .min(1)
  .max(128)
  .transform((value) => {
    // SAFETY: The preceding schema checked identity length; the brand has no runtime representation.
    return value as SheetId;
  });
const errorSchema = z
  .object({
    kind: z.literal("error"),
    code: z.enum([
      "#REF!",
      "#VALUE!",
      "#DIV/0!",
      "#NAME?",
      "#N/A",
      "#NUM!",
      "#CYCLE!",
      "#ERROR!",
    ]),
  })
  .strict()
  .readonly() satisfies z.ZodType<CellError>;
const scalarSchema = z.union([
  z.string(),
  z.number().finite(),
  z.boolean(),
  z.null(),
  errorSchema,
]) satisfies z.ZodType<Scalar>;
const styleSchema = z
  .object({
    fontWeight: z.enum(["normal", "bold"]).optional(),
    fontStyle: z.enum(["normal", "italic"]).optional(),
    color: z
      .string()
      .regex(/^#[0-9a-f]{6}$/i)
      .optional(),
    background: z
      .string()
      .regex(/^#[0-9a-f]{6}$/i)
      .optional(),
    horizontalAlign: z.enum(["left", "center", "right"]).optional(),
    numberFormat: z.string().optional(),
    wrapText: z.boolean().optional(),
    verticalAlign: z.enum(["top", "middle", "bottom"]).optional(),
  })
  .strict()
  .readonly() satisfies z.ZodType<CellStyle>;
const cellSchema = z
  .object({
    value: scalarSchema,
    formula: z.string().min(1).max(8192).optional(),
    style: styleSchema.optional(),
  })
  .strict()
  .readonly() satisfies z.ZodType<Cell>;
const positionSchema = z
  .object({
    row: z.number().int().min(0).max(1048575),
    column: z.number().int().min(0).max(16383),
  })
  .strict()
  .readonly() satisfies z.ZodType<Position>;
const rangeSchema = z
  .object({ start: positionSchema, end: positionSchema })
  .strict()
  .transform(normalizeRange)
  .transform((range) =>
    Object.freeze({
      start: Object.freeze(range.start),
      end: Object.freeze(range.end),
    }),
  );
const dimensionSchema = z
  .record(
    z.string().regex(/^(0|[1-9][0-9]*)$/),
    z.number().finite().positive().max(10000),
  )
  .readonly();
/** Parse, normalize and freeze an inclusive range within Excel grid bounds. */
export function parseCellRange(input: unknown): Result<CellRange, ModelError> {
  const parsed = rangeSchema.safeParse(input);
  return parsed.success
    ? ok(parsed.data)
    : err(new ModelError("Invalid cell range"));
}
const addressSchema = z
  .string()
  .regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/)
  .refine((value) => {
    const p = decodeAddress(value);
    return p !== undefined && p.column < 16384 && p.row < 1048576;
  });
const filterComparisonSchema = z
  .object({
    operator: z.enum([
      "equal",
      "notEqual",
      "lessThan",
      "lessThanOrEqual",
      "greaterThan",
      "greaterThanOrEqual",
    ]),
    value: z.union([z.string(), z.number().finite()]),
  })
  .strict()
  .readonly();
const filterSchema = z
  .object({
    range: rangeSchema,
    columns: z
      .array(
        z
          .object({
            column: z.number().int().min(0).max(16383),
            condition: z.discriminatedUnion("kind", [
              z
                .object({
                  kind: z.literal("values"),
                  values: z.array(z.string()).readonly(),
                  includeBlank: z.boolean(),
                })
                .strict()
                .readonly(),
              z
                .object({
                  kind: z.literal("custom"),
                  join: z.enum(["and", "or"]),
                  comparisons: z
                    .array(filterComparisonSchema)
                    .min(1)
                    .max(2)
                    .readonly(),
                })
                .strict()
                .readonly(),
            ]),
          })
          .strict()
          .readonly(),
      )
      .readonly(),
  })
  .strict()
  .readonly() satisfies z.ZodType<SheetFilter>;
const sortSchema = z
  .object({
    range: rangeSchema,
    caseSensitive: z.boolean(),
    conditions: z
      .array(
        z
          .object({
            column: z.number().int().min(0).max(16383),
            direction: z.enum(["ascending", "descending"]),
          })
          .strict()
          .readonly(),
      )
      .min(1)
      .readonly(),
  })
  .strict()
  .readonly() satisfies z.ZodType<SheetSort>;
const hiddenIndexes = (maximum: number) =>
  z
    .array(z.number().int().min(0).max(maximum))
    .transform((indexes) =>
      Object.freeze([...new Set(indexes)].sort((a, b) => a - b)),
    );
const sheetSchema = z
  .object({
    id: sheetIdSchema,
    name: z
      .string()
      .min(1)
      .max(31)
      .regex(/^[^\\/*?:\[\]]+$/),
    cells: z.record(addressSchema, cellSchema).readonly(),
    rowCount: z.number().int().min(1).max(1048576),
    columnCount: z.number().int().min(1).max(16384),
    frozen: z
      .object({
        rows: z.number().int().min(0).max(1048576),
        columns: z.number().int().min(0).max(16384),
      })
      .strict()
      .readonly()
      .optional(),
    hiddenRows: hiddenIndexes(1048575).optional(),
    hiddenColumns: hiddenIndexes(16383).optional(),
    autoFilter: filterSchema.optional(),
    sort: sortSchema.optional(),
    merges: z.array(rangeSchema).max(10000).readonly().optional(),
    defaultColumnWidth: z.number().finite().positive().max(10000).optional(),
    defaultRowHeight: z.number().finite().positive().max(10000).optional(),
    columnWidths: dimensionSchema.optional(),
    rowHeights: dimensionSchema.optional(),
  })
  .strict()
  .readonly() satisfies z.ZodType<SheetSnapshot>;
const snapshotSchema = z
  .object({
    version: z.literal(1),
    dateSystem: z.enum(["1900", "1904"]),
    sheets: z.array(sheetSchema).min(1).readonly(),
  })
  .strict()
  .superRefine((book, ctx) => {
    const ids = new Set<string>();
    const names = new Set<string>();
    for (const sheet of book.sheets) {
      if (ids.has(sheet.id) || names.has(sheet.name.toLowerCase()))
        ctx.addIssue({
          code: "custom",
          message: "Sheet IDs and names must be unique",
        });
      ids.add(sheet.id);
      names.add(sheet.name.toLowerCase());
      for (const [dimensions, extent] of [
        [sheet.columnWidths, sheet.columnCount],
        [sheet.rowHeights, sheet.rowCount],
      ] as const) {
        for (const key of Object.keys(dimensions ?? {}))
          if (Number(key) >= extent)
            ctx.addIssue({
              code: "custom",
              message: "Dimension override exceeds sheet extent",
            });
      }
      if (
        (sheet.frozen?.rows ?? 0) > sheet.rowCount ||
        (sheet.frozen?.columns ?? 0) > sheet.columnCount
      )
        ctx.addIssue({
          code: "custom",
          message: "Freeze panes exceed sheet extent",
        });
      if (
        sheet.hiddenRows?.some((index) => index >= sheet.rowCount) ||
        sheet.hiddenColumns?.some((index) => index >= sheet.columnCount)
      )
        ctx.addIssue({
          code: "custom",
          message: "Hidden dimension exceeds sheet extent",
        });
      for (const view of [sheet.autoFilter, sheet.sort])
        if (view) {
          if (
            view.range.end.row >= sheet.rowCount ||
            view.range.end.column >= sheet.columnCount
          )
            ctx.addIssue({
              code: "custom",
              message: "Filter/sort range exceeds sheet extent",
            });
          const columns = "columns" in view ? view.columns : view.conditions;
          if (
            new Set(columns.map((item) => item.column)).size !==
              columns.length ||
            columns.some(
              (item) =>
                item.column < view.range.start.column ||
                item.column > view.range.end.column,
            )
          )
            ctx.addIssue({
              code: "custom",
              message: "Invalid filter/sort columns",
            });
        }
      const merges = sheet.merges ?? [];
      for (let i = 0; i < merges.length; i++) {
        const merge = merges[i];
        if (!merge) continue;
        if (
          merge.end.row >= sheet.rowCount ||
          merge.end.column >= sheet.columnCount
        )
          ctx.addIssue({
            code: "custom",
            message: "Merge exceeds sheet extent",
          });
        if (
          merge.start.row === merge.end.row &&
          merge.start.column === merge.end.column
        )
          ctx.addIssue({
            code: "custom",
            message: "A merge must contain at least two cells",
          });
        if (
          merges
            .slice(0, i)
            .some((previous) => rangesIntersect(previous, merge))
        )
          ctx.addIssue({
            code: "custom",
            message: "Merged regions must not overlap",
          });
      }
      for (const key of Object.keys(sheet.cells)) {
        const p = decodeAddress(key);
        if (
          p &&
          (sheet.cells[key]?.value !== null || sheet.cells[key]?.formula)
        ) {
          const merge = merges.find((range) =>
            rangeContains(range, { start: p, end: p }),
          );
          if (
            merge &&
            (p.row !== merge.start.row || p.column !== merge.start.column)
          )
            ctx.addIssue({
              code: "custom",
              message: `Covered merged cell ${key} must be empty`,
            });
        }
        if (p && (p.row >= sheet.rowCount || p.column >= sheet.columnCount))
          ctx.addIssue({
            code: "custom",
            message: `Cell ${key} exceeds sheet extent`,
          });
      }
    }
  })
  .readonly() satisfies z.ZodType<WorkbookSnapshot>;
/** Parse and defensively copy/freeze external snapshot data. */
export function parseWorkbookSnapshot(
  input: unknown,
): Result<WorkbookSnapshot, ModelError> {
  const parsed = snapshotSchema.safeParse(input);
  return parsed.success
    ? ok(parsed.data)
    : err(new ModelError(parsed.error.issues.map((i) => i.message).join("; ")));
}
/** Parse a cell before accepting it in a command. */
export function parseCell(input: unknown): Result<Cell, ModelError> {
  const parsed = cellSchema.safeParse(input);
  return parsed.success
    ? ok(parsed.data)
    : err(new ModelError(parsed.error.message));
}
/** Refine a caller-selected stable identity. */
export function parseSheetId(input: string): Result<SheetId, ModelError> {
  const parsed = sheetIdSchema.safeParse(input);
  return parsed.success
    ? ok(parsed.data)
    : err(new ModelError("Invalid sheet ID"));
}
function decodeAddress(input: string): Position | undefined {
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(input);
  if (!match?.[1] || !match[2]) return undefined;
  let column = 0;
  for (const char of match[1]) column = column * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}
/** Accept A1 and absolute $A$1 references; return normalized zero-based coordinates. */
export function parseAddress(input: string): Result<Position, ModelError> {
  const match = /^\$?([a-z]{1,3})\$?([1-9][0-9]{0,6})$/i.exec(input);
  const decoded = match
    ? decodeAddress(`${match[1]?.toUpperCase()}${match[2]}`)
    : undefined;
  const parsed = positionSchema.safeParse(decoded);
  return parsed.success
    ? ok(parsed.data)
    : err(new ModelError(`Invalid cell address: ${input}`));
}
/** Encode validated zero-based grid coordinates as A1. */
export function addressOf(position: Position): string {
  let n = position.column + 1;
  let letters = "";
  while (n > 0) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return `${letters}${position.row + 1}`;
}
/** Construct a serializable formula error. */
export function cellError(code: CellError["code"]): CellError {
  return Object.freeze({ kind: "error", code });
}
/** Narrow the object member of Scalar. */
export function isCellError(value: Scalar): value is CellError {
  return typeof value === "object" && value !== null;
}
/** Create an empty single-sheet snapshot through the canonical boundary. */
export function createEmptySnapshot(): WorkbookSnapshot {
  const result = parseWorkbookSnapshot({
    version: 1,
    dateSystem: "1900",
    sheets: [
      {
        id: "sheet-1",
        name: "Sheet1",
        cells: {},
        rowCount: 100,
        columnCount: 26,
      },
    ],
  });
  if (!result.ok) throw result.error; // Defect: the fixed built-in snapshot must satisfy its own schema.
  return result.value;
}
