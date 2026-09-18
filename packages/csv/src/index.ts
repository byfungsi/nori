import {
  addressOf,
  parseWorkbookSnapshot,
  ok,
  err,
  type Cell,
  type Result,
  type WorkbookSnapshot,
} from "@nori-internal/model";
/** CSV input is decoded text; file access and encoding conversion belong to the host. */
export interface CsvParseOptions {
  readonly delimiter?: "," | ";" | "\t" | "|";
  readonly sheetName?: string;
  /** Nonempty fields are preserved as text by default. Infer converts canonical finite numbers and TRUE/FALSE only. */
  readonly valueMode?: "text" | "infer";
  readonly maxInputCharacters?: number;
  /** Counts every field, including empty fields, to bound sparse input too. */
  readonly maxCells?: number;
}
/** CSV creates one canonical worksheet, with no styles or executable formulas. */
export interface CsvImport {
  readonly snapshot: WorkbookSnapshot;
}
/** Invalid CSV, invalid options or an exceeded parsing budget. Offset is zero-based UTF-16. */
export class CsvError extends Error {
  readonly tag = "CsvError";
  constructor(
    readonly reason: "invalid-csv" | "invalid-options" | "limit",
    message: string,
    readonly offset: number = 0,
  ) {
    super(`CSV import: ${message}`);
  }
}
/** Parse quoted fields, escaped quotes, multiline values, CRLF/LF/CR and a leading BOM. */
export function parseCsv(
  text: string,
  options: CsvParseOptions = {},
): Result<CsvImport, CsvError> {
  const delimiter = options.delimiter ?? ",",
    mode = options.valueMode ?? "text";
  const maxChars = options.maxInputCharacters ?? 20_000_000,
    maxCells = options.maxCells ?? 200_000;
  if (
    ![",", ";", "\t", "|"].includes(delimiter) ||
    !["text", "infer"].includes(mode) ||
    !Number.isSafeInteger(maxChars) ||
    maxChars <= 0 ||
    !Number.isSafeInteger(maxCells) ||
    maxCells <= 0
  )
    return err(
      new CsvError(
        "invalid-options",
        "Invalid delimiter, value mode or positive integer budget",
      ),
    );
  if (typeof text !== "string")
    return err(new CsvError("invalid-csv", "Expected decoded text"));
  if (text.length > maxChars)
    return err(new CsvError("limit", "Input character budget exceeded"));
  const cells: Record<string, Cell> = {};
  let row = 0,
    column = 0,
    width = 0,
    height = 0,
    count = 0,
    field = "",
    state: "start" | "plain" | "quoted" | "closed" = "start";
  const finishField = (offset: number): CsvError | undefined => {
    if (++count > maxCells || row >= 1_048_576 || column >= 16_384)
      return new CsvError(
        "limit",
        "Cell budget or worksheet dimensions exceeded",
        offset,
      );
    if (field !== "") {
      let value: string | number | boolean = field;
      if (mode === "infer") {
        if (/^(TRUE|FALSE)$/i.test(field))
          value = field.toUpperCase() === "TRUE";
        else if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(field)) {
          const numeric = Number(field);
          // Preserve unsafe integer identifiers instead of rounding them.
          if (
            Number.isFinite(numeric) &&
            (!Number.isInteger(numeric) || Number.isSafeInteger(numeric))
          )
            value = numeric;
        }
      }
      cells[addressOf({ row, column })] = { value };
    }
    column++;
    width = Math.max(width, column);
    height = Math.max(height, row + 1);
    field = "";
    state = "start";
    return undefined;
  };
  const start = text.charCodeAt(0) === 0xfeff ? 1 : 0;
  for (let index = start; index < text.length; index++) {
    const char = text[index];
    if (char === undefined) break;
    if (state === "quoted") {
      if (char === '"') {
        if (text[index + 1] === '"') {
          field += '"';
          index++;
        } else state = "closed";
      } else field += char;
      continue;
    }
    if (char === delimiter || char === "\n" || char === "\r") {
      const error = finishField(index);
      if (error) return err(error);
      if (char !== delimiter) {
        row++;
        column = 0;
        if (char === "\r" && text[index + 1] === "\n") index++;
      }
      continue;
    }
    if (state === "closed")
      return err(
        new CsvError(
          "invalid-csv",
          "Expected delimiter or newline after closing quote",
          index,
        ),
      );
    if (char === '"') {
      if (state !== "start")
        return err(
          new CsvError("invalid-csv", "Quote inside an unquoted field", index),
        );
      state = "quoted";
    } else {
      field += char;
      state = "plain";
    }
  }
  if (state === "quoted")
    return err(
      new CsvError("invalid-csv", "Unclosed quoted field", text.length),
    );
  if (state !== "start" || column > 0 || text.at(-1) === delimiter) {
    const error = finishField(text.length);
    if (error) return err(error);
  }
  const parsed = parseWorkbookSnapshot({
    version: 1,
    dateSystem: "1900",
    sheets: [
      {
        id: "csv",
        name: options.sheetName ?? "Sheet1",
        rowCount: Math.max(1, height),
        columnCount: Math.max(1, width),
        cells,
      },
    ],
  });
  return parsed.ok
    ? ok({ snapshot: parsed.value })
    : err(new CsvError("invalid-options", parsed.error.message));
}
