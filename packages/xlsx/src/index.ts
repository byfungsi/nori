import { strFromU8, unzipSync } from "fflate";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import {
  addressOf,
  cellError,
  err,
  ok,
  parseAddress,
  parseCellRange,
  type CellRange,
  type ColumnFilter,
  type FilterComparison,
  type SheetFilter,
  type SheetSort,
  parseWorkbookSnapshot,
  type Cell,
  type CellStyle,
  type Result,
  type Scalar,
  type WorkbookSnapshot,
} from "@nori-internal/model";
/** Import diagnostics use semantic features rather than exposing XML nodes. */
export interface ImportWarning {
  readonly code: "unsupported-feature";
  readonly feature: string;
  readonly sheet?: string;
  readonly address?: string;
}
/** A canonical snapshot plus explicit lossy-import diagnostics. */
export interface XlsxImport {
  readonly snapshot: WorkbookSnapshot;
  readonly warnings: readonly ImportWarning[];
}
/** Limits apply to compressed input and declared uncompressed ZIP entries. */
export interface XlsxParseOptions {
  readonly maxInputBytes?: number;
  readonly maxUncompressedBytes?: number;
  readonly maxCells?: number;
}
/** Invalid, unsupported, or over-budget input. */
export class XlsxError extends Error {
  /** Stable error discriminator. */
  readonly tag = "XlsxError";
  /** Stable failure category for applications. */
  constructor(
    readonly reason: "invalid-file" | "unsupported" | "limit",
    message: string,
  ) {
    super(`XLSX import failed: ${message}`);
  }
}
type XmlNode = Readonly<Record<string, unknown>>;
function object(value: unknown): XmlNode {
  if (typeof value !== "object" || value === null || Array.isArray(value))
    return {};
  // SAFETY: Runtime check establishes a non-null, non-array object; all values remain unknown.
  return value as XmlNode;
}
function list(value: unknown): readonly unknown[] {
  return value === undefined ? [] : Array.isArray(value) ? value : [value];
}
function text(value: unknown): string {
  if (typeof value === "string" || typeof value === "number")
    return String(value);
  return textLeaf(object(value)["#text"]);
}
function textLeaf(value: unknown): string {
  return typeof value === "string" || typeof value === "number"
    ? String(value)
    : "";
}
function rich(value: unknown): string {
  const node = object(value);
  return (
    text(node["t"]) +
    list(node["r"])
      .map((run) => text(object(run)["t"]))
      .join("")
  );
}
function targetPath(base: string, target: string): string {
  if (target.includes("\\") || target.includes(":"))
    throw new XlsxError("unsupported", "Unsafe relationship target");
  const parts = target.startsWith("/") ? [] : base.split("/").slice(0, -1);
  for (const part of target.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") {
      if (!parts.length)
        throw new XlsxError("invalid-file", "Relationship escapes archive");
      parts.pop();
    } else parts.push(part);
  }
  return parts.join("/");
}
/** Parse XLSX bytes synchronously without browser or Node globals. File access belongs to the host. */
export function parseXlsx(
  bytes: Uint8Array,
  options: XlsxParseOptions = {},
): Result<XlsxImport, XlsxError> {
  try {
    const maxInput = options.maxInputBytes ?? 20_000_000,
      maxOutput = options.maxUncompressedBytes ?? 100_000_000,
      maxCells = options.maxCells ?? 200_000;
    if (
      ![maxInput, maxOutput, maxCells].every(
        (v) => Number.isSafeInteger(v) && v > 0,
      )
    )
      return err(
        new XlsxError("limit", "Limits must be positive safe integers"),
      );
    if (bytes.byteLength > maxInput)
      return err(new XlsxError("limit", "Compressed input exceeds limit"));
    let total = 0;
    const files = unzipSync(bytes, {
      filter: (entry) => {
        total += entry.originalSize;
        if (total > maxOutput)
          throw new XlsxError("limit", "Expanded archive exceeds limit");
        return /\.(xml|rels)$/i.test(entry.name);
      },
    });
    const parser = new XMLParser({
      ignoreAttributes: false,
      removeNSPrefix: true,
      parseTagValue: false,
      parseAttributeValue: false,
      trimValues: false,
      processEntities: true,
    });
    const xml = (path: string): XmlNode => {
      const data = files[path];
      if (!data)
        throw new XlsxError(
          "invalid-file",
          `Missing required document ${path}`,
        );
      const source = strFromU8(data);
      if (/<!DOCTYPE|<!ENTITY/i.test(source))
        throw new XlsxError(
          "unsupported",
          "Document type declarations are unsupported",
        );
      if (XMLValidator.validate(source) !== true)
        throw new XlsxError("invalid-file", `Malformed document ${path}`);
      const parsed: unknown = parser.parse(source);
      return object(parsed);
    };
    const relationships = (path: string) =>
      list(object(xml(path)["Relationships"])["Relationship"]).map(object);
    const rootRelationships = relationships("_rels/.rels");
    const office = rootRelationships.find((r) =>
      text(r["@_Type"]).endsWith("/officeDocument"),
    );
    if (!office || office["@_TargetMode"] === "External")
      throw new XlsxError("invalid-file", "No internal workbook relationship");
    const workbookPath = targetPath("", text(office["@_Target"]));
    const workbook = object(xml(workbookPath)["workbook"]);
    const slash = workbookPath.lastIndexOf("/");
    const workbookRels = relationships(
      `${workbookPath.slice(0, slash + 1)}_rels/${workbookPath.slice(slash + 1)}.rels`,
    );
    const part = (suffix: string): string | undefined => {
      const r = workbookRels.find((r) =>
        text(r["@_Type"]).endsWith("/" + suffix),
      );
      return r && r["@_TargetMode"] !== "External"
        ? targetPath(workbookPath, text(r["@_Target"]))
        : undefined;
    };
    const warnings: ImportWarning[] = [];
    const sharedPath = part("sharedStrings");
    const shared = sharedPath
      ? list(object(xml(sharedPath)["sst"])["si"]).map(rich)
      : [];
    const stylePath = part("styles");
    const styles = stylePath
      ? readStyles(object(xml(stylePath)["styleSheet"]))
      : [];
    if (Object.hasOwn(workbook, "definedNames"))
      warnings.push({ code: "unsupported-feature", feature: "named ranges" });
    if (Object.keys(files).some((p) => p.includes("pivotTable")))
      warnings.push({
        code: "unsupported-feature",
        feature: "imported pivot tables",
      });
    if (Object.keys(files).some((p) => p.includes("externalLink")))
      warnings.push({
        code: "unsupported-feature",
        feature: "external workbook links",
      });
    let cellCount = 0;
    const sheets = list(object(workbook["sheets"])["sheet"]).map(
      (input, index) => {
        const metadata = object(input);
        const name = text(metadata["@_name"]);
        const relation = workbookRels.find(
          (r) => r["@_Id"] === metadata["@_id"],
        );
        if (!relation || relation["@_TargetMode"] === "External")
          throw new XlsxError(
            "invalid-file",
            "Missing internal sheet relationship",
          );
        if (!text(relation["@_Type"]).endsWith("/worksheet"))
          throw new XlsxError("unsupported", "Only worksheets are supported");
        const sheet = object(
          xml(targetPath(workbookPath, text(relation["@_Target"])))[
            "worksheet"
          ],
        );
        for (const [tag, feature] of [
          ["drawing", "drawings"],
          ["tableParts", "Excel tables and table-level filters/sorts"],
          ["conditionalFormatting", "conditional formatting"],
          ["dataValidations", "data validation"],
        ] as const)
          if (Object.hasOwn(sheet, tag))
            warnings.push({
              code: "unsupported-feature",
              feature,
              sheet: name,
            });
        if (metadata["@_state"] && metadata["@_state"] !== "visible")
          warnings.push({
            code: "unsupported-feature",
            feature: "sheet visibility",
            sheet: name,
          });
        const warn = (feature: string) =>
          warnings.push({
            code: "unsupported-feature" as const,
            feature,
            sheet: name,
          });
        const savedView = readWorksheetView(sheet, warn);
        const hiddenRows: number[] = [],
          hiddenColumns: number[] = [];
        const merges: CellRange[] = [];
        const columnWidths: Record<string, number> = {},
          rowHeights: Record<string, number> = {};
        const format = object(sheet["sheetFormatPr"]);
        const defaultColumnWidth =
          format["@_defaultColWidth"] === undefined
            ? undefined
            : columnWidth(format["@_defaultColWidth"]);
        const defaultRowHeight =
          format["@_defaultRowHeight"] === undefined
            ? undefined
            : rowHeight(format["@_defaultRowHeight"]);
        for (const entry of list(object(sheet["cols"])["col"])) {
          const col = object(entry),
            min = Number(col["@_min"]),
            max = Number(col["@_max"]);
          if (
            !Number.isInteger(min) ||
            !Number.isInteger(max) ||
            min < 1 ||
            max < min ||
            max > 16384
          )
            throw new XlsxError(
              "invalid-file",
              "Invalid column dimension range",
            );
          if (
            ["1", "true"].includes(text(col["@_hidden"])) ||
            Number(col["@_width"]) === 0
          )
            for (let column = min - 1; column < max; column++)
              hiddenColumns.push(column);
          if (col["@_width"] !== undefined && Number(col["@_width"]) !== 0) {
            const width = columnWidth(col["@_width"]);
            for (let column = min - 1; column < max; column++)
              columnWidths[String(column)] = width;
          }
          if (col["@_style"] !== undefined) warn("column-level styles");
        }
        for (const entry of list(object(sheet["mergeCells"])["mergeCell"])) {
          const references = text(object(entry)["@_ref"]).split(":");
          if (references.length !== 2)
            throw new XlsxError("invalid-file", "Invalid merged range");
          const start = parseAddress(references[0] ?? ""),
            end = parseAddress(references[1] ?? "");
          if (!start.ok || !end.ok)
            throw new XlsxError("invalid-file", "Invalid merged range");
          const range = parseCellRange({ start: start.value, end: end.value });
          if (!range.ok)
            throw new XlsxError("invalid-file", range.error.message);
          merges.push(range.value);
        }
        const cells: Record<string, Cell> = Object.create(null);
        let rowCount = 1,
          columnCount = 1;
        for (const rowInput of list(object(sheet["sheetData"])["row"])) {
          const row = object(rowInput);
          if (row["@_ht"] !== undefined) {
            const index = Number(row["@_r"]);
            if (!Number.isInteger(index) || index < 1 || index > 1048576)
              throw new XlsxError(
                "invalid-file",
                "Invalid row dimension index",
              );
            if (Number(row["@_ht"]) === 0) hiddenRows.push(index - 1);
            else rowHeights[String(index - 1)] = rowHeight(row["@_ht"]);
          }
          if (["1", "true"].includes(text(row["@_hidden"]))) {
            const index = Number(row["@_r"]);
            if (!Number.isInteger(index) || index < 1 || index > 1048576)
              throw new XlsxError("invalid-file", "Invalid hidden row index");
            hiddenRows.push(index - 1);
          }
          for (const cellInput of list(row["c"])) {
            if (++cellCount > maxCells)
              throw new XlsxError("limit", "Cell count exceeds limit");
            const node = object(cellInput);
            const address = parseAddress(text(node["@_r"]));
            if (!address.ok)
              throw new XlsxError("invalid-file", "Invalid cell address");
            const key = addressOf(address.value);
            if (Object.hasOwn(cells, key))
              throw new XlsxError("invalid-file", "Duplicate cell address");
            rowCount = Math.max(rowCount, address.value.row + 1);
            columnCount = Math.max(columnCount, address.value.column + 1);
            const raw = text(node["v"]);
            const type = text(node["@_t"]);
            let value: Scalar = null;
            switch (type) {
              case "s": {
                const index = Number(raw);
                if (
                  !raw ||
                  !Number.isInteger(index) ||
                  index < 0 ||
                  shared[index] === undefined
                )
                  throw new XlsxError(
                    "invalid-file",
                    "Invalid shared string index",
                  );
                value = shared[index] ?? "";
                break;
              }
              case "inlineStr":
                value = rich(node["is"]);
                break;
              case "str":
                value = raw;
                break;
              case "b":
                if (raw !== "0" && raw !== "1")
                  throw new XlsxError("invalid-file", "Invalid boolean");
                value = raw === "1";
                break;
              case "e":
                value = cellError(
                  [
                    "#REF!",
                    "#VALUE!",
                    "#DIV/0!",
                    "#NAME?",
                    "#N/A",
                    "#NUM!",
                  ].includes(raw)
                    ? errorCode(raw)
                    : "#ERROR!",
                );
                break;
              case "d":
                value = raw;
                warnings.push({
                  code: "unsupported-feature",
                  feature: "ISO date cell retained as text",
                  sheet: name,
                  address: key,
                });
                break;
              case "":
              case "n":
                if (raw) {
                  value = Number(raw);
                  if (!Number.isFinite(value))
                    throw new XlsxError("invalid-file", "Invalid numeric cell");
                }
                break;
              default:
                throw new XlsxError(
                  "unsupported",
                  `Unsupported cell type ${type}`,
                );
            }
            const f = object(node["f"]);
            const formulaText = text(node["f"]);
            const formulaType = text(f["@_t"]);
            let formula: string | undefined;
            if (Object.hasOwn(node, "f")) {
              if (formulaType && formulaType !== "normal") {
                warnings.push({
                  code: "unsupported-feature",
                  feature: `${formulaType} formula (cached value only)`,
                  sheet: name,
                  address: key,
                });
              } else if (formulaText) formula = formulaText;
            }
            let style: CellStyle | undefined;
            if (node["@_s"] !== undefined) {
              const index = Number(node["@_s"]);
              if (!Number.isInteger(index) || !styles[index])
                throw new XlsxError("invalid-file", "Invalid cell style index");
              style = styles[index];
            }
            cells[key] = {
              value,
              ...(formula === undefined ? {} : { formula }),
              ...(style === undefined ? {} : { style }),
            };
          }
        }
        for (const merge of merges) {
          rowCount = Math.max(rowCount, merge.end.row + 1);
          columnCount = Math.max(columnCount, merge.end.column + 1);
        }
        for (const index of hiddenRows)
          rowCount = Math.max(rowCount, index + 1);
        for (const index of hiddenColumns)
          columnCount = Math.max(columnCount, index + 1);
        if (savedView.frozen) {
          rowCount = Math.max(rowCount, savedView.frozen.rows);
          columnCount = Math.max(columnCount, savedView.frozen.columns);
        }
        for (const view of [savedView.autoFilter, savedView.sort])
          if (view) {
            rowCount = Math.max(rowCount, view.range.end.row + 1);
            columnCount = Math.max(columnCount, view.range.end.column + 1);
          }
        for (const key of Object.keys(rowHeights))
          rowCount = Math.max(rowCount, Number(key) + 1);
        for (const key of Object.keys(columnWidths))
          columnCount = Math.max(columnCount, Number(key) + 1);
        return {
          id: `sheet-${index + 1}`,
          name,
          cells,
          rowCount,
          columnCount,
          ...savedView,
          ...(hiddenRows.length ? { hiddenRows } : {}),
          ...(hiddenColumns.length ? { hiddenColumns } : {}),
          ...(merges.length ? { merges } : {}),
          ...(Object.keys(columnWidths).length ? { columnWidths } : {}),
          ...(Object.keys(rowHeights).length ? { rowHeights } : {}),
          ...(defaultColumnWidth === undefined ? {} : { defaultColumnWidth }),
          ...(defaultRowHeight === undefined ? {} : { defaultRowHeight }),
        };
      },
    );
    const snapshot = parseWorkbookSnapshot({
      version: 1,
      dateSystem: ["1", "true"].includes(
        text(object(workbook["workbookPr"])["@_date1904"]),
      )
        ? "1904"
        : "1900",
      sheets,
    });
    return snapshot.ok
      ? ok({ snapshot: snapshot.value, warnings })
      : err(new XlsxError("invalid-file", snapshot.error.message));
  } catch (error) {
    return err(
      error instanceof XlsxError
        ? error
        : new XlsxError("invalid-file", "Unable to decode XLSX archive"),
    );
  }
}
function errorCode(
  text: string,
): "#REF!" | "#VALUE!" | "#DIV/0!" | "#NAME?" | "#N/A" | "#NUM!" {
  switch (text) {
    case "#REF!":
    case "#VALUE!":
    case "#DIV/0!":
    case "#NAME?":
    case "#N/A":
    case "#NUM!":
      return text;
    default:
      return "#VALUE!";
  }
}
function readStyles(root: XmlNode): CellStyle[] {
  const fonts = list(object(root["fonts"])["font"]).map(object),
    fills = list(object(root["fills"])["fill"]).map(object);
  const formats = new Map<number, string>([
    [0, "General"],
    [1, "0"],
    [2, "0.00"],
    [9, "0%"],
    [10, "0.00%"],
    [14, "mm-dd-yy"],
  ]);
  for (const value of list(object(root["numFmts"])["numFmt"])) {
    const f = object(value);
    formats.set(Number(f["@_numFmtId"]), text(f["@_formatCode"]));
  }
  const rgb = (node: unknown): string | undefined => {
    const raw = text(object(node)["@_rgb"]);
    return /^(?:[0-9a-f]{2})?[0-9a-f]{6}$/i.test(raw)
      ? "#" + raw.slice(-6)
      : undefined;
  };
  return list(object(root["cellXfs"])["xf"]).map((input) => {
    const xf = object(input),
      font = fonts[Number(xf["@_fontId"])] ?? {},
      fill = fills[Number(xf["@_fillId"])] ?? {};
    const color = rgb(font["color"]),
      background = rgb(object(fill["patternFill"])["fgColor"]),
      numberFormat = formats.get(Number(xf["@_numFmtId"]));
    const align = text(object(xf["alignment"])["@_horizontal"]);
    const vertical = text(object(xf["alignment"])["@_vertical"]);
    const wrap = object(xf["alignment"])["@_wrapText"];
    return {
      ...(font["b"] === undefined || object(font["b"])["@_val"] === "0"
        ? {}
        : { fontWeight: "bold" as const }),
      ...(font["i"] === undefined || object(font["i"])["@_val"] === "0"
        ? {}
        : { fontStyle: "italic" as const }),
      ...(color ? { color } : {}),
      ...(background ? { background } : {}),
      ...(numberFormat ? { numberFormat } : {}),
      ...(wrap === undefined
        ? {}
        : { wrapText: ["1", "true"].includes(text(wrap)) }),
      ...(["top", "center", "bottom"].includes(vertical)
        ? {
            verticalAlign:
              vertical === "top"
                ? ("top" as const)
                : vertical === "center"
                  ? ("middle" as const)
                  : ("bottom" as const),
          }
        : {}),
      ...(["left", "right", "center"].includes(align)
        ? { horizontalAlign: alignment(align) }
        : {}),
    };
  });
}
function alignment(value: string): "left" | "right" | "center" {
  return value === "right" ? "right" : value === "center" ? "center" : "left";
}

// OOXML widths depend on the Normal font's maximum digit width. Seven logical
// units approximates Calibri 11 at 96 dpi without importing a platform font API.
function columnWidth(input: unknown): number {
  const width = Number(input);
  if (!Number.isFinite(width) || width <= 0 || width > 255)
    throw new XlsxError("invalid-file", "Invalid column width");
  return Math.max(
    1,
    Math.floor(((256 * width + Math.floor(128 / 7)) / 256) * 7),
  );
}
function rowHeight(input: unknown): number {
  const points = Number(input);
  if (!Number.isFinite(points) || points <= 0 || points > 7500)
    throw new XlsxError("invalid-file", "Invalid row height");
  return (points * 4) / 3;
}

function viewRange(input: unknown): CellRange {
  const refs = text(input).split(":");
  if (refs.length < 1 || refs.length > 2)
    throw new XlsxError("invalid-file", "Invalid worksheet view range");
  const start = parseAddress(refs[0] ?? ""),
    end = parseAddress(refs[1] ?? refs[0] ?? "");
  if (!start.ok || !end.ok)
    throw new XlsxError("invalid-file", "Invalid worksheet view range");
  const range = parseCellRange({ start: start.value, end: end.value });
  if (!range.ok) throw new XlsxError("invalid-file", range.error.message);
  return range.value;
}
function readWorksheetView(
  sheet: XmlNode,
  warn: (feature: string) => void,
): {
  frozen?: { rows: number; columns: number };
  autoFilter?: SheetFilter;
  sort?: SheetSort;
} {
  const result: {
    frozen?: { rows: number; columns: number };
    autoFilter?: SheetFilter;
    sort?: SheetSort;
  } = {};
  const view = object(list(object(sheet["sheetViews"])["sheetView"])[0]),
    pane = object(view["pane"]);
  if (Object.keys(pane).length) {
    if (["frozen", "frozenSplit"].includes(text(pane["@_state"]))) {
      const rows = Number(pane["@_ySplit"] ?? 0),
        columns = Number(pane["@_xSplit"] ?? 0);
      if (
        !Number.isInteger(rows) ||
        !Number.isInteger(columns) ||
        rows < 0 ||
        rows > 1048576 ||
        columns < 0 ||
        columns > 16384
      )
        throw new XlsxError("invalid-file", "Invalid freeze panes");
      result.frozen = { rows, columns };
    } else warn("split panes");
  }
  const filter = object(sheet["autoFilter"]);
  if (Object.keys(filter).length) {
    const range = viewRange(filter["@_ref"]),
      columns: ColumnFilter[] = [];
    for (const item of list(filter["filterColumn"])) {
      const column = object(item),
        relative = Number(column["@_colId"]);
      if (
        !Number.isInteger(relative) ||
        relative < 0 ||
        relative > range.end.column - range.start.column
      )
        throw new XlsxError("invalid-file", "Invalid filter column");
      const index = range.start.column + relative;
      if (column["filters"] !== undefined) {
        const filters = object(column["filters"]);
        if (filters["dateGroupItem"] !== undefined) {
          warn("date-group filters (saved hidden rows retained)");
          continue;
        }
        columns.push({
          column: index,
          condition: {
            kind: "values",
            values: list(filters["filter"]).map((v) =>
              text(object(v)["@_val"]),
            ),
            includeBlank: ["1", "true"].includes(text(filters["@_blank"])),
          },
        });
      } else if (column["customFilters"] !== undefined) {
        const filters = object(column["customFilters"]),
          comparisons: FilterComparison[] = [];
        let supported = true;
        for (const entry of list(filters["customFilter"])) {
          const custom = object(entry),
            operator = text(custom["@_operator"]) || "equal",
            value = text(custom["@_val"]);
          switch (operator) {
            case "equal":
            case "notEqual":
            case "lessThan":
            case "lessThanOrEqual":
            case "greaterThan":
            case "greaterThanOrEqual":
              comparisons.push({ operator, value });
              break;
            default:
              supported = false;
          }
        }
        if (!supported || !comparisons.length || comparisons.length > 2) {
          warn("custom filter combination (saved hidden rows retained)");
          continue;
        }
        columns.push({
          column: index,
          condition: {
            kind: "custom",
            join: ["1", "true"].includes(text(filters["@_and"])) ? "and" : "or",
            comparisons,
          },
        });
      } else if (
        ["dynamicFilter", "top10", "colorFilter", "iconFilter"].some(
          (tag) => column[tag] !== undefined,
        )
      )
        warn(
          "dynamic, ranking, color or icon filters (saved hidden rows retained)",
        );
    }
    result.autoFilter = { range, columns };
  }
  const sort = object(sheet["sortState"] ?? filter["sortState"]);
  if (Object.keys(sort).length) {
    if (["1", "true"].includes(text(sort["@_columnSort"])))
      warn("left-to-right sort (saved cell order retained)");
    else {
      const range = viewRange(sort["@_ref"]),
        conditions: {
          column: number;
          direction: "ascending" | "descending";
        }[] = [];
      for (const entry of list(sort["sortCondition"])) {
        const condition = object(entry),
          by = text(condition["@_sortBy"]);
        if ((by && by !== "value") || condition["@_customList"] !== undefined) {
          warn("color, icon or custom-list sort (saved cell order retained)");
          continue;
        }
        const ref = viewRange(condition["@_ref"]);
        if (ref.start.column !== ref.end.column)
          throw new XlsxError(
            "invalid-file",
            "Sort condition must identify one column",
          );
        conditions.push({
          column: ref.start.column,
          direction: ["1", "true"].includes(text(condition["@_descending"]))
            ? "descending"
            : "ascending",
        });
      }
      if (conditions.length)
        result.sort = {
          range,
          conditions,
          caseSensitive: ["1", "true"].includes(text(sort["@_caseSensitive"])),
        };
    }
  }
  return result;
}
