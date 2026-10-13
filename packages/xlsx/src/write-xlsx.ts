import { strToU8, zipSync } from "fflate";
import {
  addressOf,
  err,
  ok,
  parseAddress,
  parseWorkbookSnapshot,
  type Cell,
  type CellRange,
  type CellStyle,
  type Position,
  type Result,
  type SheetSnapshot,
  type WorkbookSnapshot,
} from "@nori-internal/model";
import { encodeXlsxText } from "./xlsx-text";

/** Output limits count retained cells/layout/styles and serialized ZIP bytes, independently of calculation. */
export interface XlsxWriteOptions {
  readonly maxCells?: number;
  readonly maxSheets?: number;
  readonly maxStyles?: number;
  readonly maxLayoutEntries?: number;
  readonly maxUncompressedBytes?: number;
  readonly maxOutputBytes?: number;
  /** Explicitly asks Excel to recalculate on open; export itself never evaluates formulas. */
  readonly recalculateOnOpen?: boolean;
}

/** Invalid canonical data, unsupported Excel representation, or bounded serialization failure. */
export class XlsxWriteError extends Error {
  readonly tag = "XlsxWriteError";
  constructor(
    readonly reason:
      | "invalid-snapshot"
      | "invalid-options"
      | "unsupported"
      | "limit"
      | "archive",
    message: string,
  ) {
    super(`XLSX export failed: ${message}`);
  }
}

const spreadsheetNamespace =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const relationshipNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const packageRelationshipNamespace =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const xml = (text: string): string =>
  text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;")
    .replaceAll("\r", "&#13;");
const stringXml = (text: string, attribute = false): string =>
  xml(encodeXlsxText(text, attribute));
const rangeRef = (range: CellRange): string =>
  `${addressOf(range.start)}:${addressOf(range.end)}`;
const bool = (value: boolean): string => (value ? "1" : "0");
const hasLoneSurrogate = (text: string): boolean => {
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return true;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) return true;
  }
  return false;
};

function cellPosition(address: string): Position {
  const parsed = parseAddress(address);
  if (!parsed.ok)
    throw new Error("Canonical XLSX cell address invariant violated");
  return parsed.value;
}

function checkExportableSnapshot(
  snapshot: WorkbookSnapshot,
  limits: Required<Omit<XlsxWriteOptions, "recalculateOnOpen">>,
): Result<void, XlsxWriteError> {
  let cells = 0,
    layout = 0;
  if (snapshot.sheets.length > limits.maxSheets)
    return err(new XlsxWriteError("limit", "Sheet count exceeds limit"));
  for (const sheet of snapshot.sheets) {
    if (
      sheet.name.startsWith("'") ||
      sheet.name.endsWith("'") ||
      /[\u0000-\u001f]/.test(sheet.name) ||
      hasLoneSurrogate(sheet.name)
    )
      return err(
        new XlsxWriteError(
          "unsupported",
          "Sheet name cannot be represented faithfully in Excel",
        ),
      );
    cells += Object.keys(sheet.cells).length;
    layout +=
      (sheet.hiddenRows?.length ?? 0) +
      (sheet.hiddenColumns?.length ?? 0) +
      Object.keys(sheet.rowHeights ?? {}).length +
      Object.keys(sheet.columnWidths ?? {}).length +
      (sheet.merges?.length ?? 0) +
      (sheet.autoFilter?.columns.length ?? 0) +
      (sheet.sort?.conditions.length ?? 0);
    if (cells > limits.maxCells || layout > limits.maxLayoutEntries)
      return err(
        new XlsxWriteError("limit", "Cell or layout count exceeds limit"),
      );
    if (
      (sheet.frozen?.rows ?? 0) === 1048576 ||
      (sheet.frozen?.columns ?? 0) === 16384
    )
      return err(
        new XlsxWriteError(
          "unsupported",
          "Freeze panes must leave an addressable scroll origin",
        ),
      );
    for (const width of [
      sheet.defaultColumnWidth,
      ...Object.values(sheet.columnWidths ?? {}),
    ]) {
      if (width !== undefined && width > 1785)
        return err(
          new XlsxWriteError(
            "unsupported",
            "Column width exceeds Excel's character-width limit",
          ),
        );
    }
    for (const cell of Object.values(sheet.cells)) {
      if (
        typeof cell.value === "string" &&
        (cell.value.length > 32767 || hasLoneSurrogate(cell.value))
      )
        return err(
          new XlsxWriteError(
            "unsupported",
            "Cell text exceeds Excel's limit or contains invalid Unicode",
          ),
        );
      if (
        typeof cell.value === "object" &&
        cell.value !== null &&
        (cell.value.code === "#CYCLE!" || cell.value.code === "#ERROR!")
      )
        return err(
          new XlsxWriteError(
            "unsupported",
            "Nori calculation errors require an explicit Excel-compatible value before export",
          ),
        );
      if (
        cell.formula !== undefined &&
        (cell.formula === "=" ||
          /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/.test(
            cell.formula,
          ) ||
          hasLoneSurrogate(cell.formula))
      )
        return err(
          new XlsxWriteError(
            "unsupported",
            "Formula text cannot be represented faithfully in XML",
          ),
        );
      const format = cell.style?.numberFormat;
      if (
        format !== undefined &&
        (format.length === 0 ||
          format.length > 4096 ||
          hasLoneSurrogate(format))
      )
        return err(
          new XlsxWriteError(
            "unsupported",
            "Number format is empty, over budget, or invalid Unicode",
          ),
        );
    }
    for (const filter of sheet.autoFilter?.columns ?? []) {
      const values =
        filter.condition.kind === "values"
          ? filter.condition.values
          : filter.condition.comparisons.map((comparison) =>
              String(comparison.value),
            );
      layout += values.length;
      if (layout > limits.maxLayoutEntries)
        return err(
          new XlsxWriteError(
            "limit",
            "Filter value count exceeds layout limit",
          ),
        );
      if (values.some(hasLoneSurrogate))
        return err(
          new XlsxWriteError(
            "unsupported",
            "Filter text contains invalid Unicode",
          ),
        );
    }
  }
  return ok(undefined);
}

interface XlsxStyles {
  readonly entries: readonly CellStyle[];
  readonly indexes: ReadonlyMap<string, number>;
  readonly formats: ReadonlyMap<string, number>;
}
interface XlsxStrings {
  readonly values: readonly string[];
  readonly indexes: ReadonlyMap<string, number>;
  readonly count: number;
}
function collectXlsxStrings(snapshot: WorkbookSnapshot): XlsxStrings {
  const values: string[] = [];
  const indexes = new Map<string, number>();
  let count = 0;
  for (const sheet of snapshot.sheets)
    for (const address of Object.keys(sheet.cells).sort()) {
      const cell = sheet.cells[address];
      if (cell === undefined)
        throw new Error("Canonical XLSX string cell invariant violated");
      if (cell.formula !== undefined || typeof cell.value !== "string")
        continue;
      count += 1;
      if (indexes.has(cell.value)) continue;
      indexes.set(cell.value, values.length);
      values.push(cell.value);
    }
  return { values, indexes, count };
}
function* sharedStringXml(strings: XlsxStrings): Generator<string> {
  yield `${xmlHeader}<sst xmlns="${spreadsheetNamespace}" count="${strings.count}" uniqueCount="${strings.values.length}">`;
  for (const value of strings.values)
    yield `<si><t xml:space="preserve">${stringXml(value)}</t></si>`;
  yield "</sst>";
}
const styleKey = (style: CellStyle): string => JSON.stringify(style);
function collectXlsxStyles(
  snapshot: WorkbookSnapshot,
  maximum: number,
): Result<XlsxStyles, XlsxWriteError> {
  const entries: CellStyle[] = [{}];
  const indexes = new Map([[styleKey({}), 0]]);
  const formats = new Map<string, number>();
  for (const sheet of snapshot.sheets)
    for (const address of Object.keys(sheet.cells).sort()) {
      const cell = sheet.cells[address];
      if (cell === undefined)
        throw new Error("Canonical XLSX style cell invariant violated");
      if (cell.style === undefined) continue;
      const key = styleKey(cell.style);
      if (indexes.has(key)) continue;
      if (entries.length >= Math.min(maximum, 65_000))
        return err(
          new XlsxWriteError("limit", "Unique style count exceeds limit"),
        );
      indexes.set(key, entries.length);
      entries.push(cell.style);
      const format = cell.style.numberFormat;
      if (format !== undefined && format !== "General" && !formats.has(format))
        formats.set(format, 164 + formats.size);
    }
  return ok({ entries, indexes, formats });
}

function* styleXml(styles: XlsxStyles): Generator<string> {
  yield `${xmlHeader}<styleSheet xmlns="${spreadsheetNamespace}">`;
  if (styles.formats.size) {
    yield `<numFmts count="${styles.formats.size}">`;
    for (const [format, id] of styles.formats)
      yield `<numFmt numFmtId="${id}" formatCode="${stringXml(format, true)}"/>`;
    yield "</numFmts>";
  }
  yield `<fonts count="${styles.entries.length}">`;
  for (const style of styles.entries)
    yield `<font><sz val="11"/><name val="Calibri"/>${style.fontWeight === "bold" ? "<b/>" : ""}${style.fontStyle === "italic" ? "<i/>" : ""}${style.color === undefined ? "" : `<color rgb="FF${style.color.slice(1).toUpperCase()}"/>`}</font>`;
  yield `</fonts><fills count="${styles.entries.length + 2}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>`;
  for (const style of styles.entries)
    yield style.background === undefined
      ? '<fill><patternFill patternType="none"/></fill>'
      : `<fill><patternFill patternType="solid"><fgColor rgb="FF${style.background.slice(1).toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`;
  yield `</fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${styles.entries.length}">`;
  for (const [index, style] of styles.entries.entries()) {
    const numberFormat =
      style.numberFormat === undefined || style.numberFormat === "General"
        ? 0
        : styles.formats.get(style.numberFormat);
    if (numberFormat === undefined)
      throw new Error("Collected XLSX number format invariant violated");
    yield `<xf numFmtId="${numberFormat}" fontId="${index}" fillId="${index + 2}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="${bool(style.numberFormat !== undefined)}" applyAlignment="1"><alignment${style.horizontalAlign === undefined ? "" : ` horizontal="${style.horizontalAlign}"`}${style.verticalAlign === undefined ? "" : ` vertical="${style.verticalAlign === "middle" ? "center" : style.verticalAlign}"`}${style.wrapText === undefined ? "" : ` wrapText="${bool(style.wrapText)}"`}/></xf>`;
  }
  yield '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
}

function* cellXml(
  address: string,
  cell: Cell,
  styles: XlsxStyles,
  strings: XlsxStrings,
): Generator<string> {
  const style =
    cell.style === undefined ? 0 : styles.indexes.get(styleKey(cell.style));
  if (style === undefined)
    throw new Error("Collected XLSX cell style invariant violated");
  const attributes = `r="${address}"${style === 0 ? "" : ` s="${style}"`}`;
  const formula =
    cell.formula === undefined
      ? ""
      : `<f>${xml(cell.formula.startsWith("=") ? cell.formula.slice(1) : cell.formula)}</f>`;
  if (typeof cell.value === "string") {
    const stringIndex = strings.indexes.get(cell.value);
    if (!formula && stringIndex === undefined)
      throw new Error("Collected XLSX string invariant violated");
    yield formula
      ? `<c ${attributes} t="str">${formula}<v>${stringXml(cell.value)}</v></c>`
      : `<c ${attributes} t="s"><v>${stringIndex}</v></c>`;
  } else if (typeof cell.value === "number")
    yield `<c ${attributes}>${formula}<v>${cell.value}</v></c>`;
  else if (typeof cell.value === "boolean")
    yield `<c ${attributes} t="b">${formula}<v>${bool(cell.value)}</v></c>`;
  else if (cell.value === null) yield `<c ${attributes}>${formula}</c>`;
  else
    yield `<c ${attributes} t="e">${formula}<v>${xml(cell.value.code)}</v></c>`;
}

function* savedViewXml(sheet: SheetSnapshot): Generator<string> {
  const filter = sheet.autoFilter;
  if (filter !== undefined) {
    yield `<autoFilter ref="${rangeRef(filter.range)}">`;
    for (const item of filter.columns) {
      yield `<filterColumn colId="${item.column - filter.range.start.column}">`;
      if (item.condition.kind === "values") {
        yield `<filters blank="${bool(item.condition.includeBlank)}">`;
        for (const value of item.condition.values)
          yield `<filter val="${stringXml(value, true)}"/>`;
        yield "</filters>";
      } else {
        yield `<customFilters and="${bool(item.condition.join === "and")}">`;
        for (const comparison of item.condition.comparisons)
          yield `<customFilter operator="${comparison.operator}" val="${stringXml(String(comparison.value), true)}"/>`;
        yield "</customFilters>";
      }
      yield "</filterColumn>";
    }
    yield "</autoFilter>";
  }
  if (sheet.sort !== undefined) {
    yield `<sortState ref="${rangeRef(sheet.sort.range)}" caseSensitive="${bool(sheet.sort.caseSensitive)}">`;
    for (const condition of sheet.sort.conditions)
      yield `<sortCondition ref="${rangeRef({ start: { row: sheet.sort.range.start.row, column: condition.column }, end: { row: sheet.sort.range.end.row, column: condition.column } })}" descending="${bool(condition.direction === "descending")}"/>`;
    yield "</sortState>";
  }
}

function* worksheetXml(
  sheet: SheetSnapshot,
  styles: XlsxStyles,
  strings: XlsxStrings,
): Generator<string> {
  yield `${xmlHeader}<worksheet xmlns="${spreadsheetNamespace}"><dimension ref="A1:${addressOf({ row: sheet.rowCount - 1, column: sheet.columnCount - 1 })}"/><sheetViews><sheetView workbookViewId="0">`;
  const rows = sheet.frozen?.rows ?? 0,
    columns = sheet.frozen?.columns ?? 0;
  if (rows || columns)
    yield `<pane${columns ? ` xSplit="${columns}"` : ""}${rows ? ` ySplit="${rows}"` : ""} topLeftCell="${addressOf({ row: rows, column: columns })}" activePane="${rows && columns ? "bottomRight" : rows ? "bottomLeft" : "topRight"}" state="frozen"/>`;
  yield `</sheetView></sheetViews><sheetFormatPr defaultRowHeight="${((sheet.defaultRowHeight ?? 20) * 3) / 4}"${sheet.defaultColumnWidth === undefined ? "" : ` defaultColWidth="${sheet.defaultColumnWidth / 7}"`}/>`;
  const hiddenColumns = new Set(sheet.hiddenColumns ?? []);
  const columnIndexes = [
    ...new Set([
      ...hiddenColumns,
      ...Object.keys(sheet.columnWidths ?? {}).map(Number),
    ]),
  ].sort((left, right) => left - right);
  if (columnIndexes.length) {
    yield "<cols>";
    for (const column of columnIndexes) {
      const width = sheet.columnWidths?.[String(column)];
      yield `<col min="${column + 1}" max="${column + 1}"${hiddenColumns.has(column) ? ' hidden="1"' : ""}${width === undefined ? "" : ` width="${width / 7}" customWidth="1"`}/>`;
    }
    yield "</cols>";
  }
  const rowCells = new Map<
    number,
    Array<{
      readonly address: string;
      readonly column: number;
      readonly cell: Cell;
    }>
  >();
  for (const [address, cell] of Object.entries(sheet.cells)) {
    const position = cellPosition(address);
    const list = rowCells.get(position.row) ?? [];
    list.push({ address, column: position.column, cell });
    rowCells.set(position.row, list);
  }
  const hiddenRows = new Set(sheet.hiddenRows ?? []);
  const rowIndexes = [
    ...new Set([
      ...rowCells.keys(),
      ...hiddenRows,
      ...Object.keys(sheet.rowHeights ?? {}).map(Number),
    ]),
  ].sort((left, right) => left - right);
  yield "<sheetData>";
  for (const row of rowIndexes) {
    const height = sheet.rowHeights?.[String(row)];
    yield `<row r="${row + 1}"${hiddenRows.has(row) ? ' hidden="1"' : ""}${height === undefined ? "" : ` ht="${(height * 3) / 4}" customHeight="1"`}>`;
    for (const item of (rowCells.get(row) ?? []).sort(
      (left, right) => left.column - right.column,
    ))
      yield* cellXml(item.address, item.cell, styles, strings);
    yield "</row>";
  }
  yield "</sheetData>";
  yield* savedViewXml(sheet);
  if (sheet.merges?.length) {
    yield `<mergeCells count="${sheet.merges.length}">`;
    for (const merge of sheet.merges)
      yield `<mergeCell ref="${rangeRef(merge)}"/>`;
    yield "</mergeCells>";
  }
  yield "</worksheet>";
}

/** Writes supported canonical workbook data to deterministic XLSX bytes without I/O or recalculation. */
export function writeXlsx(
  snapshot: WorkbookSnapshot,
  options: XlsxWriteOptions = {},
): Result<Uint8Array, XlsxWriteError> {
  const limits = {
    maxCells: options.maxCells ?? 200_000,
    maxSheets: options.maxSheets ?? 256,
    maxStyles: options.maxStyles ?? 4096,
    maxLayoutEntries: options.maxLayoutEntries ?? 200_000,
    maxUncompressedBytes: options.maxUncompressedBytes ?? 100_000_000,
    maxOutputBytes: options.maxOutputBytes ?? 20_000_000,
  };
  if (
    !Object.values(limits).every(
      (limit) => Number.isSafeInteger(limit) && limit > 0,
    ) ||
    (options.recalculateOnOpen !== undefined &&
      options.recalculateOnOpen !== true &&
      options.recalculateOnOpen !== false)
  )
    return err(
      new XlsxWriteError(
        "invalid-options",
        "Limits must be positive safe integers and recalculation must be boolean",
      ),
    );
  const parsed = parseWorkbookSnapshot(snapshot);
  if (!parsed.ok)
    return err(new XlsxWriteError("invalid-snapshot", parsed.error.message));
  const exportable = checkExportableSnapshot(parsed.value, limits);
  if (!exportable.ok) return exportable;
  const styles = collectXlsxStyles(parsed.value, limits.maxStyles);
  if (!styles.ok) return styles;
  const book = parsed.value;
  const strings = collectXlsxStrings(book);
  const parts: Record<string, Uint8Array> = {};
  let byteSize = 0;
  const addPart = (
    path: string,
    fragments: Iterable<string>,
  ): Result<void, XlsxWriteError> => {
    const strings: string[] = [];
    for (const fragment of fragments) {
      byteSize += strToU8(fragment).byteLength;
      if (byteSize > limits.maxUncompressedBytes)
        return err(
          new XlsxWriteError("limit", "Uncompressed XML exceeds limit"),
        );
      strings.push(fragment);
    }
    parts[path] = strToU8(strings.join(""));
    return ok(undefined);
  };
  const sources: Array<readonly [string, Iterable<string>]> = [
    [
      "[Content_Types].xml",
      [
        `${xmlHeader}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
        ...book.sheets.map(
          (_sheet, index) =>
            `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        ),
        strings.count === 0
          ? ""
          : '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
        "</Types>",
      ],
    ],
    [
      "_rels/.rels",
      [
        `${xmlHeader}<Relationships xmlns="${packageRelationshipNamespace}"><Relationship Id="rIdWorkbook" Type="${relationshipNamespace}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      ],
    ],
    [
      "xl/workbook.xml",
      [
        `${xmlHeader}<workbook xmlns="${spreadsheetNamespace}" xmlns:r="${relationshipNamespace}"><workbookPr date1904="${bool(book.dateSystem === "1904")}"/><bookViews><workbookView/></bookViews><sheets>`,
        ...book.sheets.map(
          (sheet, index) =>
            `<sheet name="${stringXml(sheet.name, true)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
        ),
        `</sheets><calcPr calcId="0" calcMode="${options.recalculateOnOpen === true ? "auto" : "manual"}" fullCalcOnLoad="${bool(options.recalculateOnOpen === true)}" forceFullCalc="${bool(options.recalculateOnOpen === true)}"/></workbook>`,
      ],
    ],
    [
      "xl/_rels/workbook.xml.rels",
      [
        `${xmlHeader}<Relationships xmlns="${packageRelationshipNamespace}">`,
        ...book.sheets.map(
          (_sheet, index) =>
            `<Relationship Id="rId${index + 1}" Type="${relationshipNamespace}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
        ),
        strings.count === 0
          ? ""
          : `<Relationship Id="rIdStrings" Type="${relationshipNamespace}/sharedStrings" Target="sharedStrings.xml"/>`,
        `<Relationship Id="rIdStyles" Type="${relationshipNamespace}/styles" Target="styles.xml"/></Relationships>`,
      ],
    ],
    ["xl/styles.xml", styleXml(styles.value)],
    ...book.sheets.map((sheet, index): readonly [string, Iterable<string>] => [
      `xl/worksheets/sheet${index + 1}.xml`,
      worksheetXml(sheet, styles.value, strings),
    ]),
  ];
  if (strings.count > 0)
    sources.push(["xl/sharedStrings.xml", sharedStringXml(strings)]);
  for (const [path, fragments] of sources) {
    const part = addPart(path, fragments);
    if (!part.ok) return part;
  }
  try {
    const bytes = zipSync(parts, { level: 6, mtime: new Date(1980, 0, 1) });
    return bytes.byteLength > limits.maxOutputBytes
      ? err(new XlsxWriteError("limit", "Compressed archive exceeds limit"))
      : ok(bytes);
  } catch {
    return err(new XlsxWriteError("archive", "Unable to encode ZIP archive"));
  }
}
