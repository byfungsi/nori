declare const sheetIdBrand: unique symbol;
/** Stable sheet identity, validated at model boundaries. */
export type SheetId = string & { readonly [sheetIdBrand]: true };
/** Serializable spreadsheet calculation failure. */
export interface CellError {
  readonly kind: "error";
  readonly code:
    | "#REF!"
    | "#VALUE!"
    | "#DIV/0!"
    | "#NAME?"
    | "#N/A"
    | "#NUM!"
    | "#CYCLE!"
    | "#ERROR!";
}
/** Primitive cell values; dates remain numeric serials with formatting metadata. */
export type Scalar = string | number | boolean | null | CellError;
/** Document formatting independent of any application theme. */
export interface CellStyle {
  readonly fontWeight?: "normal" | "bold" | undefined;
  readonly fontStyle?: "normal" | "italic" | undefined;
  readonly color?: string | undefined;
  readonly background?: string | undefined;
  readonly horizontalAlign?: "left" | "center" | "right" | undefined;
  readonly wrapText?: boolean | undefined;
  readonly verticalAlign?: "top" | "middle" | "bottom" | undefined;
  readonly numberFormat?: string | undefined;
}
/** Formula text is optional; value is an imported or command-provided cache. */
export interface Cell {
  readonly value: Scalar;
  readonly formula?: string | undefined;
  readonly style?: CellStyle | undefined;
}
/** Zero-based grid indexes, validated within Excel grid limits at model boundaries. */
export interface Position {
  readonly row: number;
  readonly column: number;
}
/** Inclusive rectangular coordinates; range consumers normalize reversed corners. */
export interface CellRange {
  readonly start: Position;
  readonly end: Position;
}
/** A saved Excel-style comparison. Equal/notEqual accept Excel wildcards in text. */
export interface FilterComparison {
  readonly operator:
    | "equal"
    | "notEqual"
    | "lessThan"
    | "lessThanOrEqual"
    | "greaterThan"
    | "greaterThanOrEqual";
  readonly value: string | number;
}
/** Filter criteria use absolute zero-based column indexes. */
export interface ColumnFilter {
  readonly column: number;
  readonly condition:
    | {
        readonly kind: "values";
        readonly values: readonly string[];
        readonly includeBlank: boolean;
      }
    | {
        readonly kind: "custom";
        readonly join: "and" | "or";
        readonly comparisons: readonly FilterComparison[];
      };
}
/** The first row of the filter range is its header and is not filtered out. */
export interface SheetFilter {
  readonly range: CellRange;
  readonly columns: readonly ColumnFilter[];
}
/** Saved sort metadata; XLSX row values already carry the saved physical ordering. */
export interface SheetSort {
  readonly range: CellRange;
  readonly caseSensitive: boolean;
  readonly conditions: readonly {
    readonly column: number;
    readonly direction: "ascending" | "descending";
  }[];
}
/** Sparse A1-keyed cells and logical dimensions. */
export interface SheetSnapshot {
  readonly id: SheetId;
  readonly name: string;
  readonly cells: Readonly<Record<string, Cell>>;
  readonly rowCount: number;
  readonly columnCount: number;
  /** Nonoverlapping, inclusive merged regions. Only the top-left cell holds content. */
  readonly frozen?:
    { readonly rows: number; readonly columns: number } | undefined;
  readonly hiddenRows?: readonly number[] | undefined;
  readonly hiddenColumns?: readonly number[] | undefined;
  readonly autoFilter?: SheetFilter | undefined;
  readonly sort?: SheetSort | undefined;
  readonly merges?: readonly CellRange[] | undefined;
  /** Sparse zero-based dimension overrides in platform-neutral logical units. */
  readonly defaultColumnWidth?: number | undefined;
  readonly defaultRowHeight?: number | undefined;
  readonly columnWidths?: Readonly<Record<string, number>> | undefined;
  readonly rowHeights?: Readonly<Record<string, number>> | undefined;
}
/** Versioned canonical JSON data, separate from live runtime state and services. */
export interface WorkbookSnapshot {
  readonly version: 1;
  readonly dateSystem: "1900" | "1904";
  readonly sheets: readonly SheetSnapshot[];
}
