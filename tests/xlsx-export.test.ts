import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { strFromU8, unzipSync } from "fflate";
import { parseWorkbookSnapshot } from "@nori-internal/model";
import { parseXlsx, writeXlsx } from "@byfungsi/nori/xlsx";
import { writeXlsx as writeFromRoot } from "@byfungsi/nori";
import { unwrap } from "./helpers";
import { xlsxExportSnapshot } from "./xlsx-export-fixture";

describe("headless XLSX export", () => {
  it.each(["1900", "1904"] as const)(
    "round-trips supported workbook data and saved views in the %s date system",
    (dateSystem) => {
      const snapshot = xlsxExportSnapshot(dateSystem);
      const before = JSON.stringify(snapshot);
      const bytes = unwrap(writeXlsx(snapshot));
      const imported = unwrap(parseXlsx(bytes));
      expect(imported.warnings).toEqual([]);
      expect(imported.snapshot.dateSystem).toBe(dateSystem);
      expect(imported.snapshot.sheets.map((sheet) => sheet.name)).toEqual(
        snapshot.sheets.map((sheet) => sheet.name),
      );
      for (const [index, source] of snapshot.sheets.entries()) {
        const restored = imported.snapshot.sheets[index];
        expect(restored).toMatchObject({
          rowCount: source.rowCount,
          columnCount: source.columnCount,
        });
        for (const [address, cell] of Object.entries(source.cells))
          expect(restored?.cells[address]).toMatchObject(cell);
        if (source.autoFilter !== undefined)
          expect(restored?.autoFilter).toEqual(source.autoFilter);
        if (source.sort !== undefined)
          expect(restored?.sort).toEqual(source.sort);
        if (source.merges !== undefined)
          expect(restored?.merges).toEqual(source.merges);
        if (source.frozen !== undefined)
          expect(restored?.frozen).toEqual(source.frozen);
        if (source.hiddenRows !== undefined)
          expect(restored?.hiddenRows).toEqual(source.hiddenRows);
        if (source.hiddenColumns !== undefined)
          expect(restored?.hiddenColumns).toEqual(source.hiddenColumns);
        if (source.columnWidths !== undefined)
          expect(restored?.columnWidths).toEqual(source.columnWidths);
        if (source.rowHeights !== undefined)
          expect(restored?.rowHeights).toEqual(source.rowHeights);
      }
      expect(JSON.stringify(snapshot)).toBe(before);
      expect(unwrap(writeFromRoot(snapshot))).toEqual(bytes);
    },
  );

  it("produces identical bytes for identical data regardless of sparse-cell insertion order", () => {
    const snapshot = xlsxExportSnapshot();
    const reordered = unwrap(
      parseWorkbookSnapshot({
        ...snapshot,
        sheets: snapshot.sheets.map((sheet) => ({
          ...sheet,
          cells: Object.fromEntries(Object.entries(sheet.cells).reverse()),
        })),
      }),
    );
    expect(unwrap(writeXlsx(snapshot))).toEqual(unwrap(writeXlsx(reordered)));
  });

  it("exports caches without recalculation and marks ordinary text starting with = as text", () => {
    const original = xlsxExportSnapshot();
    const snapshot = unwrap(
      parseWorkbookSnapshot({
        ...original,
        sheets: original.sheets.map((sheet, index) =>
          index === 0
            ? {
                ...sheet,
                cells: {
                  ...sheet.cells,
                  C2: { value: 999, formula: "=SUM(A2:A3)" },
                },
              }
            : sheet,
        ),
      }),
    );
    const files = unzipSync(unwrap(writeXlsx(snapshot)));
    const workbook = files["xl/workbook.xml"],
      sheet = files["xl/worksheets/sheet1.xml"];
    if (workbook === undefined || sheet === undefined)
      throw new Error("Expected XLSX parts");
    expect(strFromU8(workbook)).toContain('calcMode="manual"');
    expect(strFromU8(sheet)).toContain("<f>SUM(A2:A3)</f><v>999</v>");
    expect(strFromU8(sheet)).toContain('r="H10" t="s"');
    const auto = unzipSync(
      unwrap(writeXlsx(snapshot, { recalculateOnOpen: true })),
    )["xl/workbook.xml"];
    if (auto === undefined) throw new Error("Expected workbook part");
    expect(strFromU8(auto)).toContain('calcMode="auto"');
  });

  it.each([
    { maxCells: 1 },
    { maxSheets: 1 },
    { maxStyles: 1 },
    { maxLayoutEntries: 1 },
    { maxUncompressedBytes: 1000 },
    { maxOutputBytes: 10 },
  ])("returns an explicit bounded-output failure for %j", (options) => {
    const result = writeXlsx(xlsxExportSnapshot(), options);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.reason).toBe("limit");
  });

  it("rejects malformed snapshots, invalid limits, and nonportable calculation errors", () => {
    const snapshot = xlsxExportSnapshot();
    // @ts-expect-error Unsupported versions are also rejected at the runtime boundary.
    const invalid = writeXlsx({ ...snapshot, version: 2 });
    expect(invalid.ok).toBe(false);
    if (!invalid.ok) expect(invalid.error.reason).toBe("invalid-snapshot");
    for (const options of [{ maxCells: 0 }, { maxOutputBytes: Number.NaN }]) {
      const result = writeXlsx(snapshot, options);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.reason).toBe("invalid-options");
    }
    const unsupported = unwrap(
      parseWorkbookSnapshot({
        ...snapshot,
        sheets: [
          {
            ...snapshot.sheets[0],
            cells: { A1: { value: { kind: "error", code: "#CYCLE!" } } },
          },
        ],
      }),
    );
    const result = writeXlsx(unsupported);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.reason).toBe("unsupported");
  });

  it("rejects overlong cells, unrepresentable widths, and invalid Unicode explicitly", () => {
    const snapshot = xlsxExportSnapshot();
    for (const patch of [
      { cells: { A1: { value: "x".repeat(32768) } } },
      { columnWidths: { "0": 1800 } },
      { cells: { A1: { value: "bad\ud800" } } },
    ]) {
      const input = unwrap(
        parseWorkbookSnapshot({
          ...snapshot,
          sheets: [{ ...snapshot.sheets[0], ...patch }],
        }),
      );
      const result = writeXlsx(input);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.reason).toBe("unsupported");
    }
  });

  it("round-trips literal escape sequences, XML text, and controls across generated strings", () => {
    fc.assert(
      fc.property(
        fc
          .array(
            fc.constantFrom(
              "a",
              "&",
              "<",
              ">",
              '"',
              "'",
              "\r",
              "\n",
              "\t",
              "\u0000",
              "_x0041_",
              "_x005F_",
              "_X0041_",
              "🦄",
              "雪",
            ),
            { maxLength: 80 },
          )
          .map((pieces) => pieces.join("")),
        (value) => {
          const snapshot = unwrap(
            parseWorkbookSnapshot({
              version: 1,
              dateSystem: "1900",
              sheets: [
                {
                  id: "one",
                  name: "Text",
                  rowCount: 1,
                  columnCount: 1,
                  cells: { A1: { value } },
                },
              ],
            }),
          );
          const restored = unwrap(
            parseXlsx(unwrap(writeXlsx(snapshot))),
          ).snapshot;
          expect(restored.sheets[0]?.cells["A1"]?.value).toBe(value);
        },
      ),
      { numRuns: 100 },
    );
  });
});
