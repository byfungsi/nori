import { mkdtempSync, writeFileSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { spawnSync } from "node:child_process";
import { describe, expect, it } from "vitest";
import { writeXlsx } from "@byfungsi/nori/xlsx";
import { unwrap } from "./helpers";
import { xlsxExportSnapshot } from "./xlsx-export-fixture";

const python = process.env["NORI_XLSX_PYTHON"];
describe.skipIf(python === undefined)(
  "independent Openpyxl XLSX export reader",
  () => {
    it("reads values, formulas/caches, Unicode, styles, dates, layout, filters and relationships", () => {
      if (python === undefined)
        throw new Error("NORI_XLSX_PYTHON must select Python with Openpyxl");
      const directory = mkdtempSync(join(tmpdir(), "nori-xlsx-export-"));
      try {
        const file = join(directory, "export.xlsx");
        writeFileSync(file, unwrap(writeXlsx(xlsxExportSnapshot())));
        const result = spawnSync(
          python,
          [
            "-c",
            `
import sys, zipfile, xml.etree.ElementTree as ET
from datetime import datetime
from openpyxl import load_workbook
path = sys.argv[1]
book = load_workbook(path, data_only=False)
cached = load_workbook(path, data_only=True)
assert book.sheetnames == ['Data & notes', 'Résumé']
sheet = book['Data & notes']
assert sheet['A1'].value == '雪 & <text> 🦄'
assert sheet['B1'].value is True
assert sheet['C1'].value == ''
assert sheet['D1'].value is None
assert sheet['A2'].value == 2.5
assert sheet['B2'].value == '#DIV/0!'
assert sheet['C2'].value == '=SUM(A2:A3)'
assert cached['Data & notes']['C2'].value == 6
assert cached['Data & notes']['D2'].value == 'cached & value'
assert cached['Data & notes']['E2'].value is False
assert cached['Data & notes']['F2'].value is None
assert sheet['H10'].value == '=SUM(A1:A9)' and sheet['H10'].data_type == 's'
assert book['Résumé']['A1'].value == "='Data & notes'!C2"
assert cached['Résumé']['A1'].value == 6
assert book.epoch == datetime(1904, 1, 1)
assert sheet['B3'].value == datetime(1904, 1, 2)
assert sheet['A1'].font.bold and sheet['A1'].font.italic
assert sheet['A1'].font.color.rgb == 'FF123456'
assert sheet['A1'].fill.fgColor.rgb == 'FFEEEEAA'
assert sheet['A1'].alignment.horizontal == 'center'
assert sheet['A1'].alignment.vertical == 'center'
assert sheet['A1'].alignment.wrap_text is True
assert sheet['A2'].number_format == '"$"#,##0.00'
assert sheet.freeze_panes == 'C2'
assert sheet.row_dimensions[4].hidden
assert sheet.column_dimensions['E'].hidden
assert sheet.column_dimensions['A'].width == 12
assert sheet.row_dimensions[2].height == 24
assert sheet.sheet_format.defaultColWidth == 10
assert sheet.sheet_format.defaultRowHeight == 18
assert str(next(iter(sheet.merged_cells.ranges))) == 'C5:D6'
assert sheet.auto_filter.ref == 'A1:F3'
assert sheet.auto_filter.filterColumn[0].filters.blank
assert sheet.auto_filter.filterColumn[1].customFilters._and is True
assert book.calculation.calcMode == 'manual'
with zipfile.ZipFile(path) as archive:
    sheet_xml = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    assert sheet_xml.find('s:dimension', ns).attrib['ref'] == 'A1:J20'
    sort = sheet_xml.find('s:sortState', ns)
    assert sort.attrib['ref'] == 'A2:F3'
    assert sort.findall('s:sortCondition', ns)[1].attrib['descending'] == '1'
print('Independent XLSX reader passed')
`,
            file,
          ],
          { encoding: "utf8" },
        );
        expect({ status: result.status, stderr: result.stderr }).toEqual({
          status: 0,
          stderr: "",
        });
        expect(result.stdout).toContain("Independent XLSX reader passed");
      } finally {
        rmSync(directory, { recursive: true, force: true });
      }
    });
  },
);
