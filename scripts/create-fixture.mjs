import { zipSync, strToU8 } from "fflate";
import { mkdir, writeFile } from "node:fs/promises";
const spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const relationships =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const office =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const files = {
  "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  "_rels/.rels": `<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  "xl/workbook.xml": `<workbook xmlns="${spreadsheet}" xmlns:r="${office}"><workbookPr date1904="0"/><sheets><sheet name="Sales" sheetId="1" r:id="rId1"/><sheet name="Summary" sheetId="2" r:id="rId2"/></sheets></workbook>`,
  "xl/_rels/workbook.xml.rels": `<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${office}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${office}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId4" Type="${office}/styles" Target="styles.xml"/></Relationships>`,
  "xl/sharedStrings.xml": `<sst xmlns="${spreadsheet}" count="5" uniqueCount="4"><si><t>Region</t></si><si><t>Revenue</t></si><si><r><t>We</t></r><r><t>st</t></r></si><si><t>East</t></si></sst>`,
  "xl/styles.xml": `<styleSheet xmlns="${spreadsheet}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FF254A39"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDF4D8"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="1" fillId="1" borderId="0" xfId="0"/></cellXfs></styleSheet>`,
  "xl/worksheets/sheet1.xml": `<worksheet xmlns="${spreadsheet}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>1200</v></c></row><row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3"><v>850</v></c></row><row r="4"><c r="A4" t="s"><v>2</v></c><c r="B4"><v>600</v></c></row><row r="6"><c r="A6" t="inlineStr"><is><t>Total</t></is></c><c r="B6" s="1"><f>SUM(B2:B4)</f><v>2650</v></c></row><row r="10"><c r="F10" t="b"><v>1</v></c></row></sheetData></worksheet>`,
  "xl/worksheets/sheet2.xml": `<worksheet xmlns="${spreadsheet}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Sales total</t></is></c><c r="B1"><f>Sales!B6</f><v>2650</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Average order</t></is></c><c r="B2"><f>AVERAGE(Sales!B2:B4)</f><v>883.3333333333334</v></c></row></sheetData></worksheet>`,
};
const bytes = zipSync(
  Object.fromEntries(
    Object.entries(files).map(([name, xml]) => [name, strToU8(xml)]),
  ),
  { level: 6 },
);
await mkdir("tests/fixtures", { recursive: true });
await mkdir("examples/react-demo/public", { recursive: true });
await writeFile("tests/fixtures/basic.xlsx", bytes);
await writeFile("examples/react-demo/public/sample.xlsx", bytes);
