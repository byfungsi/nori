import { execFileSync } from "node:child_process";
import {
  mkdtempSync,
  writeFileSync,
  readFileSync,
  rmSync,
  cpSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { resolve, join } from "node:path";
import { build } from "esbuild";
const root = process.cwd();
const temp = mkdtempSync(join(tmpdir(), "nori-consumer-"));
try {
  const packed = JSON.parse(
    execFileSync("npm", ["pack", "--json", "--pack-destination", temp], {
      cwd: resolve("packages/nori"),
      encoding: "utf8",
    }),
  );
  const archive = join(temp, packed[0].filename);
  writeFileSync(
    join(temp, "package.json"),
    JSON.stringify({ private: true, type: "module" }),
  );
  execFileSync(
    "npm",
    ["install", "--ignore-scripts", "--no-audit", "--no-fund", archive],
    { cwd: temp, stdio: "pipe" },
  );
  cpSync("tests/fixtures/basic.xlsx", join(temp, "basic.xlsx"));
  writeFileSync(
    join(temp, "smoke.mjs"),
    `
 import assert from 'node:assert/strict';
 import { readFileSync } from 'node:fs';
 import { createRequire } from 'node:module';
 import { parseXlsx,createWorkbook } from '@byfungsi/nori';
 import * as model from '@byfungsi/nori/model';
 import * as formula from '@byfungsi/nori/formula';
 import * as pivot from '@byfungsi/nori/pivot';
 import * as core from '@byfungsi/nori/core';
 import * as xlsx from '@byfungsi/nori/xlsx';
 const require=createRequire(import.meta.url);
 assert.throws(()=>require.resolve('react'));
 const imported=parseXlsx(new Uint8Array(readFileSync('basic.xlsx')));assert.equal(imported.ok,true);
 const runtime=createWorkbook(imported.value.snapshot);assert.equal(runtime.ok,true);
 assert.equal(runtime.value.getCellValue(imported.value.snapshot.sheets[1].id,'B1'),2650);
 const sheetId=runtime.value.getState().activeSheetId;
 assert.equal(runtime.value.selectCell(sheetId,{row:0,column:0}).ok,true);
 assert.equal(runtime.value.moveSelection("right",{extend:true}).ok,true);
 assert.equal(runtime.value.dispatch({type:"resizeColumn",sheetId,column:0,width:220}).ok,true);
 assert.equal(core.getColumnWidth(runtime.value.exportSnapshot().sheets[0],0),220);
 assert.ok(model.parseAddress&&formula.parseFormula&&pivot.createPivot&&core.createWorkbook&&xlsx.parseXlsx);
 `,
  );
  execFileSync("node", ["smoke.mjs"], { cwd: temp, stdio: "pipe" });
  writeFileSync(
    join(temp, "consumer.ts"),
    `
 import { createWorkbook,parseXlsx,type WorkbookSnapshot } from '@byfungsi/nori';
 import { parseAddress,type SheetId } from '@byfungsi/nori/model';
 import { createFunctionRegistry,type FormulaResult } from '@byfungsi/nori/formula';
 import { createPivot } from '@byfungsi/nori/pivot';
 const parsed=parseXlsx(new Uint8Array());
 if(parsed.ok){const snapshot:WorkbookSnapshot=parsed.value.snapshot;const runtime=createWorkbook(snapshot);if(runtime.ok){const id:SheetId=runtime.value.getState().activeSheetId;const value:FormulaResult=runtime.value.getCellValue(id,'A1');void value;}}
 // @ts-expect-error IDs must be parsed or obtained from a validated snapshot.
 const invalidId:SheetId='raw';
 void [parseAddress('A1'),createFunctionRegistry(),createPivot([], {rows:[],values:[]}),invalidId];
 `,
  );
  writeFileSync(
    join(temp, "tsconfig.json"),
    JSON.stringify({
      compilerOptions: {
        strict: true,
        noUncheckedIndexedAccess: true,
        exactOptionalPropertyTypes: true,
        noImplicitOverride: true,
        noFallthroughCasesInSwitch: true,
        target: "ES2022",
        module: "NodeNext",
        moduleResolution: "NodeNext",
        lib: ["ES2022"],
        types: [],
        skipLibCheck: false,
        noEmit: true,
      },
      include: ["consumer.ts"],
    }),
  );
  execFileSync(
    process.execPath,
    [
      resolve("node_modules/typescript/bin/tsc"),
      "-p",
      join(temp, "tsconfig.json"),
    ],
    { cwd: temp, stdio: "pipe" },
  );
  for (const entry of ["index", "core", "formula", "pivot", "model", "xlsx"]) {
    const bundled = await build({
      entryPoints: [resolve(`packages/nori/dist/${entry}.js`)],
      bundle: true,
      write: false,
      metafile: true,
      platform: "neutral",
      mainFields: ["module", "main"],
      format: "esm",
      conditions: ["import"],
    });
    const inputs = Object.keys(bundled.metafile.inputs);
    if (inputs.some((name) => /node_modules\/react\//.test(name)))
      throw Error(`${entry} pulled in React`);
    if (
      ["core", "formula", "pivot", "model"].includes(entry) &&
      inputs.some((name) =>
        /node_modules\/(fflate|fast-xml-parser)\//.test(name),
      )
    )
      throw Error(`${entry} pulled in XLSX`);
  }
  execFileSync(
    "npm",
    [
      "install",
      "--ignore-scripts",
      "--no-audit",
      "--no-fund",
      "react@^19.2.0",
      "react-dom@^19.2.0",
    ],
    { cwd: temp, stdio: "pipe" },
  );
  writeFileSync(
    join(temp, "react.mjs"),
    `
 import assert from 'node:assert/strict';
 import { createElement } from 'react';
 import { renderToString } from 'react-dom/server';
 import { createWorkbook,createEmptySnapshot } from '@byfungsi/nori';
 import { Spreadsheet,SpreadsheetPreview } from '@byfungsi/nori/react';
 const result=createWorkbook(createEmptySnapshot());assert.equal(result.ok,true);
 assert.match(renderToString(createElement(Spreadsheet,{workbook:result.value})),/Sheet1/);
 const preview=renderToString(createElement(SpreadsheetPreview,{workbook:result.value,title:"Report.xlsx",theme:"dark"}));
 assert.match(preview,/Report.xlsx/);assert.match(preview,/Sheet1/);assert.doesNotMatch(preview,/<input/);
 `,
  );
  execFileSync("node", ["react.mjs"], { cwd: temp, stdio: "pipe" });
  console.log(
    "Packed consumer passed: every subpath, headless without React, declaration types, tree isolation, and React SSR.",
  );
} catch (error) {
  if (error.stdout) console.error(String(error.stdout));
  if (error.stderr) console.error(String(error.stderr));
  throw error;
} finally {
  rmSync(temp, { recursive: true, force: true });
}
