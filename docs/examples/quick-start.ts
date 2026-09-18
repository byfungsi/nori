import { createEmptySnapshot, createWorkbook } from "@byfungsi/nori";

export function makeWorkbook() {
  const created = createWorkbook(createEmptySnapshot());
  if (!created.ok) throw created.error;
  const workbook = created.value;
  const sheetId = workbook.getState().activeSheetId;
  const edited = workbook.dispatch({
    type: "batch",
    commands: [
      { type: "setCell", sheetId, address: "A1", cell: { value: 100 } },
      { type: "setCell", sheetId, address: "A2", cell: { value: 50 } },
      {
        type: "setCell",
        sheetId,
        address: "A3",
        cell: { value: null, formula: "SUM(A1:A2)" },
      },
    ],
  });
  if (!edited.ok) throw edited.error;
  return workbook;
}

const workbook = makeWorkbook();
const sheetId = workbook.getState().activeSheetId;
console.log(workbook.getCellValue(sheetId, "A3")); // 150
const json = JSON.stringify(workbook.exportSnapshot());
const restored = createWorkbook(JSON.parse(json));
if (!restored.ok) throw restored.error;
