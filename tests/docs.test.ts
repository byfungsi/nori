import { expect, it } from "vitest";
import { createWorkbook } from "@byfungsi/nori";
import { makeWorkbook } from "../docs/examples/quick-start";
it("runs the documented quick start and rehydrates its formula", () => {
  const workbook = makeWorkbook(),
    sheetId = workbook.getState().activeSheetId;
  expect(workbook.getCellValue(sheetId, "A3")).toBe(150);
  const restored = createWorkbook(
    JSON.parse(JSON.stringify(workbook.exportSnapshot())),
  );
  expect(restored.ok).toBe(true);
  if (!restored.ok) throw restored.error;
  expect(restored.value.getCellValue(sheetId, "A3")).toBe(150);
  expect(workbook.undo()).toBe(true);
  expect(workbook.getCellValue(sheetId, "A3")).toBe(null);
});
