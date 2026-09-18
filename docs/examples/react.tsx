import { useState } from "react";
import { Spreadsheet, SpreadsheetPreview } from "@byfungsi/nori/react";
import { makeWorkbook } from "./quick-start";

export function WorkbookExample() {
  // Keep runtime identity stable across renders.
  const [workbook] = useState(makeWorkbook);
  const [readOnly, setReadOnly] = useState(false);
  return (
    <>
      <label>
        <input
          type="checkbox"
          checked={readOnly}
          onChange={(event) => setReadOnly(event.target.checked)}
        />{" "}
        Read only
      </label>
      <Spreadsheet workbook={workbook} readOnly={readOnly} />
      <SpreadsheetPreview
        workbook={workbook}
        title="report.xlsx"
        theme="dark"
      />
    </>
  );
}
