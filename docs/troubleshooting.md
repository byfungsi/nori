# Troubleshooting

| Symptom                                             | Check and remedy                                                                                                                                |
| --------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| Root import pulls in React                          | Import `/core` or the root; only `/react` should require React. Verify the installed archive is the latest build.                               |
| My formula displays `#NAME?`                        | Function is not registered. Consult the current support list; common Excel functions are still missing.                                         |
| My formula displays `#ERROR!`                       | Syntax is unsupported or malformed. Use `parseFormula` to inspect the parse Result.                                                             |
| Numbers do not update after editing snapshot fields | Snapshots are immutable. Use `dispatch`, then read `getCellValue`.                                                                              |
| Merging is rejected                                 | Covered cells must be blank and the range must not partially overlap an existing merge. Move content intentionally first.                       |
| Resizing does not work                              | Check `readOnly`, host pointer-event CSS, and whether you are using the always-read-only preview.                                               |
| Double-click does not open an editor                | Use `SheetGrid`/`Spreadsheet`, not the preview; ensure read-only is disabled. Enter/F2 also starts editing. Host overlays may intercept events. |
| Imported columns look clipped                       | Explicit widths/heights are preserved approximately; automatic font measurement and row fitting are not implemented.                            |
| Rows are missing                                    | Check explicit hidden rows and saved filter criteria. Original row numbers are retained.                                                        |
| Sorted data does not re-sort after an edit          | Saved XLSX order is preserved; interactive re-sorting is not implemented.                                                                       |
| My editor jumps back to its initial state           | Keep the live workbook instance stable across React renders.                                                                                    |
| The preview changes the wrong tab                   | Preview tabs are local by default; use its `sheetId` and `onSheetChange` for controlled state.                                                  |
| A stylesheet import fails                           | Nori has no public CSS export. Apply host CSS to documented classes/variables.                                                                  |
| A huge file blocks the UI                           | Parsing is synchronous. Move it to a host-owned worker/isolation adapter and enforce upload budgets.                                            |

If diagnosing an import issue, retain the import warnings, a minimal reproducible workbook, the package version, and expected versus actual values/layout. Do not include sensitive workbook contents in public reports.
