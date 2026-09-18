# Build with an AI agent

This guide describes the implemented Nori 0.1.0 API. It is an integration contract, not a promise of future Excel parity.

## Retrieval entry points

- [llms.txt](/llms.txt): compact reading order and per-page links.
- [llms-full.txt](/llms-full.txt): all documentation with code includes expanded.
- [docs-manifest.json](/docs-manifest.json): package version, routes, source paths and SHA-256 content hashes.
- [Public declaration index](/types/index.d.ts): generated root API; relative declaration files are served beside it. Other entries include `/types/react.d.ts`, `/types/core.d.ts`, `/types/model.d.ts`, `/types/xlsx.d.ts`, `/types/formula.d.ts`, `/types/pivot.d.ts`.
- Every page is available under `/markdown/PAGE.md`, for example <a href="./markdown/react.md" download>React source</a>. These downloads are generated from the same source as the rendered site, with no navigation markup.

The published documentation is at [byfungsi.github.io/nori](https://byfungsi.github.io/nori/). Resolve artifact paths below the site base (`/nori/` on GitHub Pages, `/` for local preview); generated indexes and manifests include that base. Local preview runs at `http://127.0.0.1:4179`.

## Integration rules

1. Import only from `@byfungsi/nori` or its documented subpaths. Never depend on internal workspace names, compiled chunk paths, or private schemas.
2. Check `result.ok` before accessing `.value`. Parsers, workbook creation, commands and pivots can fail. Do not treat them as promises; XLSX parsing is synchronous.
3. Distinguish `WorkbookSnapshot` (serializable immutable document) from `Workbook` (live runtime). Mutate through `dispatch`; never assign into snapshot cells.
4. Obtain branded sheet IDs from validated snapshots/runtime state, or use `parseSheetId`. Do not cast arbitrary strings to IDs.
5. Cell formula text omits the leading `=` in the model. Set `{value:null,formula:'SUM(A1:A2)'}`. UI editors accept the leading equals sign.
6. Read calculated values through `getCellValue`, not cached snapshot values. Results may be scalars, tagged errors or arrays.
7. Keep runtime identity stable in React. Prefer `useState` with a lazy initializer; do not create a new workbook on every render.
8. `readOnly` is a React view policy. Host-owned mutation controls must obey it; server permissions belong to the host. `SpreadsheetPreview` never edits.
9. Keep `File`, filesystem, DOM, workers and native APIs outside headless packages. Future native work should reuse core rather than importing `/react`.
10. Read import warnings and the support matrix. Do not invent `writeXlsx`, `createWorkbookStore`, `FormulaEngine`, dynamic spilling, filter menus, or React Native exports.

## Task-oriented reading order

| Task                     | Required context                                                 |
| ------------------------ | ---------------------------------------------------------------- |
| Embed a sheet            | Quick start → React → Interactions                               |
| Import attachments       | Importing → Support → API import section                         |
| Build chat attachments   | Preview → React read-only policy                                 |
| Extend calculations      | API formula section → Support → Architecture calculation section |
| Change library internals | Architecture → Contributing → affected package tests             |
| Build pivots             | Recipes → API pivot section                                      |

## Validation checklist for generated integrations

- Typecheck imports against the public package, not editor autocomplete alone.
- Exercise malformed input and Result error branches.
- Verify formula values after edits and undo.
- Unsubscribe host listeners on teardown.
- Confirm read-only actions do not mutate snapshots.
- Preserve sparse storage, styles and merge anchors when issuing commands.
- State limitations explicitly; current support and roadmap are separate documents.

The complete quick-start and React examples are checked against source types on every repository typecheck. Quick-start calculations and persistence are executed in tests. Production builds validate site links, and the docs verifier checks plain-text exports and declaration availability.
