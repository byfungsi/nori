# Verification record

Verified locally on 18 September 2026 using Node 22.20.0 for the initial milestone and Node 25.2.1 for later local checks. Supported repository environments remain Node 22.12+; Vitest 5 officially supports Node 22, 24 and 26+ rather than the local odd-numbered Node 25 runtime.

- Strict workspace typecheck and separate headless compilation with only ES2022 ambient types.
- Package dependency graph and forbidden-platform-global checks.
- 107 tests covering model validation/address properties, XLSX import/errors/limits, formula evaluation/errors/registry/graph, commands/history/selection/layout, pivots, and React upload-path rendering/editing/tab behavior.
- Public ESM/declaration build and Vite production demo build.
- Packed archive installed in a temporary standalone consumer: every public subpath resolves, headless imports work without React installed, declarations compile without DOM/Node ambient types, bundling core excludes XLSX/React, and React SSR works after installing React.
- Independent OOXML interoperability: openpyxl 3.1.5 reads the hand-authored sample, and Nori reads/recalculates a separately generated openpyxl workbook.
- Browser verification in local Chrome via agent-browser: loaded demo, uploaded basic.xlsx, edited Sales B2 from 1200 to 100, observed total 1550.00, switched to Summary and observed 1550, then undid and observed 2650. No application console errors or framework error overlay remained. At a 390px viewport the page fits horizontally and the grid scrolls within its container.
- Extended browser checks cover double-click inline editing, formula recalculation, Escape cancellation, read-only toggling, pointer drag with edge auto-scroll, Shift selection, merge/unmerge, row/column resize from headers and both sides of body borders, full-boundary hover feedback and undo, an independently generated XLSX with imported widths, hidden/filter rows, frozen panes and saved sort indicators, plus a 390px read-only chat preview. Run `npm run verify:browser` with the demo running; set `NORI_BROWSER_EXECUTABLE` if Chrome is not automatically available. Screenshots are saved under ignored `artifacts/`.
- Dependency audit reported zero vulnerabilities after locking the initial toolchain.

The production build emits two harmless annotation-placement warnings from Zod dependency comments. The package bundling check can report ignored bare imports because generated chunks are explicitly side-effect-free; the behavioral and installed-consumer checks pass.

No React Native device test, real Excel/LibreOffice corpus, adversarial archive isolation test, or large-workbook performance benchmark is claimed. See the support matrix and roadmap for limitations.

Documentation verification: VitePress production build, 17 rendered pages, local search, desktop/mobile navigation, no console errors, agent text exports with SHA-256 consistency checks, copied declaration dependency closure, and executable/typechecked examples. The docs target modern ES2022-capable browsers. VitePress 1 uses an explicit Vite 6.4.3 override to avoid older development-server advisories; the resolved dependency audit reports zero vulnerabilities.

The live-playground browser check covers embedded selection, editing and recalculation, cross-sheet totals, undo, read-only mode, reset, importing the hosted XLSX sample, and narrow mobile layout. Run `npm run docs:verify-playground`, setting `NORI_DOCS_URL` to a served documentation base URL.

CSV coverage includes quotes/escapes, embedded line endings, BOM, ragged and empty rows, optional inference, literal formula strings, invalid syntax, parsing budgets, and randomized quoted-record round trips. The installed-consumer check resolves `/csv` without React or XLSX dependencies, and the playground check imports the hosted CSV sample before calculating a total.
