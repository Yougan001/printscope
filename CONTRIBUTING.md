# Contributing to Printscope

Reports that connect a finding to an actual workbook setting are more useful than a screenshot of a blank page alone. English and Chinese are welcome.

## Report a finding

Use the [workbook issue form](https://github.com/Yougan001/printscope/issues/new?template=workbook-problem.yml). State the finding, affected sheet/cell addresses and what you observed in Excel or LibreOffice. Include the application version and browser.

Do not attach a private workbook. Make a fresh, minimal workbook with invented values and recreate just the setting that matters. Deleting visible cells from the original is not reliable anonymization: other sheets, names, comments, links and document properties may remain. Review exported reports too; they contain filenames, sheet names and addresses.

If you cannot share a safe sample, describe the relevant XML/settings and steps. A report without a workbook can still be useful, but may not be enough to reproduce the problem. Exact page-count differences are not automatically an inspection bug: Printscope does not implement Excel's rendering or printer layout.

## Develop and test

Use Node.js 22.13+ and install the committed lockfile:

```sh
npm ci
npm test
npm run lint
npm run typecheck
```

Keep inspection independent of React in `core/`. Add small synthetic workbook fixtures to the tests, with explicit expected findings and limits. ZIP changes need malformed-input, expansion and checksum cases. XML changes must work without a Node `Buffer` global and must not resolve entities or fetch relationships.

For a new diagnostic, separate observed evidence from interpretation. Unknown or unsupported rendering rules should remain explicit; candidate-visible bounds must not become an automatic replacement print area. Do not alter the uploaded workbook.

For interface changes, check sheet switching, a rejected file, cancellation, the sample's two findings, JSON export and a narrow viewport. `scripts/qa-fixtures.mjs` creates synthetic local QA files. [Verification notes](docs/testing.md) list what has and has not been tested.

Run `npm run build` when your environment supports it. The Linux Pages workflow is the release gate; emitted files after a nonzero Windows export are not a passing build. In a pull request, describe the test case and observed results without claiming broader Excel compatibility than was checked.
