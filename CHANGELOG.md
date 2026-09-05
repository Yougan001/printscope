# Changelog

## 0.1.0 — 2026-09-05

First public release.

- Inspects ordinary unencrypted XLSX files locally without modifying them.
- Separates declared, stored, populated and candidate-visible cell bounds.
- Reports explicit hiding formats, white fonts, hidden rows/columns, print areas, page breaks and distant drawing anchors.
- Shows evidence, suggested checks and interpretation limits beside a coordinate map.
- Exports metadata-only JSON reports and an inspectable sample workbook.
- Includes bounded ZIP/XML handling, cancellable worker processing, 25 automated tests and real desktop/mobile screenshots.

No formula calculation, exact pagination or automatic workbook repair. Unsupported display rules remain limitations, and candidate-visible bounds are not a recommended print area.
