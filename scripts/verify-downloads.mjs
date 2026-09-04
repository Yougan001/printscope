import { readFile } from 'node:fs/promises';
import assert from 'node:assert/strict';
import { inspectWorkbook } from '../core/workbook.mjs';

const directory = new URL('../work/downloads/', import.meta.url);
const report = JSON.parse(
  await readFile(new URL('printscope-report.json', directory), 'utf8'),
);
const workbook = inspectWorkbook(
  new Uint8Array(await readFile(new URL('monthly-report.xlsx', directory))),
);
const { file, ...downloaded } = report;
assert.equal(file, 'monthly-report.xlsx');
assert.deepEqual(downloaded, workbook);
assert.equal(workbook.storedCellCount, 13);
assert.equal(workbook.findingCount, 2);
console.log(
  'Actual browser XLSX and JSON downloads agree: 13 cells, 2 findings.',
);
