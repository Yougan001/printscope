import { mkdir, writeFile } from 'node:fs/promises';
import { sampleParts, zipParts, worksheet, REL } from '../core/sample.mjs';

const directory = new URL('../work/fixtures/', import.meta.url);
await mkdir(directory, { recursive: true });
const parts = sampleParts();
parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace(
  '</sheets>',
  '<sheet name="Clean sheet" sheetId="2" r:id="clean"/></sheets>',
);
parts['xl/_rels/workbook.xml.rels'] = parts[
  'xl/_rels/workbook.xml.rels'
].replace(
  '</Relationships>',
  `<Relationship Id="clean" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
);
parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
  '</Types>',
  '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
);
parts['xl/worksheets/sheet2.xml'] = worksheet(
  '<dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>',
);
await writeFile(new URL('two-sheets.xlsx', directory), zipParts(parts));
await writeFile(
  new URL('damaged.xlsx', directory),
  'This is not a ZIP workbook.',
);
const large = sampleParts();
large['xl/worksheets/sheet1.xml'] = worksheet(
  '<sheetData>' +
    Array.from(
      { length: 1000 },
      (_, r) => `<row r="${r + 1}">${'<c><v>1</v></c>'.repeat(100)}</row>`,
    ).join('') +
    '</sheetData>',
);
await writeFile(new URL('100k-cells.xlsx', directory), zipParts(large));
console.log('Three local QA fixtures written to work/fixtures.');
