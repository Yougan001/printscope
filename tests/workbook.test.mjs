import test from 'node:test';
import { spawnSync } from 'node:child_process';
import assert from 'node:assert/strict';
import { inspectWorkbook } from '../core/workbook.mjs';
import { parseXml, attr, child, textOf } from '../core/xml.mjs';
import { rangeLabel, printAreas, range } from '../core/ranges.mjs';
import {
  createSample,
  sampleParts,
  zipParts,
  worksheet,
  NS,
  REL,
  PKG,
} from '../core/sample.mjs';

const inspect = (body) =>
  inspectWorkbook(
    zipParts({ ...sampleParts(), 'xl/worksheets/sheet1.xml': worksheet(body) }),
  ).sheets[0];
const ids = (sheet) => sheet.findings.map((finding) => finding.id);

test('core imports and inspects without Node Buffer globals', () => {
  const script = `globalThis.Buffer = undefined;
    const {inspectWorkbook}=await import('./core/workbook.mjs');
    const {createSample}=await import('./core/sample.mjs');
    if(inspectWorkbook(createSample()).storedCellCount!==13) throw Error('Wrong result');`;
  const result = spawnSync(
    process.execPath,
    ['--input-type=module', '-e', script],
    { cwd: new URL('..', import.meta.url), encoding: 'utf8', timeout: 10000 },
  );
  assert.equal(result.status, 0, result.stderr);
});

test('real sample detects the concealed M85 value without confusing populated and candidate bounds', () => {
  const report = inspectWorkbook(createSample()),
    sheet = report.sheets[0];
  assert.equal(report.storedCellCount, 13);
  assert.equal(sheet.populatedCount, 13);
  assert.equal(rangeLabel(sheet.ranges.populated), 'A1:M85');
  assert.equal(rangeLabel(sheet.ranges.candidates), 'A1:F24');
  assert.deepEqual(ids(sheet), ['concealed', 'wide-print-area']);
  assert.match(sheet.findings[0].evidence, /M85/);
  assert.equal(sheet.pageSetup.orientation, 'portrait');
  assert.ok(!JSON.stringify(report).includes('Office supplies'));
});

test('empty, formatting-only and numeric-zero cells have distinct footprints', () => {
  const sheet = inspect(
    '<dimension ref="A1:M85"/><sheetData><row r="1"><c r="A1"><v>0</v></c></row><row r="85"><c r="M85" s="1"/></row></sheetData>',
  );
  assert.equal(sheet.cells, 2);
  assert.equal(sheet.populatedCount, 1);
  assert.equal(rangeLabel(sheet.ranges.populated), 'A1:A1');
  assert.ok(ids(sheet).includes('empty-cells'));
  assert.ok(!ids(sheet).includes('concealed'));
  assert.equal(inspect('<sheetData/>').populatedCount, 0);
});

test('hidden rows, hidden columns, row and column number formats are applied', () => {
  const sheet = inspect(
    '<cols><col min="2" max="2" hidden="1"/><col min="3" max="3" style="1"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row><row r="2" hidden="true"><c r="A2"><v>4</v></c></row><row r="3" customFormat="1" s="1"><c r="A3"><v>5</v></c></row></sheetData>',
  );
  assert.equal(rangeLabel(sheet.ranges.candidates), 'A1:A1');
  assert.match(
    sheet.findings.find((f) => f.id === 'concealed').evidence,
    /2 populated.*C1, A3/,
  );
  assert.match(
    sheet.findings.find((f) => f.id === 'hidden-data').evidence,
    /2 populated.*B1, A2/,
  );
});

test('formulas count without executing and errors honor the actual print setting', () => {
  const body =
    '<sheetData><row r="1"><c r="A1" t="str"><f>HYPERLINK("https://example.com","x")</f><v/></c><c r="B1" t="e"><v>#N/A</v></c></row></sheetData>';
  const sheet = inspect(body + '<pageSetup errors="blank"/>');
  assert.equal(sheet.formulaCount, 1);
  assert.equal(sheet.populatedCount, 2);
  assert.equal(rangeLabel(sheet.ranges.candidates), 'A1:A1');
  assert.ok(ids(sheet).includes('blank-errors'));
  assert.ok(!ids(inspect(body)).includes('blank-errors'));
});

test('shared and inline rich text distinguish empty strings from actual content', () => {
  const parts = sampleParts();
  parts['xl/_rels/workbook.xml.rels'] = parts[
    'xl/_rels/workbook.xml.rels'
  ].replace(
    '</Relationships>',
    `<Relationship Id="strings" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
  );
  parts['xl/sharedStrings.xml'] =
    `<sst xmlns="${NS}"><si><t/></si><si><r><t>Value &amp; text</t></r></si></sst>`;
  parts['xl/worksheets/sheet1.xml'] = worksheet(
    '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t/></is></c><c r="D1" t="inlineStr"><is><r><t> </t></r></is></c></row></sheetData>',
  );
  const sheet = inspectWorkbook(zipParts(parts)).sheets[0];
  assert.equal(sheet.populatedCount, 2);
  assert.equal(rangeLabel(sheet.ranges.populated), 'B1:D1');
  parts['xl/worksheets/sheet1.xml'] = worksheet(
    '<sheetData><row><c t="s"><v>2</v></c></row></sheetData>',
  );
  assert.throws(
    () => inspectWorkbook(zipParts(parts)),
    /Shared-string reference/,
  );
});

test('white RGB text is a review cue, not assumed invisible', () => {
  const parts = sampleParts();
  parts['xl/styles.xml'] = parts['xl/styles.xml'].replace(
    '<sz val="11"/>',
    '<color rgb="FFFFFFFF"/><sz val="11"/>',
  );
  const sheet = inspectWorkbook(zipParts(parts)).sheets[0];
  assert.ok(ids(sheet).includes('white-font'));
  assert.equal(rangeLabel(sheet.ranges.candidates), 'A1:F24');
});

test('applyNumberFormat=false uses the base style, not the child format', () => {
  const parts = sampleParts();
  parts['xl/styles.xml'] = parts['xl/styles.xml'].replace(
    'applyNumberFormat="1"',
    'applyNumberFormat="0"',
  );
  assert.ok(
    !ids(inspectWorkbook(zipParts(parts)).sheets[0]).includes('concealed'),
  );
});

test('manual breaks are separated from automatic records', () => {
  const sheet = inspect(
    '<sheetData/><rowBreaks><brk id="20" man="1"/><brk id="40" man="0"/></rowBreaks><colBreaks><brk id="5" man="true"/></colBreaks>',
  );
  assert.deepEqual(sheet.manualBreaks, {
    rows: [20],
    columns: [5],
    rowCount: 1,
    columnCount: 1,
  });
});

test('two-cell drawing anchors are inspected without fetching image or external bytes', () => {
  const parts = sampleParts();
  parts['xl/worksheets/sheet1.xml'] = worksheet(
    '<sheetData><row><c><v>1</v></c></row></sheetData><drawing r:id="draw"/>',
  );
  parts['xl/worksheets/_rels/sheet1.xml.rels'] =
    `<Relationships xmlns="${PKG}"><Relationship Id="draw" Type="${REL}/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="external" Type="${REL}/hyperlink" Target="http://127.0.0.1/private" TargetMode="External"/></Relationships>`;
  parts['xl/drawings/drawing1.xml'] =
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from><xdr:col>12</xdr:col><xdr:row>84</xdr:row></xdr:from><xdr:to><xdr:col>13</xdr:col><xdr:row>89</xdr:row></xdr:to></xdr:twoCellAnchor></xdr:wsDr>';
  const sheet = inspectWorkbook(zipParts(parts)).sheets[0];
  assert.equal(rangeLabel(sheet.ranges.drawingBounds), 'M85:N90');
  assert.ok(ids(sheet).includes('drawing-extent'));
  assert.ok(sheet.warnings.some((text) => /External/.test(text)));
});

test('print areas handle quoted commas, apostrophes, unions and full columns/rows', () => {
  assert.deepEqual(
    printAreas(
      "'Bob''s, report'!$A$1:$F$24,'Bob''s, report'!$M$85",
      "Bob's, report",
    ),
    [range('A1:F24'), range('M85')],
  );
  assert.equal(
    rangeLabel(printAreas('Sheet1!$A:$F', 'Sheet1')[0]),
    'A1:F1048576',
  );
  assert.equal(
    rangeLabel(printAreas('Sheet1!$1:$24', 'Sheet1')[0]),
    'A1:XFD24',
  );
  for (const value of [
    'OFFSET(A1,0,0,5,5)',
    'Other!A1',
    'Sheet1!#REF!',
    'Sheet1!XFE1',
    'Sheet1!A0',
    'Sheet1!A5:A1',
  ])
    assert.throws(() => printAreas(value, 'Sheet1'));
});

test('dynamic print areas are unknown, not replaced with a guessed rectangle', () => {
  const parts = sampleParts();
  parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace(
    "'Monthly report'!$A$1:$M$85",
    'OFFSET(A1,0,0,10,5)',
  );
  const sheet = inspectWorkbook(zipParts(parts)).sheets[0];
  assert.equal(sheet.printArea.resolved, false);
  assert.ok(sheet.warnings.some((text) => /cannot resolve/.test(text)));
  assert.ok(!ids(sheet).includes('wide-print-area'));
});

test('XML handles namespaces, predefined/numeric entities and CDATA exactly once', () => {
  const root = parseXml(
    '<x:a xmlns:x="urn:example" label="A &amp; B &quot;x&quot;"><x:t>&amp;lt;&#x4F60;&#22909;</x:t><x:t><![CDATA[&amp;]]></x:t></x:a>',
  );
  assert.equal(root.ns, 'urn:example');
  assert.equal(attr(root, 'label'), 'A & B "x"');
  assert.equal(textOf(child(root, 't')), '&lt;你好');
  assert.equal(textOf(root), '&lt;你好&amp;');
});

test('unsafe, deeply nested, malformed and reserved-name XML is refused', () => {
  for (const value of [
    '<!DOCTYPE a [<!ENTITY x "boom">]><a>&x;</a>',
    '<a>&custom;</a>',
    '<a>&#0;</a>',
    '<a><b></a>',
    '<a/><b/>',
    '<z:a/>',
    '<__proto__/>',
    '<a>'.repeat(60) + '</a>'.repeat(60),
  ])
    assert.throws(() => parseXml(value), value);
});

test('macro workbooks and internal relationship traversal are rejected', () => {
  const parts = sampleParts();
  parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
    'sheet.main+xml',
    'macroEnabled.main+xml',
  );
  assert.throws(() => inspectWorkbook(zipParts(parts)), /Macro-enabled/);
  for (const target of [
    '../../../secret.xml',
    'https://example.com/test.xml',
    '//example.com/test.xml',
  ]) {
    const source = sampleParts();
    source['xl/_rels/workbook.xml.rels'] = source[
      'xl/_rels/workbook.xml.rels'
    ].replace('worksheets/sheet1.xml', target);
    assert.throws(
      () => inspectWorkbook(zipParts(source)),
      /relationship|Relationship/,
    );
  }
});

test('out-of-range styles, duplicate positions and invalid column intervals cannot create a plausible report', () => {
  for (const body of [
    '<sheetData><row><c s="25"><v>1</v></c></row></sheetData>',
    '<sheetData><row><c r="A1"/><c r="A1"/></row></sheetData>',
    '<sheetData><row r="1"/><row r="1"/></sheetData>',
    '<cols><col min="2" max="1"/></cols><sheetData/>',
    '<sheetData><row r="1"><c r="A2"/></row></sheetData>',
  ])
    assert.throws(() => inspect(body));
});

test('100,001 cells are rejected; sparse extreme coordinates do not allocate a worksheet grid', () => {
  assert.equal(
    inspect(
      '<sheetData><row r="1048576"><c r="XFD1048576"><v>1</v></c></row></sheetData>',
    ).cells,
    1,
  );
  const rows = Array.from(
    { length: 1001 },
    (_, r) => `<row r="${r + 1}">${'<c><v>1</v></c>'.repeat(100)}</row>`,
  ).join('');
  assert.throws(() => inspect(`<sheetData>${rows}</sheetData>`), /100,000/);
});

test('strict spreadsheet namespaces are accepted; a foreign worksheet namespace is not', () => {
  const parts = sampleParts();
  for (const path of Object.keys(parts))
    parts[path] = parts[path]
      .replaceAll(NS, 'http://purl.oclc.org/ooxml/spreadsheetml/main')
      .replaceAll(
        REL,
        'http://purl.oclc.org/ooxml/officeDocument/relationships',
      );
  assert.equal(inspectWorkbook(zipParts(parts)).storedCellCount, 13);
  parts['xl/worksheets/sheet1.xml'] =
    '<worksheet xmlns="urn:other"><sheetData/></worksheet>';
  assert.throws(() => inspectWorkbook(zipParts(parts)), /Invalid worksheet/);
});
