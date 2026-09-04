import { zipSync, strToU8 } from 'fflate';

export const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PKG =
  'http://schemas.openxmlformats.org/package/2006/relationships';
export const worksheet = (body) =>
  `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS}" xmlns:r="${REL}">${body}</worksheet>`;

export function sampleParts() {
  const records = [
    [1, 'Monthly report', 1800],
    [4, 'Office supplies', 240],
    [8, 'Shipping', 120],
    [12, 'Equipment', 900],
    [18, 'Travel', 310],
    [24, 'Total', 3370],
  ];
  const rows = records
    .map(
      ([row, label, value]) =>
        `<row r="${row}"><c r="A${row}" t="inlineStr"><is><t>${label}</t></is></c><c r="F${row}"><v>${value}</v></c></row>`,
    )
    .join('');
  return {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${NS}" xmlns:r="${REL}"><sheets><sheet name="Monthly report" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Monthly report'!$A$1:$M$85</definedName></definedNames></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode=";;;"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    'xl/worksheets/sheet1.xml': worksheet(
      `<dimension ref="A1:M85"/><sheetData>${rows}<row r="85"><c r="M85" s="1"><v>1</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait"/>`,
    ),
  };
}

export function zipParts(parts) {
  return zipSync(
    Object.fromEntries(
      Object.entries(parts).map(([name, value]) => [name, strToU8(value)]),
    ),
    { level: 6 },
  );
}

export const createSample = () => zipParts(sampleParts());
