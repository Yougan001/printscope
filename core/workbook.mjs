import { openZip } from './zip.mjs';
import { parseXml, attr, child, children, textOf } from './xml.mjs';
import {
  cellAddress,
  columnName,
  contains,
  include,
  printAreas,
  range,
  rangeLabel,
} from './ranges.mjs';

export const WORKBOOK_LIMITS = Object.freeze({
  sheets: 50,
  cells: 100_000,
  styles: 20_000,
  strings: 100_000,
});
const sheetNamespaces = new Set([
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://purl.oclc.org/ooxml/spreadsheetml/main',
]);
const packageNamespace =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const typeNamespace =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const relNamespaces = [
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
];
const workbookType =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const yes = (value) => value === '1' || value === 'true';
const relationshipId = (node) =>
  relNamespaces.map((ns) => attr(node, 'id', ns)).find(Boolean);
const relationIs = (entry, type) =>
  relNamespaces.some((ns) => entry.type === `${ns}/${type}`);

function number(value, fallback, maximum = 1048576) {
  if (value === undefined) return fallback;
  if (!/^[0-9]+$/.test(value) || Number(value) > maximum)
    throw new Error('Invalid workbook numeric attribute.');
  return Number(value);
}

function targetPart(source, target) {
  let decoded;
  try {
    decoded = decodeURIComponent(target);
  } catch {
    throw new Error('Invalid relationship target encoding.');
  }
  if (
    !decoded ||
    /[\\:#?\u0000-\u001f]/.test(decoded) ||
    decoded.startsWith('//')
  )
    throw new Error('Unsafe internal workbook relationship.');
  const stack = decoded.startsWith('/') ? [] : source.split('/').slice(0, -1);
  for (const part of decoded.split('/')) {
    if (part === '..') {
      if (!stack.length) throw new Error('Relationship escapes the workbook.');
      stack.pop();
    } else if (part && part !== '.') stack.push(part);
  }
  return stack.join('/');
}

function richTextPresent(node) {
  return (
    children(node, 't').some((text) => textOf(text).length > 0) ||
    children(node, 'r').some((run) =>
      children(run, 't').some((text) => textOf(text).length > 0),
    )
  );
}

export function inspectWorkbook(bytes) {
  const zip = openZip(bytes),
    budget = { nodes: 0, cells: 0 };
  const xml = (path) => parseXml(zip.read(path), budget);
  function relationships(source) {
    const slash = source.lastIndexOf('/');
    const path = `${source.slice(0, slash + 1)}_rels/${source.slice(slash + 1)}.rels`;
    if (!zip.has(path)) return new Map();
    const root = xml(path);
    if (root.name !== 'Relationships' || root.ns !== packageNamespace)
      throw new Error('Invalid relationships part.');
    const result = new Map();
    for (const entry of children(root, 'Relationship')) {
      const id = attr(entry, 'Id'),
        type = attr(entry, 'Type'),
        target = attr(entry, 'Target');
      if (!id || !type || !target || result.has(id))
        throw new Error('Missing or duplicate workbook relationship.');
      const mode = attr(entry, 'TargetMode');
      if (mode && mode !== 'Internal' && mode !== 'External')
        throw new Error('Unknown relationship mode.');
      result.set(id, {
        type,
        external: mode === 'External',
        target: mode === 'External' ? null : targetPart(source, target),
      });
    }
    return result;
  }
  const types = xml('[Content_Types].xml');
  if (types.name !== 'Types' || types.ns !== typeNamespace)
    throw new Error('Invalid workbook content types.');
  const overrides = new Map();
  for (const entry of children(types, 'Override')) {
    const path = attr(entry, 'PartName'),
      type = attr(entry, 'ContentType');
    if (!path || !type || overrides.has(path))
      throw new Error('Invalid content-type declarations.');
    overrides.set(path, type);
  }
  if (
    [...children(types, 'Override'), ...children(types, 'Default')].some(
      (entry) =>
        /macroEnabled|vbaProject/i.test(attr(entry, 'ContentType') ?? ''),
    ) ||
    zip.names.some((path) => /vbaProject\.bin$/i.test(path))
  )
    throw new Error(
      'Macro-enabled workbooks are not supported. Save a separate .xlsx copy first.',
    );
  const office = [...relationships('').values()].filter((entry) =>
    relationIs(entry, 'officeDocument'),
  );
  if (
    office.length !== 1 ||
    office[0].external ||
    overrides.get(`/${office[0].target}`) !== workbookType
  )
    throw new Error(
      'Use an ordinary .xlsx workbook, not .xls, .xlsb, .xlsm or an encrypted file.',
    );
  const workbookPath = office[0].target,
    workbook = xml(workbookPath),
    links = relationships(workbookPath);
  if (workbook.name !== 'workbook' || !sheetNamespaces.has(workbook.ns))
    throw new Error('Unsupported workbook namespace.');
  const sheetNodes = children(child(workbook, 'sheets'), 'sheet');
  if (!sheetNodes.length || sheetNodes.length > WORKBOOK_LIMITS.sheets)
    throw new Error('Use a workbook with 1–50 sheets.');
  function optionalPart(type, rootName) {
    const matching = [...links.values()].filter((entry) =>
      relationIs(entry, type),
    );
    if (matching.length > 1 || matching.some((entry) => entry.external))
      throw new Error(`Invalid ${type} relationship.`);
    if (!matching.length) return null;
    const root = xml(matching[0].target);
    if (root.name !== rootName || root.ns !== workbook.ns)
      throw new Error(`Invalid ${type} part.`);
    return root;
  }
  const shared = children(optionalPart('sharedStrings', 'sst'), 'si');
  if (shared.length > WORKBOOK_LIMITS.strings)
    throw new Error('Shared-string table exceeds 100,000 entries.');
  const stringPresence = shared.map(richTextPresent);
  const styleRoot = optionalPart('styles', 'styleSheet');
  const formats = new Map();
  for (const item of children(child(styleRoot, 'numFmts'), 'numFmt')) {
    const id = number(attr(item, 'numFmtId'), -1, 65535);
    if (id < 0 || formats.has(id))
      throw new Error('Invalid custom number-format identifier.');
    formats.set(id, attr(item, 'formatCode') ?? '');
  }
  const fonts = children(child(styleRoot, 'fonts'), 'font');
  const bases = children(child(styleRoot, 'cellStyleXfs'), 'xf');
  const xfs = children(child(styleRoot, 'cellXfs'), 'xf');
  if (
    xfs.length > WORKBOOK_LIMITS.styles ||
    bases.length > WORKBOOK_LIMITS.styles ||
    fonts.length > WORKBOOK_LIMITS.styles
  )
    throw new Error('Workbook style table exceeds the inspection limit.');
  const styles = (xfs.length ? xfs : [null]).map((xf) => {
    const base = bases[number(attr(xf, 'xfId'), 0, WORKBOOK_LIMITS.styles)];
    const effective = (name, flag) =>
      attr(xf, flag) === '0' || attr(xf, flag) === 'false'
        ? attr(base, name)
        : (attr(xf, name) ?? attr(base, name));
    const code =
      formats.get(
        number(effective('numFmtId', 'applyNumberFormat'), 0, 65535),
      ) ?? '';
    const font =
      fonts[
        number(effective('fontId', 'applyFont'), 0, WORKBOOK_LIMITS.styles)
      ];
    const rgb = attr(child(font, 'color'), 'rgb');
    return {
      concealed: code === ';;;',
      whiteFont: /^(?:FF)?FFFFFF$/i.test(rgb ?? ''),
    };
  });
  const definitions = children(child(workbook, 'definedNames'), 'definedName');
  const names = new Set(),
    targets = new Set();
  const sheets = sheetNodes.map((sheet, index) => {
    const name = attr(sheet, 'name');
    if (!name || name.length > 128 || names.has(name))
      throw new Error('Missing, duplicate or overlong sheet name.');
    names.add(name);
    const entry = links.get(relationshipId(sheet));
    if (!entry || entry.external)
      throw new Error('Worksheet relationship is missing or external.');
    if (!relationIs(entry, 'worksheet'))
      return {
        name,
        state: attr(sheet, 'state') ?? 'visible',
        supported: false,
        reason:
          'This sheet is not a cell worksheet (for example, a chart sheet).',
      };
    if (targets.has(entry.target))
      throw new Error('Multiple worksheets refer to the same part.');
    targets.add(entry.target);
    const root = xml(entry.target);
    if (
      root.name !== 'worksheet' ||
      root.ns !== workbook.ns ||
      !child(root, 'sheetData')
    )
      throw new Error('Invalid worksheet part.');
    const findings = [],
      warnings = [];
    const add = (id, title, evidence, action, level = 'review') =>
      findings.push({ id, title, evidence, action, level });
    const groups = new Map();
    const mark = (id, reference) => {
      if (!groups.has(id)) groups.set(id, { count: 0, examples: [] });
      const group = groups.get(id);
      group.count++;
      if (group.examples.length < 12) group.examples.push(reference);
    };
    const colHidden = new Uint8Array(16385),
      colStyles = new Uint32Array(16385),
      colSeen = new Uint8Array(16385);
    for (const col of children(child(root, 'cols'), 'col')) {
      const min = number(attr(col, 'min'), 0, 16384),
        max = number(attr(col, 'max'), 0, 16384);
      if (!min || max < min) throw new Error('Invalid column interval.');
      for (let column = min; column <= max; column++) {
        if (colSeen[column])
          throw new Error('Overlapping column definitions are not supported.');
        colSeen[column] = 1;
        colHidden[column] = yes(attr(col, 'hidden')) ? 1 : 0;
        colStyles[column] = number(
          attr(col, 'style'),
          0,
          WORKBOOK_LIMITS.styles,
        );
      }
    }
    const setup = child(root, 'pageSetup');
    const rowsSeen = new Set();
    let storedCells = null,
      populated = null,
      candidates = null,
      cells = 0,
      populatedCount = 0,
      previousRow = 0,
      formulaCount = 0;
    const hiddenRows = [];
    for (const row of children(child(root, 'sheetData'), 'row')) {
      const rowNumber = number(attr(row, 'r'), previousRow + 1);
      if (!rowNumber || rowsSeen.has(rowNumber))
        throw new Error('Invalid or duplicate worksheet row.');
      rowsSeen.add(rowNumber);
      previousRow = rowNumber;
      const hiddenRow = yes(attr(row, 'hidden'));
      if (hiddenRow) hiddenRows.push(rowNumber);
      const rowStyle = yes(attr(row, 'customFormat'))
        ? number(attr(row, 's'), 0, WORKBOOK_LIMITS.styles)
        : null;
      let previousColumn = 0;
      const cellsSeen = new Set();
      for (const cell of children(row, 'c')) {
        if (++budget.cells > WORKBOOK_LIMITS.cells)
          throw new Error('Workbook exceeds 100,000 stored cells.');
        cells++;
        const point = attr(cell, 'r')
          ? cellAddress(attr(cell, 'r'))
          : { row: rowNumber, column: previousColumn + 1 };
        if (
          point.row !== rowNumber ||
          point.column > 16384 ||
          cellsSeen.has(point.column)
        )
          throw new Error('Invalid or duplicate cell position.');
        cellsSeen.add(point.column);
        previousColumn = point.column;
        const reference = `${columnName(point.column)}${point.row}`;
        storedCells = include(storedCells, point.row, point.column);
        const type = attr(cell, 't'),
          value = textOf(child(cell, 'v')),
          formula = !!child(cell, 'f');
        if (formula) formulaCount++;
        let hasValue = value.length > 0 || formula;
        if (type === 'inlineStr')
          hasValue = richTextPresent(child(cell, 'is')) || formula;
        if (type === 's') {
          const stringIndex = number(value, undefined, WORKBOOK_LIMITS.strings);
          if (stringIndex === undefined || stringIndex >= stringPresence.length)
            throw new Error(
              'Shared-string reference is missing or out of range.',
            );
          hasValue = stringPresence[stringIndex] || formula;
        }
        const styleIndex = number(
          attr(cell, 's'),
          rowStyle ?? colStyles[point.column],
          WORKBOOK_LIMITS.styles,
        );
        const style = styles[styleIndex];
        if (!style) throw new Error('Cell style reference is out of range.');
        if (!hasValue) continue;
        populatedCount++;
        populated = include(populated, point.row, point.column);
        const printBlankError =
          type === 'e' && attr(setup, 'errors') === 'blank';
        if (style.concealed) mark('concealed', reference);
        if (style.whiteFont) mark('white-font', reference);
        if (printBlankError) mark('blank-errors', reference);
        if (hiddenRow || colHidden[point.column])
          mark('hidden-data', reference);
        if (
          !style.concealed &&
          !printBlankError &&
          !hiddenRow &&
          !colHidden[point.column]
        )
          candidates = include(candidates, point.row, point.column);
      }
    }
    let dimension = null;
    if (attr(child(root, 'dimension'), 'ref')) {
      try {
        dimension = range(attr(child(root, 'dimension'), 'ref'));
      } catch {
        warnings.push('The stored dimension is not a valid A1 range.');
      }
    }
    if (dimension && storedCells && !contains(dimension, storedCells))
      warnings.push(
        'Stored cell coordinates extend outside the declared worksheet dimension.',
      );
    if (storedCells && populated && !contains(populated, storedCells))
      add(
        'empty-cells',
        'Stored cells extend beyond populated cells',
        `${rangeLabel(storedCells)} contains cell records; populated cells occupy ${rangeLabel(populated)}. Formatting-only cells can extend the stored footprint.`,
        'Use Ctrl+End in Excel and inspect the extra rows or columns. Clear unused formatting only after checking that it is not intentional.',
      );
    if (dimension && (!populated || !contains(populated, dimension)))
      add(
        'dimension',
        'Declared used range extends beyond populated cells',
        `The worksheet declares ${rangeLabel(dimension)}; populated cell bounds are ${rangeLabel(populated) ?? 'empty'}. This metadata can be stale.`,
        'Compare Ctrl+End with your actual report. The declared range alone does not prove extra printed pages.',
        'info',
      );
    const groupDetails = {
      concealed: [
        'Values use the display-hiding format ;;;',
        'Review Format Cells → Number. Change the format only if these values should be visible.',
      ],
      'white-font': [
        'Values use an explicit white font',
        'Inspect the cell fill and conditional formatting in Excel. White text is only invisible on a matching background.',
      ],
      'blank-errors': [
        'Error cells are configured to print as blank',
        'Open Page Setup → Sheet → Cell errors as, then inspect and fix the formulas.',
      ],
      'hidden-data': [
        'Hidden rows or columns contain data',
        'Review hidden rows and columns. Do not delete them without checking formulas and dependencies.',
      ],
    };
    for (const [id, group] of groups) {
      const [title, action] = groupDetails[id];
      add(
        id,
        title,
        `${group.count} populated cell(s): ${group.examples.join(', ')}${group.count > group.examples.length ? ' …' : ''}.`,
        action,
        id === 'hidden-data' ? 'info' : 'review',
      );
    }
    const areaDefinitions = definitions.filter(
      (definition) =>
        attr(definition, 'name') === '_xlnm.Print_Area' &&
        attr(definition, 'localSheetId') === String(index),
    );
    if (areaDefinitions.length > 1)
      throw new Error('Duplicate print-area definitions.');
    let areas = [],
      areaExpression = null;
    if (areaDefinitions.length) {
      areaExpression = textOf(areaDefinitions[0]);
      if (areaExpression.length > 10000)
        throw new Error('Print-area expression exceeds the inspection limit.');
      try {
        areas = printAreas(areaExpression, name);
      } catch {
        warnings.push(
          'The print area uses an expression that this version cannot resolve. Review it in Excel.',
        );
      }
      if (areas.length > 1)
        add(
          'multiple-areas',
          'Multiple separate print areas are defined',
          areas.map(rangeLabel).join(', '),
          'Excel prints separate areas separately. Review Page Layout → Print Area and keep only the intended report areas.',
        );
      if (areas.length && areas.some((area) => !contains(candidates, area)))
        add(
          'wide-print-area',
          'Print area extends beyond candidate-visible cell bounds',
          `Print area: ${areas.map(rangeLabel).join(', ')}. Candidate-visible cell bounds: ${rangeLabel(candidates) ?? 'empty'}. Drawings and formatting may intentionally occupy the remaining area.`,
          'Review the print area in Page Break Preview. Select the intended report, then use Page Layout → Print Area → Set Print Area if appropriate.',
        );
      if (populated && areas.length === 1 && !contains(areas[0], populated))
        add(
          'excluded-data',
          'Some populated cells are outside the print area',
          `Populated bounds: ${rangeLabel(populated)}. Print area: ${rangeLabel(areas[0])}.`,
          'Check whether this exclusion is intentional before changing the print area.',
          'info',
        );
    }
    const breaks = { rows: [], columns: [] };
    for (const [tag, key, max] of [
      ['rowBreaks', 'rows', 1048576],
      ['colBreaks', 'columns', 16384],
    ]) {
      for (const entry of children(child(root, tag), 'brk'))
        if (yes(attr(entry, 'man')))
          breaks[key].push(number(attr(entry, 'id'), 0, max));
    }
    if (breaks.rows.length || breaks.columns.length)
      add(
        'manual-breaks',
        'Manual page breaks are present',
        `${breaks.rows.length} row break(s), ${breaks.columns.length} column break(s). ${hiddenRows.length || colHidden.includes(1) ? 'This sheet also has hidden rows or columns.' : ''}`,
        'Inspect Page Break Preview. Use Page Layout → Breaks → Reset All Page Breaks only if these breaks are no longer wanted.',
      );
    const sheetLinks = relationships(entry.target);
    let drawingCount = 0,
      drawingBounds = null;
    for (const drawing of children(root, 'drawing')) {
      const link = sheetLinks.get(relationshipId(drawing));
      if (!link || link.external || !relationIs(link, 'drawing')) {
        warnings.push('A drawing reference could not be inspected.');
        continue;
      }
      const document = xml(link.target);
      if (
        document.name !== 'wsDr' ||
        ![
          'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
          'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing',
        ].includes(document.ns)
      ) {
        warnings.push('An unsupported drawing part was skipped.');
        continue;
      }
      for (const anchor of document.children.filter(
        (node) => typeof node !== 'string',
      )) {
        drawingCount++;
        const from = child(anchor, 'from'),
          to = child(anchor, 'to');
        if (!from) {
          warnings.push(
            'An absolute-positioned drawing has no cell anchor; its extent is not estimated.',
          );
          continue;
        }
        for (const marker of [from, to].filter(Boolean)) {
          const row = number(textOf(child(marker, 'row')), 0, 1048575) + 1;
          const column = number(textOf(child(marker, 'col')), 0, 16383) + 1;
          drawingBounds = include(drawingBounds, row, column);
        }
        if (!to)
          warnings.push(
            'A one-cell drawing anchor is known, but its size is not converted into worksheet cells.',
          );
      }
    }
    if (drawingBounds && !contains(candidates, drawingBounds))
      add(
        'drawing-extent',
        'Drawing anchors extend beyond candidate-visible cells',
        `Drawing anchor bounds: ${rangeLabel(drawingBounds)}. Cell bounds: ${rangeLabel(candidates) ?? 'empty'}.`,
        'Use Excel’s Selection Pane or Go To Special → Objects to inspect distant objects. Confirm each object’s Print object setting before removing it.',
      );
    if (
      children(root, 'legacyDrawing').length ||
      children(root, 'legacyDrawingHF').length
    )
      warnings.push(
        'Legacy VML drawings and header/footer drawings are not inspected.',
      );
    if (children(root, 'conditionalFormatting').length)
      warnings.push(
        'Conditional formatting is present; its display effects are not evaluated.',
      );
    if (children(root, 'mergeCells').length)
      warnings.push(
        'Merged cells are present; candidate-visible bounds describe cell values, not merged visual extents.',
      );
    if (yes(attr(child(root, 'sheetFormatPr'), 'zeroHeight')))
      warnings.push(
        'Default row height is hidden; implicit row visibility is not included in candidate-visible bounds.',
      );
    if (
      [...links.values(), ...sheetLinks.values()].some((item) => item.external)
    )
      warnings.push('External relationships are not fetched or evaluated.');
    if (formulaCount)
      warnings.push(
        'Formulas are not calculated. Formula cells count as populated even when a cached result is empty.',
      );
    return {
      name,
      state: attr(sheet, 'state') ?? 'visible',
      supported: true,
      cells,
      populatedCount,
      formulaCount,
      ranges: { dimension, storedCells, populated, candidates, drawingBounds },
      printArea: {
        expression: areaExpression,
        ranges: areas,
        resolved: !areaExpression || areas.length > 0,
      },
      pageSetup: {
        orientation: attr(setup, 'orientation') ?? 'default',
        paperSize: attr(setup, 'paperSize') ?? null,
        scale: attr(setup, 'scale') ?? null,
        fitToPage: yes(
          attr(child(child(root, 'sheetPr'), 'pageSetUpPr'), 'fitToPage'),
        ),
        fitToWidth: attr(setup, 'fitToWidth') ?? null,
        fitToHeight: attr(setup, 'fitToHeight') ?? null,
        errors: attr(setup, 'errors') ?? 'displayed',
      },
      manualBreaks: {
        rows: breaks.rows.slice(0, 100),
        columns: breaks.columns.slice(0, 100),
        rowCount: breaks.rows.length,
        columnCount: breaks.columns.length,
      },
      drawingCount,
      findings,
      warnings: [...new Set(warnings)],
    };
  });
  return {
    version: 1,
    sheets,
    storedCellCount: budget.cells,
    findingCount: sheets.reduce(
      (count, sheet) => count + (sheet.findings?.length ?? 0),
      0,
    ),
    limitations: [
      'This is a read-only structure inspection, not Excel pagination or a workbook repair.',
      'No formulas, macros, external links, printer drivers or conditional formats are executed.',
      'Candidate-visible bounds exclude only known ;;; formats, hidden row/column records and errors printed blank. They are not a recommended print area.',
      'Theme colors, inherited display effects, object print settings, row heights, merged extents, headers and printer metrics can affect the final output.',
    ],
  };
}
