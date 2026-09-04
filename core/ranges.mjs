export function cellAddress(value) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$/.exec(value ?? '');
  if (!match)
    throw new Error(`Unsupported cell address: ${String(value).slice(0, 80)}`);
  let column = 0;
  for (const letter of match[1])
    column = column * 26 + letter.charCodeAt(0) - 64;
  const row = Number(match[2]);
  if (column > 16384 || row > 1048576)
    throw new Error('Cell address exceeds worksheet limits.');
  return { row, column };
}

export function columnName(number) {
  let result = '';
  while (number > 0) {
    number--;
    result = String.fromCharCode(65 + (number % 26)) + result;
    number = Math.floor(number / 26);
  }
  return result;
}

export function range(value) {
  const parts = value.split(':');
  if (parts.length > 2) throw new Error('Invalid cell range.');
  const start = cellAddress(parts[0]),
    end = cellAddress(parts[1] ?? parts[0]);
  if (end.row < start.row || end.column < start.column)
    throw new Error('Reversed cell range.');
  return {
    top: start.row,
    left: start.column,
    bottom: end.row,
    right: end.column,
  };
}

export const rangeLabel = (box) =>
  box
    ? `${columnName(box.left)}${box.top}:${columnName(box.right)}${box.bottom}`
    : null;
export const contains = (outer, inner) =>
  !!outer &&
  !!inner &&
  outer.top <= inner.top &&
  outer.left <= inner.left &&
  outer.bottom >= inner.bottom &&
  outer.right >= inner.right;
export function include(box, row, column) {
  if (!box) return { top: row, bottom: row, left: column, right: column };
  box.top = Math.min(box.top, row);
  box.bottom = Math.max(box.bottom, row);
  box.left = Math.min(box.left, column);
  box.right = Math.max(box.right, column);
  return box;
}

export function printAreas(expression, sheetName) {
  const parts = [];
  let quoted = false,
    start = 0;
  for (let index = 0; index < expression.length; index++) {
    if (expression[index] === "'") {
      if (quoted && expression[index + 1] === "'") index++;
      else quoted = !quoted;
    } else if (expression[index] === ',' && !quoted) {
      parts.push(expression.slice(start, index));
      start = index + 1;
    }
  }
  parts.push(expression.slice(start));
  if (quoted || parts.length > 100)
    throw new Error('Unsupported print-area expression.');
  return parts.map((part) => {
    const match = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(part.trim());
    if (!match || (match[1]?.replace(/''/g, "'") ?? match[2]) !== sheetName)
      throw new Error(
        'Only local, explicit print-area references are supported.',
      );
    const reference = match[3];
    if (/^\$?[A-Z]{1,3}:\$?[A-Z]{1,3}$/.test(reference)) {
      const [left, right] = reference.replaceAll('$', '').split(':');
      return range(`${left}1:${right}1048576`);
    }
    if (/^\$?[1-9][0-9]*:\$?[1-9][0-9]*$/.test(reference)) {
      const [top, bottom] = reference.replaceAll('$', '').split(':');
      return range(`A${top}:XFD${bottom}`);
    }
    return range(reference);
  });
}
