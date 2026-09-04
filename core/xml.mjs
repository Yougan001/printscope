import { XMLParser, XMLValidator } from 'fast-xml-parser';

const predefined = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

function decode(text) {
  return text.replace(/&([^;\s]*);?/g, (match, name) => {
    if (!match.endsWith(';')) throw new Error('Incomplete XML entity.');
    if (Object.hasOwn(predefined, name)) return predefined[name];
    if (!/^#(?:[0-9]+|x[0-9a-fA-F]+)$/.test(name))
      throw new Error('Custom XML entities are not supported.');
    const point =
      name[1] === 'x' ? parseInt(name.slice(2), 16) : Number(name.slice(1));
    if (
      !(
        point === 9 ||
        point === 10 ||
        point === 13 ||
        (point >= 32 && point <= 0xd7ff) ||
        (point >= 0xe000 && point <= 0xfffd) ||
        (point >= 0x10000 && point <= 0x10ffff)
      )
    )
      throw new Error('Invalid XML character.');
    return String.fromCodePoint(point);
  });
}

export function parseXml(source, budget = { nodes: 0 }) {
  if (/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source))
    throw new Error(
      'DOCTYPE declarations and custom entities are not supported.',
    );
  // The standalone validator currently imports Node-only Buffer code in browsers.
  // eslint-disable-next-line typescript/no-deprecated
  const valid = XMLValidator.validate(source);
  if (valid !== true) throw new Error(`Invalid workbook XML: ${valid.err.msg}`);
  const parser = new XMLParser({
    preserveOrder: true,
    ignoreAttributes: false,
    parseTagValue: false,
    parseAttributeValue: false,
    trimValues: false,
    processEntities: false,
    ignoreDeclaration: true,
    ignorePiTags: true,
    cdataPropName: '#cdata',
    maxNestedTags: 48,
    strictReservedNames: true,
    updateTag(name) {
      if (++budget.nodes > 600_000)
        throw new Error('Workbook exceeds the XML node limit.');
      return name;
    },
  });
  function expand(records, inherited) {
    const output = [];
    for (const record of records) {
      const key = Object.keys(record).find((name) => name !== ':@');
      if (key === '#text') {
        output.push(decode(record[key]));
        continue;
      }
      if (key === '#cdata') {
        output.push(record[key].map((value) => value['#text'] ?? '').join(''));
        continue;
      }
      if (!key || key.startsWith('?') || key.startsWith('#')) continue;
      const raw = record[':@'] ?? {};
      let namespaces = inherited;
      if (
        Object.keys(raw).some(
          (name) => name === '@_xmlns' || name.startsWith('@_xmlns:'),
        )
      ) {
        namespaces = { ...inherited };
        for (const [name, value] of Object.entries(raw)) {
          if (name === '@_xmlns') namespaces[''] = decode(value);
          else if (name.startsWith('@_xmlns:'))
            namespaces[name.slice(8)] = decode(value);
        }
      }
      const resolve = (name, attribute = false) => {
        const parts = name.split(':');
        if (parts.length > 2 || (parts.length === 2 && !namespaces[parts[0]]))
          throw new Error('Undeclared XML namespace.');
        return {
          name: parts.at(-1),
          ns:
            parts.length === 2
              ? namespaces[parts[0]]
              : attribute
                ? ''
                : (namespaces[''] ?? ''),
        };
      };
      const attributes = [];
      for (const [name, value] of Object.entries(raw)) {
        if (name === '@_xmlns' || name.startsWith('@_xmlns:')) continue;
        attributes.push({
          ...resolve(name.slice(2), true),
          value: decode(value),
        });
      }
      output.push({
        ...resolve(key),
        attributes,
        children: expand(record[key], namespaces),
      });
    }
    return output;
  }
  const roots = expand(parser.parse(source), {
    xml: 'http://www.w3.org/XML/1998/namespace',
  });
  const elements = roots.filter((value) => typeof value !== 'string');
  if (
    elements.length !== 1 ||
    roots.some((value) => typeof value === 'string' && value.trim())
  )
    throw new Error('Workbook XML must have one root element.');
  return elements[0];
}

export const attr = (node, name, namespace = '') =>
  node?.attributes.find((item) => item.name === name && item.ns === namespace)
    ?.value;
export const children = (node, name, namespace = node?.ns) =>
  node?.children.filter(
    (item) =>
      typeof item !== 'string' && item.name === name && item.ns === namespace,
  ) ?? [];
export const child = (node, name) => children(node, name)[0];
export function textOf(node) {
  return (
    node?.children
      .map((item) => (typeof item === 'string' ? item : textOf(item)))
      .join('') ?? ''
  );
}
