import { Inflate } from 'fflate';

export const ZIP_LIMITS = Object.freeze({
  archiveBytes: 20 * 1024 * 1024,
  entries: 4000,
  xmlBytes: 12 * 1024 * 1024,
  totalReadBytes: 64 * 1024 * 1024,
});
const decoder = new TextDecoder('utf-8', { fatal: true });
const crcTable = Uint32Array.from({ length: 256 }, (_, value) => {
  for (let bit = 0; bit < 8; bit++)
    value = (value >>> 1) ^ (value & 1 ? 0xedb88320 : 0);
  return value >>> 0;
});

function crc32(data) {
  let crc = 0xffffffff;
  for (const byte of data) crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 255];
  return (crc ^ 0xffffffff) >>> 0;
}

export function openZip(input) {
  if (
    !(input instanceof Uint8Array) ||
    input.byteLength < 22 ||
    input.byteLength > ZIP_LIMITS.archiveBytes
  )
    throw new Error('Use an unencrypted .xlsx file, no larger than 20 MiB.');
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  const u16 = (offset) => view.getUint16(offset, true);
  const u32 = (offset) => view.getUint32(offset, true);
  let end = -1;
  for (
    let offset = input.length - 22;
    offset >= Math.max(0, input.length - 65557);
    offset--
  ) {
    if (
      u32(offset) === 0x06054b50 &&
      offset + 22 + u16(offset + 20) === input.length
    ) {
      end = offset;
      break;
    }
  }
  if (end < 0)
    throw new Error(
      'The file is not a supported ZIP-based workbook, or it is incomplete.',
    );
  const count = u16(end + 10),
    directorySize = u32(end + 12),
    directoryStart = u32(end + 16);
  if (
    u16(end + 4) ||
    u16(end + 6) ||
    count !== u16(end + 8) ||
    count > ZIP_LIMITS.entries ||
    directoryStart + directorySize !== end
  )
    throw new Error(
      'Split, ZIP64 or oversized ZIP directories are not supported.',
    );
  const entries = new Map();
  let cursor = directoryStart;
  const ranges = [];
  for (let index = 0; index < count; index++) {
    if (cursor + 46 > end || u32(cursor) !== 0x02014b50)
      throw new Error('Invalid ZIP directory entry.');
    const flags = u16(cursor + 8),
      method = u16(cursor + 10),
      crc = u32(cursor + 16);
    const compressedSize = u32(cursor + 20),
      originalSize = u32(cursor + 24);
    const nameLength = u16(cursor + 28),
      extraLength = u16(cursor + 30),
      commentLength = u16(cursor + 32);
    const localOffset = u32(cursor + 42),
      next = cursor + 46 + nameLength + extraLength + commentLength;
    if (
      next > end ||
      !nameLength ||
      u16(cursor + 34) ||
      flags & 0x2041 ||
      ![0, 8].includes(method) ||
      localOffset + 30 > directoryStart
    )
      throw new Error(
        'Encrypted, split or unsupported ZIP entries cannot be inspected.',
      );
    const nameBytes = input.subarray(cursor + 46, cursor + 46 + nameLength);
    if (!(flags & 0x800) && nameBytes.some((byte) => byte > 127))
      throw new Error('Non-UTF-8 ZIP entry names are not supported.');
    const name = decoder.decode(nameBytes);
    if (
      name.startsWith('/') ||
      name.includes('\\') ||
      name.includes('\0') ||
      name.split('/').some((part) => part === '.' || part === '..') ||
      entries.has(name)
    )
      throw new Error('Unsafe or duplicate ZIP entry name.');
    if (
      u32(localOffset) !== 0x04034b50 ||
      u16(localOffset + 6) !== flags ||
      u16(localOffset + 8) !== method
    )
      throw new Error('ZIP local and directory headers disagree.');
    const localNameLength = u16(localOffset + 26),
      localExtraLength = u16(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength,
      dataEnd = dataStart + compressedSize;
    if (
      dataEnd > directoryStart ||
      localNameLength !== nameLength ||
      decoder.decode(
        input.subarray(localOffset + 30, localOffset + 30 + localNameLength),
      ) !== name
    )
      throw new Error('ZIP entry offsets or names are inconsistent.');
    if (method === 0 && compressedSize !== originalSize)
      throw new Error('Stored ZIP entry sizes disagree.');
    entries.set(name, { name, method, crc, dataStart, dataEnd, originalSize });
    ranges.push({ start: localOffset, end: dataEnd });
    cursor = next;
  }
  if (cursor !== end)
    throw new Error('ZIP directory size does not match its entries.');
  ranges.sort((a, b) => a.start - b.start);
  for (let index = 1; index < ranges.length; index++)
    if (ranges[index].start < ranges[index - 1].end)
      throw new Error('Overlapping ZIP entries are not supported.');
  let totalRead = 0;
  const cache = new Map();

  function read(name) {
    if (cache.has(name)) return cache.get(name);
    const entry = entries.get(name);
    if (!entry) throw new Error(`Workbook part is missing: ${name}`);
    if (
      entry.originalSize > ZIP_LIMITS.xmlBytes ||
      totalRead + entry.originalSize > ZIP_LIMITS.totalReadBytes
    )
      throw new Error('Workbook XML exceeds the inspection size limit.');
    totalRead += entry.originalSize;
    const source = input.subarray(entry.dataStart, entry.dataEnd);
    let output;
    if (entry.method === 0) output = source;
    else {
      output = new Uint8Array(entry.originalSize);
      let written = 0,
        finished = false;
      const inflater = new Inflate((chunk, final) => {
        if (written + chunk.length > output.length)
          throw new Error('Inflated ZIP data exceeds its declared size.');
        output.set(chunk, written);
        written += chunk.length;
        finished = final;
      });
      // Small compressed chunks bound each inflater callback before size validation.
      for (let offset = 0; offset < source.length; offset += 1024)
        inflater.push(
          source.subarray(offset, offset + 1024),
          offset + 1024 >= source.length,
        );
      if (!finished || written !== entry.originalSize)
        throw new Error('Inflated ZIP data is incomplete.');
    }
    if (crc32(output) !== entry.crc)
      throw new Error('ZIP checksum mismatch. The workbook may be damaged.');
    const text = decoder.decode(output);
    cache.set(name, text);
    return text;
  }
  return { names: [...entries.keys()], has: (name) => entries.has(name), read };
}
