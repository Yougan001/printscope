import test from 'node:test';
import assert from 'node:assert/strict';
import { zipSync, strToU8 } from 'fflate';
import { openZip, ZIP_LIMITS } from '../core/zip.mjs';

test('stored and deflated XML parts read exactly, including Unicode', () => {
  for (const level of [0, 6]) {
    const text = '<sheet>你好 &amp; ' + 'data'.repeat(100000) + '</sheet>';
    const archive = openZip(
      zipSync({ 'xl/worksheets/sheet1.xml': strToU8(text) }, { level }),
    );
    assert.deepEqual(archive.names, ['xl/worksheets/sheet1.xml']);
    assert.equal(archive.read(archive.names[0]), text);
    assert.equal(archive.read(archive.names[0]), text);
  }
});

test('arbitrary binary, truncated ZIP, missing parts and traversal names fail clearly', () => {
  assert.throws(() => openZip(new Uint8Array(100)), /not a supported/);
  const bytes = zipSync({ 'a.xml': strToU8('a') });
  assert.throws(() => openZip(bytes.subarray(0, -1)), /incomplete/);
  assert.throws(() => openZip(bytes).read('missing.xml'), /missing/);
  assert.throws(() => openZip(zipSync({ '../a.xml': strToU8('a') })), /Unsafe/);
});

test('two ZIP entries cannot claim the same workbook part name', () => {
  const bytes = zipSync(
    { 'a.xml': strToU8('a'), 'b.xml': strToU8('b') },
    { level: 0 },
  );
  const view = new DataView(bytes.buffer);
  for (let offset = 0; offset < bytes.length - 46; offset++) {
    const signature = view.getUint32(offset, true);
    const nameOffset =
      signature === 0x04034b50
        ? offset + 30
        : signature === 0x02014b50
          ? offset + 46
          : -1;
    if (nameOffset > 0 && bytes[nameOffset] === 98) bytes[nameOffset] = 97;
  }
  assert.throws(() => openZip(bytes), /duplicate/);
});

test('checksum corruption is detected even in otherwise valid XML', () => {
  const bytes = zipSync({ 'a.xml': strToU8('<a>text</a>') }, { level: 0 });
  const view = new DataView(bytes.buffer);
  const start = 30 + view.getUint16(26, true) + view.getUint16(28, true);
  bytes[start + 3] ^= 1;
  assert.throws(() => openZip(bytes).read('a.xml'), /checksum/);
});

test('declared XML sizes are bounded before allocating decompression output', () => {
  const bytes = zipSync({ 'a.xml': strToU8('<a/>') });
  const view = new DataView(bytes.buffer);
  const directory = view.getUint32(bytes.length - 6, true);
  view.setUint32(directory + 24, ZIP_LIMITS.xmlBytes + 1, true);
  assert.throws(() => openZip(bytes).read('a.xml'), /size limit/);
});

test('actual expansion cannot overrun a falsely small declared size', () => {
  const bytes = zipSync({ 'a.xml': strToU8('x'.repeat(1_000_000)) });
  const view = new DataView(bytes.buffer);
  const directory = view.getUint32(bytes.length - 6, true);
  view.setUint32(directory + 24, 4, true);
  assert.throws(() => openZip(bytes).read('a.xml'), /declared size/);
});

test('encrypted entries and invalid header offsets are refused', () => {
  const bytes = zipSync({ 'a.xml': strToU8('<a/>') });
  const view = new DataView(bytes.buffer);
  const directory = view.getUint32(bytes.length - 6, true);
  const encrypted = bytes.slice();
  new DataView(encrypted.buffer).setUint16(directory + 8, 1, true);
  assert.throws(() => openZip(encrypted), /Encrypted/);
  view.setUint32(directory + 42, bytes.length, true);
  assert.throws(() => openZip(bytes), /unsupported/);
});
