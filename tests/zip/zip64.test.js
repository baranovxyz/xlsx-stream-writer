const test = require("node:test");
const assert = require("node:assert/strict");

const { __testing } = require("../../dist/zip/writer");

const { buildCentralHeader, buildEndRecords, buildLocalHeader } = __testing;

const U32_MAX = 0xffffffff;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const SIG_EOCD = 0x06054b50;

const name = bytes => new TextEncoder().encode(bytes);

function view(bytes) {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

/** Read the ZIP64 extended information extra field out of a central header. */
function zip64Extra(header, nameLength) {
  const v = view(header);
  const extraLength = v.getUint16(30, true);
  if (!extraLength) return null;
  const start = 46 + nameLength;
  assert.equal(v.getUint16(start, true), 0x0001, "extra field should be the ZIP64 record");
  const size = v.getUint16(start + 2, true);
  const values = [];
  for (let i = 0; i + 8 <= size; i += 8) {
    values.push(Number(v.getBigUint64(start + 4 + i, true)));
  }
  return values;
}

const entry = overrides => ({
  nameBytes: name("big.bin"),
  crc: 0x12345678,
  compressedSize: 100,
  uncompressedSize: 200,
  localOffset: 300,
  streaming: false,
  ...overrides,
});

test("a central header under every limit carries no ZIP64 extra field", () => {
  const header = buildCentralHeader(entry());
  assert.equal(zip64Extra(header, 7), null);
  assert.equal(view(header).getUint32(20, true), 100);
  assert.equal(view(header).getUint32(24, true), 200);
  assert.equal(view(header).getUint32(42, true), 300);
});

test("an oversized uncompressed size moves into the ZIP64 extra field", () => {
  const size = 5 * 1024 * 1024 * 1024; // 5 GiB
  const header = buildCentralHeader(entry({ uncompressedSize: size }));

  assert.equal(view(header).getUint32(24, true), U32_MAX, "base field should saturate");
  assert.deepEqual(zip64Extra(header, 7), [size]);
});

test("saturated values appear in the order the format fixes", () => {
  const uncompressed = 6 * 1024 * 1024 * 1024;
  const compressed = 5 * 1024 * 1024 * 1024;
  const offset = 4.5 * 1024 * 1024 * 1024;
  const header = buildCentralHeader(
    entry({ uncompressedSize: uncompressed, compressedSize: compressed, localOffset: offset }),
  );

  // Uncompressed, then compressed, then the local header offset.
  assert.deepEqual(zip64Extra(header, 7), [uncompressed, compressed, offset]);
  assert.equal(view(header).getUint32(20, true), U32_MAX);
  assert.equal(view(header).getUint32(24, true), U32_MAX);
  assert.equal(view(header).getUint32(42, true), U32_MAX);
});

test("only the saturated fields are moved, and only those", () => {
  const offset = 5 * 1024 * 1024 * 1024;
  const header = buildCentralHeader(entry({ localOffset: offset }));

  // Sizes stay in their base fields; just the offset moves.
  assert.deepEqual(zip64Extra(header, 7), [offset]);
  assert.equal(view(header).getUint32(20, true), 100);
  assert.equal(view(header).getUint32(24, true), 200);
});

test("a streamed entry advertises ZIP64 in its local header", () => {
  const header = buildLocalHeader(name("big.bin"), {
    streaming: true,
    crc: 0,
    compressedSize: 0,
    uncompressedSize: 0,
  });
  const v = view(header);

  assert.equal(v.getUint16(4, true), 45, "version needed should be 4.5");
  assert.ok(v.getUint16(6, true) & 0x08, "data descriptor flag should be set");
  assert.equal(v.getUint16(28, true), 20, "a 20-byte ZIP64 extra field should follow the name");
  // The extra field must be present with zeroed sizes; that is what tells a
  // reader the trailing descriptor holds 8-byte values rather than 4.
  const extra = 30 + name("big.bin").length;
  assert.equal(v.getUint16(extra, true), 0x0001);
  assert.equal(v.getUint16(extra + 2, true), 16);
});

test("a small archive ends with a plain EOCD", () => {
  const end = buildEndRecords(7, 1000, 500);
  assert.equal(end.length, 22);
  const v = view(end);
  assert.equal(v.getUint32(0, true), SIG_EOCD);
  assert.equal(v.getUint16(10, true), 7);
  assert.equal(v.getUint32(12, true), 500);
  assert.equal(v.getUint32(16, true), 1000);
});

test("more than 65535 entries adds the ZIP64 end records", () => {
  const entryCount = 70000;
  const cdOffset = 1000;
  const cdSize = 500;
  const end = buildEndRecords(entryCount, cdOffset, cdSize);
  const v = view(end);

  assert.equal(end.length, 56 + 20 + 22);
  assert.equal(v.getUint32(0, true), SIG_ZIP64_EOCD);
  assert.equal(Number(v.getBigUint64(4, true)), 44, "record size excludes its first 12 bytes");
  assert.equal(Number(v.getBigUint64(32, true)), entryCount);
  assert.equal(Number(v.getBigUint64(40, true)), cdSize);
  assert.equal(Number(v.getBigUint64(48, true)), cdOffset);

  assert.equal(v.getUint32(56, true), SIG_ZIP64_LOCATOR);
  assert.equal(
    Number(v.getBigUint64(64, true)),
    cdOffset + cdSize,
    "the locator must point at the ZIP64 EOCD, which sits right after the directory",
  );
  assert.equal(v.getUint32(72, true), 1, "total disks");

  // The plain EOCD still follows, with its fields saturated.
  assert.equal(v.getUint32(76, true), SIG_EOCD);
  assert.equal(v.getUint16(86, true), 0xffff, "entry count saturates");
});

test("a central directory past 4 GiB adds the ZIP64 end records", () => {
  const cdOffset = 5 * 1024 * 1024 * 1024;
  const end = buildEndRecords(7, cdOffset, 500);
  const v = view(end);

  assert.equal(v.getUint32(0, true), SIG_ZIP64_EOCD);
  assert.equal(Number(v.getBigUint64(48, true)), cdOffset);
  assert.equal(v.getUint32(76, true), SIG_EOCD);
  assert.equal(v.getUint32(76 + 16, true), U32_MAX, "the EOCD offset saturates");
  assert.equal(v.getUint16(76 + 10, true), 7, "the entry count still fits");
});
