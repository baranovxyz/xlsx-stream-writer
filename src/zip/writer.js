/**
 * A streaming ZIP writer, sufficient for the OPC packages Excel reads.
 *
 * Entries are emitted in the order given, and each entry's source is only
 * consumed when its turn comes. That ordering is a guarantee the workbook
 * depends on: `xl/sharedStrings.xml` can only be written once the sheet has
 * been walked and the string table is complete.
 *
 * Each entry takes one of two shapes:
 *
 * - Under `BUFFER_LIMIT` uncompressed bytes, the compressed output is held in
 *   memory and the local header carries exact sizes. This is the plainest,
 *   most widely readable form, and it covers the great majority of workbooks.
 * - Above it, the final size cannot be known before the header is written, so
 *   the entry switches to a streamed form: sizes deferred to a trailing data
 *   descriptor, with ZIP64 records so it stays correct past 4 GiB.
 */

const { crc32 } = require("./crc32");
// Bundlers swap this for ./deflate.browser.js through the "browser" field in
// package.json, so node:zlib never reaches a browser build.
const { deflateRaw } = require("./deflate.node");

const SIG_LOCAL = 0x04034b50;
const SIG_DESCRIPTOR = 0x08074b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;

const METHOD_DEFLATE = 8;
const FLAG_DATA_DESCRIPTOR = 0x08;
const FLAG_UTF8_NAMES = 0x800;
const VERSION_DEFAULT = 20;
const VERSION_ZIP64 = 45;

const U16_MAX = 0xffff;
const U32_MAX = 0xffffffff;

// The DOS timestamp for 1980-01-01 00:00 — the earliest the format can express.
// Fixed rather than "now", so the same rows always produce the same archive.
const DOS_TIME = 0;
const DOS_DATE = 0x0021;

const BUFFER_LIMIT = 8 * 1024 * 1024;

const encoder = new TextEncoder();

function record(length) {
  const bytes = new Uint8Array(length);
  return { bytes, view: new DataView(bytes.buffer) };
}

async function* toByteChunks(source) {
  if (typeof source === "string") {
    yield encoder.encode(source);
    return;
  }
  for await (const chunk of source) {
    if (typeof chunk === "string") yield encoder.encode(chunk);
    else if (chunk instanceof Uint8Array) yield chunk;
    else yield new Uint8Array(chunk);
  }
}

function buildLocalHeader(nameBytes, entry) {
  // A streamed entry declares its ZIP64 extra field up front: that is how a
  // reader knows the trailing descriptor holds 8-byte sizes rather than 4.
  const extraLength = entry.streaming ? 20 : 0;
  const { bytes, view } = record(30 + nameBytes.length + extraLength);

  view.setUint32(0, SIG_LOCAL, true);
  view.setUint16(4, entry.streaming ? VERSION_ZIP64 : VERSION_DEFAULT, true);
  view.setUint16(6, FLAG_UTF8_NAMES | (entry.streaming ? FLAG_DATA_DESCRIPTOR : 0), true);
  view.setUint16(8, METHOD_DEFLATE, true);
  view.setUint16(10, DOS_TIME, true);
  view.setUint16(12, DOS_DATE, true);
  view.setUint32(14, entry.streaming ? 0 : entry.crc, true);
  view.setUint32(18, entry.streaming ? 0 : entry.compressedSize, true);
  view.setUint32(22, entry.streaming ? 0 : entry.uncompressedSize, true);
  view.setUint16(26, nameBytes.length, true);
  view.setUint16(28, extraLength, true);
  bytes.set(nameBytes, 30);

  if (entry.streaming) {
    const extra = 30 + nameBytes.length;
    view.setUint16(extra, 0x0001, true);
    view.setUint16(extra + 2, 16, true);
    view.setBigUint64(extra + 4, 0n, true); // uncompressed, filled in the descriptor
    view.setBigUint64(extra + 12, 0n, true); // compressed, likewise
  }
  return bytes;
}

function buildDataDescriptor(entry) {
  const { bytes, view } = record(24);
  view.setUint32(0, SIG_DESCRIPTOR, true);
  view.setUint32(4, entry.crc, true);
  view.setBigUint64(8, BigInt(entry.compressedSize), true);
  view.setBigUint64(16, BigInt(entry.uncompressedSize), true);
  return bytes;
}

function buildCentralHeader(entry) {
  const saturated = [];
  if (entry.uncompressedSize >= U32_MAX) saturated.push(entry.uncompressedSize);
  if (entry.compressedSize >= U32_MAX) saturated.push(entry.compressedSize);
  if (entry.localOffset >= U32_MAX) saturated.push(entry.localOffset);
  const extraLength = saturated.length ? 4 + saturated.length * 8 : 0;
  const needsZip64 = entry.streaming || saturated.length > 0;

  const { bytes, view } = record(46 + entry.nameBytes.length + extraLength);
  view.setUint32(0, SIG_CENTRAL, true);
  view.setUint16(4, needsZip64 ? VERSION_ZIP64 : VERSION_DEFAULT, true);
  view.setUint16(6, needsZip64 ? VERSION_ZIP64 : VERSION_DEFAULT, true);
  view.setUint16(8, FLAG_UTF8_NAMES | (entry.streaming ? FLAG_DATA_DESCRIPTOR : 0), true);
  view.setUint16(10, METHOD_DEFLATE, true);
  view.setUint16(12, DOS_TIME, true);
  view.setUint16(14, DOS_DATE, true);
  view.setUint32(16, entry.crc, true);
  view.setUint32(20, Math.min(entry.compressedSize, U32_MAX), true);
  view.setUint32(24, Math.min(entry.uncompressedSize, U32_MAX), true);
  view.setUint16(28, entry.nameBytes.length, true);
  view.setUint16(30, extraLength, true);
  view.setUint16(32, 0, true); // comment length
  view.setUint16(34, 0, true); // disk number start
  view.setUint16(36, 0, true); // internal attributes
  view.setUint32(38, 0, true); // external attributes
  view.setUint32(42, Math.min(entry.localOffset, U32_MAX), true);
  bytes.set(entry.nameBytes, 46);

  if (extraLength) {
    const extra = 46 + entry.nameBytes.length;
    view.setUint16(extra, 0x0001, true);
    view.setUint16(extra + 2, saturated.length * 8, true);
    // Saturated values appear here in a fixed order: uncompressed, compressed,
    // then local header offset — each present only if its 32-bit slot overflowed.
    saturated.forEach((value, i) => view.setBigUint64(extra + 4 + i * 8, BigInt(value), true));
  }
  return bytes;
}

function buildEndRecords(entryCount, cdOffset, cdSize) {
  const needsZip64 =
    entryCount > U16_MAX || cdOffset >= U32_MAX || cdSize >= U32_MAX;
  const { bytes, view } = record(needsZip64 ? 56 + 20 + 22 : 22);
  let pos = 0;

  if (needsZip64) {
    view.setUint32(0, SIG_ZIP64_EOCD, true);
    view.setBigUint64(4, 44n, true); // size of the rest of this record
    view.setUint16(12, VERSION_ZIP64, true);
    view.setUint16(14, VERSION_ZIP64, true);
    view.setUint32(16, 0, true); // this disk
    view.setUint32(20, 0, true); // disk with the start of the central directory
    view.setBigUint64(24, BigInt(entryCount), true);
    view.setBigUint64(32, BigInt(entryCount), true);
    view.setBigUint64(40, BigInt(cdSize), true);
    view.setBigUint64(48, BigInt(cdOffset), true);

    view.setUint32(56, SIG_ZIP64_LOCATOR, true);
    view.setUint32(60, 0, true);
    view.setBigUint64(64, BigInt(cdOffset + cdSize), true);
    view.setUint32(72, 1, true); // total number of disks
    pos = 76;
  }

  view.setUint32(pos, SIG_EOCD, true);
  view.setUint16(pos + 4, 0, true);
  view.setUint16(pos + 6, 0, true);
  view.setUint16(pos + 8, Math.min(entryCount, U16_MAX), true);
  view.setUint16(pos + 10, Math.min(entryCount, U16_MAX), true);
  view.setUint32(pos + 12, Math.min(cdSize, U32_MAX), true);
  view.setUint32(pos + 16, Math.min(cdOffset, U32_MAX), true);
  view.setUint16(pos + 20, 0, true);
  return bytes;
}

/**
 * @param {Array<{name: string, source: string | AsyncIterable}>} entries
 * @param {{level?: number}} options
 * @returns {AsyncGenerator<Uint8Array>} the archive, chunk by chunk
 */
async function* writeZip(entries, options = {}) {
  const level = typeof options.level === "number" ? options.level : 4;
  // Overridable so tests can exercise the streamed branch without building a
  // multi-megabyte fixture for every case.
  const bufferLimit =
    typeof options.bufferLimit === "number" ? options.bufferLimit : BUFFER_LIMIT;
  // Normally supplied by the "browser" field swap above; overridable so tests
  // can drive the browser adapter through this same code path.
  const compress = options.deflateRaw || deflateRaw;
  const written = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name);
    let crc = 0;
    let uncompressedSize = 0;

    const counted = (async function* () {
      for await (const chunk of toByteChunks(entry.source)) {
        crc = crc32(chunk, crc);
        uncompressedSize += chunk.length;
        yield chunk;
      }
    })();

    const compressed = compress(counted, level);
    const buffered = [];
    let compressedSize = 0;
    let streaming = false;

    // Buffer until the entry either finishes or proves too big to hold.
    while (true) {
      const next = await compressed.next();
      if (next.done) break;
      buffered.push(next.value);
      compressedSize += next.value.length;
      if (uncompressedSize > bufferLimit) {
        streaming = true;
        break;
      }
    }

    const localOffset = offset;
    const header = buildLocalHeader(nameBytes, {
      streaming,
      crc,
      compressedSize,
      uncompressedSize,
    });
    yield header;
    offset += header.length;

    for (const chunk of buffered) {
      yield chunk;
      offset += chunk.length;
    }

    if (streaming) {
      while (true) {
        const next = await compressed.next();
        if (next.done) break;
        compressedSize += next.value.length;
        yield next.value;
        offset += next.value.length;
      }
      const descriptor = buildDataDescriptor({ crc, compressedSize, uncompressedSize });
      yield descriptor;
      offset += descriptor.length;
    }

    written.push({ nameBytes, crc, compressedSize, uncompressedSize, localOffset, streaming });
  }

  const cdOffset = offset;
  let cdSize = 0;
  for (const entry of written) {
    const central = buildCentralHeader(entry);
    yield central;
    cdSize += central.length;
  }
  yield buildEndRecords(written.length, cdOffset, cdSize);
}

module.exports = { writeZip, BUFFER_LIMIT };
