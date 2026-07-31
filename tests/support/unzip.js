/**
 * Minimal ZIP reader for tests, built only on node:zlib.
 *
 * It reads the central directory rather than walking local headers, so it works
 * on archives written in streaming mode — where local headers carry zeroes and
 * the real sizes live in a trailing data descriptor. That is exactly the shape
 * this package produces, and it is the shape Excel has to cope with.
 *
 * This doubles as the acceptance harness for the built-in ZIP writer: anything
 * this reader rejects is something a real unzip implementation may reject too.
 */

const { inflateRawSync, crc32 } = require("node:zlib");

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;

const U32_MAX = 0xffffffff;
const U16_MAX = 0xffff;

function findEocdOffset(buf) {
  // The EOCD is last, but a trailing comment of up to 65535 bytes may follow it.
  const earliest = Math.max(0, buf.length - (22 + U16_MAX));
  for (let i = buf.length - 22; i >= earliest; i--) {
    if (buf.readUInt32LE(i) === SIG_EOCD) return i;
  }
  throw new Error("not a zip archive: end of central directory not found");
}

function readCentralDirectoryLocation(buf) {
  const eocd = findEocdOffset(buf);
  let entryCount = buf.readUInt16LE(eocd + 10);
  let cdOffset = buf.readUInt32LE(eocd + 16);
  let cdSize = buf.readUInt32LE(eocd + 12);

  const needsZip64 =
    entryCount === U16_MAX || cdOffset === U32_MAX || cdSize === U32_MAX;
  if (!needsZip64) return { cdOffset, cdSize, entryCount, zip64: false };

  const locator = eocd - 20;
  if (locator < 0 || buf.readUInt32LE(locator) !== SIG_ZIP64_LOCATOR) {
    throw new Error("zip64 values present but the zip64 EOCD locator is missing");
  }
  const zip64Eocd = Number(buf.readBigUInt64LE(locator + 8));
  if (buf.readUInt32LE(zip64Eocd) !== SIG_ZIP64_EOCD) {
    throw new Error("zip64 EOCD locator does not point at a zip64 EOCD record");
  }
  entryCount = Number(buf.readBigUInt64LE(zip64Eocd + 32));
  cdSize = Number(buf.readBigUInt64LE(zip64Eocd + 40));
  cdOffset = Number(buf.readBigUInt64LE(zip64Eocd + 48));
  return { cdOffset, cdSize, entryCount, zip64: true };
}

/**
 * Pull the 64-bit replacements out of the zip64 extended information extra
 * field. Present values appear in a fixed order, but only for the fields whose
 * 32-bit slot was saturated to 0xffffffff.
 */
function readZip64Extra(extra, saturated) {
  const out = {};
  let pos = 0;
  while (pos + 4 <= extra.length) {
    const headerId = extra.readUInt16LE(pos);
    const size = extra.readUInt16LE(pos + 2);
    const body = extra.subarray(pos + 4, pos + 4 + size);
    if (headerId === 0x0001) {
      let cursor = 0;
      for (const field of ["uncompressedSize", "compressedSize", "localHeaderOffset"]) {
        if (!saturated[field]) continue;
        if (cursor + 8 > body.length) break;
        out[field] = Number(body.readBigUInt64LE(cursor));
        cursor += 8;
      }
      break;
    }
    pos += 4 + size;
  }
  return out;
}

function readEntryData(buf, entry) {
  if (buf.readUInt32LE(entry.localHeaderOffset) !== SIG_LOCAL) {
    throw new Error(`${entry.name}: central directory points at a non-local-header offset`);
  }
  const nameLength = buf.readUInt16LE(entry.localHeaderOffset + 26);
  const extraLength = buf.readUInt16LE(entry.localHeaderOffset + 28);
  const start = entry.localHeaderOffset + 30 + nameLength + extraLength;
  const compressed = buf.subarray(start, start + entry.compressedSize);

  let data;
  if (entry.method === 0) data = Buffer.from(compressed);
  else if (entry.method === 8) data = inflateRawSync(compressed);
  else throw new Error(`${entry.name}: unsupported compression method ${entry.method}`);

  if (data.length !== entry.uncompressedSize) {
    throw new Error(
      `${entry.name}: uncompressed size mismatch — header says ${entry.uncompressedSize}, got ${data.length}`,
    );
  }
  const actualCrc = crc32(data) >>> 0;
  if (actualCrc !== entry.crc) {
    throw new Error(
      `${entry.name}: CRC mismatch — header says ${entry.crc.toString(16)}, got ${actualCrc.toString(16)}`,
    );
  }
  return data;
}

/**
 * @param {Buffer} buf a complete zip archive
 * @returns {{ entries: Array, names: string[], get(name): Buffer, text(name): string }}
 */
function readZip(buf) {
  const { cdOffset, entryCount } = readCentralDirectoryLocation(buf);
  const entries = [];
  let pos = cdOffset;

  for (let i = 0; i < entryCount; i++) {
    if (buf.readUInt32LE(pos) !== SIG_CENTRAL) {
      throw new Error(`central directory entry ${i} has a bad signature`);
    }
    const flags = buf.readUInt16LE(pos + 8);
    const method = buf.readUInt16LE(pos + 10);
    const crc = buf.readUInt32LE(pos + 16) >>> 0;
    const nameLength = buf.readUInt16LE(pos + 28);
    const extraLength = buf.readUInt16LE(pos + 30);
    const commentLength = buf.readUInt16LE(pos + 32);

    const saturated = {
      compressedSize: buf.readUInt32LE(pos + 20) === U32_MAX,
      uncompressedSize: buf.readUInt32LE(pos + 24) === U32_MAX,
      localHeaderOffset: buf.readUInt32LE(pos + 42) === U32_MAX,
    };
    const extra = buf.subarray(pos + 46 + nameLength, pos + 46 + nameLength + extraLength);
    const zip64 = readZip64Extra(extra, saturated);

    const entry = {
      name: buf.toString("utf8", pos + 46, pos + 46 + nameLength),
      flags,
      method,
      crc,
      compressedSize: zip64.compressedSize ?? buf.readUInt32LE(pos + 20),
      uncompressedSize: zip64.uncompressedSize ?? buf.readUInt32LE(pos + 24),
      localHeaderOffset: zip64.localHeaderOffset ?? buf.readUInt32LE(pos + 42),
      usesDataDescriptor: (flags & 0x08) !== 0,
      utf8NameFlag: (flags & 0x800) !== 0,
      zip64: Object.keys(zip64).length > 0,
    };
    entry.data = entry.name.endsWith("/") ? Buffer.alloc(0) : readEntryData(buf, entry);
    entries.push(entry);

    pos += 46 + nameLength + extraLength + commentLength;
  }

  const byName = new Map(entries.map(entry => [entry.name, entry]));
  return {
    entries,
    names: entries.map(entry => entry.name),
    get(name) {
      const entry = byName.get(name);
      if (!entry) {
        throw new Error(`archive has no entry "${name}"; it has: ${entries.map(e => e.name).join(", ")}`);
      }
      return entry.data;
    },
    text(name) {
      return this.get(name).toString("utf8");
    },
    entry(name) {
      return byName.get(name);
    },
  };
}

module.exports = { readZip };
