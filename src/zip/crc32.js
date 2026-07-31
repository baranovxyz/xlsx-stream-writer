// CRC-32 (IEEE 802.3), the checksum every ZIP entry carries.

const TABLE = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let bit = 0; bit < 8; bit++) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  TABLE[i] = c >>> 0;
}

/**
 * @param {Uint8Array} bytes
 * @param {number} previous running value from an earlier chunk, 0 to start
 * @returns {number} unsigned 32-bit checksum
 */
function crc32(bytes, previous = 0) {
  let c = ~previous >>> 0;
  for (let i = 0; i < bytes.length; i++) {
    c = (TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8)) >>> 0;
  }
  return ~c >>> 0;
}

module.exports = { crc32 };
