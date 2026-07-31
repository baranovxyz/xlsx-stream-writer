// Raw-deflate compression on Node, via the built-in zlib binding.
//
// Preferred over CompressionStream here because zlib exposes the compression
// level, and level 4 is the trade-off this package has always shipped: most of
// the ratio for a fraction of the CPU on very large sheets.

const { createDeflateRaw } = require("node:zlib");
const { Readable } = require("node:stream");

/**
 * @param {AsyncIterable<Uint8Array>} source
 * @param {number} level zlib compression level, 0-9
 * @returns {AsyncGenerator<Uint8Array>}
 */
async function* deflateRaw(source, level) {
  const deflater = createDeflateRaw({ level });
  const input = Readable.from(source);

  // .pipe() does not forward errors, so a failing source would leave the
  // deflater open and the archive would never finish.
  input.on("error", error => deflater.destroy(error));
  input.pipe(deflater);

  yield* deflater;
}

module.exports = { deflateRaw };
