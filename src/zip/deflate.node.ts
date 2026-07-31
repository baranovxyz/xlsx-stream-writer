// Raw-deflate compression on Node, via the built-in zlib binding.
//
// Preferred over CompressionStream here because zlib exposes the compression
// level, and level 4 is the trade-off this package has always shipped: most of
// the ratio for a fraction of the CPU on very large sheets.

import { createDeflateRaw } from "node:zlib";
import { Readable } from "node:stream";

/**
 * @param source uncompressed bytes
 * @param level zlib compression level, 0-9
 */
export async function* deflateRaw(
  source: AsyncIterable<Uint8Array>,
  level: number,
): AsyncGenerator<Uint8Array> {
  const deflater = createDeflateRaw({ level });
  const input = Readable.from(source);

  // .pipe() does not forward errors, so a failing source would leave the
  // deflater open and the archive would never finish.
  input.on("error", error => deflater.destroy(error));
  input.pipe(deflater);

  yield* deflater;
}
