// Raw-deflate compression in the browser, via the platform CompressionStream.
//
// Available in Chrome 80+, Firefox 113+ and Safari 16.4+. It exposes no level
// control, so `level` is accepted and ignored; the browser's default is close
// to zlib level 6.

/**
 * @param {AsyncIterable<Uint8Array>} source
 * @param {number} level accepted for parity with the Node adapter; unused
 * @returns {AsyncGenerator<Uint8Array>}
 */
async function* deflateRaw(source, level) {
  if (typeof CompressionStream === "undefined") {
    throw new Error(
      "This environment has no CompressionStream, which xlsx-stream-writer needs to compress the workbook",
    );
  }

  const compressor = new CompressionStream("deflate-raw");
  const writer = compressor.writable.getWriter();

  // Feed the compressor concurrently with draining it; writing everything first
  // would deadlock as soon as the internal queue fills. The pump settles rather
  // than rejects, so a source failure cannot escape as an unhandled rejection
  // when the read loop tears down first — it is re-thrown below instead.
  let pumpError = null;
  const pump = (async () => {
    try {
      for await (const chunk of source) await writer.write(chunk);
      await writer.close();
    } catch (error) {
      pumpError = error;
      try {
        await writer.abort(error);
      } catch {
        // The stream is already coming down; the original error is what matters.
      }
    }
  })();

  const reader = compressor.readable.getReader();
  let drained = false;
  try {
    while (true) {
      let result;
      try {
        result = await reader.read();
      } catch (readError) {
        await pump;
        throw pumpError || readError;
      }
      if (result.done) {
        drained = true;
        break;
      }
      yield result.value;
    }
  } finally {
    // An abandoned consumer must not leave the pump writing forever.
    if (!drained) {
      try {
        await reader.cancel();
      } catch {
        // Nothing left to do if the stream is already closed.
      }
    }
    reader.releaseLock();
  }

  await pump;
  if (pumpError) throw pumpError;
}

module.exports = { deflateRaw };
