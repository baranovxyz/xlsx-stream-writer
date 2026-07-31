/**
 * Stream adapters. Everything the package consumes is normalised to an async
 * iterable, which is the one abstraction Node streams, web streams, generators
 * and plain arrays can all agree on — and which needs no polyfill in either
 * environment.
 */

function isAsyncIterable(value) {
  return Boolean(value) && typeof value[Symbol.asyncIterator] === "function";
}

function isWebReadableStream(value) {
  return Boolean(value) && typeof value.getReader === "function";
}

function isNodeReadable(value) {
  return (
    Boolean(value) && typeof value.pipe === "function" && typeof value.on === "function"
  );
}

/**
 * Iterate a web ReadableStream without relying on async iteration support,
 * which browsers only gained recently.
 */
async function* iterateWebStream(stream) {
  const reader = stream.getReader();
  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) return;
      yield value;
    }
  } finally {
    reader.releaseLock();
  }
}

/** Bridge an older Node readable that predates async iteration support. */
function iterateNodeReadable(stream) {
  const chunks = [];
  let waiting = null;
  let ended = false;
  let failure = null;

  const wake = () => {
    if (!waiting) return;
    const resolve = waiting;
    waiting = null;
    resolve();
  };
  stream.on("data", chunk => {
    chunks.push(chunk);
    stream.pause();
    wake();
  });
  stream.on("end", () => {
    ended = true;
    wake();
  });
  stream.on("error", error => {
    failure = error;
    wake();
  });

  return (async function* () {
    while (true) {
      if (chunks.length) {
        yield chunks.shift();
        stream.resume();
        continue;
      }
      if (failure) throw failure;
      if (ended) return;
      await new Promise(resolve => {
        waiting = resolve;
        stream.resume();
      });
    }
  })();
}

/**
 * @param {Array|AsyncIterable|Iterable|ReadableStream|import("node:stream").Readable} source
 * @returns {AsyncIterable}
 */
function toAsyncIterable(source, what = "value") {
  if (Array.isArray(source)) {
    return (async function* () {
      for (const item of source) yield item;
    })();
  }
  if (isAsyncIterable(source)) return source;
  if (isWebReadableStream(source)) return iterateWebStream(source);
  if (isNodeReadable(source)) return iterateNodeReadable(source);
  if (source && typeof source !== "string" && typeof source[Symbol.iterator] === "function") {
    return (async function* () {
      for (const item of source) yield item;
    })();
  }
  throw new TypeError(
    `${what} must be an array, a readable stream, or an iterable`,
  );
}

/** Wrap an async iterable as a web ReadableStream, pulling one chunk at a time. */
function toWebReadableStream(iterable) {
  const iterator = iterable[Symbol.asyncIterator]();
  return new ReadableStream({
    async pull(controller) {
      try {
        const { value, done } = await iterator.next();
        if (done) controller.close();
        else controller.enqueue(value);
      } catch (error) {
        controller.error(error);
      }
    },
    cancel(reason) {
      if (iterator.return) return iterator.return(reason);
    },
  });
}

async function collect(iterable) {
  const chunks = [];
  for await (const chunk of iterable) chunks.push(chunk);
  return chunks;
}

module.exports = {
  toAsyncIterable,
  toWebReadableStream,
  iterateWebStream,
  collect,
  isNodeReadable,
  isWebReadableStream,
};
