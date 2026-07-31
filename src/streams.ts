/**
 * Stream adapters. Everything the package consumes is normalised to an async
 * iterable, which is the one abstraction Node streams, web streams, generators
 * and plain arrays can all agree on — and which needs no polyfill in either
 * environment.
 */

/** Anything `addRows` or a zip entry source will accept. */
export type StreamLike<T> =
  | readonly T[]
  | AsyncIterable<T>
  | Iterable<T>
  | ReadableStream<T>
  | NodeReadableLike<T>;

/** The shape of a Node readable, structurally — no node:stream import needed. */
export interface NodeReadableLike<T> {
  pipe: (...args: never[]) => unknown;
  on: (event: string, listener: (arg: T) => void) => unknown;
  pause?: () => unknown;
  resume?: () => unknown;
}

function isAsyncIterable<T>(value: unknown): value is AsyncIterable<T> {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as AsyncIterable<T>)[Symbol.asyncIterator] === "function"
  );
}

export function isWebReadableStream<T>(value: unknown): value is ReadableStream<T> {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as ReadableStream<T>).getReader === "function"
  );
}

export function isNodeReadable<T>(value: unknown): value is NodeReadableLike<T> {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as NodeReadableLike<T>).pipe === "function" &&
    typeof (value as NodeReadableLike<T>).on === "function"
  );
}

function isSyncIterable<T>(value: unknown): value is Iterable<T> {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as Iterable<T>)[Symbol.iterator] === "function"
  );
}

/**
 * Iterate a web ReadableStream without relying on async iteration support,
 * which browsers only gained recently.
 */
export async function* iterateWebStream<T>(stream: ReadableStream<T>): AsyncGenerator<T> {
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
function iterateNodeReadable<T>(stream: NodeReadableLike<T>): AsyncGenerator<T> {
  const chunks: T[] = [];
  let waiting: (() => void) | null = null;
  let ended = false;
  let failure: unknown = null;

  const wake = () => {
    if (!waiting) return;
    const resolve = waiting;
    waiting = null;
    resolve();
  };
  stream.on("data", chunk => {
    chunks.push(chunk);
    stream.pause?.();
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
        yield chunks.shift() as T;
        stream.resume?.();
        continue;
      }
      if (failure) throw failure;
      if (ended) return;
      await new Promise<void>(resolve => {
        waiting = resolve;
        stream.resume?.();
      });
    }
  })();
}

export function toAsyncIterable<T>(source: StreamLike<T>, what = "value"): AsyncIterable<T> {
  if (Array.isArray(source)) {
    const items = source as readonly T[];
    return (async function* () {
      for (const item of items) yield item;
    })();
  }
  if (isAsyncIterable<T>(source)) return source;
  if (isWebReadableStream<T>(source)) return iterateWebStream(source);
  if (isNodeReadable<T>(source)) return iterateNodeReadable(source);
  if (typeof source !== "string" && isSyncIterable<T>(source)) {
    const items = source;
    return (async function* () {
      for (const item of items) yield item;
    })();
  }
  throw new TypeError(`${what} must be an array, a readable stream, or an iterable`);
}

/** Wrap an async iterable as a web ReadableStream, pulling one chunk at a time. */
export function toWebReadableStream<T>(iterable: AsyncIterable<T>): ReadableStream<T> {
  const iterator = iterable[Symbol.asyncIterator]();
  return new ReadableStream<T>({
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
      if (iterator.return) return iterator.return(reason).then(() => undefined);
      return undefined;
    },
  });
}
