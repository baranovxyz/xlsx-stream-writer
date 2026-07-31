const Readable = require("stream-browserify").Readable;
const Writable = require("stream-browserify").Writable;

function getCellAddress(rowIndex, colIndex) {
  let colAddress = "";
  let input = (colIndex - 1).toString(26);
  while (input.length) {
    const a = input.charCodeAt(input.length - 1);
    colAddress =
      String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colAddress;
    input =
      input.length > 1
        ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26)
        : "";
  }
  return colAddress + rowIndex;
}

function getXmlFromXmlStream(xmlStream) {
  return new Promise((resolve, reject) => {
    const ws = Writable();
    let xml = "";
    ws._write = function(chunk, enc, next) {
      xml += chunk.toString();
      next();
    };
    xmlStream.pipe(ws);
    xmlStream.on("error", reject);
    ws.on("finish", () => resolve(xml));
    ws.on("error", reject);
  });
}

function wrapRowsInStream(rows) {
  const rs = Readable({ objectMode: true });
  let c = 0;
  rs._read = function() {
    if (c === rows.length) rs.push(null);
    else rs.push(rows[c]);
    c++;
  };
  return rs;
}

function isNodeReadable(value) {
  return (
    Boolean(value) &&
    typeof value.pipe === "function" &&
    typeof value.on === "function"
  );
}

function isWebReadableStream(value) {
  return Boolean(value) && typeof value.getReader === "function";
}

function getIterator(value) {
  if (!value || typeof value === "string") return null;
  if (typeof value[Symbol.asyncIterator] === "function") return value[Symbol.asyncIterator]();
  if (typeof value[Symbol.iterator] === "function") return value[Symbol.iterator]();
  return null;
}

/**
 * Adapt any iterator — sync or async — into an object-mode readable, pulling one
 * row at a time so a slow consumer cannot be outrun by a fast source.
 */
function streamFromIterator(iterator) {
  const rs = Readable({ objectMode: true });
  let pending = false;
  rs._read = function() {
    if (pending) return;
    pending = true;
    Promise.resolve(iterator.next()).then(
      result => {
        pending = false;
        rs.push(result.done ? null : result.value);
      },
      error => {
        pending = false;
        rs.emit("error", error);
      },
    );
  };
  return rs;
}

/**
 * Normalise whatever the caller passed into an object-mode readable of rows.
 *
 * Node readables are passed straight through, which keeps their own
 * backpressure and error semantics intact. Everything else is adapted.
 */
function toRowsStream(rowsOrStream) {
  if (Array.isArray(rowsOrStream)) return wrapRowsInStream(rowsOrStream);
  if (isNodeReadable(rowsOrStream)) return rowsOrStream;
  if (isWebReadableStream(rowsOrStream)) {
    const reader = rowsOrStream.getReader();
    return streamFromIterator({ next: () => reader.read() });
  }
  const iterator = getIterator(rowsOrStream);
  if (iterator) return streamFromIterator(iterator);

  throw new TypeError(
    "Rows must be an array of arrays, a readable stream of arrays, or an iterable of arrays",
  );
}

// XML 1.0 has no representation for these code points — not even a numeric
// character reference — so they have to be dropped rather than escaped. Left in
// place they produce a file Excel refuses to open.
const ILLEGAL_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

// Matches a well-formed pair first, so only unpaired halves reach the replacer.
const SURROGATES = /[\uD800-\uDBFF][\uDC00-\uDFFF]|[\uD800-\uDFFF]/g;

// One cheap test decides whether any of the work below is needed. Most cells in
// a large export are plain text, and this is on the hot path for every one.
const NEEDS_WORK = /[&<>"'\u0000-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/;

function sanitize(str) {
  return str
    .replace(ILLEGAL_XML_CHARS, "")
    .replace(SURROGATES, match => (match.length === 2 ? match : "\uFFFD"));
}

function escapeXml(str = "") {
  if (!NEEDS_WORK.test(str)) return str;
  return sanitize(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeXmlExtended(str = "") {
  if (!NEEDS_WORK.test(str)) return str;
  return sanitize(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

module.exports = {
  getCellAddress,
  wrapRowsInStream,
  getXmlFromXmlStream,
  toRowsStream,
  escapeXml,
  escapeXmlExtended,
};
