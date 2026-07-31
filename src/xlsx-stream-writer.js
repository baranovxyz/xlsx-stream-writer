const xmlParts = require("./xml/parts");
const xmlBlobs = require("./xml/blobs");
const { getCellAddress, escapeXml } = require("./helpers");
const { toAsyncIterable, toWebReadableStream } = require("./streams");
const { writeZip } = require("./zip/writer");
const getStyles = require("./styles").getStyles;

const defaultOptions = {
  inlineStrings: false,
  styles: [],
  styleIdFunc: (value, columnId, rowId) => 0,
  compressionLevel: 4,
};

// The grid Excel itself supports. Exceeding either limit produces a file Excel
// refuses to open, so it is better to say so at the offending row.
const MAX_ROWS = 1048576;
const MAX_COLUMNS = 16384;

// Excel counts days from 1899-12-30 — two days behind 1900-01-01, because it
// preserves Lotus 1-2-3's belief that 1900 was a leap year.
const EXCEL_EPOCH_OFFSET_DAYS = 25569;
const MS_PER_DAY = 86400000;

const SHEET_PART = "xl/worksheets/sheet1.xml";
const SHARED_STRINGS_PART = "xl/sharedStrings.xml";

class XlsxStreamWriter {
  constructor(options) {
    // Spread rather than assign into the defaults: Object.assign(defaultOptions,
    // options) mutates the shared module-level object, so the next writer built
    // in the same process would inherit this one's options.
    this.options = { ...defaultOptions, ...options };

    if (this.options.styles != null && !Array.isArray(this.options.styles)) {
      throw new TypeError("options.styles must be an array of style objects");
    }
    if (typeof this.options.styleIdFunc !== "function") {
      throw new TypeError("options.styleIdFunc must be a function");
    }

    this.sharedStringsArr = [];
    this.sharedStringsMap = {};

    this._rows = null;
    this._rowsAdded = false;
    this._consumed = false;
    this._sharedStringRefs = 0;

    this.xlsx = {
      "[Content_Types].xml": cleanUpXml(xmlBlobs.contentTypes),
      "_rels/.rels": cleanUpXml(xmlBlobs.rels),
      "xl/workbook.xml": cleanUpXml(xmlBlobs.workbook),
      "xl/styles.xml": cleanUpXml(getStyles(this.options.styles)),
      "xl/_rels/workbook.xml.rels": cleanUpXml(xmlBlobs.workbookRels),
    };
  }

  /**
   * Add rows to the workbook.
   * @param {Array | Readable | ReadableStream | Iterable | AsyncIterable} rowsOrStream
   *   array of arrays, a readable stream of arrays, or any iterable of arrays
   * @return {undefined}
   */
  addRows(rowsOrStream) {
    if (this._rowsAdded) {
      throw new Error(
        "addRows() can only be called once per writer — the previous row stream would be discarded. Create a new XlsxStreamWriter for another workbook.",
      );
    }
    this._rows = toAsyncIterable(rowsOrStream, "Rows");
    this._rowsAdded = true;
  }

  /** The worksheet part, as a web ReadableStream of XML text. */
  get sheetXmlStream() {
    if (!this._rowsAdded) return null;
    if (!this._sheetStream) this._sheetStream = toWebReadableStream(this._sheetXml());
    return this._sheetStream;
  }

  /**
   * The shared-string table, as a web ReadableStream of XML text.
   *
   * Only read this once `sheetXmlStream` has been drained: the table is built
   * as the sheet is walked, and this stream begins emitting — counts included —
   * as soon as you touch it.
   */
  get sharedStringsXmlStream() {
    if (!this._sharedStringsStream) {
      this._sharedStringsStream = toWebReadableStream(this._sharedStringsXml());
    }
    return this._sharedStringsStream;
  }

  async *_sheetXml() {
    yield xmlParts.sheetHeader;
    let rowIndex = 0;
    for await (const row of this._rows) {
      yield this._getRowXml(row, rowIndex);
      rowIndex++;
    }
    yield xmlParts.sheetFooter;
  }

  async *_sharedStringsXml() {
    yield xmlParts.getSharedStringsHeader(
      this.sharedStringsArr.length,
      this._sharedStringRefs,
    );
    for (const value of this.sharedStringsArr) {
      yield xmlParts.getSharedStringXml(escapeXml(String(value)));
    }
    yield xmlParts.sharedStringsFooter;
  }

  _getRowXml(row, rowIndex) {
    if (rowIndex >= MAX_ROWS) {
      throw new RangeError(
        `Row ${rowIndex + 1} exceeds the Excel worksheet limit of ${MAX_ROWS} rows`,
      );
    }
    if (!Array.isArray(row)) {
      throw new TypeError(`Row ${rowIndex + 1} is not an array of cell values`);
    }
    if (row.length > MAX_COLUMNS) {
      throw new RangeError(
        `Row ${rowIndex + 1} has ${row.length} cells, over the Excel worksheet limit of ${MAX_COLUMNS} columns`,
      );
    }

    let rowXml = xmlParts.getRowStart(rowIndex);
    row.forEach((cellValue, colIndex) => {
      const cellAddress = getCellAddress(rowIndex + 1, colIndex + 1);
      const styleId = this.options.styleIdFunc(cellValue, colIndex, rowIndex);
      rowXml += this._getCellXml(cellValue, cellAddress, styleId);
    });
    rowXml += xmlParts.rowEnd;
    return rowXml;
  }

  _getCellXml(value, address, styleId = 0) {
    if (value === null || typeof value === "undefined") {
      return xmlParts.getBlankCellXml(address, styleId);
    }
    if (typeof value === "number") {
      // NaN and ±Infinity have no SpreadsheetML representation; emitting them
      // literally makes Excel treat the sheet as damaged.
      return Number.isFinite(value)
        ? xmlParts.getNumberCellXml(value, address, styleId)
        : xmlParts.getBlankCellXml(address, styleId);
    }
    if (typeof value === "bigint") {
      return xmlParts.getNumberCellXml(value.toString(), address, styleId);
    }
    if (typeof value === "boolean") {
      return xmlParts.getBooleanCellXml(value, address, styleId);
    }
    if (value instanceof Date) {
      const serial = value.getTime() / MS_PER_DAY + EXCEL_EPOCH_OFFSET_DAYS;
      return Number.isFinite(serial)
        ? xmlParts.getNumberCellXml(serial, address, styleId)
        : xmlParts.getBlankCellXml(address, styleId);
    }
    return this._getStringCellXml(value, address, styleId);
  }

  _getStringCellXml(value, address, styleId) {
    const stringValue = String(value);

    // Anything whose toString is still Object.prototype's would land in the
    // sheet as "[object Object]". Values that define their own toString —
    // decimal and date libraries, for instance — keep working.
    if (/^\[object [A-Za-z]+\]$/.test(stringValue)) {
      throw new TypeError(
        `Cell ${address} received a ${stringValue.slice(8, -1)} with no meaningful string form; convert it before adding the row`,
      );
    }

    if (this.options.inlineStrings) {
      return xmlParts.getInlineStringCellXml(escapeXml(stringValue), address, styleId);
    }
    this._sharedStringRefs++;
    return xmlParts.getStringCellXml(this._lookupString(stringValue), address, styleId);
  }

  _lookupString(value) {
    let sharedStringIndex = this.sharedStringsMap[value];
    if (typeof sharedStringIndex !== "undefined") return sharedStringIndex;
    sharedStringIndex = this.sharedStringsArr.length;
    this.sharedStringsMap[value] = sharedStringIndex;
    this.sharedStringsArr.push(value);
    return sharedStringIndex;
  }

  _entries() {
    const entries = Object.keys(this.xlsx).map(name => ({
      name,
      source: this.xlsx[name],
    }));
    // Async generators, deliberately not the ReadableStream getters above: a
    // ReadableStream starts pulling the moment it is constructed, which would
    // emit the shared-string header — and its count — before the sheet had been
    // walked. A generator does nothing until the zip writer reaches its entry,
    // so the ordering the workbook depends on is a contract here rather than
    // the accident it was under JSZip.
    entries.push({ name: SHEET_PART, source: this._sheetXml() });
    entries.push({ name: SHARED_STRINGS_PART, source: this._sharedStringsXml() });
    return entries;
  }

  _claim() {
    if (!this._rowsAdded) throw new Error("call addRows() before building the workbook");
    if (this._consumed) {
      throw new Error(
        "this writer has already produced a workbook — the row stream has been consumed",
      );
    }
    this._consumed = true;
  }

  /**
   * The workbook as a stream of byte chunks, without ever holding the whole
   * archive in memory.
   * @returns {ReadableStream<Uint8Array>}
   */
  getStream() {
    this._claim();
    return toWebReadableStream(
      writeZip(this._entries(), { level: this.options.compressionLevel }),
    );
  }

  /**
   * The whole workbook at once: a Buffer in Node.js, a Blob in the browser.
   * @returns {Promise<Buffer|Blob>}
   */
  async getFile() {
    this._claim();
    const chunks = [];
    for await (const chunk of writeZip(this._entries(), {
      level: this.options.compressionLevel,
    })) {
      chunks.push(chunk);
    }

    const isBrowser =
      typeof window !== "undefined" &&
      {}.toString.call(window) === "[object Window]";

    return isBrowser
      ? new Blob(chunks, {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        })
      : Buffer.concat(chunks);
  }
}

function cleanUpXml(xml) {
  return xml.replace(/>\s+</g, "><").trim();
}

module.exports = XlsxStreamWriter;
