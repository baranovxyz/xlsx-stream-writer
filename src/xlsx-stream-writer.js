const PassThrough = require("stream-browserify").PassThrough;
const Readable = require("stream-browserify").Readable;
const JSZip = require("jszip");
const xmlParts = require("./xml/parts");
const xmlBlobs = require("./xml/blobs");
const { getCellAddress, toRowsStream, escapeXml } = require("./helpers");
const getStyles = require("./styles").getStyles;

const defaultOptions = {
  inlineStrings: false,
  styles: [],
  styleIdFunc: (value, columnId, rowId) => 0,
};

// The grid Excel itself supports. Exceeding either limit produces a file Excel
// refuses to open, so it is better to say so at the offending row.
const MAX_ROWS = 1048576;
const MAX_COLUMNS = 16384;

// Excel counts days from 1899-12-30 — two days behind 1900-01-01, because it
// preserves Lotus 1-2-3's belief that 1900 was a leap year.
const EXCEL_EPOCH_OFFSET_DAYS = 25569;
const MS_PER_DAY = 86400000;

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

    this.sheetXmlStream = null;
    this.sharedStringsXmlStream = null;
    this.sharedStringsArr = [];
    this.sharedStringsMap = {};
    this.sharedStringsHashMap = {};

    this._rowsAdded = false;
    this._fileRequested = false;
    this._error = null;
    this._rejectFile = null;
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
   * Add rows to xlsx.
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
    const rowsStream = toRowsStream(rowsOrStream);
    this._rowsAdded = true;

    const rowsToXml = this._getRowsToXmlTransformStream();

    // .pipe() does not forward errors, so without this a failing source leaves
    // the destination open forever and getFile() never settles.
    rowsStream.on("error", error => this._fail(error));
    rowsToXml.on("error", error => this._fail(error));

    if (this.options.inlineStrings) {
      const tsToString = this._getToStringTransformStream();
      tsToString.on("error", error => this._fail(error));
      this.sheetXmlStream = rowsStream.pipe(rowsToXml).pipe(tsToString);
    } else {
      this.sheetXmlStream = rowsStream.pipe(rowsToXml);
    }

    // Keep a failure before getFile() from crashing the process as an unhandled
    // 'error' event; _fail() has already recorded it for the promise to reject.
    this.sheetXmlStream.on("error", () => {});

    this.sharedStringsXmlStream = this._getSharedStringsXmlStream();
  }

  _fail(error) {
    if (this._error) return;
    this._error = error;
    if (this._rejectFile) this._rejectFile(error);
  }

  _getToStringTransformStream() {
    const ts = PassThrough();
    ts._transform = (data, encoding, callback) => {
      ts.push(data.toString(), "utf8");
      callback();
    };
    return ts;
  }

  _getRowsToXmlTransformStream() {
    const ts = PassThrough({ objectMode: true });
    let c = 0;
    ts._transform = (data, encoding, callback) => {
      try {
        if (c === 0) ts.push(xmlParts.sheetHeader, "utf8");
        ts.push(this._getRowXml(data, c), "utf8");
        c++;
        callback();
      } catch (error) {
        callback(error);
      }
    };

    ts._flush = cb => {
      // An empty workbook still needs its header, or the sheet part is a bare
      // closing tag and the whole file is malformed.
      if (c === 0) ts.push(xmlParts.sheetHeader, "utf8");
      ts.push(xmlParts.sheetFooter, "utf8");
      cb();
    };
    return ts;
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

  _getSharedStringsXmlStream() {
    const rs = Readable();
    let c = 0;
    rs._read = () => {
      if (c === 0) {
        rs.push(
          xmlParts.getSharedStringsHeader(
            this.sharedStringsArr.length,
            this._sharedStringRefs,
          ),
        );
      }
      if (c === this.sharedStringsArr.length) {
        rs.push(xmlParts.sharedStringsFooter);
        rs.push(null);
      } else
        rs.push(
          xmlParts.getSharedStringXml(escapeXml(String(this.sharedStringsArr[c]))),
        );
      c++;
    };
    return rs;
  }

  _clearSharedStrings() {
    this.sharedStringsMap = {};
    this.sharedStringsArr = [];
    this._sharedStringRefs = 0;
  }

  // returns blob in a browser, buffer in nodejs
  getFile() {
    if (!this._rowsAdded) {
      throw new Error("call addRows() before getFile()");
    }
    if (this._fileRequested) {
      throw new Error(
        "getFile() can only be called once — the row stream has already been consumed",
      );
    }
    this._fileRequested = true;

    this._clearSharedStrings();
    const zip = new JSZip();
    // add all static files
    Object.keys(this.xlsx).forEach(key => zip.file(key, this.xlsx[key]));

    // add "xl/worksheets/sheet1.xml"
    zip.file("xl/worksheets/sheet1.xml", this.sheetXmlStream);
    // add "xl/sharedStrings.xml"
    zip.file("xl/sharedStrings.xml", this.sharedStringsXmlStream);

    const isBrowser =
      typeof window !== "undefined" &&
      {}.toString.call(window) === "[object Window]";

    const generateOptions = {
      type: isBrowser ? "blob" : "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 4 },
      streamFiles: true,
    };
    if (!isBrowser) generateOptions.platform = process.platform;

    return new Promise((resolve, reject) => {
      if (this._error) {
        reject(this._error);
        return;
      }
      this._rejectFile = reject;
      zip.generateAsync(generateOptions).then(resolve, reject);
    });
  }
}

function cleanUpXml(xml) {
  return xml.replace(/>\s+</g, "><").trim();
}

module.exports = XlsxStreamWriter;
