const Readable = require("stream-browserify").Readable;
const PassThrough = require("stream-browserify").PassThrough;
const JSZip = require("jszip");
const xmlParts = require("./xml/parts");
const xmlBlobs = require("./xml/blobs");
const { getCellAddress, wrapRowsInStream } = require("./helpers");
// const { crc32 } = require("crc");

const defaultOptions = {
  inlineStrings: false,
};

class XlsxStreamWriter {
  constructor(options) {
    this.options = Object.assign(defaultOptions, options);
    this.sheetXmlStream = null;

    this.sharedStringsXmlStream = null;
    this.sharedStringsArr = [];
    this.sharedStringsMap = {};
    this.sharedStringsHashMap = {};

    this.xlsx = {
      "[Content_Types].xml": cleanUpXml(xmlBlobs.contentTypes),
      "_rels/.rels": cleanUpXml(xmlBlobs.rels),
      "xl/workbook.xml": cleanUpXml(xmlBlobs.workbook),
      "xl/styles.xml": cleanUpXml(xmlBlobs.styles),
      "xl/_rels/workbook.xml.rels": cleanUpXml(xmlBlobs.workbookRels),
    };
  }

  /**
   * Add rows to xlsx.
   * @param {Array | Readable} rowsOrStream array of arrays or readable stream of arrays
   * @return {undefined}
   */
  addRows(rowsOrStream) {
    let rowsStream;
    if (rowsOrStream instanceof Readable) rowsStream = rowsOrStream;
    else if (Array.isArray(rowsOrStream))
      rowsStream = wrapRowsInStream(rowsOrStream);
    else
      throw Error(
        "Argument must be an array of arrays or a readable stream of arrays",
      );
    const rowsToXml = this._getRowsToXmlTransformStream();
    const tsToString = this._getToStringTransforStream();
    // TODO why do we need to call .toString in case we want to inline strings?
    this.sheetXmlStream = this.options.inlineStrings
      ? rowsStream.pipe(rowsToXml).pipe(tsToString)
      : rowsStream.pipe(rowsToXml);
    this.sharedStringsXmlStream = this._getSharedStringsXmlStream();
  }

  _getToStringTransforStream() {
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
      if (c === 0) {
        ts.push(xmlParts.sheetHeader, "utf8");
      }
      const rowXml = this._getRowXml(data, c);
      // console.log(rowXml);
      ts.push(rowXml.toString(), "utf8");
      c++;
      callback();
    };

    ts._flush = cb => {
      ts.push(xmlParts.sheetFooter, "utf8");
      cb();
    };
    return ts;
  }

  _getRowXml(row, rowIndex) {
    let rowXml = xmlParts.getRowStart(rowIndex);
    row.forEach((cellValue, colIndex) => {
      const cellAddress = getCellAddress(rowIndex + 1, colIndex + 1);
      rowXml += this._getCellXml(cellValue, cellAddress);
    });
    rowXml += xmlParts.rowEnd;
    return rowXml;
  }

  _getCellXml(value, address) {
    let cellXml;
    if (Number.isNaN(value) || value === null || typeof value === "undefined")
      cellXml = xmlParts.getStringCellXml("", address);
    else if (typeof value === "number")
      cellXml = xmlParts.getNumberCellXml(value, address);
    else cellXml = this._getStringCellXml(value, address);
    return cellXml;
  }

  _getStringCellXml(value, address) {
    const stringValue = String(value);
    // console.log(value, stringValue);
    return this.options.inlineStrings
      ? xmlParts.getInlineStringCellXml(escapeXml(String(value)), address)
      : xmlParts.getStringCellXml(this._lookupString(stringValue), address);
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
        rs.push(xmlParts.getSharedStringsHeader(this.sharedStringsArr.length));
      }
      if (c === this.sharedStringsArr.length) {
        rs.push(xmlParts.sharedStringsFooter);
        rs.push(null);
      } else
        rs.push(
          xmlParts.getSharedStringXml(
            escapeXml(String(this.sharedStringsArr[c])),
          ),
        );
      c++;
    };
    return rs;
  }

  _clearSharedStrings() {
    this.sharedStringsMap = {};
    this.sharedStringsArr = [];
  }

  // returns blob in a browser, buffer in nodejs
  getFile() {
    this._clearSharedStrings();
    const zip = new JSZip();
    // add all static files
    Object.keys(this.xlsx).forEach(key => zip.file(key, this.xlsx[key]));

    // add "xl/worksheets/sheet1.xml"
    zip.file("xl/worksheets/sheet1.xml", this.sheetXmlStream);
    // add "xl/sharedStrings.xml"
    zip.file("xl/sharedStrings.xml", this.sharedStringsXmlStream);
    this._clearSharedStrings();

    const isBrowser =
      typeof window !== "undefined" &&
      {}.toString.call(window) === "[object Window]";

    return new Promise((resolve, reject) => {
      if (isBrowser) {
        zip
          .generateAsync({
            type: "blob",
            compression: "DEFLATE",
            compressionOptions: {
              level: 4,
            },
            streamFiles: true,
          })
          .then(resolve)
          .catch(reject);
      } else {
        zip
          .generateAsync({
            type: "nodebuffer",
            platform: process.platform,
            compression: "DEFLATE",
            compressionOptions: {
              level: 4,
            },
            streamFiles: true,
          })
          .then(resolve)
          .catch(reject);
        // zip
        //   .generateNodeStream({
        //     type: "nodebuffer",
        //     platform: process.platform,
        //   })
        //   .pipe(require('fs').createWriteStream("test.xlsx"))
        //   .on("finish", () => {
        //     // JSZip generates a readable stream with a "end" event,
        //     // but is piped here in a writable stream which emits a "finish" event.
        //     console.log(`test xlsx written.`);
        //   });
      }
    });
  }
}

function cleanUpXml(xml) {
  return xml.replace(/>\s+</g, "><").trim();
}

function escapeXml(str = "") {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

module.exports = XlsxStreamWriter;
