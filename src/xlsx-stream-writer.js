const Readable = require("stream-browserify").Readable;
const PassThrough = require("stream-browserify").PassThrough;
const JSZip = require("jszip");
const xmlParts = require("./xml/parts");
const xmlBlobs = require("./xml/blobs");
const { getCellAddress } = require("./helpers");
// const { crc32 } = require("crc");

class XlsxStreamWriter {
  constructor() {
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
    this.sheetXmlStream = rowsStream.pipe(rowsToXml);
    this.sharedStringsXmlStream = this._getSharedStringsXmlStream();
  }

  _getRowsToXmlTransformStream() {
    const ts = PassThrough({ objectMode: true });
    let c = 0;
    ts._transform = (data, encoding, callback) => {
      if (c === 0) {
        ts.push(xmlParts.sheetHeader);
      }
      const rowXml = this._getRowXml(data, c);
      ts.push(rowXml);
      c++;
      callback();
    };

    ts._flush = cb => {
      ts.push(xmlParts.sheetFooter);
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
    return xmlParts.getStringCellXml(this._lookupString(stringValue), address);
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
