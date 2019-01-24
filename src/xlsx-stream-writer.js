const Readable = require("stream-browserify").Readable;
const PassThrough = require("stream-browserify").PassThrough;
const JSZip = require("jszip");
const xmlParts = require("./xml/parts");
const xmlBlobs = require("./xml/blobs");
const { getCellAddress, getRowXml } = require("./xml/helpers");

class XlsxWriter {
  constructor() {
    // https://github.com/substack/stream-handbook
    // If the readable stream you're piping from writes strings, they will be converted
    // into Buffers unless you create your writable stream
    // with Writable({ decodeStrings: false }).
    this.sharedStrings = [];
    this.sharedStringsMap = {};

    this.currentRow = 0;
    this.sheetEnded = false;

    this.sheetXmlStream = null;
    this.sharedStringsXmlStream = null;
    this.sharedStringsXml = "";
    // this.sharedStringsXml = "";
    // this.sheetStringXml = "";
    this.xlsx = {
      "[Content_Types].xml": cleanUpXml(xmlBlobs.contentTypes),
      "_rels/.rels": cleanUpXml(xmlBlobs.rels),
      "xl/workbook.xml": cleanUpXml(xmlBlobs.workbook),
      "xl/styles.xml": cleanUpXml(xmlBlobs.styles),
      "xl/_rels/workbook.xml.rels": cleanUpXml(xmlBlobs.workbookRels),
    };

    this._startSheet();
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
        // console.log("push sheet header");
        ts.push(xmlParts.sheetHeader);
      }
      // console.log("push data:", JSON.stringify(data).slice(0, 100));
      const rowXml = getRowXml.bind(this)(data, c);
      // console.log(rowXml);
      ts.push(rowXml);
      c++;
      callback();
    };

    ts._flush = cb => {
      // console.log("push sheet footer");
      ts.push(xmlParts.sheetFooter);
      cb();
    };
    return ts;
  }

  _getSharedStringsXmlStream() {
    const rs = Readable();
    let c = 0;
    rs._read = () => {
      if (c === 0) {
        rs.push(xmlParts.getSharedStringsHeader(this.sharedStrings.length));
      }
      if (c === this.sharedStrings.length) {
        rs.push(xmlParts.sharedStringsFooter);
        rs.push(null);
      } else
        rs.push(
          xmlParts.getSharedStringXml(escapeXml(String(this.sharedStrings[c]))),
        );
      c++;
    };
    return rs;
  }

  addRow(row) {
    this._startRow();
    row.forEach((value, index) => this._addCell(value, index + 1));
    this._endRow();
  }

  endSheet() {
    this._endSheet();
    this.sharedStringsMap = {};
    this.sheetEnded = true;
  }

  _startSheet() {
    // const sheetRange = getRange(numRows, numColumns);
    this.sheetStringXml = xmlParts.sheetHeader;
  }

  _endSheet() {
    this.sheetStringXml += xmlParts.sheetFooter;
  }

  _startRow() {
    this.rowBuffer = xmlParts.getRowStart(this.currentRow);
    this.currentRow++;
  }

  _endRow() {
    this.sheetStringXml += this.rowBuffer + xmlParts.rowEnd;
  }

  _addCell(value, colIndex) {
    const cellAddress = getCellAddress(this.currentRow, colIndex);
    let cellXml;
    if (Number.isNaN(value) || value === null || typeof value === "undefined")
      cellXml = xmlParts.getStringCellXml("", cellAddress);
    else if (typeof value === "number")
      cellXml = xmlParts.getNumberCellXml(value, cellAddress);
    else
      cellXml = xmlParts.getStringCellXml(
        this._lookupString(String(value)),
        cellAddress,
      );

    this.rowBuffer += cellXml;
  }

  _lookupString(value) {
    let sharedStringIndex = this.sharedStringsMap[value];
    if (typeof sharedStringIndex !== "undefined") return sharedStringIndex;
    sharedStringIndex = this.sharedStrings.length;
    this.sharedStringsMap[value] = sharedStringIndex;
    this.sharedStrings.push(value);
    return sharedStringIndex;
  }

  // returns blob in a browser, buffer in nodejs
  getFile() {
    this.endSheet();
    const zip = new JSZip();
    // add all static files
    Object.keys(this.xlsx).forEach(key => zip.file(key, this.xlsx[key]));

    // add "xl/worksheets/sheet1.xml"
    zip.file("xl/worksheets/sheet1.xml", this.sheetXmlStream);
    // add "xl/sharedStrings.xml"
    zip.file("xl/sharedStrings.xml", this.sharedStringsXmlStream);
    // clean shared strings
    this.sharedStrings = [];
    
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

module.exports = XlsxWriter;