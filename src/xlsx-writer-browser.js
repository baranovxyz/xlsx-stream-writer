const Writable = require("stream-browserify").Writable;
const JSZip = require("jszip");
const xmlParts = require("./xml-parts");
const xmlBlobs = require("./xml-blobs");
const { getCellAddress, getCellXml, getRowXml } = require("./helpers");

class XlsxWriter extends Writable {
  constructor() {
    // https://github.com/substack/stream-handbook
    // If the readable stream you're piping from writes strings, they will be converted
    // into Buffers unless you create your writable stream
    // with Writable({ decodeStrings: false }).
    super({ decodeStrings: true });
    this.sharedStrings = [];
    this.sharedStringsMap = {};

    this.currentRow = 0;
    this.sheetEnded = false;

    this.sharedStringsXml = "";
    this.sheetStringXml = "";
    this.xlsx = {
      "[Content_Types].xml": cleanUpXml(xmlBlobs.contentTypes),
      "_rels/.rels": cleanUpXml(xmlBlobs.rels),
      "xl/workbook.xml": cleanUpXml(xmlBlobs.workbook),
      "xl/styles.xml": cleanUpXml(xmlBlobs.styles),
      "xl/_rels/workbook.xml.rels": cleanUpXml(xmlBlobs.workbookRels),
    };

    this.on("pipe", () => {
      console.log("pipe");
    });

    this.on("unpipe", () => {
      console.log("unpipe");
    });

    this.on("end", () => {
      console.log("end");
    });

    this.on("finish", () => {
      console.log("finish");
    });

    this.on("close", () => {
      console.log("close");
    });

    this.on("drain", () => {
      console.log("drain");
    });

    this._startSheet();
  }

  // write rows here
  _write(chunk, enc, next) {
    if (chunk === null) console.log("i got null!");
    console.log(chunk);
    next();
  }

  addRow(row) {
    this._startRow();
    row.forEach((value, index) => this._addCell(value, index + 1));
    this._endRow();
  }

  end() {
    console.log("end is called!");
    this._endSheet();
    this._processSharedStrings();
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

  _processSharedStrings() {
    // clean up map asap
    this.sharedStringsMap = {};
    this.sharedStringsXml = xmlParts.getSharedStringsHeader(
      this.sharedStrings.length,
    );
    this.sharedStrings.map(text => {
      this.sharedStringsXml += xmlParts.getSharedStringXml(
        escapeXml(String(text)),
      );
    });
    this.sharedStringsXml += xmlParts.sharedStringsFooter;
    // clean up array asap
    this.sharedStrings = [];
  }

  // returns blob in a browser, buffer in nodejs
  getFile() {
    if (!this.sheetEnded) {
      this.end();
      console.warn("Sheet was ended, because getBlob() was called.");
    }
    const zip = new JSZip();
    // add all static files
    Object.keys(this.xlsx).forEach(key => zip.file(key, this.xlsx[key]));
    // add "xl/sharedStrings.xml"
    zip.file("xl/sharedStrings.xml", this.sharedStringsXml);
    // add "xl/worksheets/sheet1.xml"
    zip.file("xl/worksheets/sheet1.xml", this.sheetStringXml);

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

module.exports = { XlsxWriter, getRowXml };
