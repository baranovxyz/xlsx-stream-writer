const fs = require("fs");
const fse = require("fs-extra");
const path = require("path");
const JSZip = require("jszip");

const blobs = require("./xml-parts");

const numberRegex = /^[1-9\.][\d\.]+$/;

class XlsxWriter {
  constructor(out) {
    this.out = out;
    this.strings = [];
    this.stringMap = {};
    this.stringIndex = 0;
    this.currentRow = 0;

    this.haveHeader = false;
    this.prepared = false;

    this.tempPath = "";

    this.sheetStream = null;

    this.cellMap = [];
    this.cellLabelMap = {};
  }

  addRow(obj) {
    // console.log('add row', this);
    if (!this.prepared) throw "Should call prepare() first!";
    if (!this.haveHeader) {
      this._startRow();
      let col = 1;
      Object.keys(obj).map(key => {
        this._addCell(key, col);
        this.cellMap.push(key);
        col++;
      });
      this._endRow();
      this.haveHeader = true;
    }
    this._startRow();
    this.cellMap.forEach((key, col) => this._addCell(obj[key] || "", col + 1));
    this._endRow();
  }

  prepare(rows, columns) {
    // Add one extra row for the header
    // console.log('prepare', this);
    const dimensions = this.dimensions(rows + 1, columns);
    console.log({ dimensions });
    this.tempPath = path.join(__dirname, "temp");

    console.log("temp path", this.tempPath);
    fse.removeSync(this.tempPath);
    fse.mkdirSync(this.tempPath);

    fse.ensureDirSync(this._filename("_rels"));
    fse.ensureDirSync(this._filename("xl"));
    fse.ensureDirSync(this._filename("xl", "_rels"));
    fse.ensureDirSync(this._filename("xl", "worksheets"));
    fs.writeFileSync(this._filename("[Content_Types].xml"), blobs.contentTypes);
    fs.writeFileSync(this._filename("_rels", ".rels"), blobs.rels);
    fs.writeFileSync(this._filename("xl", "workbook.xml"), blobs.workbook);
    fs.writeFileSync(this._filename("xl", "styles.xml"), blobs.styles);
    fs.writeFileSync(
      this._filename("xl", "_rels", "workbook.xml.rels"),
      blobs.workbookRels,
    );
    console.log(
      "sheet write stream",
      this._filename("xl", "worksheets", "sheet1.xml"),
    );
    this.sheetStream = fs.createWriteStream(
      this._filename("xl", "worksheets", "sheet1.xml"),
    );
    this.sheetStream.write(blobs.sheetHeader(dimensions));
    this.prepared = true;
    return true;
  }

  async _endSheet() {
    return new Promise((resolve, reject) => {
      this.sheetStream.write(blobs.sheetFooter);
      this.sheetStream.end(() => resolve());
    });
  }

  async pack() {
    if (!this.prepared) throw "Should call prepare() first!";
    await this._endSheet();

    const zipfile = new JSZip();

    let stringTable = "";
    this.strings.map(text => {
      stringTable += blobs.string(this.escapeXml(String(text)));
    });
    fs.writeFileSync(
      this._filename("xl", "sharedStrings.xml"),
      blobs.stringsHeader(this.strings.length) +
        stringTable +
        blobs.stringsFooter,
    );

    const readFile = filePath => fs.readFileSync(filePath);

    zipfile.file(
      "[Content_Types].xml",
      readFile(this._filename("[Content_Types].xml")),
    );
    zipfile.file("_rels/.rels", readFile(this._filename("_rels", ".rels")));
    zipfile.file(
      "xl/workbook.xml",
      readFile(this._filename("xl", "workbook.xml")),
    );
    zipfile.file("xl/styles.xml", readFile(this._filename("xl", "styles.xml")));
    zipfile.file(
      "xl/sharedStrings.xml",
      readFile(this._filename("xl", "sharedStrings.xml")),
    );
    zipfile.file(
      "xl/_rels/workbook.xml.rels",
      readFile(this._filename("xl", "_rels", "workbook.xml.rels")),
    );
    zipfile.file(
      "xl/worksheets/sheet1.xml",
      readFile(this._filename("xl", "worksheets", "sheet1.xml")),
    );
    await zipfile
      .generateNodeStream({ type: "nodebuffer", streamFiles: true })
      .pipe(fs.createWriteStream(this.out))
      .on("finish", () => {
        // JSZip generates a readable stream with a "end" event,
        // but is piped here in a writable stream which emits a "finish" event.
        console.log(`${this.out} written.`);
      });
  }

  dimensions(rows, columns) {
    return "A1:" + this.cell(rows, columns);
  }

  cell(row, col) {
    let colIndex = "";
    if (this.cellLabelMap[col]) colIndex = this.cellLabelMap[col];
    else {
      if (col === 0) {
        row = 1;
        col = 1;
      }
      let input = (+col - 1).toString(26);
      while (input.length) {
        const a = input.charCodeAt(input.length - 1);
        colIndex =
          String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colIndex;
        input =
          input.length > 1
            ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26)
            : "";
      }
      this.cellLabelMap[col] = colIndex;
    }
    return colIndex + row;
  }

  _filename(...args) {
    return path.join(this.tempPath, ...args);
  }

  _startRow() {
    this.rowBuffer = blobs.startRow(this.currentRow);
    this.currentRow++;
  }

  _lookupString(value) {
    if (!this.stringMap[value]) {
      this.stringMap[value] = this.stringIndex;
      this.strings.push(value);
      this.stringIndex += 1;
    }
    return this.stringMap[value];
  }

  _addCell(value = "", col) {
    const cell = this.cell(this.currentRow, col);
    if (numberRegex.test(value))
      this.rowBuffer += blobs.numberCell(value, cell);
    else this.rowBuffer += blobs.cell(this._lookupString(value), cell);
  }

  _endRow() {
    this.sheetStream.write(this.rowBuffer + blobs.endRow);
  }

  escapeXml(str = "") {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }
}

module.exports = { XlsxWriter };
