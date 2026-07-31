import * as xmlParts from "./xml/parts";
import * as xmlBlobs from "./xml/blobs";
import { getCellAddress, escapeXml } from "./helpers";
import { toAsyncIterable, toWebReadableStream, type StreamLike } from "./streams";
import { writeZip, type ZipEntry } from "./zip/writer";
import { getStyles, type CellStyle } from "./styles";

/** A value this package knows how to put in a cell. */
type CellValue = string | number | bigint | boolean | Date | null | undefined | object;

type Row = readonly CellValue[];

interface XlsxStreamWriterOptions {
  /** Write strings into the sheet rather than into a shared-string table. */
  inlineStrings?: boolean;
  /** Cell formats, referenced by index from `styleIdFunc`. Index 0 is the default. */
  styles?: readonly CellStyle[];
  /** Choose a style index per cell. */
  styleIdFunc?: (value: CellValue, columnId: number, rowId: number) => number;
  /** Deflate level, 0-9. Node only; browsers expose no level control. */
  compressionLevel?: number;
}

type ResolvedOptions = Required<XlsxStreamWriterOptions>;

const defaultOptions: ResolvedOptions = {
  inlineStrings: false,
  styles: [],
  styleIdFunc: () => 0,
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
  readonly options: ResolvedOptions;
  readonly xlsx: Record<string, string>;

  sharedStringsArr: string[] = [];
  // A Map, not a plain object: a cell holding "constructor", "toString" or
  // "__proto__" would otherwise look up an inherited member of Object.prototype
  // rather than miss, and that member would be written into the sheet as the
  // string's index — corrupting the workbook and losing the string table.
  sharedStringsMap = new Map<string, number>();

  private rows: AsyncIterable<Row> | null = null;
  private rowsAdded = false;
  private consumed = false;
  private sheetComplete = false;
  private sharedStringRefs = 0;
  private sheetStream: ReadableStream<string> | null = null;
  private sharedStringsStream: ReadableStream<string> | null = null;
  /** cellXfs entries: the implicit default plus one per declared style. */
  private readonly styleCount: number;

  constructor(options?: XlsxStreamWriterOptions) {
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

    this.styleCount = (this.options.styles?.length ?? 0) + 1;

    this.xlsx = {
      "[Content_Types].xml": xmlBlobs.contentTypes,
      "_rels/.rels": xmlBlobs.rels,
      "xl/workbook.xml": xmlBlobs.workbook,
      "xl/styles.xml": getStyles(this.options.styles),
      "xl/_rels/workbook.xml.rels": xmlBlobs.workbookRels,
    };
  }

  /**
   * Add rows to the workbook: an array of arrays, a readable stream of arrays,
   * or any iterable or async iterable of arrays.
   */
  addRows(rowsOrStream: StreamLike<Row>): void {
    if (this.rowsAdded) {
      throw new Error(
        "addRows() can only be called once per writer — the previous row stream would be discarded. Create a new XlsxStreamWriter for another workbook.",
      );
    }
    this.rows = toAsyncIterable(rowsOrStream, "Rows");
    this.rowsAdded = true;
  }

  /** The worksheet part, as a stream of XML text. */
  get sheetXmlStream(): ReadableStream<string> | null {
    if (!this.rowsAdded) return null;
    if (!this.sheetStream) this.sheetStream = toWebReadableStream(this.sheetXml());
    return this.sheetStream;
  }

  /**
   * The shared-string table, as a stream of XML text.
   *
   * Only read this once `sheetXmlStream` has been drained: the table is built
   * as the sheet is walked, and this stream begins emitting — counts included —
   * as soon as you touch it.
   */
  get sharedStringsXmlStream(): ReadableStream<string> {
    if (!this.sharedStringsStream) {
      this.sharedStringsStream = toWebReadableStream(this.sharedStringsXml());
    }
    return this.sharedStringsStream;
  }

  private async *sheetXml(): AsyncGenerator<string> {
    yield xmlParts.sheetHeader;
    let rowIndex = 0;
    for await (const row of this.rows!) {
      yield this.getRowXml(row, rowIndex);
      rowIndex++;
    }
    this.sheetComplete = true;
    yield xmlParts.sheetFooter;
  }

  private async *sharedStringsXml(): AsyncGenerator<string> {
    // The table is filled as the sheet is walked, so emitting it early would
    // declare a count of zero and list nothing — a valid-looking part that is
    // quietly wrong. The archive writer never trips this, because it consumes
    // entries in order.
    if (!this.sheetComplete) {
      throw new Error(
        "the shared-string table is only complete once the worksheet has been read — drain sheetXmlStream first",
      );
    }
    yield xmlParts.getSharedStringsHeader(this.sharedStringsArr.length, this.sharedStringRefs);
    for (const value of this.sharedStringsArr) {
      yield xmlParts.getSharedStringXml(escapeXml(String(value)));
    }
    yield xmlParts.sharedStringsFooter;
  }

  /** @internal exposed for tests; the row-limit guard is impractical to reach otherwise */
  _getRowXml(row: Row, rowIndex: number): string {
    return this.getRowXml(row, rowIndex);
  }

  private getRowXml(row: Row, rowIndex: number): string {
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
      const styleId = this.resolveStyleId(cellValue, colIndex, rowIndex, cellAddress);
      rowXml += this.getCellXml(cellValue, cellAddress, styleId);
    });
    rowXml += xmlParts.rowEnd;
    return rowXml;
  }

  /**
   * A style id is interpolated straight into an attribute and has to index a
   * real `cellXfs` entry, so anything else is rejected at the cell rather than
   * left to surface as a repair prompt when the file is opened.
   */
  private resolveStyleId(
    value: CellValue,
    colIndex: number,
    rowIndex: number,
    address: string,
  ): number {
    const styleId = this.options.styleIdFunc(value, colIndex, rowIndex);
    if (!Number.isInteger(styleId) || (styleId as number) < 0) {
      // Name the type too: a boxed Number or a numeric string prints like a
      // perfectly good index, which makes the bare value baffling on its own.
      throw new TypeError(
        `styleIdFunc returned ${JSON.stringify(styleId)} (${typeof styleId}) for cell ${address}; it must return a non-negative integer`,
      );
    }
    if (styleId >= this.styleCount) {
      throw new RangeError(
        `styleIdFunc returned style ${styleId} for cell ${address}, but only ${this.styleCount} styles are defined (0 is the default, and options.styles adds ${this.styleCount - 1} more)`,
      );
    }
    return styleId;
  }

  private getCellXml(value: CellValue, address: string, styleId = 0): string {
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
    return this.getStringCellXml(value, address, styleId);
  }

  private getStringCellXml(value: CellValue, address: string, styleId: number): string {
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
    this.sharedStringRefs++;
    return xmlParts.getStringCellXml(this.lookupString(stringValue), address, styleId);
  }

  private lookupString(value: string): number {
    const existing = this.sharedStringsMap.get(value);
    if (existing !== undefined) return existing;
    const index = this.sharedStringsArr.length;
    this.sharedStringsMap.set(value, index);
    this.sharedStringsArr.push(value);
    return index;
  }

  private entries(): ZipEntry[] {
    const entries: ZipEntry[] = Object.keys(this.xlsx).map(name => ({
      name,
      source: this.xlsx[name]!,
    }));
    // Async generators, deliberately not the ReadableStream getters above: a
    // ReadableStream starts pulling the moment it is constructed, which would
    // emit the shared-string header — and its count — before the sheet had been
    // walked. A generator does nothing until the zip writer reaches its entry,
    // so the ordering the workbook depends on is a contract here rather than
    // the accident it was under JSZip.
    entries.push({ name: SHEET_PART, source: this.sheetXml() });
    entries.push({ name: SHARED_STRINGS_PART, source: this.sharedStringsXml() });
    return entries;
  }

  private claim(): void {
    if (!this.rowsAdded) throw new Error("call addRows() before building the workbook");
    if (this.consumed) {
      throw new Error(
        "this writer has already produced a workbook — the row stream has been consumed",
      );
    }
    if (this.sheetStream) {
      // Reading sheetXmlStream walks the rows, and rows can only be walked
      // once. Without this the workbook would come out silently empty.
      throw new Error(
        "sheetXmlStream has already been handed out, and reading it consumes the rows — inspect the XML streams or build a workbook, not both",
      );
    }
    this.consumed = true;
  }

  /**
   * The workbook as a stream of byte chunks, without ever holding the whole
   * archive in memory.
   */
  getStream(): ReadableStream<Uint8Array> {
    this.claim();
    return toWebReadableStream(
      writeZip(this.entries(), { level: this.options.compressionLevel }),
    );
  }

  /** The whole workbook at once: a Buffer in Node.js, a Blob in the browser. */
  async getFile(): Promise<Buffer | Blob> {
    this.claim();
    const chunks: Uint8Array[] = [];
    for await (const chunk of writeZip(this.entries(), {
      level: this.options.compressionLevel,
    })) {
      chunks.push(chunk);
    }

    const isBrowser =
      typeof window !== "undefined" && {}.toString.call(window) === "[object Window]";

    return isBrowser
      ? new Blob(chunks as BlobPart[], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        })
      : Buffer.concat(chunks);
  }
}

// Merged with the class so consumers of this CommonJS package can still write
// `import type { CellValue } from "xlsx-stream-writer"` alongside `export =`.
declare namespace XlsxStreamWriter {
  export type { CellValue, Row, XlsxStreamWriterOptions, CellStyle };
}

export = XlsxStreamWriter;
