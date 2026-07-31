const test = require("node:test");
const assert = require("node:assert/strict");
const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const XlsxStreamWriter = require("../../dist");
const { readZip } = require("../support/unzip");
const { parseCsv } = require("../support/csv");
const { BUFFER_LIMIT } = require("../../dist/zip/writer");

// Every other check in this suite reads the archive and the XML. None of them
// prove a spreadsheet application accepts the result, which is the only claim
// that matters to a caller. LibreOffice is the practical proxy for Excel: it is
// scriptable, so it can run here, and it shares no lineage with this writer.
//
// Skipped silently when LibreOffice is absent, the same way the system unzip
// checks are, so the suite stays portable. docs/verifying-output.md records what
// this establishes and what it still does not.

const STREAMED_ROW_COUNT = 120000;

function* streamedRows() {
  yield ["Name", "Location", "Amount", "Active"];
  for (let i = 1; i < STREAMED_ROW_COUNT; i++) {
    yield [`Customer number ${i}`, `City of origin ${i % 997}`, i * 1.5, i % 2 === 0];
  }
}

// Values chosen because they have broken spreadsheet readers, not because they
// look realistic. Each is [label, written value, what LibreOffice should read
// back] — the third differs from the second only where CSV has no other way to
// spell it.
const HOSTILE = [
  ["xml-lt-gt", "<sheet>"],
  ["xml-amp", "a & b"],
  ["xml-quote", 'say "hi"'],
  ["xml-apos", "it's"],
  ["xml-cdata", "]]><!--x-->"],
  ["attr-escape", '"/><f>=1+1</f><x a="'],
  // The bug class this repo has hit twice: a caller-supplied string that names a
  // member of Object.prototype. If the shared-string table were a plain object,
  // these would find an inherited member where an index belonged.
  ["proto-key", "__proto__"],
  ["ctor-key", "constructor"],
  ["tostring-key", "toString"],
  ["valueof-key", "valueOf"],
  ["hasownprop-key", "hasOwnProperty"],
  ["proto-key-repeated", "__proto__"],
  // Must come back as text. If either is evaluated, the writer emitted a formula
  // where it was handed a string.
  ["formula-like", "=1+1"],
  ["dde-like", "=cmd|' /C calc'!A0"],
  ["leading-space", "   pad"],
  ["trailing-space", "pad   "],
  ["emoji", "\u{1F600}\u{1F1F7}\u{1F1FA}"],
  ["rtl", "שלום"],
  ["combining", "é"],
  ["cjk", "你好"],
  ["long-string", "x".repeat(32767)],
  ["number-negative", -42, "-42"],
  ["number-float", 1.5, "1.5"],
  ["number-zero", 0, "0"],
  ["bool-true", true, "TRUE"],
  ["bool-false", false, "FALSE"],
];

const TEXT_ROWS = [
  ["Name", "Location"],
  ["Иван", "Москва"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const FIXTURES = [
  {
    name: "shared-strings",
    async build() {
      const xlsx = new XlsxStreamWriter();
      xlsx.addRows(TEXT_ROWS);
      return xlsx.getFile();
    },
  },
  {
    name: "inline-strings",
    async build() {
      const xlsx = new XlsxStreamWriter({ inlineStrings: true });
      xlsx.addRows(TEXT_ROWS);
      return xlsx.getFile();
    },
  },
  {
    name: "hostile-values",
    async build() {
      const xlsx = new XlsxStreamWriter();
      xlsx.addRows([["Label", "Value"], ...HOSTILE.map(([label, value]) => [label, value])]);
      return xlsx.getFile();
    },
  },
  {
    name: "streamed-sheet",
    async build() {
      const xlsx = new XlsxStreamWriter();
      xlsx.addRows(streamedRows());
      return xlsx.getFile();
    },
  },
];

// One LibreOffice invocation converts every fixture: startup dominates, so
// batching costs a few seconds where one process per fixture would cost tens.
const converted = (async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "xlsx-calc-"));
  const built = new Map();

  for (const fixture of FIXTURES) {
    const buffer = await fixture.build();
    const file = path.join(dir, `${fixture.name}.xlsx`);
    fs.writeFileSync(file, buffer);
    built.set(fixture.name, { buffer, file });
  }

  // A private profile: a LibreOffice already running for the human at this
  // machine would otherwise refuse to start a second instance.
  const profile = path.join(dir, "profile");
  const args = [
    "--headless",
    `-env:UserInstallation=${pathToFileURL(profile).href}`,
    "--convert-to",
    // Comma-separated, quote-delimited, UTF-8 — so non-ASCII survives the proxy
    // rather than failing in the reader and looking like a writer bug.
    "csv:Text - txt - csv (StarCalc):44,34,76",
    "--outdir",
    dir,
    ...FIXTURES.map(fixture => built.get(fixture.name).file),
  ];

  try {
    execFileSync("soffice", args, { encoding: "utf8", stdio: "pipe" });
  } catch (error) {
    if (error.code === "ENOENT") {
      fs.rmSync(dir, { recursive: true, force: true });
      // Absent LibreOffice means these tests pass by asserting nothing, which is
      // fine on a contributor's machine and useless in CI. The workflow sets
      // XLSX_REQUIRE_SPREADSHEET so that a runner that loses LibreOffice fails
      // rather than going quietly green.
      if (process.env.XLSX_REQUIRE_SPREADSHEET) {
        throw new Error(
          "XLSX_REQUIRE_SPREADSHEET is set but LibreOffice is not installed, so " +
            "no workbook was opened in a spreadsheet application.",
        );
      }
      return null;
    }
    throw error;
  }

  const results = new Map();
  for (const fixture of FIXTURES) {
    const csv = path.join(dir, `${fixture.name}.csv`);
    if (!fs.existsSync(csv)) {
      // The failure mode docs/verifying-output.md warns about: a LibreOffice
      // without its spreadsheet component fails to load anything at all, which
      // reads as a corrupt workbook and is not one.
      throw new Error(
        `LibreOffice produced no CSV for ${fixture.name}. If it failed for every ` +
          `fixture, check that the spreadsheet component is installed ` +
          `(apt install libreoffice-calc) before suspecting the output.`,
      );
    }
    results.set(fixture.name, {
      ...built.get(fixture.name),
      rows: parseCsv(fs.readFileSync(csv, "utf8")),
    });
  }

  fs.rmSync(dir, { recursive: true, force: true });
  return results;
})();

async function fixture(name) {
  const all = await converted;
  return all === null ? null : all.get(name);
}

for (const name of ["shared-strings", "inline-strings"]) {
  test(`LibreOffice reads every cell of a ${name} workbook`, { timeout: 180000 }, async () => {
    const result = await fixture(name);
    if (!result) return;

    assert.deepEqual(result.rows, TEXT_ROWS, "the sheet should read back exactly as written");
  });
}

test("LibreOffice reads back values chosen to break readers", { timeout: 180000 }, async () => {
  const result = await fixture("hostile-values");
  if (!result) return;

  const [header, ...rows] = result.rows;
  assert.deepEqual(header, ["Label", "Value"]);
  assert.equal(rows.length, HOSTILE.length, "every row should survive");

  for (const [i, [label, written, rendered]] of HOSTILE.entries()) {
    const expected = rendered === undefined ? written : rendered;
    assert.equal(rows[i][0], label, `row ${i} label`);
    assert.equal(rows[i][1], expected, `${label} should read back as written`);
  }
});

test("LibreOffice reads a sheet written in the streamed shape", { timeout: 180000 }, async () => {
  const result = await fixture("streamed-sheet");
  if (!result) return;

  // The point of this fixture: the archive shape most likely to be rejected is
  // the one being read. Sizes are not known when the local header is written,
  // so they trail the data instead.
  const sheet = readZip(result.buffer).entry("xl/worksheets/sheet1.xml");
  assert.ok(
    sheet.uncompressedSize > BUFFER_LIMIT,
    `sheet should exceed the ${BUFFER_LIMIT} byte limit, was ${sheet.uncompressedSize}`,
  );
  assert.equal(sheet.usesDataDescriptor, true, "an oversized part must defer its sizes");

  const [header, ...rows] = result.rows;
  assert.deepEqual(header, ["Name", "Location", "Amount", "Active"]);
  assert.equal(rows.length, STREAMED_ROW_COUNT - 1, "every row should survive the round trip");
  assert.deepEqual(rows[0], ["Customer number 1", "City of origin 1", "1.5", "FALSE"]);
  assert.deepEqual(rows.at(-1), [
    `Customer number ${STREAMED_ROW_COUNT - 1}`,
    `City of origin ${(STREAMED_ROW_COUNT - 1) % 997}`,
    String((STREAMED_ROW_COUNT - 1) * 1.5),
    (STREAMED_ROW_COUNT - 1) % 2 === 0 ? "TRUE" : "FALSE",
  ]);
});
