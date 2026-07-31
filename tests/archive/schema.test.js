const test = require("node:test");
const assert = require("node:assert/strict");
const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const XlsxStreamWriter = require("../../dist");

// The spreadsheet checks prove a reader accepts the output. They cannot prove
// it conforms to the format as specified, because LibreOffice is tolerant by
// design — it silently coerces a broken shared-string reference to index 0
// rather than refusing the file. Excel is the strict reader that matters and
// cannot run here, so this uses the next best thing: Microsoft's own schema
// validator, from the Open XML SDK.
//
// Skipped silently when the .NET SDK is absent, the same way the LibreOffice
// and system-unzip checks are, so the suite stays portable. CI installs .NET
// and sets XLSX_REQUIRE_SCHEMA. See docs/verifying-output.md.

const PROJECT = path.join(__dirname, "..", "schema");

// Chosen for structural variety rather than realism: each fixture puts a
// different part, or a different shape of the same part, in front of the
// validator. Deliberately defined here rather than shared with the spreadsheet
// checks — two independent fixture sets catch more than one reused twice.
const FIXTURES = [
  {
    name: "shared-strings",
    build: () => {
      const xlsx = new XlsxStreamWriter();
      xlsx.addRows([
        ["Name", "Location"],
        ["Alpha", "Berlin"],
        ["Alpha", "Berlin"],
      ]);
      return xlsx.getFile();
    },
  },
  {
    name: "inline-strings",
    build: () => {
      const xlsx = new XlsxStreamWriter({ inlineStrings: true });
      xlsx.addRows([
        ["Name", "Location"],
        ["Alpha", "Berlin"],
      ]);
      return xlsx.getFile();
    },
  },
  {
    name: "mixed-types",
    build: () => {
      const xlsx = new XlsxStreamWriter();
      xlsx.addRows([
        ["str", "num", "bool", "date", "empty", "big"],
        ["text", 1.5, true, new Date(Date.UTC(2026, 0, 1)), null, 10n],
        ["", -0.25, false, new Date(Date.UTC(1999, 11, 31)), undefined, 0],
      ]);
      return xlsx.getFile();
    },
  },
  {
    name: "styles",
    build: () => {
      const xlsx = new XlsxStreamWriter({
        styles: [{ fill: "FFFF0000" }, { format: "dd.mm.yyyy" }],
        styleIdFunc: (value, columnId, rowId) => (rowId === 0 ? 1 : 2),
      });
      xlsx.addRows([
        ["Header", "Header"],
        ["Alpha", new Date(Date.UTC(2026, 5, 15))],
      ]);
      return xlsx.getFile();
    },
  },
  {
    // The values that have broken readers before, including the
    // Object.prototype names this package has been bitten by twice.
    name: "hostile-strings",
    build: () => {
      const xlsx = new XlsxStreamWriter();
      xlsx.addRows([
        ["label", "value"],
        ["xml", "<sheet> & \"quoted\" 'apos'"],
        ["cdata", "]]><!--x-->"],
        ["attr", '"/><f>=1+1</f><x a="'],
        ["proto", "__proto__"],
        ["ctor", "constructor"],
        ["tostring", "toString"],
        ["formula", "=1+1"],
        ["emoji", "\u{1F600}"],
        ["rtl", "שלום"],
        ["cjk", "漢字"],
        ["wide", "x".repeat(32767)],
      ]);
      return xlsx.getFile();
    },
  },
  {
    // Past the writer's buffering limit, so local headers carry zeroes and the
    // real sizes trail the data. The archive shape most likely to be rejected.
    name: "streamed-shape",
    build: () => {
      const xlsx = new XlsxStreamWriter();
      function* rows() {
        yield ["Name", "Amount"];
        for (let i = 1; i < 40000; i++) yield [`Customer ${i}`, i * 1.5];
      }
      xlsx.addRows(rows());
      return xlsx.getFile();
    },
  },
];

const validated = (async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "xlsx-schema-"));
  const files = [];

  for (const fixture of FIXTURES) {
    const buffer = await fixture.build();
    const file = path.join(dir, `${fixture.name}.xlsx`);
    fs.writeFileSync(file, buffer);
    files.push(file);
  }

  // One invocation for every fixture: restoring and building the validator
  // dominates, so per-file processes would cost far more than the validation.
  //
  // Everything after `--` reaches the program as a file path, so nothing but a
  // path may follow it. `dotnet run` forwards flags it does not recognise
  // rather than rejecting them, so a stray `--nologo` arrives as a filename and
  // reports itself as an unopenable workbook.
  const args = ["run", "--project", PROJECT, "-c", "Release", "--", ...files];

  let stdout;
  let failures = 0;

  try {
    stdout = execFileSync("dotnet", args, { encoding: "utf8", stdio: "pipe" });
  } catch (error) {
    if (error.code === "ENOENT") {
      fs.rmSync(dir, { recursive: true, force: true });
      // Absent .NET means these tests pass while asserting nothing, which is
      // fine locally and useless in CI. The workflow sets XLSX_REQUIRE_SCHEMA
      // so a runner that loses .NET fails rather than going quietly green.
      if (process.env.XLSX_REQUIRE_SCHEMA) {
        throw new Error(
          "XLSX_REQUIRE_SCHEMA is set but the .NET SDK is not installed, so no " +
            "workbook was checked against the Open XML schema.",
        );
      }
      return null;
    }

    // A non-zero exit is the validator reporting errors, which is a result
    // rather than a crash — keep the output so the assertion can print it.
    if (typeof error.status === "number" && error.status > 0 && error.status <= 100) {
      stdout = String(error.stdout ?? "");
      failures = error.status;
    } else {
      throw error;
    }
  }

  fs.rmSync(dir, { recursive: true, force: true });
  return { stdout, failures };
})();

test("every generated workbook conforms to the Open XML schema", { timeout: 600000 }, async () => {
  const result = await validated;

  if (result === null) {
    return; // .NET absent; see the comment on the skip above.
  }

  assert.equal(
    result.failures,
    0,
    `Microsoft's Open XML validator rejected ${result.failures} workbook(s). ` +
      `Unlike Excel, it says exactly what is wrong:\n\n${result.stdout}`,
  );

  // The validator counts what it was handed, so a count that disagrees with the
  // fixture list means the argument vector did — an extra flag arriving as a
  // filename, or a fixture silently not reaching it.
  assert.match(
    result.stdout,
    new RegExp(`^${FIXTURES.length} file\\(s\\) conform to the Office2007 schema$`, "m"),
    `The validator did not report exactly ${FIXTURES.length} files, so it was not ` +
      `handed what this test thinks it was:\n\n${result.stdout}`,
  );

  // A pass that validated nothing would be indistinguishable from a real pass.
  for (const fixture of FIXTURES) {
    assert.match(
      result.stdout,
      new RegExp(`^ok\\s+${fixture.name}\\.xlsx$`, "m"),
      `${fixture.name}.xlsx was not reported as validated:\n\n${result.stdout}`,
    );
  }
});
