const test = require("node:test");
const assert = require("node:assert/strict");
const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const JSZip = require("jszip");
const XlsxStreamWriter = require("../../dist");
const { rows } = require("../helpers");
const { PARTS } = require("../support/golden");

// tests/support/unzip.js is ours, so it could in principle share a blind spot
// with the writer. These checks read the same bytes with implementations that
// have no such relationship.
const archive = (async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(rows);
  return xlsx.getFile();
})();

test("an independent zip implementation reads every part", async () => {
  const zip = await JSZip.loadAsync(await archive);
  for (const [name, expected] of Object.entries(PARTS)) {
    assert.equal(await zip.file(name).async("string"), expected, `${name} via JSZip`);
  }
});

test("the system unzip accepts the archive", async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "xlsx-")), "book.xlsx");
  fs.writeFileSync(file, await archive);
  try {
    const output = execFileSync("unzip", ["-t", file], { encoding: "utf8" });
    assert.match(output, /No errors detected/);
  } catch (error) {
    if (error.code === "ENOENT") return; // unzip is not installed; nothing to assert
    throw error;
  } finally {
    fs.rmSync(path.dirname(file), { recursive: true, force: true });
  }
});
