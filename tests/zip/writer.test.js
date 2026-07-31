const test = require("node:test");
const assert = require("node:assert/strict");
const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const zlib = require("node:zlib");

const JSZip = require("jszip");
const { writeZip } = require("../../src/zip/writer");
const { crc32 } = require("../../src/zip/crc32");
const { readZip } = require("../support/unzip");

async function build(entries, options) {
  const chunks = [];
  for await (const chunk of writeZip(entries, options)) chunks.push(chunk);
  return Buffer.concat(chunks);
}

async function* chunked(text, size) {
  for (let i = 0; i < text.length; i += size) yield text.slice(i, i + size);
}

function withUnzip(buffer, assertion) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "zip-"));
  const file = path.join(dir, "test.zip");
  fs.writeFileSync(file, buffer);
  try {
    assertion(execFileSync("unzip", ["-t", file], { encoding: "utf8" }));
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test("crc32 matches zlib for every input shape", () => {
  const cases = [
    Buffer.alloc(0),
    Buffer.from("a"),
    Buffer.from("The quick brown fox jumps over the lazy dog"),
    Buffer.from(Array.from({ length: 1000 }, (_, i) => i % 256)),
  ];
  for (const input of cases) {
    assert.equal(crc32(input), zlib.crc32(input) >>> 0, `crc32 of ${input.length} bytes`);
  }
});

test("crc32 chunked equals crc32 whole", () => {
  const data = Buffer.from(Array.from({ length: 5000 }, (_, i) => (i * 31) % 256));
  const whole = crc32(data);
  let running = 0;
  for (let i = 0; i < data.length; i += 97) running = crc32(data.subarray(i, i + 97), running);
  assert.equal(running, whole);
});

test("a buffered entry round-trips through three readers", async () => {
  const text = "hello ".repeat(100);
  const buffer = await build([{ name: "a.txt", source: text }]);

  assert.equal(readZip(buffer).text("a.txt"), text);
  assert.equal(await (await JSZip.loadAsync(buffer)).file("a.txt").async("string"), text);
  withUnzip(buffer, output => assert.match(output, /No errors detected/));
});

test("a streamed entry uses a data descriptor and ZIP64, and still round-trips", async () => {
  const text = "streamed content ".repeat(2000);
  // bufferLimit forces the branch a >8 MiB part would otherwise take.
  const buffer = await build([{ name: "big.txt", source: chunked(text, 512) }], {
    bufferLimit: 64,
  });

  const entry = readZip(buffer).entry("big.txt");
  assert.equal(entry.usesDataDescriptor, true);
  assert.equal(entry.uncompressedSize, Buffer.byteLength(text));

  assert.equal(readZip(buffer).text("big.txt"), text);
  assert.equal(await (await JSZip.loadAsync(buffer)).file("big.txt").async("string"), text);
  withUnzip(buffer, output => assert.match(output, /No errors detected/));
});

test("buffered and streamed entries can be mixed in one archive", async () => {
  const small = "small";
  const large = "large ".repeat(3000);
  const buffer = await build(
    [
      { name: "small.txt", source: small },
      { name: "large.txt", source: chunked(large, 256) },
      { name: "after.txt", source: "written after a streamed entry" },
    ],
    { bufferLimit: 64 },
  );

  const zip = readZip(buffer);
  assert.deepEqual(zip.names, ["small.txt", "large.txt", "after.txt"]);
  assert.equal(zip.entry("small.txt").usesDataDescriptor, false);
  assert.equal(zip.entry("large.txt").usesDataDescriptor, true);
  // The entry after a streamed one must still be found, which only works if
  // the streamed entry's local header offset accounting was right.
  assert.equal(zip.text("after.txt"), "written after a streamed entry");

  const viaJsZip = await JSZip.loadAsync(buffer);
  assert.equal(await viaJsZip.file("large.txt").async("string"), large);
  assert.equal(await viaJsZip.file("after.txt").async("string"), "written after a streamed entry");
  withUnzip(buffer, output => assert.match(output, /No errors detected/));
});

test("entries are written in the order given", async () => {
  const order = [];
  const tracked = name =>
    (async function* () {
      order.push(name);
      yield name;
    })();

  const buffer = await build([
    { name: "first", source: tracked("first") },
    { name: "second", source: tracked("second") },
    { name: "third", source: tracked("third") },
  ]);

  // Sources must be consumed lazily, in order: the workbook relies on the
  // shared-string table being generated only after the sheet has been walked.
  assert.deepEqual(order, ["first", "second", "third"]);
  assert.deepEqual(readZip(buffer).names, ["first", "second", "third"]);
});

test("an empty entry is valid", async () => {
  const buffer = await build([{ name: "empty.txt", source: "" }]);
  assert.equal(readZip(buffer).text("empty.txt"), "");
  withUnzip(buffer, output => assert.match(output, /No errors detected/));
});

test("non-ASCII entry names survive", async () => {
  const buffer = await build([{ name: "файл.txt", source: "содержимое" }]);
  const zip = readZip(buffer);
  assert.deepEqual(zip.names, ["файл.txt"]);
  assert.equal(zip.entry("файл.txt").utf8NameFlag, true);
  assert.equal(await (await JSZip.loadAsync(buffer)).file("файл.txt").async("string"), "содержимое");
});

test("a failing source aborts the archive rather than truncating it", async () => {
  async function* explode() {
    yield "partial";
    throw new Error("source blew up");
  }
  await assert.rejects(build([{ name: "bad.txt", source: explode() }]), /source blew up/);
});
