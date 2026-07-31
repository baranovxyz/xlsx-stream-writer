const test = require("node:test");
const assert = require("node:assert/strict");
const { inflateRawSync } = require("node:zlib");

const nodeAdapter = require("../../src/zip/deflate.node");
const browserAdapter = require("../../src/zip/deflate.browser");

async function deflateAll(adapter, chunks, level = 4) {
  async function* source() {
    for (const chunk of chunks) yield Buffer.from(chunk);
  }
  const out = [];
  for await (const chunk of adapter.deflateRaw(source(), level)) out.push(Buffer.from(chunk));
  return Buffer.concat(out);
}

const SAMPLE = "the quick brown fox ".repeat(500);

// Both adapters must produce a raw deflate stream — not zlib- or gzip-wrapped —
// because that is what a ZIP entry stores. Byte-for-byte equality between them
// is not required; inflating to the same input is.
test("the node adapter produces inflatable raw deflate", async () => {
  const compressed = await deflateAll(nodeAdapter, [SAMPLE]);
  assert.equal(inflateRawSync(compressed).toString(), SAMPLE);
  assert.ok(compressed.length < SAMPLE.length, "should actually compress");
});

test("the browser adapter produces inflatable raw deflate", async () => {
  const compressed = await deflateAll(browserAdapter, [SAMPLE]);
  assert.equal(inflateRawSync(compressed).toString(), SAMPLE);
  assert.ok(compressed.length < SAMPLE.length, "should actually compress");
});

test("both adapters handle many small chunks", async () => {
  const chunks = Array.from({ length: 500 }, (_, i) => `row ${i}\n`);
  const joined = chunks.join("");
  for (const [name, adapter] of [["node", nodeAdapter], ["browser", browserAdapter]]) {
    const compressed = await deflateAll(adapter, chunks);
    assert.equal(inflateRawSync(compressed).toString(), joined, `${name} adapter`);
  }
});

test("both adapters handle empty input", async () => {
  for (const [name, adapter] of [["node", nodeAdapter], ["browser", browserAdapter]]) {
    const compressed = await deflateAll(adapter, []);
    assert.equal(inflateRawSync(compressed).length, 0, `${name} adapter`);
  }
});

test("both adapters surface a failing source", async () => {
  async function* explode() {
    yield Buffer.from("partial");
    throw new Error("source blew up");
  }
  for (const [name, adapter] of [["node", nodeAdapter], ["browser", browserAdapter]]) {
    await assert.rejects(
      (async () => {
        for await (const _ of adapter.deflateRaw(explode(), 4)) void _;
      })(),
      /source blew up/,
      `${name} adapter`,
    );
  }
});

test("the node adapter honours the compression level", async () => {
  const fast = await deflateAll(nodeAdapter, [SAMPLE], 1);
  const best = await deflateAll(nodeAdapter, [SAMPLE], 9);
  assert.ok(best.length <= fast.length, "level 9 should not be larger than level 1");
  assert.equal(inflateRawSync(fast).toString(), SAMPLE);
  assert.equal(inflateRawSync(best).toString(), SAMPLE);
});
