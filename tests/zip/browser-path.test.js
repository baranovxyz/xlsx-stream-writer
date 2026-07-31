const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const JSZip = require("jszip");
const { writeZip } = require("../../src/zip/writer");
const { deflateRaw } = require("../../src/zip/deflate.browser");
const { readZip } = require("../support/unzip");

const packageJson = require("../../package.json");

// In a browser bundle the "browser" field swaps the zlib adapter for the
// CompressionStream one. Nothing else about the writer changes, so driving the
// same code with that adapter is what a browser build would do.
async function buildWithBrowserAdapter(entries, options = {}) {
  const chunks = [];
  for await (const chunk of writeZip(entries, { ...options, deflateRaw })) {
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

test("the browser field maps a file that exists", () => {
  const map = packageJson.browser;
  assert.ok(map, "package.json should declare a browser field");
  for (const [from, to] of Object.entries(map)) {
    assert.ok(fs.existsSync(path.join(__dirname, "../..", from)), `${from} should exist`);
    assert.ok(fs.existsSync(path.join(__dirname, "../..", to)), `${to} should exist`);
  }
});

test("no source file outside the swapped adapter imports node builtins", () => {
  const root = path.join(__dirname, "../../src");
  const allowed = new Set([path.join(root, "zip", "deflate.node.js")]);
  const offenders = [];

  (function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(p);
      else if (p.endsWith(".js") && !allowed.has(p)) {
        const source = fs.readFileSync(p, "utf8");
        // A stray node: import would break any browser bundle, and the browser
        // field only redirects the one adapter.
        if (/require\(["']node:/.test(source)) offenders.push(path.relative(root, p));
      }
    }
  })(root);

  assert.deepEqual(offenders, []);
});

test("the browser adapter produces an archive three readers accept", async () => {
  const text = "browser path ".repeat(200);
  const buffer = await buildWithBrowserAdapter([
    { name: "a.txt", source: text },
    { name: "b.txt", source: "second entry" },
  ]);

  const zip = readZip(buffer);
  assert.deepEqual(zip.names, ["a.txt", "b.txt"]);
  assert.equal(zip.text("a.txt"), text);

  const viaJsZip = await JSZip.loadAsync(buffer);
  assert.equal(await viaJsZip.file("a.txt").async("string"), text);
});

test("the browser adapter also handles the streamed path", async () => {
  const text = "streamed in the browser ".repeat(1000);
  async function* chunks() {
    for (let i = 0; i < text.length; i += 256) yield text.slice(i, i + 256);
  }
  const buffer = await buildWithBrowserAdapter([{ name: "big.txt", source: chunks() }], {
    bufferLimit: 64,
  });

  const entry = readZip(buffer).entry("big.txt");
  assert.equal(entry.usesDataDescriptor, true);
  assert.equal(readZip(buffer).text("big.txt"), text);
  assert.equal(await (await JSZip.loadAsync(buffer)).file("big.txt").async("string"), text);
});
