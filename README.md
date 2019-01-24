This was rewritten from coffee script https://github.com/rubenv/node-xlsx-writer and 
changed to work both in browser and nodejs. Api is completely different from rubenv 
implementation.

It is actually capable of streaming rows into xlsx file both in browser and nodejs.

It uses JSZip to compress resulting structure. Lucky for us JSZip is capable of 
processing readable streams, so we just stream rows into xlxs file (which is a zip file).

Plans:
- improve api
- add tests
- make browser build, put on some cdn
- optimize shared string stuff
- (maybe) use web workers to build xlsx in browser
- (maybe) implement some specifis for nodejs

You can add rows:
```javascript
const XlsxStreamWriter = require("xlsx-stream-writer");
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const xlsx = new XlsxStreamWriter();
xlsx.addRows(rows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
```

Or add readable stream of rows:
```javascript
const XlsxStreamWriter = require("xlsx-stream-writer");
const Readable = require("stream-browserify").Readable;
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

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
const streamOfRows = wrapRowsInStream(rows);

const xlsx = new XlsxStreamWriter();
xlsx.addRows(streamOfRows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
```
