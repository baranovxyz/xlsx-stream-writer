This was rewritten from coffee script https://github.com/rubenv/node-xlsx-writer and 
changed to work in browser. Api is different from rubenv implementation.

Does not actually stream data, more like a future plan.
Though, it should be efficient enough...

It uses JSZip to compress resulting structure, so will work both in nodejs and in browser.

Plans:
- test for large files (when string size in browser will be an issue)
- add tests
- implement a better version for nodejs
- (maybe) implement or find some kind of streaming zip module for browser

```javascript
const XLSX = require("xlsx-stream-writer");
const fs = require("fs");

const rows = [["Name", "Location"], ["Bob", "Sweden"], ["Alice", "France"]];

const xlsx = new XLSX();
rows.map(row => xlsx.addRow(row));
xlsx
  .getFile()
  .then(buffer => {
    fs.writeFileSync("test_file.xlsx", buffer);
  })
  .catch(console.error);
```
